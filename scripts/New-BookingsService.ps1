#Requires -Modules Microsoft.Graph.Bookings

<#
.SYNOPSIS
    Defines the Bookings service that represents reserving a Hotel Office or
    Cubicle Office room.

.DESCRIPTION
    Creates a bookingService on a Microsoft Bookings business - the "service"
    employees pick when they reserve a room, e.g. "Hotel Office" or "Cubicle
    Office". This is a one-time setup step per resource category, done once a
    business exists and before rooms are linked to it: New-BookingsStaffLink.ps1
    adds each room mailbox as a staff member capable of providing this service.

    The advance-booking window is enforced here, on the Bookings side, via
    schedulingPolicy.maximumAdvance (ISO 8601 duration, e.g. "P14D") - this is
    what actually limits how far out an employee can pick a slot in the Bookings
    UI. Set-CalendarProcessing -BookingWindowInDays (applied to the room mailboxes
    by New-RoomResource.ps1 / Set-RoomBookingPolicy.ps1) is the mailbox-side
    backstop for the same limit, described in CLAUDE.md.

    Idempotent: if a service with the given display name already exists on the
    business, the script reports it and does not create a duplicate.

.PARAMETER BookingBusinessId
    The id (typically the business's email address, e.g. "rooms@contoso.com")
    of the Microsoft Bookings business to define the service on.

.PARAMETER ResourceType
    Which room-booking category this service represents: HotelOffice or CubicleOffice.
    Determines the default display name ("Hotel Office" / "Cubicle Office").

.PARAMETER BookingWindowDays
    Maximum number of days in advance the service can be booked. Defaults to 14
    (2 weeks), the framework-wide policy. Should match the -BookingWindowDays
    used when provisioning the rooms themselves.

.PARAMETER DurationMinutes
    Length of a single reservation slot, in minutes. Defaults to 480 (a full
    8-hour workday), since these are day-use desk/cubicle bookings rather than
    short meetings.

.PARAMETER StaffMemberIds
    Optional. Bookings staff member IDs (as returned by New-BookingsStaffLink.ps1)
    to associate with this service from the start. Staff can also be linked to a
    service later by updating the service's staffMemberIds.

.EXAMPLE
    ./New-BookingsService.ps1 -BookingBusinessId "rooms@contoso.com" -ResourceType HotelOffice

.EXAMPLE
    ./New-BookingsService.ps1 -BookingBusinessId "rooms@contoso.com" -ResourceType CubicleOffice -BookingWindowDays 14 -DurationMinutes 240
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string]$BookingBusinessId,

    [Parameter(Mandatory)]
    [ValidateSet('HotelOffice', 'CubicleOffice')]
    [string]$ResourceType,

    [ValidateRange(1, 365)]
    [int]$BookingWindowDays = 14,

    [ValidateRange(15, 1440)]
    [int]$DurationMinutes = 480,

    [string[]]$StaffMemberIds
)

$ErrorActionPreference = 'Stop'

$displayName = switch ($ResourceType) {
    'HotelOffice'   { 'Hotel Office' }
    'CubicleOffice' { 'Cubicle Office' }
}

if (-not (Get-MgContext)) {
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
    Connect-MgGraph -Scopes 'Bookings.ReadWrite.All' -NoWelcome
}

$business = Get-MgBookingBusiness -BookingBusinessId $BookingBusinessId -ErrorAction SilentlyContinue
if (-not $business) {
    throw "Bookings business '$BookingBusinessId' not found. Verify the id (usually the business's email address) and that Bookings.ReadWrite.All access has been granted."
}

$existingService = Get-MgBookingBusinessService -BookingBusinessId $BookingBusinessId -All |
    Where-Object { $_.DisplayName -eq $displayName }
if ($existingService) {
    Write-Host "Service '$displayName' already exists on '$($business.DisplayName)' - nothing to do." -ForegroundColor Yellow
    return $existingService
}

if ($PSCmdlet.ShouldProcess($displayName, "Create Bookings service on '$($business.DisplayName)' ($BookingWindowDays-day window, ${DurationMinutes}min duration)")) {
    $body = @{
        '@odata.type'      = '#microsoft.graph.bookingService'
        displayName        = $displayName
        description        = "Reserve a $displayName for up to $BookingWindowDays days in advance."
        defaultDuration    = "PT${DurationMinutes}M"
        isLocationOnline   = $false
        schedulingPolicy   = @{
            '@odata.type'     = '#microsoft.graph.bookingSchedulingPolicy'
            maximumAdvance    = "P${BookingWindowDays}D"
            minimumLeadTime   = 'PT0M'
            timeSlotInterval  = "PT${DurationMinutes}M"
            allowStaffSelection = $true
        }
    }
    if ($StaffMemberIds) {
        $body['staffMemberIds'] = $StaffMemberIds
    }

    $service = New-MgBookingBusinessService -BookingBusinessId $BookingBusinessId -BodyParameter $body

    Write-Host "Created Bookings service '$displayName' on '$($business.DisplayName)' with a $BookingWindowDays-day booking window." -ForegroundColor Green
    $service
}
