#Requires -Modules ExchangeOnlineManagement, Microsoft.Graph.Bookings

<#
.SYNOPSIS
    Registers an Exchange Online room mailbox as a staff member on a Microsoft
    Bookings business, so employees can reserve it like an appointment.

.DESCRIPTION
    Looks up the room mailbox (created by New-RoomResource.ps1) and adds it as a
    bookingStaffMember on the given Bookings business via Microsoft Graph. The
    room's own calendar remains the source of truth: availabilityIsAffectedByPersonalCalendar
    is set to $true so Bookings checks the room mailbox's real calendar (which
    already reflects the 14-day window enforced by Set-RoomBookingPolicy.ps1)
    before allowing a reservation, rather than only tracking availability inside
    Bookings itself. useBusinessHours is set to $true so rooms are only offered
    for booking during the business's configured hours of operation.

    Idempotent: if a staff member with the room's email address already exists
    on the business, the script reports it and does not create a duplicate.

.PARAMETER Identity
    The room mailbox to link (name, alias, or email), as created by New-RoomResource.ps1.

.PARAMETER BookingBusinessId
    The id (typically the business's email address, e.g. "rooms@contoso.com")
    of the Microsoft Bookings business to add the room to.

.PARAMETER Role
    The Bookings staff role to assign. Defaults to 'guest', the least-privileged
    role, since the room mailbox isn't a real person who should administer the business.

.EXAMPLE
    ./New-BookingsStaffLink.ps1 -Identity "HotelOffice-12A" -BookingBusinessId "rooms@contoso.com"
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string]$Identity,

    [Parameter(Mandatory)]
    [string]$BookingBusinessId,

    [ValidateSet('guest', 'viewer', 'scheduler', 'teamMember', 'administrator', 'externalGuest')]
    [string]$Role = 'guest'
)

$ErrorActionPreference = 'Stop'

$exoSession = Get-ConnectionInformation -ErrorAction SilentlyContinue |
    Where-Object { $_.State -eq 'Connected' }
if (-not $exoSession) {
    Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
    Connect-ExchangeOnline -ShowBanner:$false
}

if (-not (Get-MgContext)) {
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
    Connect-MgGraph -Scopes 'Bookings.ReadWrite.All' -NoWelcome
}

$room = Get-Mailbox -Identity $Identity -ErrorAction Stop
if ($room.RecipientTypeDetails -ne 'RoomMailbox') {
    throw "'$Identity' is a $($room.RecipientTypeDetails), not a RoomMailbox."
}
if (-not $room.CustomAttribute1) {
    Write-Warning "'$($room.Name)' has no CustomAttribute1 resource-type tag - it wasn't provisioned by New-RoomResource.ps1, or the tag was cleared."
}

$business = Get-MgBookingBusiness -BookingBusinessId $BookingBusinessId -ErrorAction SilentlyContinue
if (-not $business) {
    throw "Bookings business '$BookingBusinessId' not found. Verify the id (usually the business's email address) and that Bookings.ReadWrite.All access has been granted."
}

$existingStaff = Get-MgBookingBusinessStaffMember -BookingBusinessId $BookingBusinessId -All |
    Where-Object { $_.EmailAddress -eq $room.PrimarySmtpAddress }
if ($existingStaff) {
    Write-Host "'$($room.Name)' is already a staff member on '$($business.DisplayName)' - nothing to do." -ForegroundColor Yellow
    return $existingStaff
}

if ($PSCmdlet.ShouldProcess($room.Name, "Add as Bookings staff member on '$($business.DisplayName)'")) {
    $body = @{
        '@odata.type'                              = '#microsoft.graph.bookingStaffMember'
        displayName                                = $room.DisplayName
        emailAddress                                = $room.PrimarySmtpAddress
        'role@odata.type'                           = '#microsoft.graph.bookingStaffRole'
        role                                        = $Role
        availabilityIsAffectedByPersonalCalendar    = $true
        useBusinessHours                            = $true
        isEmailNotificationEnabled                  = $false
    }

    $staffMember = New-MgBookingBusinessStaffMember -BookingBusinessId $BookingBusinessId -BodyParameter $body

    Write-Host "Added '$($room.Name)' as a Bookings staff member on '$($business.DisplayName)'." -ForegroundColor Green
    $staffMember
}
