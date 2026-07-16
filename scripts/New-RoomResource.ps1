#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
    Creates one or more Exchange Online room mailboxes for the room-booking framework.

.DESCRIPTION
    Provisions a room mailbox tagged as either "Hotel Office" or "Cubicle Office",
    adds it to the matching room list (a RoomList distribution group) so it shows
    up in Outlook's room finder grouped by resource type, and applies the
    framework's advance-booking window via Set-CalendarProcessing. The resource
    type is also written to CustomAttribute1, which downstream scripts
    (New-BookingsStaffLink.ps1, Get-RoomInventory.ps1) use to identify rooms
    belonging to this framework and their category.

    Rooms can be created one at a time via -Name/-ResourceType/etc., or in bulk
    via -CsvPath. This script does not link the room into Microsoft Bookings
    (see New-BookingsStaffLink.ps1).

.PARAMETER Name
    Short identifier for the room, used as the mailbox alias/name (e.g. "HotelOffice-12A").
    Must be unique in the tenant.

.PARAMETER DisplayName
    Friendly name shown in Outlook/room finder (e.g. "Hotel Office 12A - 3rd Floor").
    Defaults to Name if not specified.

.PARAMETER ResourceType
    Which room-booking category this resource belongs to: HotelOffice or CubicleOffice.

.PARAMETER Capacity
    Optional seating capacity for the room (defaults to 1, since these are single-occupant desks).

.PARAMETER Location
    Optional free-text location, e.g. "Building A, Floor 3". Stored on the mailbox Office field.

.PARAMETER CsvPath
    Path to a CSV file describing multiple rooms to create in bulk, instead of -Name/-ResourceType/etc.
    Required columns: Name, ResourceType. Optional columns: DisplayName, Capacity, Location,
    BookingWindowDays (overrides -BookingWindowDays for that row only). See docs/sample-rooms.csv
    for an example. Rows that fail (e.g. duplicate name) are reported as warnings and skipped;
    the rest of the file still runs.

.PARAMETER BookingWindowDays
    Maximum number of days in advance the room can be booked. Defaults to 14 (2 weeks),
    the framework-wide policy. Applied immediately via Set-CalendarProcessing so a
    newly created room is never left without a booking window. Can be overridden per
    row in -CsvPath.

.EXAMPLE
    ./New-RoomResource.ps1 -Name "HotelOffice-12A" -DisplayName "Hotel Office 12A - 3rd Floor" -ResourceType HotelOffice -Location "Building A, Floor 3"

.EXAMPLE
    ./New-RoomResource.ps1 -Name "Cubicle-204" -ResourceType CubicleOffice -Capacity 1 -BookingWindowDays 14

.EXAMPLE
    ./New-RoomResource.ps1 -CsvPath ./docs/sample-rooms.csv
#>

[CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'SingleRoom')]
param(
    [Parameter(Mandatory, ParameterSetName = 'SingleRoom')]
    [ValidatePattern('^[A-Za-z0-9\-_]+$')]
    [string]$Name,

    [Parameter(ParameterSetName = 'SingleRoom')]
    [string]$DisplayName,

    [Parameter(Mandatory, ParameterSetName = 'SingleRoom')]
    [ValidateSet('HotelOffice', 'CubicleOffice')]
    [string]$ResourceType,

    [Parameter(ParameterSetName = 'SingleRoom')]
    [ValidateRange(1, 100)]
    [int]$Capacity = 1,

    [Parameter(ParameterSetName = 'SingleRoom')]
    [string]$Location,

    [Parameter(Mandatory, ParameterSetName = 'FromCsv')]
    [ValidateScript({ Test-Path -Path $_ -PathType Leaf })]
    [string]$CsvPath,

    [ValidateRange(1, 365)]
    [int]$BookingWindowDays = 14
)

$ErrorActionPreference = 'Stop'

function New-SingleRoomResource {
    param(
        [string]$Name,
        [string]$DisplayName,
        [string]$ResourceType,
        [int]$Capacity,
        [string]$Location,
        [int]$BookingWindowDays
    )

    # Human-readable room list / category name used both as the RoomList
    # distribution group name and the CustomAttribute1 tag prefix.
    $categoryName = switch ($ResourceType) {
        'HotelOffice'   { 'Hotel Office' }
        'CubicleOffice' { 'Cubicle Office' }
        default         { throw "ResourceType must be 'HotelOffice' or 'CubicleOffice', got '$ResourceType'." }
    }

    if (-not $DisplayName) {
        $DisplayName = $Name
    }

    $existing = Get-Mailbox -Identity $Name -ErrorAction SilentlyContinue
    if ($existing) {
        throw "A mailbox named '$Name' already exists (RecipientTypeDetails: $($existing.RecipientTypeDetails)). Choose a different Name."
    }

    if (-not $PSCmdlet.ShouldProcess($Name, "Create room mailbox ($categoryName), $BookingWindowDays-day booking window")) {
        return
    }

    Write-Host "Creating room mailbox '$Name' ($categoryName)..." -ForegroundColor Cyan
    $mailbox = New-Mailbox -Name $Name -DisplayName $DisplayName -Room -Alias $Name

    Write-Host "Tagging mailbox with resource type '$ResourceType' and capacity $Capacity..." -ForegroundColor Cyan
    Set-Mailbox -Identity $mailbox.Identity `
        -CustomAttribute1 $ResourceType `
        -ResourceCapacity $Capacity `
        -Office $Location

    Write-Host "Applying $BookingWindowDays-day booking window..." -ForegroundColor Cyan
    Set-CalendarProcessing -Identity $mailbox.Identity `
        -AutomateProcessing AutoAccept `
        -BookingWindowInDays $BookingWindowDays `
        -AllowConflicts $false

    # Ensure the room list for this category exists, then add the room to it
    # so it's grouped correctly in Outlook's room finder.
    $roomList = Get-DistributionGroup -Identity $categoryName -ErrorAction SilentlyContinue
    if (-not $roomList) {
        Write-Host "Room list '$categoryName' not found - creating it..." -ForegroundColor Cyan
        $roomList = New-DistributionGroup -Name $categoryName -DisplayName $categoryName -RoomList
    }
    Add-DistributionGroupMember -Identity $roomList.Identity -Member $mailbox.Identity -ErrorAction SilentlyContinue

    Write-Host "Room resource '$Name' created and added to room list '$categoryName'." -ForegroundColor Green
    Get-Mailbox -Identity $mailbox.Identity |
        Select-Object Name, DisplayName, PrimarySmtpAddress, RecipientTypeDetails, CustomAttribute1, @{Name = 'BookingWindowInDays'; Expression = { $BookingWindowDays } }
}

# Reuse an existing Exchange Online session if one is already connected.
$activeSession = Get-ConnectionInformation -ErrorAction SilentlyContinue |
    Where-Object { $_.State -eq 'Connected' }
if (-not $activeSession) {
    Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
    Connect-ExchangeOnline -ShowBanner:$false
}

if ($PSCmdlet.ParameterSetName -eq 'FromCsv') {
    $rows = Import-Csv -Path $CsvPath
    $requiredColumns = 'Name', 'ResourceType'
    $missingColumns = $requiredColumns | Where-Object { $_ -notin $rows[0].PSObject.Properties.Name }
    if ($missingColumns) {
        throw "CSV file '$CsvPath' is missing required column(s): $($missingColumns -join ', ')"
    }

    foreach ($row in $rows) {
        try {
            $rowWindow = if ($row.PSObject.Properties.Name -contains 'BookingWindowDays' -and $row.BookingWindowDays) {
                [int]$row.BookingWindowDays
            }
            else {
                $BookingWindowDays
            }
            $rowCapacity = if ($row.PSObject.Properties.Name -contains 'Capacity' -and $row.Capacity) { [int]$row.Capacity } else { 1 }

            New-SingleRoomResource -Name $row.Name -DisplayName $row.DisplayName -ResourceType $row.ResourceType `
                -Capacity $rowCapacity -Location $row.Location -BookingWindowDays $rowWindow
        }
        catch {
            Write-Warning "Failed to create room '$($row.Name)': $_"
        }
    }
}
else {
    New-SingleRoomResource -Name $Name -DisplayName $DisplayName -ResourceType $ResourceType `
        -Capacity $Capacity -Location $Location -BookingWindowDays $BookingWindowDays
}
