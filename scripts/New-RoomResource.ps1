#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
    Creates an Exchange Online room mailbox for the room-booking framework.

.DESCRIPTION
    Provisions a room mailbox tagged as either "Hotel Office" or "Cubicle Office"
    and adds it to the matching room list (a RoomList distribution group), so it
    shows up in Outlook's room finder grouped by resource type. The resource type
    is also written to CustomAttribute1, which downstream scripts
    (New-BookingsStaffLink.ps1, Get-RoomInventory.ps1) use to identify rooms
    belonging to this framework and their category.

    This script only creates and tags the mailbox. It does not configure the
    booking window (see Set-RoomBookingPolicy.ps1) or link the room into
    Microsoft Bookings (see New-BookingsStaffLink.ps1).

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

.EXAMPLE
    ./New-RoomResource.ps1 -Name "HotelOffice-12A" -DisplayName "Hotel Office 12A - 3rd Floor" -ResourceType HotelOffice -Location "Building A, Floor 3"

.EXAMPLE
    ./New-RoomResource.ps1 -Name "Cubicle-204" -ResourceType CubicleOffice -Capacity 1
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [ValidatePattern('^[A-Za-z0-9\-_]+$')]
    [string]$Name,

    [string]$DisplayName,

    [Parameter(Mandatory)]
    [ValidateSet('HotelOffice', 'CubicleOffice')]
    [string]$ResourceType,

    [ValidateRange(1, 100)]
    [int]$Capacity = 1,

    [string]$Location
)

$ErrorActionPreference = 'Stop'

# Human-readable room list / category name used both as the RoomList
# distribution group name and the CustomAttribute1 tag prefix.
$categoryName = switch ($ResourceType) {
    'HotelOffice'   { 'Hotel Office' }
    'CubicleOffice' { 'Cubicle Office' }
}

if (-not $DisplayName) {
    $DisplayName = $Name
}

# Reuse an existing Exchange Online session if one is already connected.
$activeSession = Get-ConnectionInformation -ErrorAction SilentlyContinue |
    Where-Object { $_.State -eq 'Connected' }
if (-not $activeSession) {
    Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
    Connect-ExchangeOnline -ShowBanner:$false
}

$existing = Get-Mailbox -Identity $Name -ErrorAction SilentlyContinue
if ($existing) {
    throw "A mailbox named '$Name' already exists (RecipientTypeDetails: $($existing.RecipientTypeDetails)). Choose a different -Name."
}

if ($PSCmdlet.ShouldProcess($Name, "Create room mailbox ($categoryName)")) {
    Write-Host "Creating room mailbox '$Name' ($categoryName)..." -ForegroundColor Cyan

    $mailbox = New-Mailbox -Name $Name -DisplayName $DisplayName -Room -Alias $Name

    Write-Host "Tagging mailbox with resource type '$ResourceType' and capacity $Capacity..." -ForegroundColor Cyan
    Set-Mailbox -Identity $mailbox.Identity `
        -CustomAttribute1 $ResourceType `
        -ResourceCapacity $Capacity `
        -Office $Location

    # Ensure the room list for this category exists, then add the room to it
    # so it's grouped correctly in Outlook's room finder.
    $roomList = Get-DistributionGroup -Identity $categoryName -ErrorAction SilentlyContinue
    if (-not $roomList) {
        Write-Host "Room list '$categoryName' not found - creating it..." -ForegroundColor Cyan
        $roomList = New-DistributionGroup -Name $categoryName -DisplayName $categoryName -RoomList
    }
    Add-DistributionGroupMember -Identity $roomList.Identity -Member $mailbox.Identity -ErrorAction SilentlyContinue

    Write-Host "Room resource '$Name' created and added to room list '$categoryName'." -ForegroundColor Green
    Get-Mailbox -Identity $mailbox.Identity | Select-Object Name, DisplayName, PrimarySmtpAddress, RecipientTypeDetails, CustomAttribute1
}
