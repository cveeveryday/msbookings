#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
    Reports the current room resources in the booking framework and their configuration.

.DESCRIPTION
    Lists Exchange Online room mailboxes tagged by New-RoomResource.ps1
    (CustomAttribute1 = HotelOffice or CubicleOffice), along with the booking
    window applied by Set-RoomBookingPolicy.ps1. Optionally cross-references
    each room against a Microsoft Bookings business to report whether
    New-BookingsStaffLink.ps1 has been run for it, so gaps in the pipeline
    (a room mailbox that was never linked into Bookings, or vice versa) are
    easy to spot in one report.

.PARAMETER ResourceType
    Restrict the report to a single category. Defaults to reporting both.

.PARAMETER BookingBusinessId
    Optional. The id (typically the business's email address, e.g.
    "rooms@contoso.com") of a Microsoft Bookings business to check each room
    against. Requires the Microsoft.Graph.Bookings module. If omitted, the
    Bookings link status is not checked.

.EXAMPLE
    ./Get-RoomInventory.ps1

.EXAMPLE
    ./Get-RoomInventory.ps1 -ResourceType HotelOffice -BookingBusinessId "rooms@contoso.com"
#>

[CmdletBinding()]
param(
    [ValidateSet('HotelOffice', 'CubicleOffice')]
    [string]$ResourceType,

    [string]$BookingBusinessId
)

$ErrorActionPreference = 'Stop'

$exoSession = Get-ConnectionInformation -ErrorAction SilentlyContinue |
    Where-Object { $_.State -eq 'Connected' }
if (-not $exoSession) {
    Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
    Connect-ExchangeOnline -ShowBanner:$false
}

$rooms = Get-Mailbox -RecipientTypeDetails RoomMailbox -ResultSize Unlimited |
    Where-Object { $_.CustomAttribute1 -in @('HotelOffice', 'CubicleOffice') }
if ($ResourceType) {
    $rooms = $rooms | Where-Object { $_.CustomAttribute1 -eq $ResourceType }
}

if (-not $rooms) {
    Write-Warning "No tagged room resources found$(if ($ResourceType) { " for resource type '$ResourceType'" })."
    return
}

$staffByEmail = $null
if ($BookingBusinessId) {
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Bookings)) {
        Write-Warning "Microsoft.Graph.Bookings module not found - skipping Bookings link check."
    }
    else {
        if (-not (Get-MgContext)) {
            Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
            Connect-MgGraph -Scopes 'Bookings.Read.All' -NoWelcome
        }
        # Fetch staff once and index by email rather than querying per-room, since
        # a booking business rarely exceeds 100 staff members (Bookings' own limit).
        $staffByEmail = @{}
        Get-MgBookingBusinessStaffMember -BookingBusinessId $BookingBusinessId -All |
            ForEach-Object { $staffByEmail[$_.EmailAddress] = $_ }
    }
}

$rooms | ForEach-Object {
    $room = $_
    $calProcessing = Get-CalendarProcessing -Identity $room.Identity

    $result = [ordered]@{
        Name                = $room.Name
        DisplayName         = $room.DisplayName
        PrimarySmtpAddress  = $room.PrimarySmtpAddress
        ResourceType        = $room.CustomAttribute1
        Capacity            = $room.ResourceCapacity
        Location            = $room.Office
        AutomateProcessing  = $calProcessing.AutomateProcessing
        BookingWindowInDays = $calProcessing.BookingWindowInDays
    }

    if ($staffByEmail) {
        $result.LinkedToBookings = $staffByEmail.ContainsKey($room.PrimarySmtpAddress)
        $result.BookingsRole = $staffByEmail[$room.PrimarySmtpAddress].Role
    }

    [pscustomobject]$result
} | Sort-Object ResourceType, Name
