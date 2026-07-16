#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
    Applies the room-booking framework's calendar processing policy, including the
    advance-booking window, to one or more Exchange Online room mailboxes.

.DESCRIPTION
    Sets Set-CalendarProcessing on the target room mailboxes so they auto-accept
    bookings and reject anything more than BookingWindowDays out. This is a hard
    backstop on the mailbox itself - Microsoft Bookings (configured separately via
    New-BookingsStaffLink.ps1) is expected to enforce the same limit in its own
    booking policy, but a booking made directly against the room mailbox (e.g. via
    Outlook room finder) must still be rejected by the mailbox if it falls outside
    the window.

    Note: New-RoomResource.ps1 already applies -BookingWindowDays at creation time.
    Use this script to change the window later, or to (re)apply it to rooms created
    outside this framework.

    Targets can be specified explicitly (-Identity), resolved by resource category
    (-ResourceType, via the CustomAttribute1 tag set by New-RoomResource.ps1), or
    supplied in bulk (-CsvPath), with an optional per-row window override.

.PARAMETER Identity
    One or more room mailbox identities (name, alias, or email) to apply the policy to.

.PARAMETER ResourceType
    Apply the policy to every room mailbox tagged with this category
    (CustomAttribute1 = HotelOffice or CubicleOffice) instead of specifying -Identity.

.PARAMETER CsvPath
    Path to a CSV file listing rooms to apply the policy to, instead of -Identity/-ResourceType.
    Required column: Identity (name, alias, or email). Optional column: BookingWindowDays
    (overrides -BookingWindowDays for that row only). See docs/sample-rooms.csv for an example
    (the same file works here, since it includes a Name column - see -Identity note below).

.PARAMETER BookingWindowDays
    Maximum number of days in advance a booking is allowed. Defaults to 14 (2 weeks),
    the framework-wide policy. Can be overridden per row in -CsvPath.

.EXAMPLE
    ./Set-RoomBookingPolicy.ps1 -ResourceType HotelOffice

.EXAMPLE
    ./Set-RoomBookingPolicy.ps1 -Identity "HotelOffice-12A", "Cubicle-204" -BookingWindowDays 14

.EXAMPLE
    ./Set-RoomBookingPolicy.ps1 -CsvPath ./docs/sample-rooms.csv -BookingWindowDays 14
#>

[CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'ByResourceType')]
param(
    [Parameter(Mandatory, ParameterSetName = 'ByIdentity')]
    [string[]]$Identity,

    [Parameter(Mandatory, ParameterSetName = 'ByResourceType')]
    [ValidateSet('HotelOffice', 'CubicleOffice')]
    [string]$ResourceType,

    [Parameter(Mandatory, ParameterSetName = 'FromCsv')]
    [ValidateScript({ Test-Path -Path $_ -PathType Leaf })]
    [string]$CsvPath,

    [ValidateRange(1, 365)]
    [int]$BookingWindowDays = 14
)

$ErrorActionPreference = 'Stop'

$activeSession = Get-ConnectionInformation -ErrorAction SilentlyContinue |
    Where-Object { $_.State -eq 'Connected' }
if (-not $activeSession) {
    Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
    Connect-ExchangeOnline -ShowBanner:$false
}

# Per-mailbox window overrides collected from -CsvPath (Identity -> BookingWindowDays).
# Empty for the -Identity and -ResourceType parameter sets, which always use the single
# script-wide -BookingWindowDays value.
$windowOverrides = @{}

switch ($PSCmdlet.ParameterSetName) {
    'ByResourceType' {
        $rooms = Get-Mailbox -RecipientTypeDetails RoomMailbox -ResultSize Unlimited |
            Where-Object { $_.CustomAttribute1 -eq $ResourceType }
        if (-not $rooms) {
            Write-Warning "No room mailboxes found with CustomAttribute1 = '$ResourceType'."
            return
        }
    }
    'ByIdentity' {
        $rooms = $Identity | ForEach-Object {
            $mbx = Get-Mailbox -Identity $_ -ErrorAction SilentlyContinue
            if (-not $mbx) {
                Write-Warning "Mailbox '$_' not found - skipping."
            }
            elseif ($mbx.RecipientTypeDetails -ne 'RoomMailbox') {
                Write-Warning "'$_' is a $($mbx.RecipientTypeDetails), not a RoomMailbox - skipping."
            }
            else {
                $mbx
            }
        }
    }
    'FromCsv' {
        $csvRows = Import-Csv -Path $CsvPath
        if ('Identity' -notin $csvRows[0].PSObject.Properties.Name -and 'Name' -notin $csvRows[0].PSObject.Properties.Name) {
            throw "CSV file '$CsvPath' must have an 'Identity' or 'Name' column."
        }

        $rooms = $csvRows | ForEach-Object {
            $rowIdentity = if ($_.PSObject.Properties.Name -contains 'Identity' -and $_.Identity) { $_.Identity } else { $_.Name }
            $mbx = Get-Mailbox -Identity $rowIdentity -ErrorAction SilentlyContinue
            if (-not $mbx) {
                Write-Warning "Mailbox '$rowIdentity' not found - skipping."
                return
            }
            if ($mbx.RecipientTypeDetails -ne 'RoomMailbox') {
                Write-Warning "'$rowIdentity' is a $($mbx.RecipientTypeDetails), not a RoomMailbox - skipping."
                return
            }
            if ($_.PSObject.Properties.Name -contains 'BookingWindowDays' -and $_.BookingWindowDays) {
                $windowOverrides[$mbx.Identity.ToString()] = [int]$_.BookingWindowDays
            }
            $mbx
        }
    }
}

foreach ($room in $rooms) {
    $window = if ($windowOverrides.ContainsKey($room.Identity.ToString())) { $windowOverrides[$room.Identity.ToString()] } else { $BookingWindowDays }

    if ($PSCmdlet.ShouldProcess($room.Name, "Set booking window to $window days")) {
        Set-CalendarProcessing -Identity $room.Identity `
            -AutomateProcessing AutoAccept `
            -BookingWindowInDays $window `
            -AllowConflicts $false

        Write-Host "Applied $window-day booking window to '$($room.Name)'." -ForegroundColor Green
    }
}

if ($rooms) {
    $rooms | ForEach-Object { Get-CalendarProcessing -Identity $_.Identity } |
        Select-Object Identity, AutomateProcessing, BookingWindowInDays, AllowConflicts
}
