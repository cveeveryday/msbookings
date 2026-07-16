#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
    Applies the room-booking framework's calendar processing policy, including the
    14-day advance-booking window, to one or more Exchange Online room mailboxes.

.DESCRIPTION
    Sets Set-CalendarProcessing on the target room mailboxes so they auto-accept
    bookings and reject anything more than BookingWindowDays out. This is a hard
    backstop on the mailbox itself - Microsoft Bookings (configured separately via
    New-BookingsStaffLink.ps1) is expected to enforce the same 14-day limit in its
    own booking policy, but a booking made directly against the room mailbox
    (e.g. via Outlook room finder) must still be rejected by the mailbox if it
    falls outside the window.

    Targets can be specified explicitly (-Identity) or resolved by resource
    category (-ResourceType), which looks up all room mailboxes tagged via
    CustomAttribute1 by New-RoomResource.ps1.

.PARAMETER Identity
    One or more room mailbox identities (name, alias, or email) to apply the policy to.

.PARAMETER ResourceType
    Apply the policy to every room mailbox tagged with this category
    (CustomAttribute1 = HotelOffice or CubicleOffice) instead of specifying -Identity.

.PARAMETER BookingWindowDays
    Maximum number of days in advance a booking is allowed. Defaults to 14 (2 weeks),
    the framework-wide policy.

.EXAMPLE
    ./Set-RoomBookingPolicy.ps1 -ResourceType HotelOffice

.EXAMPLE
    ./Set-RoomBookingPolicy.ps1 -Identity "HotelOffice-12A", "Cubicle-204" -BookingWindowDays 14
#>

[CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'ByResourceType')]
param(
    [Parameter(Mandatory, ParameterSetName = 'ByIdentity')]
    [string[]]$Identity,

    [Parameter(Mandatory, ParameterSetName = 'ByResourceType')]
    [ValidateSet('HotelOffice', 'CubicleOffice')]
    [string]$ResourceType,

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

if ($PSCmdlet.ParameterSetName -eq 'ByResourceType') {
    $rooms = Get-Mailbox -RecipientTypeDetails RoomMailbox -ResultSize Unlimited |
        Where-Object { $_.CustomAttribute1 -eq $ResourceType }
    if (-not $rooms) {
        Write-Warning "No room mailboxes found with CustomAttribute1 = '$ResourceType'."
        return
    }
}
else {
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

foreach ($room in $rooms) {
    if ($PSCmdlet.ShouldProcess($room.Name, "Set booking window to $BookingWindowDays days")) {
        Set-CalendarProcessing -Identity $room.Identity `
            -AutomateProcessing AutoAccept `
            -BookingWindowInDays $BookingWindowDays `
            -AllowConflicts $false

        Write-Host "Applied $BookingWindowDays-day booking window to '$($room.Name)'." -ForegroundColor Green
    }
}

if ($rooms) {
    $rooms | ForEach-Object { Get-CalendarProcessing -Identity $_.Identity } |
        Select-Object Identity, AutomateProcessing, BookingWindowInDays, AllowConflicts
}
