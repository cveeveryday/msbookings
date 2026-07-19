# msbookings

PowerShell scripts to provision and manage Exchange Online room resources for a self-service employee room-booking framework, reserved through Microsoft Bookings.

## What this is

Two categories of bookable resource, each backed by an Exchange Online room mailbox:

- **Hotel Office** — hot-desking offices
- **Cubicle Office** — cubicle-style desks

Employees can reserve either up to a configurable advance-booking window — **14 days (2 weeks) by default**. Each room mailbox is registered as a **staff member** on a Microsoft Bookings business, so reserving a room looks like booking an appointment — employees pick a room and a time slot through the Bookings page/app.

See [CLAUDE.md](./CLAUDE.md) for the full architecture and design rationale.

## Scripts

| Script | Purpose |
| --- | --- |
| `scripts/New-RoomResource.ps1` | Creates a room mailbox, tags it, applies the booking window, and sets Place metadata (Building, Floor, Capacity, wheelchair accessibility) via `Set-Place`. Supports `-CsvPath` for bulk creation. |
| `scripts/Set-RoomBookingPolicy.ps1` | (Re)applies the booking window to existing rooms, by identity, by resource type, or in bulk via `-CsvPath`. |
| `scripts/New-BookingsService.ps1` | Defines the Bookings service for a resource category (e.g. "Hotel Office") and sets its booking window. |
| `scripts/New-BookingsStaffLink.ps1` | Registers a room mailbox as a staff member on a Microsoft Bookings business. |
| `scripts/Get-RoomInventory.ps1` | Reports current room resources and their configuration. |

### Setup order

1. Create the Bookings business itself (Bookings admin UI, or `New-MgBookingBusiness` — not yet scripted here).
2. `New-RoomResource.ps1` — create the room mailboxes.
3. `New-BookingsService.ps1` — once per resource category, to define the bookable service and its window.
4. `New-BookingsStaffLink.ps1` — once per room, to link it in as staff.

### Bulk creation via CSV

```powershell
./scripts/New-RoomResource.ps1 -CsvPath ./docs/sample-rooms.csv
```

See [`docs/sample-rooms.csv`](./docs/sample-rooms.csv) for the expected columns (`Name`, `ResourceType`, and optionally `DisplayName`, `Capacity`, `Location`, `BookingWindowDays`, `Building`, `Floor`, `IsWheelChairAccessible`).

### Setting the booking window

The advance-booking window is a parameter, not a hardcoded value:

```powershell
./scripts/New-RoomResource.ps1 -Name "Cubicle-206" -ResourceType CubicleOffice -BookingWindowDays 14
./scripts/Set-RoomBookingPolicy.ps1 -ResourceType HotelOffice -BookingWindowDays 21
./scripts/New-BookingsService.ps1 -BookingBusinessId "rooms@contoso.com" -ResourceType HotelOffice -BookingWindowDays 21
```

Keep the value consistent across all three — `New-BookingsService.ps1` sets the limit employees actually experience in the Bookings UI, while the mailbox-side window is a backstop.

## Requirements

- `ExchangeOnlineManagement` PowerShell module
- `Microsoft.Graph.Bookings` PowerShell module
- An account with permissions to create room mailboxes in Exchange Online and manage a Microsoft Bookings business via Graph

## Status

All five core scripts implemented. Not yet tested end-to-end against a live tenant.
