# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

---

## Project Identity

**Name:** msbookings
**Purpose:** PowerShell scripts to provision and manage Exchange Online room resources for a self-service employee room-booking framework, reserved through Microsoft Bookings.
**Core Design Principle:** Exchange Online room mailboxes are the source of truth for the resources themselves (identity, calendar, booking-window enforcement). Microsoft Bookings is the employee-facing reservation surface layered on top — each room mailbox is registered as a **staff member** on a Bookings business, so reserving a desk/cubicle looks like booking an appointment with that staff member.

---

## Booking Model

Two resource categories, each backed by Exchange Online room mailboxes:

- **Hotel Office** — hot-desking offices, booked ad hoc by any employee.
- **Cubicle Office** — cubicle-style desks, booked ad hoc by any employee.

**Booking window:** Employees may reserve either resource type **up to a configurable number of days in advance — 14 days (2 weeks) by default**. This is enforced in two places:
1. **Microsoft Bookings**, via `schedulingPolicy.maximumAdvance` on the `bookingService` (set by `New-BookingsService.ps1`), governs the actual booking UX/policy employees interact with.
2. **`Set-CalendarProcessing -BookingWindowInDays`** on the underlying room mailbox is a hard backstop, in case a booking is attempted directly against the mailbox calendar (e.g. via Outlook room finder) rather than through Bookings.

The window is a `-BookingWindowDays` parameter (default `14`) on `New-RoomResource.ps1` (applied at creation time), `Set-RoomBookingPolicy.ps1` (to change it later), and `New-BookingsService.ps1` (the Bookings-side limit) — not a hardcoded value, so it can be overridden globally or per room without editing the scripts. Keep the value consistent across all three or the mailbox backstop and the Bookings UX will disagree.

**Room mailbox → Bookings staff/service pattern:** A normal Bookings business associates *people* (staff) with *services*, not physical resources. To repurpose it for room booking:
- A `bookingService` per category ("Hotel Office" / "Cubicle Office") is defined once via `New-BookingsService.ps1` — this is what employees pick as the thing they're booking, and where the Bookings-side advance-booking window lives.
- Each room mailbox is added as Bookings "staff" via `New-BookingsStaffLink.ps1` (Graph `solutions/bookingBusinesses/{id}/staffMembers`) and associated with the matching service.

Employees pick a service (room type), then a room (staff member) and a time slot, the way they'd normally pick a person to meet with.

---

## Auth

Scripts use **interactive delegated auth**, run by an admin:
- `Connect-ExchangeOnline` (interactive) for room mailbox creation and calendar processing config.
- `Connect-MgGraph` (interactive, delegated scopes) for Bookings business/staff/service management.

No app registration or certificate-based app-only auth is set up yet. If this framework needs to run unattended (e.g. scheduled provisioning), that would require adding an Entra app registration with a certificate and switching to app-only auth — treat that as a future enhancement, not the current design.

---

## Current State

Five core scripts implemented: `New-RoomResource.ps1`, `Set-RoomBookingPolicy.ps1`, `New-BookingsService.ps1`, `New-BookingsStaffLink.ps1`, `Get-RoomInventory.ps1`. Not yet tested end-to-end against a live tenant.

**Provisioning order:** the Bookings business itself must already exist (created manually in the Bookings admin UI, or via `New-MgBookingBusiness` — not yet scripted here), then: 1) `New-RoomResource.ps1` to create room mailboxes, 2) `New-BookingsService.ps1` once per resource category to define the bookable service, 3) `New-BookingsStaffLink.ps1` per room to link it to its service.

---

## Directory Contract

```
msbookings/
├── CLAUDE.md
├── README.md
├── scripts/
│   ├── New-RoomResource.ps1          ← creates room mailbox(es) (Hotel Office | Cubicle Office); -CsvPath for bulk
│   ├── Set-RoomBookingPolicy.ps1     ← (re)applies the booking window + calendar processing rules; -CsvPath for bulk
│   ├── New-BookingsService.ps1       ← defines the Bookings service for a resource category + its booking window
│   ├── New-BookingsStaffLink.ps1     ← registers a room mailbox as staff on the Bookings business
│   └── Get-RoomInventory.ps1         ← reports current room resources and their configuration
└── docs/
    ├── architecture.md
    └── sample-rooms.csv              ← example bulk-import file for -CsvPath
```

## Bulk Provisioning via CSV

`New-RoomResource.ps1` and `Set-RoomBookingPolicy.ps1` both accept `-CsvPath` as an alternative to their single-room parameters, for provisioning/updating many rooms in one run. See `docs/sample-rooms.csv` for the format:

- **`New-RoomResource.ps1 -CsvPath`** required columns: `Name`, `ResourceType`. Optional: `DisplayName`, `Capacity`, `Location`, `BookingWindowDays` (per-row override of `-BookingWindowDays`).
- **`Set-RoomBookingPolicy.ps1 -CsvPath`** required column: `Identity` (or `Name`, so the same CSV can be reused). Optional: `BookingWindowDays` (per-row override).

A row that fails (e.g. duplicate name, mailbox not found) is reported as a warning and skipped — the rest of the file still runs.
