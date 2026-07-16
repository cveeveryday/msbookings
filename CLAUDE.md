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

**Booking window:** Employees may reserve either resource type **up to 14 days (2 weeks) in advance** — no further out. This is enforced in two places:
1. **Microsoft Bookings** governs the actual booking UX/policy employees interact with.
2. **`Set-CalendarProcessing -BookingWindowInDays 14`** on the underlying room mailbox is a hard backstop, in case a booking is attempted directly against the mailbox calendar (e.g. via Outlook room finder) rather than through Bookings.

**Room mailbox → Bookings staff pattern:** A normal Bookings business associates *people* (staff) with *services*, not physical resources. To repurpose it for room booking, each room mailbox is added as Bookings "staff" (via Graph `solutions/bookingBusinesses/{id}/staffMembers`), and a Bookings "service" represents the act of reserving that room type. Employees pick a room (staff member) and a time slot the way they'd normally pick a person to meet with.

---

## Auth

Scripts use **interactive delegated auth**, run by an admin:
- `Connect-ExchangeOnline` (interactive) for room mailbox creation and calendar processing config.
- `Connect-MgGraph` (interactive, delegated scopes) for Bookings business/staff/service management.

No app registration or certificate-based app-only auth is set up yet. If this framework needs to run unattended (e.g. scheduled provisioning), that would require adding an Entra app registration with a certificate and switching to app-only auth — treat that as a future enhancement, not the current design.

---

## Current State

Bootstrapped. No scripts implemented yet.

---

## Directory Contract

```
msbookings/
├── CLAUDE.md
├── README.md
├── scripts/
│   ├── New-RoomResource.ps1          ← creates a room mailbox (Hotel Office | Cubicle Office)
│   ├── Set-RoomBookingPolicy.ps1     ← applies the 14-day booking window + calendar processing rules
│   ├── New-BookingsStaffLink.ps1     ← registers a room mailbox as staff on the Bookings business
│   └── Get-RoomInventory.ps1         ← reports current room resources and their configuration
└── docs/
    └── architecture.md
```
