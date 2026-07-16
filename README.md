# msbookings

PowerShell scripts to provision and manage Exchange Online room resources for a self-service employee room-booking framework, reserved through Microsoft Bookings.

## What this is

Two categories of bookable resource, each backed by an Exchange Online room mailbox:

- **Hotel Office** — hot-desking offices
- **Cubicle Office** — cubicle-style desks

Employees can reserve either up to **2 weeks (14 days) in advance**. Each room mailbox is registered as a **staff member** on a Microsoft Bookings business, so reserving a room looks like booking an appointment — employees pick a room and a time slot through the Bookings page/app.

See [CLAUDE.md](./CLAUDE.md) for the full architecture and design rationale.

## Requirements

- `ExchangeOnlineManagement` PowerShell module
- `Microsoft.Graph` PowerShell module
- An account with permissions to create room mailboxes in Exchange Online and manage a Microsoft Bookings business via Graph

## Status

Bootstrapped — no scripts implemented yet.
