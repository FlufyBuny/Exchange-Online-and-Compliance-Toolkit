# EPSILON v1.1 – Release Notes

## Overview

EPSILON v1.1 represents the first stable release of the toolkit, consolidating multiple development phases (ALPHA, DELTA, and GAMMA) into a unified, menu-driven PowerShell solution for Exchange Online and Compliance management.

This release establishes the core structure and functionality that future versions build upon.

---

## Development Evolution

EPSILON v1.1 is the result of several internal iterations:

- ALPHA
  - Initial concept and basic Exchange Online interaction

- DELTA (v0.01 → v0.8)
  - Introduction of core menu system
  - Archiving and mailbox management features
  - Compliance search and purge functionality
  - Iterative improvements and feature expansion

- GAMMA (v0.9 → v1.0)
  - Stabilization of core features
  - Refinement of menu structure and usability
  - Preparation for production use

---

## Core Features

### Exchange Online

- Mailbox management (create, remove, view details)
- Archive mailbox enablement
- Auto-expanding archive support
- Managed Folder Assistant execution
- Inbox rule visibility and management
- Mail user and contact management

---

### Compliance (Purview)

- Compliance search creation
- Search execution and monitoring
- Email purge functionality
- Subject-based search filtering

---

## Improvements

- Consolidated multiple scripts into a single toolkit
- Standardized menu-driven interface
- Improved usability for day-to-day administration
- Streamlined common administrative workflows

---

## Known Limitations

- Inbox rule creation timestamps are not available via standard cmdlets
- Some Exchange operations (e.g., Managed Folder Assistant) may return intermittent service-side errors

---

## Notes

This release serves as the baseline for EPSILON moving forward, providing a stable and consistent foundation for continued development and feature enhancements.
