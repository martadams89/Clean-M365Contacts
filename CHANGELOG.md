# Changelog

All notable changes to this project are documented here. The format is based on
[Keep a Changelog](https://keepachangelog.com/en/1.0.0/), and this project
adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.1.0](https://github.com/martadams89/Clean-M365Contacts/compare/v1.0.0...v1.1.0) (2026-07-04)


### Features

* automated releases, Renovate, PowerShell 7-only CI ([09babf3](https://github.com/martadams89/Clean-M365Contacts/commit/09babf33f001bec4e8f900ca5d8f7c5832632978))


### Bug Fixes

* **ci:** commit release-please-config.json so releases can run ([#1](https://github.com/martadams89/Clean-M365Contacts/issues/1)) ([b09bbe8](https://github.com/martadams89/Clean-M365Contacts/commit/b09bbe80cd6f4e9f6ed305eb6fdd91f08ce3d14f))
* scope shared Pester variable to satisfy PSScriptAnalyzer ([cecdeaf](https://github.com/martadams89/Clean-M365Contacts/commit/cecdeaf4ccb1d3aec47f360b62cb47fd38388c89))

## [1.0.0] - 2026-07-04

### Added

- Interactive menu-driven de-duplication of personal contacts in a Microsoft 365
  mailbox via the Microsoft Graph REST API (works with new Outlook / OWA).
- Delegated device-code sign-in against the public Graph CLI client - no modules,
  no app registration.
- Dry-run preview and CSV export before any deletion.
- Completeness scoring so the richest record in a duplicate set is kept.
- Three matching modes: Name + Email, Name only, Email only.
- Safety guard so contacts with an empty match key are never auto-flagged.
- Recoverable deletes (contacts move to Deleted Items) with a typed `DELETE`
  confirmation and a CSV audit log.
- DPAPI-encrypted refresh-token cache for silent re-runs; sign-out clears it.
- Runs on Windows PowerShell 5.1 (incl. the ISE) and PowerShell 7+.
- PSScriptAnalyzer settings, Pester unit tests, and a GitHub Actions CI pipeline
  covering PowerShell 7 (Windows + Linux) and Windows PowerShell 5.1.

[1.0.0]: https://github.com/YOUR-USERNAME/Clean-M365Contacts/releases/tag/v1.0.0
