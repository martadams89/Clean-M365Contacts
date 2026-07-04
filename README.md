# Clean-M365Contacts

Remove **duplicate contacts** from a **Microsoft 365 / Office 365** mailbox — including **new Outlook** and **Outlook on the web**, where the old COM/VBA and "merge duplicates" tricks no longer work.

A single, dependency-free PowerShell script with an interactive menu, a dry-run preview, and a safe delete that sends duplicates to Deleted Items (recoverable). It talks to the **Microsoft Graph API** directly, so there are **no modules to install** and it runs on both **Windows PowerShell 5.1 (incl. the ISE)** and **PowerShell 7+**.

[![CI](https://github.com/martadams89/Clean-M365Contacts/actions/workflows/ci.yml/badge.svg)](https://github.com/martadams89/Clean-M365Contacts/actions/workflows/ci.yml)
[![PowerShell 5.1 | 7+](https://img.shields.io/badge/PowerShell-5.1%20%7C%207%2B-5391FE?logo=powershell&logoColor=white)](https://learn.microsoft.com/powershell/)
[![Lint: PSScriptAnalyzer](https://img.shields.io/badge/lint-PSScriptAnalyzer-brightgreen?logo=powershell&logoColor=white)](https://github.com/PowerShell/PSScriptAnalyzer)
[![Tested with Pester](https://img.shields.io/badge/tested%20with-Pester-5391FE)](https://pester.dev/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![PRs Welcome](https://img.shields.io/badge/PRs-welcome-brightgreen.svg)](CONTRIBUTING.md)
[![Platform: Windows](https://img.shields.io/badge/platform-Windows-0078D6?logo=windows&logoColor=white)](#requirements)

---

## Why this exists

Cleaning up duplicate contacts after a mailbox migration or import used to mean the Outlook desktop "Clean Up" tools or a VBA macro. **New Outlook and OWA removed those options**, and the Exchange Online PowerShell module deliberately can't touch personal contacts inside a mailbox. That leaves two awkward paths — deleting by hand, or writing Graph plumbing from scratch. This script is that plumbing, packaged so you can just run it.

## Features

- **Interactive menu** — sign in, scan, preview, export, delete.
- **Dry run first** — see exactly which contact is kept and which are deleted before anything changes.
- **Keeps the richest record** — when it finds a duplicate set it keeps the most complete contact (most filled-in fields) and removes the stubs, rather than keeping whichever happened to be first.
- **Three matching modes** — `Name + Email` (safest), `Name only`, or `Email only`.
- **Recoverable deletes** — removed contacts go to the mailbox's Deleted Items.
- **CSV audit trail** — preview and deletion logs written to your Desktop.
- **Silent re-runs** — the refresh token is cached (DPAPI-encrypted, per-user/per-machine) so you don't re-authenticate on every run.
- **No modules, no app registration** — uses delegated device-code sign-in against Microsoft's public Graph CLI client.

## Requirements

- Windows with **Windows PowerShell 5.1** (built in, works in the ISE) **or PowerShell 7+**.
- The signed-in account must have a mailbox and be able to consent to the delegated `Contacts.ReadWrite` and `User.Read` scopes (both are low-risk, user-consentable in most tenants).
- Internet access to `login.microsoftonline.com` and `graph.microsoft.com`.

> **How access works:** the script uses **delegated** sign-in, so it only ever touches the mailbox of the account you sign in with. Run it on the user's machine, sign in as that user, and it is impossible to hit the wrong mailbox.

## Quick start

```powershell
# PowerShell 7
pwsh -File .\Clean-M365Contacts.ps1

# Windows PowerShell 5.1 / ISE
powershell -ExecutionPolicy Bypass -File .\Clean-M365Contacts.ps1
```

Then:

1. Choose **1** to sign in. A browser opens to `microsoft.com/devicelogin` and the code is copied to your clipboard — sign in as the target user.
2. Choose **3** to preview duplicates (nothing is changed).
3. Optionally choose **4** to export the preview to CSV.
4. Choose **5**, type `DELETE` to confirm, and the duplicates are removed.

## The menu

```
 1. Sign in / switch account
 2. Scan (show counts)
 3. Preview duplicates (dry run)
 4. Export preview to CSV
 5. Delete duplicates
 6. Change match mode
 7. Sign out (clears saved session on this machine)
 0. Exit
```

## Matching modes

| Mode          | A duplicate is...                              | Use when                                        |
| ------------- | ---------------------------------------------- | ----------------------------------------------- |
| `Name + Email`| same display name **and** same first email     | the default — clean double-imports              |
| `Name`        | same display name (email ignored)              | the same person imported with different emails  |
| `Email`       | same first email (name ignored)                | records where names vary but the address is key |

Contacts whose match key would be empty (blank name / no email) are **never** auto-flagged.

## How it works

- Signs in with the **OAuth 2.0 device-code flow** against the public **Microsoft Graph Command Line Tools** client (`14d82eec-...`), requesting delegated `Contacts.ReadWrite` + `User.Read`.
- Reads `/me/contacts` (paged), groups them by the chosen key, scores each group for completeness, and marks all but the best in each group for deletion.
- Deletes via `DELETE /me/contacts/{id}` — which moves the item to Deleted Items.
- Caches only the **refresh token**, encrypted with Windows **DPAPI** (`ConvertFrom-SecureString`), scoped to the current user on the current machine.

## Security notes

- No client secret, no certificate, no app registration — it's a public client doing an interactive delegated sign-in.
- The cached refresh token lives at `%LOCALAPPDATA%\ContactCleanup\session.dat`, DPAPI-encrypted; it can't be read by another user or copied to another machine. Menu option **7** deletes it.
- Deletions are recoverable from Deleted Items, but review the dry run first.
- Conditional Access sign-in-frequency policies can still force a fresh prompt regardless of the cache.

## Development

```powershell
# Lint
Install-Module PSScriptAnalyzer -Scope CurrentUser
Invoke-ScriptAnalyzer -Path . -Recurse -Settings ./PSScriptAnalyzerSettings.psd1

# Test
Install-Module Pester -MinimumVersion 5.5.0 -Scope CurrentUser
Invoke-Pester -Path ./tests
```

CI runs PSScriptAnalyzer plus the Pester suite on **PowerShell 7 (Windows + Linux)** and **Windows PowerShell 5.1** on every push and pull request. See [`.github/workflows/ci.yml`](.github/workflows/ci.yml).

## Contributing

Issues and PRs welcome — see [CONTRIBUTING.md](CONTRIBUTING.md). If this saved you a migration headache, a star helps others find it.

## License

[MIT](LICENSE)

---

<sub>Keywords: remove duplicate contacts Microsoft 365, delete duplicate contacts new Outlook, Office 365 dedupe contacts PowerShell, Outlook on the web duplicate contacts, Microsoft Graph delete contacts script, clean up contacts after migration, Exchange Online personal contacts PowerShell.</sub>
