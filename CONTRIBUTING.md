# Contributing

Thanks for helping improve Clean-M365Contacts. Bug reports, feature ideas, and pull requests are all welcome.

## Reporting bugs

Open an issue and include:

- Your PowerShell version (`$PSVersionTable.PSVersion`) and host (ISE, console, VS Code, PowerShell 7).
- What you ran and what happened, including the exact error text.
- Whether the mailbox is classic Outlook, new Outlook, or OWA.

Please **never paste tokens, device codes, or the contents of `session.dat`** into an issue.

## Development setup

No build step — it's a single script. To work on it you'll want the linter and test runner:

```powershell
Install-Module PSScriptAnalyzer -Scope CurrentUser
Install-Module Pester -MinimumVersion 5.5.0 -Scope CurrentUser
```

## Before you open a PR

Run both locally; CI runs the same checks on PowerShell 7 (Windows + Linux) and Windows PowerShell 5.1.

```powershell
# Lint - must be clean (no errors or warnings)
Invoke-ScriptAnalyzer -Path . -Recurse -Settings ./PSScriptAnalyzerSettings.psd1

# Tests - must pass
Invoke-Pester -Path ./tests
```

If you add or change logic in a pure function (key generation, scoring, duplicate detection, JWT decoding), please add or update a test for it in `tests/`.

## Coding guidelines

- Keep it dependency-free: standard cmdlets and `Invoke-RestMethod` only, no external modules required to run.
- Stay compatible with **Windows PowerShell 5.1** — avoid PS7-only syntax (ternary `? :`, `??`, `-Parallel`, `Clean {}` blocks).
- Use **approved verbs** for function names (`Get-Verb`).
- Keep the file **ASCII**, saved as **UTF-8 with BOM** so Windows PowerShell 5.1 parses non-ASCII correctly.
- Anything that changes a mailbox must stay behind the dry-run/confirmation pattern.

## Scope

This tool targets **personal contacts** in a single signed-in mailbox via delegated Graph access. Tenant-wide, app-only, or directory-object contact management is intentionally out of scope.
