# Purview DLP Report

A PowerShell 7 tool that collects Microsoft Purview DLP Activity Explorer events and writes CSV and HTML reports.

This folder contains everything needed to run the tool: `Invoke-PurviewDlpReport.ps1`, the module and its C# engine, the configuration, the report template, the SQLite library and both guides. Tests and build tools stay outside it, in the repository.

> [!IMPORTANT]
> Files downloaded from the Internet may be blocked by Windows. Unblock them once, from this folder:
>
> ```powershell
> Get-ChildItem . -Recurse -File | Unblock-File
> ```

## Requirements

- PowerShell 7.4 or later; 7.6 for ExchangeOnlineManagement 3.10 or later.
- ExchangeOnlineManagement 3.9.0 or later, but not 3.10.0 in certificate mode.
- Windows Terminal for emoji and colours.
- A Purview role that can read Activity Explorer.
- SQLite is bundled in `lib\sqlite`.

## Quick start

```powershell
notepad .\config\PurviewDlpReport.config.psd1      # TenantId, PolicyId, RuleId, authentication

.\Invoke-PurviewDlpReport.ps1 -Mode Status          # checks the configuration, no connection
.\Invoke-PurviewDlpReport.ps1                       # report of the last 7 days
.\Invoke-PurviewDlpReport.ps1 -Range PreviousMonth  # previous calendar month
.\Invoke-PurviewDlpReport.ps1 -Mode Collect         # daily collection (scheduled task)
```

## Content

| Item | Role |
|---|---|
| `config\` | Configuration file to fill in. |
| `docs\` | User and developer guides, Markdown and self-contained HTML. |
| `lib\` | Bundled SQLite libraries. |
| `src\` | C# engine source, compiled on first use. |
| `templates\` | HTML report template. |
| `Invoke-PurviewDlpReport.ps1` | Entry script. |
| `PurviewDlpReport.psd1` | Module manifest. |
| `PurviewDlpReport.psm1` | Module loader. |
| `README.md` | This package readme. |
| `LICENSE` | MIT license. |
| `THIRD-PARTY-NOTICES.md` | Notices for bundled components. |

## Documentation

- [User guide](docs/PurviewDlpReport-UserGuide.md) - also `docs/PurviewDlpReport-UserGuide.html`, a single file to open locally
- [Developer guide](docs/PurviewDlpReport-Guide.md) - also `docs/PurviewDlpReport-Guide.html`

Project page, releases and change log: https://github.com/Nico77600/PurviewDlpReport

License: [MIT](LICENSE). Third-party components: [THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md).
