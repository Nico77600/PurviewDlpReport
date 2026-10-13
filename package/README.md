# Purview DLP Report

A PowerShell 7 tool that collects Microsoft Purview DLP Activity Explorer events and writes CSV and HTML reports.

This folder contains everything needed to run the tool: `Invoke-PurviewDlpReport.ps1`, the module and its C# engine, the configuration, the report template, the SQLite library and both guides. Tests and build tools stay outside it, in the repository.

> [!IMPORTANT]
> Files downloaded from the Internet may be blocked by Windows. Unblock them once, from this folder:
>
> ```powershell
> Get-ChildItem . -Recurse -File | Unblock-File
> ```

## Signed package integrity (2.1.2)

The five signed PowerShell files (launcher, module, module metadata, verifier and `PackageIntegrity.psd1`) require valid Nicolas Fabert signatures and RFC3161 timestamps. Expected signer: `0F28E442889F17B5F4D251E8F96E7CC6E2C9C6D6`. The verifier authenticates sixteen stable raw SHA256 entries, including the C# source, original vendor DLLs, template, README and both HTML guides. Missing or changed files, unexpected stable code and protected reparse points fail before code loading; there is no unsigned fallback.

Edit only `config\PurviewDlpReport.config.psd1`. `data`, `logs`, `reports` and old `bin` output are not trusted code. Repository-only `LICENSE`, the two Markdown guides and the ten named `docs\images\*.png` sources are an explicit, unhashed documentation boundary; they are not loaded at runtime or copied by the packager. Unknown documentation names and executable files are not exempt.

Review the publisher and trusted chain before unblocking downloads. Follow your organization's execution policy; do not disable it or integrity verification to resolve an error. Chain/revocation checks can need access to issuer services; an unavailable or untrusted chain is a blocking failure, not a reason to import arbitrary trust roots. See the HTML developer guide for source edits and certificate rotation.

Verified C# compiles in memory once per process. Startup recompilation and warm checks have overhead; unsigned engine caches are never loaded. PowerShell 7.4+ is required (business tests used 7.6.6); Windows PowerShell 5.1 is for signature inspection only. Offline and limited functional testing are not broad performance or production validation.

Coverage measures queried time. Check `report-metadata.json` and the exit code: a rounded 100% or a folder without `.pending` does not prove complete coverage. Tenant/environment display is disabled by default; company branding is independent of the fixed author.

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
| `src\` | Authenticated C# engine source, compiled in memory once per process. |
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

## Disclaimer

This Script is a Personal project.
It's provided "AS-IS". It's not an official Microsoft product so no support can be expected from Microsoft.

As any scripts you must read carefully the documentation and test it first in a Test environment before any run in Production.
