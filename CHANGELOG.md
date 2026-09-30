# Changelog — Purview DLP Report

All notable changes are listed here. Versions follow MAJOR.MINOR.PATCH (see the guide, Annex E).
Author: Nicolas Fabert.

## [2.1.1] — 2026-09-30

First public release.

### Added
- Guide, **Annex F — Least-privilege access for the application**: service principal and dedicated Purview role group with the single role **Information Protection Reader**, step by step for the administrator. Lab results of 2026-09-30: `Exchange.ManageAsApp` alone opens the session but gives no access to Activity Explorer; Security Reader (Entra) 69 commands, Information Protection Analyst 41, **Information Protection Reader 13 and no write command**; five other read roles do not give the export. Pitfalls documented: `New-RoleGroup` without `-DisplayName`, Object ID of the enterprise application, transient "not present in the role definition" errors after re-creating a role group.
- `tools\New-DlpPackage.ps1`: copies only the files needed to run into `..\package\PurviewDlpReport-<version>`, ready to be zipped — no database (created empty at the first run), tenant values of the configuration emptied and checked absent from the package.

### Fixed
- **ExchangeOnlineManagement 3.10.0 in certificate mode**: its `Connect-IPPSSession` fails with "Object reference not set to an instance of an object" (reproduced in the lab; fixed in 3.10.1, validated). The tool no longer takes blindly the most recent installed version: it uses the version already loaded, otherwise the most recent one without a known issue for the authentication mode, says which version it skipped, and gives the command to run when none is usable. New test.
- **Icons in the classic Windows console** (PowerShell 7 started outside Windows Terminal): 16 of the symbols (✔ ✖ ⚠ ◆ ◷ ▤ …) did not exist in the console fonts and were shown as empty boxes. The symbol style now uses only characters of Consolas and Lucida Console (√ • ▲ × » ■ ► ↔ ↓ ≡ …); the emoji style of Windows Terminal / VS Code is unchanged. Frames keep their rounded corners (present in Consolas); the console font is read at run time and square corners are used only with Lucida Console or the raster font, which have no rounded corners. New test: every character written outside the emoji style belongs to the console font repertoire (code page 437 and Latin-1, rounded corners only where the font has them).
- `-Mode Status` on a new installation (no database yet) ended with an error; it now confirms that the configuration is valid and the engine ready, and shows the next command (exit code 0). New test.

### Changed
- `Target.PolicyName` and `Target.RuleName` are optional (display only): without them, the console shows the rule ID and the report subtitle no name. The automated tests use fictitious identifiers and a test copy of the configuration file, so they no longer depend on a tenant.
- **Retention**: database 180 days (was 400), logs 30 days (was 90). The configuration comment now gives the measured size, 1.5 to 2.5 GB per month at 30-50k messages per day (about 9 to 15 GB for 180 days), instead of 1 to 1.5 GB.
- Module manifest `PurviewDlpReport.psd1` reduced to the entries in use (the empty template entries and gallery metadata are removed). It is now the only export list (the `Export-ModuleMember` duplicate in the `.psm1` is removed), limited to the 29 functions called by the script and the tests; the helper functions stay internal to the module.
- Minimum **ExchangeOnlineManagement 3.9.0** (checked at start, with the install command in the message).
- Guide: project background (why messages sent to more than 25 recipients are reported, audit mode first, then blocking by the DLP policy or by `Set-Mailbox -RecipientLimits`, visibility and exception list). Annex B compares both sources on the same 48-hour window and explains the effect of the number of DLP policies without future figures; `TenantId` / `Organization` and `DisableWAM` explained.
- Documentation screenshots regenerated from an anonymized copy of the lab data: no lab domain, tenant name or application ID remains.

### Documented
- `DisableWAM` must stay `$true`: without it, PowerShell 7 crashed during the interactive sign-in (`0xC0000005` in `msalruntime.dll`, ExchangeOnlineManagement 3.9.2).
- `Connect-IPPSSession -Organization` rejects a tenant GUID ("Organization cannot be a Guid"): the certificate mode needs the `xxx.onmicrosoft.com` name, while `TenantId` remains the safety check of both modes.

## [2.1.0] — 2026-09-29

### Added
- **Certificate (app-only) mode validated end to end** in the lab: dedicated application `Purview DLP Report - collector`, read-only **Security Reader** role, `-Mode Collect` with exit code 0.
- Console redesign: title card, numbered step pills with icons, a real table for the collection windows, coverage bars in `-Mode Status`, framed summary card. Emoji in Windows Terminal / VS Code, Unicode symbols elsewhere; `PDR_ICONS` (Emoji | Symbols | Ascii), `NO_COLOR` and `PDR_FORCE_COLOR` environment variables. The log file never contains colours or icons.
- Guide redesign: parts, one card per chapter with icon, callouts, visual steps and diagrams, sticky navigation, copy buttons, light/dark theme — built by `tools\Build-Documentation.ps1` from the Markdown source.

### Fixed
- Durations lost their decimals (`[Math]::Max(0, $x)` selected the integer overload).
- `-Mode Status` described the first collected day as "collectable 0 more day(s)" when its earlier hours were already beyond the 30-day retention.

### Documented
- `Exchange.ManageAsApp` must be granted on **Office 365 Exchange Online** (owner of the `ps.compliance.protection.outlook.com` audience in the lab tenant), not only on Microsoft Exchange Online Protection; how to check the token `roles` claim.
- PowerShell pitfalls: `-f` precedence, case-insensitive variables (`$C`/`$c`, `$matches`), flattened nested arrays.

## [2.0.0] — 2026-09-29

Complete redesign. Replaces version 1 (unified audit log script) and the Activity Explorer / Graph prototypes.

### Added
- Single entry point `Invoke-PurviewDlpReport.ps1` with three modes: `Report` (default), `Collect` (scheduled task), `Status`.
- Configuration file `config\PurviewDlpReport.config.psd1`, checked at start (all errors listed at once); a few report settings can be overridden on the command line.
- Collection from **Microsoft Purview Activity Explorer** only (`Export-ActivityExplorerData`, filters Exchange / DLPRuleMatch / policy / rule at the source).
- Local **SQLite** database (Microsoft.Data.Sqlite, bundled): history beyond the 30 days of Activity Explorer, restartable collection without duplicates (`RecordIdentity` key), re-collection of recent hours (`SettlingHours`), detection of missing periods and of periods that will soon be lost, retention purge.
- C# engine (`src\PurviewDlpReport.Engine.cs`, compiled automatically): normalization, storage, CSV and HTML writing.
- Report periods: Last24Hours, Last7Days, Last30Days, PreviousMonth, Month, Day, Custom (report time zone, summer/winter time handled).
- Business columns only: detection time, sender, recipients (optional: `IncludeRecipientDetails`), subject, recipient count, Message ID. One row per unique Message ID.
- Configurable splitting above `MaxRowsPerFile`: by week (default), day or rows; numbered parts for oversized days/weeks; file names show the period.
- New HTML report: one self-contained file per period, compressed data, virtual scrolling (150,000 rows open in ~2 s), filters, Top 10 senders, message details with recipient list, export of the filtered view, navigation between the files of a split report.
- Modern console output: numbered steps, per-window lines, progress bar with remaining time, final summary, daily log file.
- Interactive (administrator) and certificate (application) authentication; tenant and account verification.
- Lock preventing two simultaneous collections.
- Pester test suite (36 tests) and offline replay tool.
- Administrator guide (Markdown + HTML) with troubleshooting annex.

### Changed (compared with the prototypes)
- Local processing of 150,000 records: 10 min 30 s → 44 s; report writing: 3 min 11 s → 7 s; HTML: 15 files / 357 MB → 1 file / 17 MB.
- Pages of results are passed to the engine in a `PageData` object: passing large strings to .NET methods from PowerShell triggers an AMSI scan (~0.5 s per page).

### Known limitations
- Certificate (app-only) authentication with `Export-ActivityExplorerData` is not yet validated in the lab.
- `PageSize` above 1000 is allowed but untested.
