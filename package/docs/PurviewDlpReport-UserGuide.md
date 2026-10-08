---
title: Purview DLP Report
subtitle: User guide
version: 2.1.1
author: Nicolas Fabert
updated: 2026-10-08
---

# Purview DLP Report — User guide

> What is needed before the first report, then the commands used every day: **which messages matched the DLP rule?**, **who sent them, to whom and when?**, and **is the daily collection up to date?** The configuration keys, the application registration, the lab measurements and the internals are in the [developer guide](PurviewDlpReport-Guide.md).

> [!IMPORTANT]
> Files downloaded from the Internet may be blocked by Windows and fail to run. Before using this project, unblock every file in the downloaded folder:
>
> ```powershell
> Get-ChildItem "C:\Chemin\Du\Dossier" -Recurse -File -Force | Unblock-File
> ```
>
> Replace the example path with the folder where you downloaded or extracted this project.
>
> The `Install-Module` commands in this documentation use `-Force`, so they also update or reinstall a module that is already installed. If an older version still conflicts, close every PowerShell window, open a new one (as administrator for `-Scope AllUsers`), run `Uninstall-Module <ModuleName> -AllVersions -Force`, then run the `Install-Module` command again.

```cards
checklist | Prerequisites | PowerShell 7.4+, ExchangeOnlineManagement 3.9 or later, an account allowed to read Activity Explorer.
terminal | Everyday use | One command for a period, a daily collection, or the content of the database.
file | Results | CSV and self-contained HTML files written locally under `reports\`.
```

# Part I · Start here

<!-- icon: checklist -->
## 1. Prerequisites

| Item | Requirement |
|---|---|
| Operating system | Windows 10/11 or Windows Server 2016+, x64 or ARM64 |
| PowerShell | **PowerShell 7.4 or later** (`pwsh`), not Windows PowerShell 5.1. PowerShell 7.6 for ExchangeOnlineManagement 3.10 or later |
| Module | **ExchangeOnlineManagement 3.9.0 or later — not 3.10.0 in certificate mode** (bug fixed in 3.10.1; the tool skips it) |
| Console | **Windows Terminal** — emoji and colours |
| SQLite | Supplied in `lib\sqlite` — nothing to install |
| Browser | Recent Edge, Chrome or Firefox, to open the HTML report |
| Network | HTTPS to `*.protection.outlook.com`, `*.compliance.protection.outlook.com`, `login.microsoftonline.com` |
| Disk | About **1.5 to 2.5 GB per month** of database at 30–50k messages/day — **9 to 15 GB** with the default retention of 180 days — plus the reports |

Install or update the Exchange Online module:

```powershell
# Install or update the Exchange Online module: 3.10.1 or later with PowerShell 7.6,
# 3.9.2 with PowerShell 7.4 or 7.5
Install-Module ExchangeOnlineManagement -MinimumVersion 3.10.1 -Scope CurrentUser -Force
Install-Module ExchangeOnlineManagement -RequiredVersion 3.9.2 -Scope CurrentUser -Force
```

```cards
people | Administrator account (interactive) | One of the role groups that can read Activity Explorer: **Compliance Administrator**, **Security Administrator**, **Security Reader**, or the Purview *Information Protection* role groups.
key | Application (unattended) | A dedicated app registration with a certificate, the **Exchange.ManageAsApp** permission, and the Purview role **Information Protection Reader** through a dedicated role group — [developer guide, chapter 7](PurviewDlpReport-Guide.md#7-unattended-execution) and Annex F.
```

> [!IMPORTANT]
> The tool is **read-only** for Microsoft 365: it never sends e-mail and never changes a setting. Reports stay on the local disk; the administrator sends or shares them.

<!-- icon: download -->
## 2. One-time setup

```steps
Copy the package | Copy the package folder (`PurviewDlpReport-<version>`) to the server, for example `D:\Tools\PurviewDlpReport`.
Unblock the files | `Get-ChildItem D:\Tools\PurviewDlpReport -Recurse -File -Force | Unblock-File` (files copied from the Internet or a share).
Edit the configuration | `config\PurviewDlpReport.config.psd1` — the *Tenant*, *Target* and *Authentication* values are empty in the package: fill them in.
Check without connecting | `.\Invoke-PurviewDlpReport.ps1 -Mode Status` — checks the configuration and builds the engine (about 2 s, once). On a new installation it ends with *Ready for the first collection*.
Run the first report | `.\Invoke-PurviewDlpReport.ps1` — collects what is missing for the last 7 days, then writes the CSV and HTML files.
```

The values to fill in are:

| Setting | Meaning |
|---|---|
| `Tenant.TenantId` | Microsoft Entra tenant GUID — a safety check: the tool stops if it is connected to another tenant |
| `Tenant.Organization` | Initial domain `xxx.onmicrosoft.com` — **certificate mode only** |
| `Target.PolicyId` | DLP policy identifier, as shown in an Activity Explorer event |
| `Target.RuleId` | DLP rule identifier as shown in Activity Explorer = the rule **ImmutableId** |
| `Target.PolicyName`, `Target.RuleName` | Optional, display only |
| `Authentication.Mode` | `Interactive` (an administrator signs in) or `Certificate` (scheduled task) |
| `Authentication.UserPrincipalName` | Interactive: expected account (`''` accepts any account of the tenant) |
| `Authentication.AppId`, `Authentication.CertificateThumbprint` | Certificate mode: application (client) ID and certificate thumbprint |

Find the two identifiers of the rule:

```powershell
# Find the identifiers (Security & Compliance PowerShell)
Connect-IPPSSession
Get-DlpCompliancePolicy -Identity '<policy name>' | Select-Object Name, ImmutableId, Guid
Get-DlpComplianceRule   -Identity '<rule name>'   | Select-Object Name, ImmutableId, Guid
```

> [!CAUTION]
> The rule **Guid** is *not* the value used by Activity Explorer — use **ImmutableId**. When in doubt, open one event of the rule in Activity Explorer and copy `PolicyId` / `RuleId`.

> [!TIP]
> When a value is wrong, the tool lists **all** the problems at once and stops before doing anything. Relative paths (`.\data`) are relative to the tool folder. Every configuration key is described in the [developer guide, chapter 6](PurviewDlpReport-Guide.md#6-configuration).

# Part II · Everyday use

<!-- icon: terminal -->
## 3. Build a report

Open **PowerShell 7** (`pwsh`) in **Windows Terminal**, go to the tool folder, run one command:

| I want… | Command |
|---|---|
| The last 7 days (default) | `.\Invoke-PurviewDlpReport.ps1` |
| The last 24 hours | `.\Invoke-PurviewDlpReport.ps1 -Range Last24Hours` |
| The previous calendar month | `.\Invoke-PurviewDlpReport.ps1 -Range PreviousMonth` |
| A given month | `.\Invoke-PurviewDlpReport.ps1 -Range Month -Month 2026-08` |
| One day | `.\Invoke-PurviewDlpReport.ps1 -Range Day -Date 2026-09-28` |
| A custom range | `.\Invoke-PurviewDlpReport.ps1 -Range Custom -Start '2026-09-28 08:00' -End '2026-09-29'` |
| Counts only, smaller files | `.\Invoke-PurviewDlpReport.ps1 -Range PreviousMonth -IncludeRecipientDetails:$false` |
| A report without connecting | `.\Invoke-PurviewDlpReport.ps1 -NoCollect` |
| The daily collection | `.\Invoke-PurviewDlpReport.ps1 -Mode Collect` |
| What the database contains | `.\Invoke-PurviewDlpReport.ps1 -Mode Status` |
| The full help | `Get-Help .\Invoke-PurviewDlpReport.ps1 -Full` |

Unless `-NoCollect` is used, Report mode first collects the missing part of the period, then writes the files. The default period is `Report.DefaultRange` in the configuration (`Last7Days`).

### What you see

![A report: title card, numbered steps, one row per collected window, files written, summary card](images/console-report.png)

```cards
info | Banner | Mode, period, rule, database and log file.
checklist | Numbered steps | 1 database · 2 plan · 3 connection · 4 collection · 5 report.
download | Collection table | One row per window: events, new events, pages, duration, rate — with a progress bar and the remaining time.
target | Summary card | Green when complete, amber when part of the period is missing, red on error.
```

Icons: ✅ done · 🔹 information · ⚠️ attention · ❌ error · ⏩ step skipped.

> [!NOTE]
> **Activity Explorer is not real time.** Microsoft states 60 to 90 minutes for Exchange events. The tool collects the **last 6 hours again** at every run, so the most recent hours of a report are provisional.

<!-- icon: calendar -->
## 4. Schedule the daily collection

Activity Explorer keeps **30 days**: a day that was not collected within 30 days is lost for this tool. The daily collection is therefore mandatory in production, and uses an **application with a certificate** (no one is there to sign in):

```powershell
$action = New-ScheduledTaskAction -Execute 'pwsh.exe' `
    -Argument '-NoProfile -NonInteractive -File "D:\Tools\PurviewDlpReport\Invoke-PurviewDlpReport.ps1" -Mode Collect' `
    -WorkingDirectory 'D:\Tools\PurviewDlpReport'
$trigger  = New-ScheduledTaskTrigger -Daily -At 06:00
$settings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Hours 6) -StartWhenAvailable -MultipleInstances IgnoreNew
Register-ScheduledTask -TaskName 'Purview DLP Report - daily collection' -Action $action -Trigger $trigger `
    -Settings $settings -User 'DOMAIN\svc-dlpreport' -Password '<password>' -RunLevel Limited
```

```cards
key | Certificate store | The certificate must be in **CurrentUser\My** of the account that runs the task (or LocalMachine\My with read access to the private key).
clock | First run | Collects `BackfillDays` (30) days: about **1 hour per 350,000 events** — several hours in production. Next runs: minutes.
shield | Safety | Two collections never run at the same time (lock file). Running at 06:00 **and** 18:00 limits the risk of losing a day.
```

The registration of the application, its permissions and the Purview role group are described step by step in the [developer guide, chapter 7](PurviewDlpReport-Guide.md#7-unattended-execution) and Annex F.

<!-- icon: chart -->
## 5. Read the results

Each report creates a folder `reports\<date>_<range>\` with the files:

| File | Content |
|---|---|
| `.html` | Self-contained report: summary tiles, distribution by recipient count, Top 10 senders, filters, message details and **Export view to CSV** |
| `.csv` | Complete export, UTF-8 with BOM and `;` separator — double-click opens it in Excel |

![HTML report: summary tiles, distribution, Top 10 senders, filters and messages](images/report-overview.png)

Click a row to see every recipient, with a **Copy recipients** button. The report holds one row per unique Message ID:

| Column | Meaning |
|---|---|
| Detection time | When Purview matched the rule, in the report time zone (`Report.TimeZone`), to the second |
| Sender | Sender address |
| Recipients | All recipient addresses — optional (`IncludeRecipientDetails`) |
| Subject | Message subject |
| Recipient count | Number of recipients reported by the DLP rule |
| Message ID | Internet Message ID, the unique key of a message |

**To, Cc and Bcc are not separated** and Bcc is included: protect the CSV and HTML files as sensitive data. A rule match is not a sent mail, and the detection time is close to, but not, the send time.

> [!TIP]
> Files are written in `<folder>.pending`, renamed only when everything is complete: **a folder without `.pending` is always a complete report**. Up to `MaxRowsPerFile` messages (500,000 by default) each report is one CSV and one HTML file; above that, files are split by local **week**, by **day** or every *MaxRowsPerFile* **rows**.

> [!TIP]
> **To send by e-mail**: use `-IncludeRecipientDetails:$false` (CSV about 6× smaller), or zip the CSV (about 10× smaller).

<!-- icon: info -->
## 6. Status, logs and exit codes

Status shows what the database holds, day by day, without connecting:

```powershell
.\Invoke-PurviewDlpReport.ps1 -Mode Status
```

![Status: one line per day with a coverage bar](images/console-status.png)

Each day shows its messages, a **coverage bar** and a state: *Complete*, *Complete, refreshed at next run*, *Missing, collectable N more days*, or *no longer collectable*.

Everything shown on screen is also written to `logs\PurviewDlpReport_yyyyMMdd.log`, without colours or icons; logs are kept 30 days.

| Exit code | Meaning |
|---|---|
| `0` | Success |
| `1` | Failure — message on screen and in the log. What was already collected is kept. |
| `2` | Report written but **incomplete**: the summary lists the missing ranges, the HTML shows the coverage. |

> [!TIP]
> **Interrupted?** Close the window or press <kbd>Ctrl</kbd>+<kbd>C</kbd> at any time: every page received is already in the database. Run the **same command again** — only the missing windows are collected, without duplicates.

# Part III · Troubleshoot

<!-- icon: lifebuoy -->
## 7. Common situations

| Symptom | Cause | What to do |
|---|---|---|
| `Invalid configuration (...)` and a list | Wrong values in the .psd1 | Fix every line listed. Nothing was done. |
| `Export-ActivityExplorerData is not available for …` | No Purview role allowing Activity Explorer | Interactive: assign a role (chapter 1), wait ~30 min, sign in again. Application: role group with *Information Protection Reader*. |
| `Connected to tenant …` / `Signed in as …` | Wrong account in the sign-in window | Use the account of `Authentication.UserPrincipalName`. |
| No sign-in window | Session without desktop, blocked pop-up | Run on a desktop session; use certificate mode for tasks. |
| `UnAuthorized` in certificate mode | The token has no `Exchange.ManageAsApp` role | Grant `Exchange.ManageAsApp` on **Office 365 Exchange Online** (and Exchange Online Protection) with admin consent. |
| *"Object reference not set to an instance of an object"* at the connection, certificate mode | ExchangeOnlineManagement **3.10.0** | `Install-Module ExchangeOnlineManagement -MinimumVersion 3.10.1 -Scope CurrentUser -Force`, then open a new PowerShell window. |
| *"… is already loaded in this PowerShell session and cannot be used"* | A version that must not be used was loaded earlier in the same window | Open a new PowerShell window. |
| `Another execution is already collecting` | The scheduled task is running | Wait, or use `-NoCollect`. |
| `Missing … collectable for N more day(s)` | The daily collection did not run | Run `-Mode Collect` before the deadline; check the task. |
| "Report written, but INCOMPLETE", exit code 2 | Part of the period is not in the database | Run again. Periods older than 30 days are lost. |
| A folder ends with `.pending` | The execution stopped while writing | Delete it, run again. |
| Times shifted by 1–2 h | Report time zone vs UTC | Times are in `Report.TimeZone`. |
| "This browser cannot open the compressed data" | Old browser | Use Edge/Chrome/Firefox, or the CSV. |
| Squares instead of icons in the console | The console is not Windows Terminal | Run the tool in Windows Terminal. |

For the full troubleshooting list, the known issue with ExchangeOnlineManagement 3.10.0, the configuration reference, the least-privilege application and the internals, continue with the [developer guide](PurviewDlpReport-Guide.md).
