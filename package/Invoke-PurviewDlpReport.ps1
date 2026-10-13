#Requires -Version 7.4

<#
.SYNOPSIS
    Purview DLP Report - Exchange Online messages matched by a Microsoft Purview DLP rule
    (by default: messages sent to more than 25 recipients). One row per Message ID.

.DESCRIPTION
    The tool works in two stages:

      1. COLLECT  Activity Explorer (Export-ActivityExplorerData) -> local SQLite database.
                  Activity Explorer keeps only 30 days: collect at least once a day
                  (scheduled task, -Mode Collect). The database keeps the history.
      2. REPORT   SQLite database -> CSV + HTML files in a local folder.
                  Before writing, the missing part of the period (if any) is collected.

    Everything is set in config\PurviewDlpReport.config.psd1; the command line only
    chooses what to do. The tool is read-only for the tenant: it never sends e-mail and
    never changes any Microsoft 365 setting.

.PARAMETER Mode
    Report  (default) Collects what is missing for the period, then writes the report.
    Collect           Collects the last days into the database (scheduled task). No report.
    Status            Shows what the database contains, day by day. No connection.

.PARAMETER Range
    Period of the report (Report mode). Default: Report.DefaultRange in the configuration.
      Last24Hours, Last7Days, Last30Days : rolling windows ending now
      PreviousMonth                      : the previous calendar month
      Month  -Month 2026-08              : a calendar month
      Day    -Date 2026-09-28            : a calendar day
      Custom -Start '2026-09-01 08:00' -End '2026-09-03'   (report time zone unless an offset is given)

.PARAMETER ConfigPath
    Configuration file. Default: config\PurviewDlpReport.config.psd1 next to this script.

.PARAMETER IncludeRecipientDetails
    Overrides Report.IncludeRecipientDetails for this execution.
    -IncludeRecipientDetails        : recipient addresses in the files
    -IncludeRecipientDetails:$false : recipient count only (smaller files)

.PARAMETER SplitBy
    Overrides Report.SplitBy (Rows, Day or Week) for this execution.

.PARAMETER MaxRowsPerFile
    Overrides Report.MaxRowsPerFile for this execution.

.PARAMETER OutputPath
    Overrides Report.OutputPath for this execution.

.PARAMETER NoCollect
    Report mode: do not connect to Microsoft 365; use only the data already in the database.

.EXAMPLE
    .\Invoke-PurviewDlpReport.ps1
    Report of the default period (last 7 days), collecting what is missing first.

.EXAMPLE
    .\Invoke-PurviewDlpReport.ps1 -Range PreviousMonth -IncludeRecipientDetails:$false
    Report of the previous month with the recipient count only.

.EXAMPLE
    .\Invoke-PurviewDlpReport.ps1 -Mode Collect
    Daily collection (scheduled task).

.EXAMPLE
    .\Invoke-PurviewDlpReport.ps1 -Mode Status
    What is in the database, day by day, and what is missing.

.NOTES
    Author  : Nicolas Fabert
    Version : 2.1.2
    Exit codes : 0 = success, 1 = failure, 2 = finished but incomplete (see the summary).
    Documentation : docs\PurviewDlpReport-Guide.md (or .html)

    This Script is a Personal project.
    It's provided "AS-IS". It's not an official Microsoft product so no support can be expected from Microsoft.

    As any scripts you must read carefully the documentation and test it first in a Test environment before any run in Production.
#>
[CmdletBinding()]
param(
    [ValidateSet('Report', 'Collect', 'Status')]
    [string]$Mode = 'Report',

    [ValidateSet('Last24Hours', 'Last7Days', 'Last30Days', 'PreviousMonth', 'Month', 'Day', 'Custom')]
    [string]$Range,
    [string]$Month,
    [string]$Date,
    [string]$Start,
    [string]$End,

    [string]$ConfigPath = (Join-Path $PSScriptRoot 'config\PurviewDlpReport.config.psd1'),
    [switch]$IncludeRecipientDetails,
    [ValidateSet('Rows', 'Day', 'Week')]
    [string]$SplitBy,
    [ValidateRange(1000, 1048575)]
    [int]$MaxRowsPerFile,
    [string]$OutputPath,
    [switch]$NoCollect
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
# Numbers and dates are displayed the same way on every server (1,234.5), whatever the regional settings.
$previousCulture = [Threading.Thread]::CurrentThread.CurrentCulture
[Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]::GetCultureInfo('en-US')
$clock = [Diagnostics.Stopwatch]::StartNew()
$exitCode = 1
$runStatus = 'Failed'
$store = $null; $connection = $null; $lock = $null; $runId = $null; $details = [ordered]@{}

try {
    # do { } while ($false): 'break' ends the execution early; the finally block always runs.
    do {
        $integrityVerifier = Join-Path $PSScriptRoot 'Test-PackageIntegrity.ps1'
        $integrityCursor = [IO.Path]::GetFullPath($integrityVerifier)
        while ($integrityCursor) {
            $integrityItem = Microsoft.PowerShell.Management\Get-Item -LiteralPath $integrityCursor -Force -ErrorAction Stop
            if ($integrityItem.Attributes -band [IO.FileAttributes]::ReparsePoint) {
                throw "Package integrity: reparse point: $integrityCursor"
            }
            $integrityCursor = [IO.Path]::GetDirectoryName($integrityCursor)
        }
        $integritySignature = Microsoft.PowerShell.Security\Get-AuthenticodeSignature -LiteralPath $integrityVerifier -ErrorAction Stop
        if ([string]$integritySignature.Status -cne 'Valid' -or $null -eq $integritySignature.SignerCertificate -or
            $integritySignature.SignerCertificate.Thumbprint -cne '0F28E442889F17B5F4D251E8F96E7CC6E2C9C6D6' -or
            $null -eq $integritySignature.TimeStamperCertificate) {
            throw 'Package integrity: verifier must have the trusted, timestamped owner signature.'
        }
        & $integrityVerifier -Root $PSScriptRoot
        Import-Module (Join-Path $PSScriptRoot 'PurviewDlpReport.psd1') -Force

        # ---------------------------------------------------------------------------------
        # Configuration file, then the command-line overrides.
        # ---------------------------------------------------------------------------------
        $settings = Import-DlpConfiguration -Path $ConfigPath -Root $PSScriptRoot
        if ($PSBoundParameters.ContainsKey('IncludeRecipientDetails')) { $settings.Report.IncludeRecipientDetails = [bool]$IncludeRecipientDetails }
        if ($SplitBy) { $settings.Report.SplitBy = $SplitBy }
        if ($MaxRowsPerFile) { $settings.Report.MaxRowsPerFile = $MaxRowsPerFile }
        if ($OutputPath) { $settings.Report.OutputPath = [IO.Path]::GetFullPath($OutputPath, (Get-Location).Path) }
        if ($Mode -eq 'Report' -and -not $Range) { $Range = $settings.Report.DefaultRange }
        $zone = $settings.Zone

        $logPath = Start-DlpLog -Directory $settings.Logging.Path -RetentionDays $settings.Logging.RetentionDays
        Initialize-DlpEngine -Root $PSScriptRoot

        $collecting = $Mode -eq 'Collect' -or ($Mode -eq 'Report' -and -not $NoCollect)
        $totalSteps = switch ($Mode) { 'Status' { 2 } 'Collect' { 4 } default { 5 } }
        $period = $null
        if ($Mode -eq 'Report') {
            $period = Resolve-DlpPeriod -Range $Range -Month $Month -Date $Date -Start $Start -End $End -Zone $zone
        }
        $dot = [char]0x00B7
        $modeText = $Mode + $(if ($period) { " $dot $Range" } elseif ($Mode -eq 'Collect') { " $dot last $($settings.Collection.BackfillDays) days" } else { '' })
        $banner = [ordered]@{ Mode = @('Info', $modeText) }
        if ($period) { $banner['Period'] = @('Calendar', "$(Format-DlpRange $period.StartMs $period.EndMs $zone)  ($($settings.Report.TimeZone), end excluded)") }
        $ruleText = if ($settings.Target.RuleName) { $settings.Target.RuleName + $(if ($settings.Target.PolicyName) { "  ($($settings.Target.PolicyName))" }) } else { "rule $($settings.Target.RuleId)" }
        $banner['Rule'] = @('Target', $ruleText)
        $banner['Database'] = @('Database', $settings.Storage.DatabasePath)
        $banner['Log'] = @('Log', $logPath)
        Write-DlpBanner -Title 'Purview DLP Report' -Subtitle "Exchange Online DLP matches $dot one row per Message ID" -Details $banner

        # ---------------------------------------------------------------------------------
        # Step 1 - Database (and lock when this execution collects).
        # ---------------------------------------------------------------------------------
        Write-DlpStep 1 $totalSteps 'Opening the database' -Icon Database
        if ($Mode -eq 'Status' -and -not (Test-Path -LiteralPath $settings.Storage.DatabasePath)) {
            # New installation: nothing collected yet. The configuration is valid and the engine is built.
            Write-DlpItem Info 'No database yet: the first collection (-Mode Collect) or the first report creates it.'
            Write-DlpSummary -Title 'Ready for the first collection' -Values ([ordered]@{
                Config        = @('Ok', 'valid')
                Engine        = @('Ok', 'ready')
                Database      = @('Database', 'none yet')
                Next          = @('Info', '.\Invoke-PurviewDlpReport.ps1 -Mode Collect')
            }) -Status Ok
            $exitCode = 0
            $runStatus = 'Completed'
            break
        }
        if ($Mode -eq 'Status') {
            $store = Open-DlpStore -Settings $settings -ReadOnly
        } else {
            if ($collecting) { $lock = Enter-DlpLock -Path ($settings.Storage.DatabasePath + '.lock') }
            $store = Open-DlpStore -Settings $settings
            if ($collecting) {
                $closed = $store.CloseAbandonedWork()
                if ($closed) { Write-DlpItem Warn "$closed collection window(s) left unfinished by an interrupted execution will be collected again." }
            }
            $runId = $store.StartRun($Mode, $null, [Environment]::MachineName, (Get-Module PurviewDlpReport).Version.ToString(), $null)
        }
        $targetId = if ($Mode -eq 'Status') { $store.FindTarget($settings.Target.PolicyId, $settings.Target.RuleId) }
                    else { $store.GetOrCreateTarget($settings.Target.PolicyId, $settings.Target.RuleId, $settings.Target.PolicyName, $settings.Target.RuleName) }
        $stats = $store.GetStatistics($targetId)
        Write-DlpItem Ok ("{0} events {3} {1} messages {3} {2}" -f (Format-DlpNumber $stats.Events), (Format-DlpNumber $stats.Messages), (Format-DlpBytes $stats.FileBytes), [char]0x00B7)

        # ---------------------------------------------------------------------------------
        # Status mode: day-by-day view, then stop.
        # ---------------------------------------------------------------------------------
        if ($Mode -eq 'Status') {
            Write-DlpStep 2 $totalSteps 'Collected data, day by day' -Icon Calendar
            Show-DlpStatus -Store $store -TargetId $targetId -Settings $settings
            $expiring = @(if ($targetId) { Get-DlpExpiringGaps -Store $store -TargetId $targetId -Settings $settings })
            foreach ($gap in $expiring) { Write-DlpItem Warn ("Missing {0}: Activity Explorer keeps it for about {1} more day(s). Run -Mode Collect." -f (Format-DlpRange $gap.Start $gap.End $zone), $gap.DaysLeft) }
            $exitCode = 0
            $runStatus = 'Completed'
            break
        }

        # ---------------------------------------------------------------------------------
        # Step 2 - What must be collected?
        # ---------------------------------------------------------------------------------
        $nowMs = [DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds()
        if ($Mode -eq 'Collect') {
            Write-DlpStep 2 $totalSteps 'Planning the collection' -Icon Plan
            $windowStart = [Math]::Max($nowMs - [long]$settings.Collection.BackfillDays * 86400000, $nowMs - [long]$settings.Collection.SourceRetentionDays * 86400000 + 3600000)
            $plan = Get-DlpCollectionPlan -Store $store -TargetId $targetId -StartMs $windowStart -EndMs $nowMs -Settings $settings -NowMs $nowMs
        } else {
            Write-DlpStep 2 $totalSteps 'Checking the data available for the period' -Icon Plan
            $plan = Get-DlpCollectionPlan -Store $store -TargetId $targetId -StartMs $period.StartMs -EndMs $period.EndMs -Settings $settings -NowMs $nowMs
        }
        $slices = $plan.Slices
        if ($slices.Count) {
            $refresh = if ($plan.RefreshMs -gt 0) { " (of which {0} already collected, refreshed to catch late events)" -f (Format-DlpDuration ($plan.RefreshMs / 1000)) } else { '' }
            Write-DlpItem Info ("To collect: {0} window(s), {1} in total{2}" -f $slices.Count, (Format-DlpDuration ($plan.CollectableMs / 1000)), $refresh)
        } else {
            Write-DlpItem Ok 'Everything is already in the database.'
        }
        foreach ($gap in $plan.Unrecoverable) {
            Write-DlpItem Warn ("Not in the database and older than the {0} days kept by Activity Explorer: {1}" -f $settings.Collection.SourceRetentionDays, (Format-DlpRange $gap.Start $gap.End $zone))
        }
        if ($settings.Collection.SettlingHours -eq 0) {
            Write-DlpItem Warn 'SettlingHours is 0: late Activity Explorer events will not be recovered by a later collection.'
        }
        if ($Mode -eq 'Report' -and $NoCollect -and $plan.NewMs -gt 0) {
            Write-DlpItem Warn '-NoCollect: the missing part will not be collected; the report will be incomplete.'
        } elseif ($Mode -eq 'Report' -and $NoCollect -and $plan.RefreshMs -gt 0) {
            Write-DlpItem Info '-NoCollect: recent data is not refreshed; events that reached Activity Explorer late may be missing.'
        }

        # ---------------------------------------------------------------------------------
        # Steps 3 and 4 - Connection and collection (only when something is missing).
        # ---------------------------------------------------------------------------------
        $collection = $null
        if ($collecting -and $slices.Count) {
            Write-DlpStep 3 $totalSteps 'Connecting to Security & Compliance PowerShell' -Icon Key
            if ($settings.Authentication.Mode -eq 'Interactive') { Write-DlpItem Info 'A sign-in window may open: sign in with the administrator account.' -Icon People }
            $connectClock = [Diagnostics.Stopwatch]::StartNew()
            $connection = Connect-DlpActivityExplorer -Settings $settings
            $legacyBinding = $store.BindAuthenticatedTenant($connection.TenantId)
            if ($legacyBinding) { Write-DlpItem Warn 'This legacy database had no tenant binding. Future collections are now restricted to the verified tenant; the tenant of historical rows cannot be independently verified. Use a separate database for each tenant.' }
            $store.UpdateRunAccount($runId, $connection.Account)
            Write-DlpItem Ok ("{0}  {3} {1} mode {3} tenant verified {3} {2}" -f $connection.Account, $connection.Mode.ToLowerInvariant(), (Format-DlpDuration $connectClock.Elapsed.TotalSeconds), [char]0x00B7)

            Write-DlpStep 4 $totalSteps 'Collecting Activity Explorer data' -Icon Download
            $collection = Invoke-DlpCollection -Store $store -RunId $runId -TargetId $targetId -Settings $settings -Slices $slices -Exporter $connection.Exporter
            $details['Collection'] = $collection
            if ($collection.Quarantined) { Write-DlpItem Warn ("{0} record(s) could not be read and were set aside (table 'quarantine'): {1}" -f $collection.Quarantined, (($collection.QuarantineReasons.GetEnumerator() | ForEach-Object { "$($_.Key) x$($_.Value)" }) -join ', ')) }
            if ($collection.Conflicts) { Write-DlpItem Info ("{0} event(s) had a sender, subject or recipient list different from an earlier event of the same message; the first values are kept." -f $collection.Conflicts) }
            Write-DlpCollectionDiagnostics -Collection $collection -Zone $zone
            if ($collection.Completed) {
                Write-DlpItem Ok ("Collection complete: {0} events received, {1} new, in {2} (Activity Explorer: {3})" -f (Format-DlpNumber $collection.Records), (Format-DlpNumber $collection.NewEvents), (Format-DlpDuration $collection.Seconds), (Format-DlpDuration $collection.SourceSeconds))
            } else {
                Write-DlpItem Fail ("Collection stopped: {0}" -f $collection.Error)
                Write-DlpItem Info 'What was received is kept. Run the same command again to continue from the missing windows.'
            }
        } elseif ($collecting) {
            Write-DlpStep 3 $totalSteps 'Connecting to Security & Compliance PowerShell' -Icon Key
            Write-DlpItem Skip 'Not needed.'
            Write-DlpStep 4 $totalSteps 'Collecting Activity Explorer data' -Icon Download
            Write-DlpItem Skip 'Nothing to collect.'
        } else {
            Write-DlpStep 3 $totalSteps 'Connecting to Security & Compliance PowerShell' -Icon Key
            Write-DlpItem Skip '-NoCollect: no connection.'
            Write-DlpStep 4 $totalSteps 'Collecting Activity Explorer data' -Icon Download
            Write-DlpItem Skip '-NoCollect: the report uses the data already collected.'
        }
        if ($collecting) {
            $purge = Invoke-DlpRetention -Store $store -Settings $settings
            if ($purge -and $purge.Events) { Write-DlpItem Info ("Retention ({0} days): {1} old events deleted." -f $settings.Storage.RetentionDays, (Format-DlpNumber $purge.Events)) }
        }
        $collectionFailed = $collection -and -not $collection.Completed

        # ---------------------------------------------------------------------------------
        # Collect mode ends here.
        # ---------------------------------------------------------------------------------
        if ($Mode -eq 'Collect') {
            $expiring = @(Get-DlpExpiringGaps -Store $store -TargetId $targetId -Settings $settings)
            foreach ($gap in $expiring) { Write-DlpItem Warn ("Still missing {0}: collectable for about {1} more day(s)." -f (Format-DlpRange $gap.Start $gap.End $zone), $gap.DaysLeft) }
            $stats = $store.GetStatistics($targetId)
            $summary = [ordered]@{
                'Result'   = @($(if ($collectionFailed) { 'Fail' } else { 'Ok' }), $(if ($collectionFailed) { 'Incomplete - run again to continue' } else { 'Collection complete' }))
                'Events'   = @('Download', $(if ($collection) { '{0} received, {1} new' -f (Format-DlpNumber $collection.Records), (Format-DlpNumber $collection.NewEvents) } else { 'nothing to collect' }))
                'Database' = @('Database', ('{0} events {3} {1} messages {3} {2}' -f (Format-DlpNumber $stats.Events), (Format-DlpNumber $stats.Messages), (Format-DlpBytes $stats.FileBytes), [char]0x00B7))
                'Duration' = @('Clock', (Format-DlpDuration $clock.Elapsed.TotalSeconds))
                'Log'      = @('Log', $logPath)
            }
            $status = if ($collectionFailed) { 'Fail' } elseif ($expiring.Count) { 'Warn' } else { 'Ok' }
            Write-DlpSummary -Title $(if ($collectionFailed) { 'Collection incomplete' } else { 'Collection finished' }) -Values $summary -Status $status
            $exitCode = if ($collectionFailed) { 1 } else { 0 }
            $runStatus = if ($collectionFailed) { 'Incomplete' } else { 'Completed' }
            break
        }

        # ---------------------------------------------------------------------------------
        # Step 5 - Report files.
        # ---------------------------------------------------------------------------------
        Write-DlpStep 5 $totalSteps 'Writing the report' -Icon Report
        $coverage = Get-DlpCoverage -Store $store -TargetId $targetId -StartMs $period.StartMs -EndMs $period.EndMs -SettlingMs ([long]$settings.Collection.SettlingHours * 3600000)
        $reportClock = [Diagnostics.Stopwatch]::StartNew()
        $report = New-DlpReport -Store $store -TargetId $targetId -Settings $settings -Period $period -Coverage $coverage -CollectionFailed:([bool]$collectionFailed)
        foreach ($file in $report.Result.Files) {
            Write-DlpItem Ok ("{0,-4} {1}   {2} messages {4} {3}" -f $file.Kind, (Split-Path $file.Path -Leaf), (Format-DlpNumber $file.Rows), (Format-DlpBytes $file.Bytes), [char]0x00B7) -Icon File
        }
        Write-DlpItem Info ("Written in {0}" -f (Format-DlpDuration $reportClock.Elapsed.TotalSeconds)) -Icon Clock
        $r = $report.Result
        if ($r.MessagesWithoutCount) { Write-DlpItem Info ("{0} message(s) without a reliable recipient count (column left empty)." -f (Format-DlpNumber $r.MessagesWithoutCount)) }
        $complete = $coverage.MissingMs -eq 0 -and -not $collectionFailed
        $summary = [ordered]@{
            'Period'   = @('Calendar', "$(Format-DlpRange $period.StartMs $period.EndMs $zone)  ($($settings.Report.TimeZone))")
            'Messages' = @('Mail', ('{0} unique messages {2} {1} senders' -f (Format-DlpNumber $r.Messages), (Format-DlpNumber $r.Senders), [char]0x00B7))
            'Coverage' = @($(if ($coverage.MissingMs -eq 0) { 'Chart' } else { 'Warn' }), $(if ($coverage.MissingMs -eq 0) { '{0}% of queried time{1}' -f $coverage.Percent, $(if ($coverage.ProvisionalMs -gt 0) { "; settled: $($coverage.FinalPercent)% (recent data is provisional)" } else { '' }) } else { '{0}% of queried time (rounded) - missing: {1} ({2} ms)' -f $coverage.Percent, ((@($coverage.Gaps) | Select-Object -First 3 | ForEach-Object { Format-DlpRange $_.Start $_.End $zone }) -join ', '), $coverage.MissingMs }))
            'Files'    = @('File', ('{0} file(s) {2} recipients {1}' -f $r.Files.Count, $(if ($settings.Report.IncludeRecipientDetails) { 'listed' } else { 'counted only' }), [char]0x00B7))
            'Folder'   = @('Folder', $report.Directory)
            'Duration' = @('Clock', (Format-DlpDuration $clock.Elapsed.TotalSeconds))
            'Log'      = @('Log', $logPath)
        }
        $summary['Metadata'] = @('File', 'report-metadata.json (coverage status; CSV columns unchanged)')
        Write-DlpSummary -Title $(if ($complete) { 'Report ready' } else { 'Report written, but INCOMPLETE' }) -Values $summary -Status $(if ($complete) { 'Ok' } else { 'Warn' })
        $details['Report'] = [ordered]@{ Directory = $report.Directory; Messages = $r.Messages; Files = @($r.Files | ForEach-Object { $_.Path }); CoverageComplete = $complete; MissingMs = $coverage.MissingMs; ProvisionalMs = $coverage.ProvisionalMs; CoveragePercent = $coverage.Percent; SettledCoveragePercent = $coverage.FinalPercent }
        $exitCode = if ($complete) { 0 } else { 2 }
        $runStatus = if ($complete) { 'Completed' } else { 'Incomplete' }
    } while ($false)
}
catch {
    $message = $_.Exception.Message
    Write-Host ''
    if (Get-Command Write-DlpSummary -ErrorAction SilentlyContinue) {
        Write-DlpSummary -Title 'Execution stopped' -Values ([ordered]@{ Error = @('Fail', $message); Duration = @('Clock', (Format-DlpDuration $clock.Elapsed.TotalSeconds)) }) -Status Fail
    } else {
        Write-Host "  [ERROR] $message"
    }
    if (Get-Command Write-DlpLog -ErrorAction SilentlyContinue) { Write-DlpLog -Level 'ERROR' -Message ($message + "`n" + $_.ScriptStackTrace) }
    $details['Error'] = $message
    $exitCode = 1
}
finally {
    if ($store -and $runId) {
        try { $store.FinishRun($runId, $runStatus, ($details | ConvertTo-Json -Depth 6 -Compress)) } catch { }
    }
    if ($store) { $store.Dispose() }
    if (Get-Command Disconnect-DlpActivityExplorer -ErrorAction SilentlyContinue) { Disconnect-DlpActivityExplorer -Connection $connection }
    if (Get-Command Exit-DlpLock -ErrorAction SilentlyContinue) { Exit-DlpLock -Lock $lock }
    if (Get-Command Write-DlpLog -ErrorAction SilentlyContinue) {
        Write-DlpLog -Level 'INFO' -Message ("Exit code {0} after {1:0.0} s" -f $exitCode, $clock.Elapsed.TotalSeconds)
        Stop-DlpLog
    }
    [Threading.Thread]::CurrentThread.CurrentCulture = $previousCulture
}
exit $exitCode

# SIG # Begin signature block
# MIIghQYJKoZIhvcNAQcCoIIgdjCCIHICAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBVh/RF+CK94BpR
# eafSl8mNxEG5qnIwQS+Hr5wNevSshaCCDW8wggZyMIIEWqADAgECAghkM1HTxzif
# CDANBgkqhkiG9w0BAQsFADB8MQswCQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMx
# EDAOBgNVBAcMB0hvdXN0b24xGDAWBgNVBAoMD1NTTCBDb3Jwb3JhdGlvbjExMC8G
# A1UEAwwoU1NMLmNvbSBSb290IENlcnRpZmljYXRpb24gQXV0aG9yaXR5IFJTQTAe
# Fw0xNjA2MjQyMDQ0MzBaFw0zMTA2MjQyMDQ0MzBaMHgxCzAJBgNVBAYTAlVTMQ4w
# DAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91c3RvbjERMA8GA1UECgwIU1NMIENv
# cnAxNDAyBgNVBAMMK1NTTC5jb20gQ29kZSBTaWduaW5nIEludGVybWVkaWF0ZSBD
# QSBSU0EgUjEwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQCfgxNzqrDG
# bSHL24t6h3TQcdyOl3Ka5LuINLTdgAPGL0WkdJq/Hg9Q6p5tePOf+lEmqT2d0bKU
# Vz77OYkbkStW72fL5gvjDjmMxjX0jD3dJekBrBdCfVgWQNz51ShEHZVkMGE6ZPKX
# 13NMfXsjAm3zdetVPW+qLcSvvnSsXf5qtvzqXHnpD0OctVIFD+8+sbGP0EmtpuNC
# GVQ/8y8Ooct8/hP5IznaJRy4PgBKOm8yMDdkHseudQfYVdIYyQ6KvKNc8HwKp4WB
# wg6vj5lc02AlvINaaRwlE81y9eucgJvcLGfE3ckJmNVz68Qho+Uyjj4vUpjGYDdk
# jLJvSlRyGMwnh/rNdaJjIUy1PWT9K6abVa8mTGC0uVz+q0O9rdATZlAfC9KJpv/X
# gAbxwxECMzNhF/dWH44vO2jnFfF3VkopngPawismYTJboFblSSmNNqf1x1KiVgMg
# Lzh4gL32Bq5BNMuURb2bx4kYHwu6/6muakCZE93vUN8BuvIE1tAx3zQ4XldbyDge
# VtSsSKbt//m4wTvtwiS+RGCnd83VPZhZtEPqqmB9zcLlL/Hr9dQg1Zc0bl0EawUR
# 0tOSjAknRO1PNTFGfnQZBWLsiePqI3CY5NEv1IoTGEaTZeVYc9NMPSd6Ij/D+KNV
# t/nmh4LsRR7Fbjp8sU65q2j3m2PVkUG8qQIDAQABo4H7MIH4MA8GA1UdEwEB/wQF
# MAMBAf8wHwYDVR0jBBgwFoAU3QQJB6L1en1SUxKSle44gCUNplkwMAYIKwYBBQUH
# AQEEJDAiMCAGCCsGAQUFBzABhhRodHRwOi8vb2NzcHMuc3NsLmNvbTARBgNVHSAE
# CjAIMAYGBFUdIAAwEwYDVR0lBAwwCgYIKwYBBQUHAwMwOwYDVR0fBDQwMjAwoC6g
# LIYqaHR0cDovL2NybHMuc3NsLmNvbS9zc2wuY29tLXJzYS1Sb290Q0EuY3JsMB0G
# A1UdDgQWBBRUwv4QlQCTzWr158DX2bJLuI8M4zAOBgNVHQ8BAf8EBAMCAYYwDQYJ
# KoZIhvcNAQELBQADggIBAPUPJodwr5miyvXWyfCNZj05gtOII9iCv49UhCe204MH
# 154niU2EjlTRIO5gQ9tXQjzHsJX2vszqoz2OTwbGK1mGf+tzG8rlQCbgPW/M9r1x
# xs19DiBAOdYF0q+UCL9/wlG3K7V7gyHwY9rlnOFpLnUdTsthHvWlM98CnRXZ7WmT
# V7pGRS6AvGW+5xI+3kf/kJwQrfZWsqTU+tb8LryXIbN2g9KR+gZQ0bGAKID+260P
# Z+34fdzZcFt6umi1s0pmF4/n8OdX3Wn+vF7h1YyfE7uVmhX7eSuF1W0+Z0duGwdc
# +1RFDxYRLhHDsLy1bhwzV5Qe/kI0Ro4xUE7bM1eV+jjk5hLbq1guRbfZIsr0WkdJ
# LCjoT4xCPGRo6eZDrBmRqccTgl/8cQo3t51Qezxd96JSgjXktefTCm9r/o35pNfV
# HUvnfWII+NnXrJlJ27WEQRQu9i5gl1NLmv7xiHp0up516eDap8nMLDt7TAp4z5T3
# NmC2gzyKVMtODWgqlBF1JhTqIDfM63kXdlV4cW3iSTgzN9vkbFnHI2LmvM4uVEv9
# XgMqyN0eS3FE0HU+MWJliymm7STheh2ENH+kF3y0rH0/NVjLw78a3Z9UVm1F5VPz
# iIorMaPKPlDRADTsJwjDZ8Zc6Gi/zy4WZbg8Zv87spWrmo2dzJTw7XhQf+xkR6Od
# MIIG9TCCBN2gAwIBAgIQBof4ljXihJpmCqqt2mhAWzANBgkqhkiG9w0BAQsFADB4
# MQswCQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNVBAcMB0hvdXN0b24x
# ETAPBgNVBAoMCFNTTCBDb3JwMTQwMgYDVQQDDCtTU0wuY29tIENvZGUgU2lnbmlu
# ZyBJbnRlcm1lZGlhdGUgQ0EgUlNBIFIxMB4XDTI2MTAxMzAyNDgwMloXDTI3MTAx
# MzAyNDgwMlowXDELMAkGA1UEBhMCRlIxGzAZBgNVBAcMEkNoYW50ZWxvdXAtRW4t
# QnJpZTEXMBUGA1UECgwOTmljb2xhcyBGYWJlcnQxFzAVBgNVBAMMDk5pY29sYXMg
# RmFiZXJ0MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAsg6CT0mFQ4jB
# V2plIS4uWFSoVqRHd82Y7XYZ4CUTe/7nPBRUlXXdKTZMMHzGgiG8dEXNE8K1Xe90
# Dgl/FnorzAFlrNWSwf6t27UAUw9xkqtRDVVpVgpixpWkAwjcQeLG+YMqBkGbFi35
# TGvS23lPFWvLp4DxARrv15MtZ/NjNTCAAGUlMP2hFVBknEmrnYjEGyeJ93sl20RS
# WXQQLNqlenBdkGCrQW4XjHtRjJAEqAC68hN6nWevopSQ8CLUV7A1RFfHFarDrzrL
# zr0/NjIdSyIfM9hShln31NDrktB6Ups2XIyrtPDNB7fjkW4jpLlOqQGfYKsU9ErE
# 4SnFsh0yOOjitegyh1RvZ9wH++3DF4CmJfTli5l/gEIm8xCCryxI+K5EogW3wu1y
# u8NT2hcvHjdh7OpRyPyFAuLZ8MfX7kPAJpnuCeOuktzY/20eVD8bDglj+4zgqzsn
# cR/t6T/7w1/HxpRyrVoKvEyRx+mphYh5jMdEa0BjxOL0vAJpdGZSfXGmPm/yPqRn
# 5STTkMtzkVQXPbcvgKM0luzFcQKPHk3qdHLEXa8gf93ktYDbLE5Wd2UHrXz3SZxN
# URHPdIXCr1qgrqQve7041kMpzrkDI5DnG2lvRG4108gp02bHG0muKyZlDdCOra7h
# Zonu2nO79sFallruzzNi0o994PoWksUCAwEAAaOCAZUwggGRMAwGA1UdEwEB/wQC
# MAAwHwYDVR0jBBgwFoAUVML+EJUAk81q9efA19myS7iPDOMwegYIKwYBBQUHAQEE
# bjBsMEgGCCsGAQUFBzAChjxodHRwOi8vY2VydC5zc2wuY29tL1NTTGNvbS1TdWJD
# QS1Db2RlU2lnbmluZy1SU0EtNDA5Ni1SMS5jZXIwIAYIKwYBBQUHMAGGFGh0dHA6
# Ly9vY3Nwcy5zc2wuY29tMFEGA1UdIARKMEgwCAYGZ4EMAQQBMDwGDCsGAQQBgqkw
# AQMDATAsMCoGCCsGAQUFBwIBFh5odHRwczovL3d3dy5zc2wuY29tL3JlcG9zaXRv
# cnkwEwYDVR0lBAwwCgYIKwYBBQUHAwMwTQYDVR0fBEYwRDBCoECgPoY8aHR0cDov
# L2NybHMuc3NsLmNvbS9TU0xjb20tU3ViQ0EtQ29kZVNpZ25pbmctUlNBLTQwOTYt
# UjEuY3JsMB0GA1UdDgQWBBRoUx3630l9wA1fUpkKkw2GUIxl/jAOBgNVHQ8BAf8E
# BAMCB4AwDQYJKoZIhvcNAQELBQADggIBAE3oj6cNuwlDToi1BwjgsA5djfRN4DIi
# BYDFZIRYaIRe/rEdnItz3aqQQdFuwXnJA/MQEmi2WAGek8GeZF+R41s8FbcDxBUE
# BoSt98TuBE+0xgxQWDiPhhXaX1Aw9V0GF4ZnlfCSw1eu2KiASaKn7Bo0PLbEAXff
# vftRXof/kyqtw8iID6I4Rhrf5DI3WVZQodakUUmWnbnziWvEZOWbNOfzwP+7BwEt
# 7jQANfbmEfb4p75rfy5dQVwniTDSvdo6xkrbe1Cr93GS7cJta18IG5CoIVUmdK+s
# W38B1sLuy8Mi0MstMXV3004yaRuzLdNaAM0f9WAFerxemiQC0kdmA44/lWV3lVkM
# WJSwkHBkkSg8Eh0hxcGzTH2xmtDhLJHNNLRwkI9u2Ur+HbLCt50oBgSV8BcCs/1i
# qi884iC/yuX+HIv9YgTUNZ8c4lLaq4VuliJpXLu2buHIsTkO0v0+5oCM91zjTVZm
# ULeH38N0Vr83psiJVuBSxQyt5nRVrpYz7FACq5gtXELkyViq/rGOlJKaRrgQ6WWq
# eoIXqKhVFAdP/jAEm2M8ezoVYsUgPyKY2OcuecV9coulAcxUIUydx6qYiOfsGj5u
# Dy7l4aYjvB771wkUMzUQwxzcDYXJtMMTaWRe6n9EemR67jc73u5CJyFg3PfA5WR1
# wZislcdCZuS6MYISbDCCEmgCAQEwgYwweDELMAkGA1UEBhMCVVMxDjAMBgNVBAgM
# BVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMREwDwYDVQQKDAhTU0wgQ29ycDE0MDIG
# A1UEAwwrU1NMLmNvbSBDb2RlIFNpZ25pbmcgSW50ZXJtZWRpYXRlIENBIFJTQSBS
# MQIQBof4ljXihJpmCqqt2mhAWzANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3
# AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisG
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCA9CDNAUYr+
# ibqx6AMpbnlv1H27TkCmtakiLLLVSEEKqzANBgkqhkiG9w0BAQEFAASCAgAYfLcd
# MqtSJSx1Id317qKOs68MQFJ6ycfWgVJ+WhOiilB587QyHzMLyPt/l3q3VPKZrLw+
# mvE3OIm5NXPAEFW50Qs0gUd4O4S20fH1RKfJkcAe+9HpHzR7yq74a8uPP1IhAdxI
# SKHmY3rHFXJshjx6N2eJCog1dT2CQuZP2iwa2jz0dHAaDa6VHvAq+Bp3TOz+ns3D
# 3206C9YlBk1bZAoKQjWW1we0s7Gfl+UAIgDU+4Y3+izNfN5RNlzwcOmtSFi/45QM
# om61tDuEpGsUykjk+LT/j/Bpw5E9jKPIQ2MudNaT0GsJIcgP94FIY+eC9B+4+BXm
# Kfq4YNEBGite5wg2p1+YHjIyfDZZGrtHPpCvZKnZEjvOYGRVfSGuAEW199M2Mmgb
# hYw6vb6cAFqBmbUEOp6pC9Mv1EiAwxnTT3HVZOKOztkokxtbYnkP1I/h0D5nDAS1
# fB7WvftAY/uC6aSlQ8t0AHQkJDmnklblp6fQkdAqedRpgDM7ETh+rQlJsoGNKPuB
# +SXNdd2jsEKMCMqi9j2CycEo64oA8KHhs434sdZM2d9Snelhc8Vf3VFgfwL+7Ey7
# 4i7mcK8Ie2gc7Yc3Fzo+Kh4s6tYEEJI49HVdS7q5mGBBIo0dscILJW4yvMRoIKR0
# pxjK16IdmaS8WKUwatVv1TcnNGdcA0XD0AkVs6GCDykwgg8lBgorBgEEAYI3AwMB
# MYIPFTCCDxEGCSqGSIb3DQEHAqCCDwIwgg7+AgEDMQ0wCwYJYIZIAWUDBAIBMIGJ
# BgsqhkiG9w0BCRABBKB6BHgwdgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEIC9Gf/hqBxOGne5Rim6Ih2p+QEtdjrpn2TCLFpBnBnYyAghZUZguYL5z
# lhgPMjAyNjEwMTMxMTQyNDhaMAMCAQECEAVJA+lOup+1ZvTmnR+KEhygggwAMIIE
# /DCCAuSgAwIBAgIQH2sWYtIuG2xd8cDBoGAOODANBgkqhkiG9w0BAQsFADBzMQsw
# CQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNVBAcMB0hvdXN0b24xETAP
# BgNVBAoMCFNTTCBDb3JwMS8wLQYDVQQDDCZTU0wuY29tIFRpbWVzdGFtcGluZyBJ
# c3N1aW5nIFJTQSBDQSBSMTAeFw0yNTAyMTgxNjMyMDJaFw0zNDExMTIxODUwMDVa
# MG4xCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91c3Rv
# bjERMA8GA1UECgwIU1NMIENvcnAxKjAoBgNVBAMMIVNTTC5jb20gVGltZXN0YW1w
# aW5nIFVuaXQgMjAyNSBFMTBZMBMGByqGSM49AgEGCCqGSM49AwEHA0IABBv7UVHH
# b+ZVluxPXlfE3M6tg0Xnq8dic+O3vPOCRpalUM1vO9A+GRzSVVjyygHhYrBw62XL
# Fh1kv7e+yRd/aTajggFaMIIBVjAfBgNVHSMEGDAWgBQMnRAljpqnG5mHQ88IfuG9
# gZD0zzBRBggrBgEFBQcBAQRFMEMwQQYIKwYBBQUHMAKGNWh0dHA6Ly9jZXJ0LnNz
# bC5jb20vU1NMLmNvbS10aW1lU3RhbXBpbmctSS1SU0EtUjEuY2VyMFEGA1UdIARK
# MEgwPAYMKwYBBAGCqTABAwYBMCwwKgYIKwYBBQUHAgEWHmh0dHBzOi8vd3d3LnNz
# bC5jb20vcmVwb3NpdG9yeTAIBgZngQwBBAIwFgYDVR0lAQH/BAwwCgYIKwYBBQUH
# AwgwRgYDVR0fBD8wPTA7oDmgN4Y1aHR0cDovL2NybHMuc3NsLmNvbS9TU0wuY29t
# LXRpbWVTdGFtcGluZy1JLVJTQS1SMS5jcmwwHQYDVR0OBBYEFM582cAEgMUkEGoJ
# 6hyrJT0R/ajSMA4GA1UdDwEB/wQEAwIHgDANBgkqhkiG9w0BAQsFAAOCAgEAgHN1
# JIW9ZlNip07FfY9r1LJwLKTl1h/y9r1BjDgHMEz7U7J7mAy32JxA1KLyjCCJqNku
# RfhDHGFEfZtjTOru7AQMDSDiUDvCCc3jaz5py9jGvwIe5npkLbXqXhA+SqCzG719
# USigwhHkYr7v8NiXoaY0Xu03p0BiZZreZlpngmGB9+86N94FqIFx2I5Z5iuu7d4j
# 7aAOZBNxfU9jFDIIGDtWfWeCpF0Q8rLE/3DHDlzJKs5Il3i/VoOl7rTU938oSrxM
# www6GDBlKX8rGHxvkjgIb9MXkVglJSDHsNAZDeCS/inaxKHA+ChuB8K4OxIxa6k1
# e+eHBypakAY3wuN7w7PlPmhN6k5IdqZn3HZHs2VjoR841Z9wmVPFdGDdkBZ7XOmb
# 5OZUaKufsvtER0VbY2DkzaqAaItt1kc+I+FUz7PiVz2PUzpegkftRRxvrPyIL08b
# lJTnbMQ0XrqDN0rEfuTLv4XhhnKyJaXhxCUdjO/cRumMzZ35QTUMItBWy1xGp0iy
# oFVnAya2pPsVMuAI+sy1zxlOS9l5iSxvKJ8gpnpOgjqYTa7u1eyZo+4JlcGoiiR1
# 7LthMTrF1q62tO5xOpI8txUZ6gtKGtkJV6wdj+vZLcwKZ2M93GqVpOo6UBlGjx1s
# KsVZ7AIUzJZli4NJ0M2KYcrfNnlebkYKpDnSqWwwggb8MIIE5KADAgECAhBtUhhw
# h+gjTYVgANCAj5NWMA0GCSqGSIb3DQEBCwUAMHwxCzAJBgNVBAYTAlVTMQ4wDAYD
# VQQIDAVUZXhhczEQMA4GA1UEBwwHSG91c3RvbjEYMBYGA1UECgwPU1NMIENvcnBv
# cmF0aW9uMTEwLwYDVQQDDChTU0wuY29tIFJvb3QgQ2VydGlmaWNhdGlvbiBBdXRo
# b3JpdHkgUlNBMB4XDTE5MTExMzE4NTAwNVoXDTM0MTExMjE4NTAwNVowczELMAkG
# A1UEBhMCVVMxDjAMBgNVBAgMBVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMREwDwYD
# VQQKDAhTU0wgQ29ycDEvMC0GA1UEAwwmU1NMLmNvbSBUaW1lc3RhbXBpbmcgSXNz
# dWluZyBSU0EgQ0EgUjEwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQCu
# URAT0vk8IKAghd7JUBxkyeH9xek0/wp/MUjoclrFXqhh/fGH91Fc+7fm0MHCE7A+
# wmOiqBj9ODrJAYGq3rm33jCnHSsCBNWAQYyoauLq8IjqsS1JlXL29qDNMMdwZ8UN
# zQS7vWZMDJ40JSGNphMGTIA2qn2bohGtgRc4p1395ESypUOaGvJ3t0FNL3BuKmb6
# YctMcQUF2sqooMzd89h0E6ujdvBDo6ZwNnWoxj7YmfWjSXg33A5GuY9ym4QZM5OE
# Vgo8ebz/B+gyhyCLNNhh4Mb/4xvCTCMVmNYrBviGgdPZYrym8Zb84TQCmSuX0JlL
# La6WK1aO6qlwISbb9bVGh866ekKblC/XRP20gAu1CjvcYciUgNTrGFg8f8AJgQPO
# Cc1/CCdaJSYwhJpSdheKOnQgESgNmYZPhFOC6IKaMAUXk5U1tjTcFCgFvvArXtK4
# azAWUOO1Y3fdldIBL6LjkzLUCYJNkFXqhsBVcPMuB0nUDWvLJfPimstjJ8lF4S6E
# CxWnlWi7OElVwTnt1GtRqeY9ydvvGLntU+FecK7DbqHDUd366UreMkSBtzevAc9a
# qoZPnjVMjvFqV1pYOjzmTiVHZtAc80bAfFe5LLfJzPI6DntNyqobpwTevQpHqPDN
# 9qqNO83r3kaw8A9j+HZiSw2AX5cGdQP0kG0vhzfgBwIDAQABo4IBgTCCAX0wEgYD
# VR0TAQH/BAgwBgEB/wIBADAfBgNVHSMEGDAWgBTdBAkHovV6fVJTEpKV7jiAJQ2m
# WTCBgwYIKwYBBQUHAQEEdzB1MFEGCCsGAQUFBzAChkVodHRwOi8vd3d3LnNzbC5j
# b20vcmVwb3NpdG9yeS9TU0xjb21Sb290Q2VydGlmaWNhdGlvbkF1dGhvcml0eVJT
# QS5jcnQwIAYIKwYBBQUHMAGGFGh0dHA6Ly9vY3Nwcy5zc2wuY29tMD8GA1UdIAQ4
# MDYwNAYEVR0gADAsMCoGCCsGAQUFBwIBFh5odHRwczovL3d3dy5zc2wuY29tL3Jl
# cG9zaXRvcnkwEwYDVR0lBAwwCgYIKwYBBQUHAwgwOwYDVR0fBDQwMjAwoC6gLIYq
# aHR0cDovL2NybHMuc3NsLmNvbS9zc2wuY29tLXJzYS1Sb290Q0EuY3JsMB0GA1Ud
# DgQWBBQMnRAljpqnG5mHQ88IfuG9gZD0zzAOBgNVHQ8BAf8EBAMCAYYwDQYJKoZI
# hvcNAQELBQADggIBAJIZdQ2mWkLPGQfZ8vyU+sCb8BXpRJZaL3Ez3VDlE3uZk3cP
# xPtybVfLuqaci0W6SB22JTMttCiQMnIVOsXWnIuAbD/aFTcUkTLBI3xys+wEajzX
# aXJYWACDS47BRjDtYlDW14gLJxf8W6DQoH3jHDGGy8kGJFOlDKG7/YrK7UGfHtBA
# EDVe6lyZ+FtCsrk7dD/IiL/+Q3Q6SFASJLQ2XI89ihFugdYL77CiDNXrI2MFspQG
# swXEAGpHuaQDTHUp/LdR3TyrIsLlnzoLskUGswF/KF8+kpWUiKJNC4rPWtNrxlbX
# YRGgdEdx8SMjUTDClldcrknlFxbqHsVmr9xkT2QtFmG+dEq1v5fsIK0vHaHrWjMM
# maJ9i+4qGJSD0stYfQ6v0PddT7EpGxGd867Ada6FZyHwbuQSadMb0K0P0OC2r7rw
# qBUe0BaMqTa6LWzWItgBjGcObXeMxmbQqlEz2YtAcErkZvh0WABDDE4U8GyV/32F
# daAvJgTfe9MiL2nSBioYe/g5mHUSWAay/Ip1RQmQCvmF9sNfqlhJwkjy/1U1ibUk
# TIUBX3HgymyQvqQTZLLys6pL2tCdWcjI9YuLw30rgZm8+K387L7ycUvqrmQ3ZJlu
# jHl3r1hgV76s3WwMPgKk1bAEFMj+rRXimSC+Ev30hXZdqyMdl/il5Ksd0vhGMYIC
# WDCCAlQCAQEwgYcwczELMAkGA1UEBhMCVVMxDjAMBgNVBAgMBVRleGFzMRAwDgYD
# VQQHDAdIb3VzdG9uMREwDwYDVQQKDAhTU0wgQ29ycDEvMC0GA1UEAwwmU1NMLmNv
# bSBUaW1lc3RhbXBpbmcgSXNzdWluZyBSU0EgQ0EgUjECEB9rFmLSLhtsXfHAwaBg
# DjgwCwYJYIZIAWUDBAIBoIIBYTAaBgkqhkiG9w0BCQMxDQYLKoZIhvcNAQkQAQQw
# HAYJKoZIhvcNAQkFMQ8XDTI2MTAxMzExNDI0OFowKAYJKoZIhvcNAQk0MRswGTAL
# BglghkgBZQMEAgGhCgYIKoZIzj0EAwIwLwYJKoZIhvcNAQkEMSIEIOY+Y0kpUMWm
# YBYYjWz5HU3UMXO2IQfTqo1Ax5zDNRGnMIHJBgsqhkiG9w0BCRACLzGBuTCBtjCB
# szCBsAQgVCr5oWqNci5mEUl4iumUwYqaruWmXLNEolSa+Wx5x4swgYswd6R1MHMx
# CzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91c3RvbjER
# MA8GA1UECgwIU1NMIENvcnAxLzAtBgNVBAMMJlNTTC5jb20gVGltZXN0YW1waW5n
# IElzc3VpbmcgUlNBIENBIFIxAhAfaxZi0i4bbF3xwMGgYA44MAoGCCqGSM49BAMC
# BEcwRQIgVkaW5eq+BWBo767YpkoPtjzhgbauPL0lVZMfUoZFSxICIQD7RMGbqo7h
# jtNuX9LBeV5RDMUyV+3U7oMuhjcec8XBnA==
# SIG # End signature block
