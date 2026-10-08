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
    Version : 2.1.1
    Exit codes : 0 = success, 1 = failure, 2 = finished but incomplete (see the summary).
    Documentation : docs\PurviewDlpReport-Guide.md (or .html)
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
            $store.UpdateRunAccount($runId, $connection.Account)
            Write-DlpItem Ok ("{0}  {3} {1} mode {3} tenant verified {3} {2}" -f $connection.Account, $connection.Mode.ToLowerInvariant(), (Format-DlpDuration $connectClock.Elapsed.TotalSeconds), [char]0x00B7)

            Write-DlpStep 4 $totalSteps 'Collecting Activity Explorer data' -Icon Download
            $collection = Invoke-DlpCollection -Store $store -RunId $runId -TargetId $targetId -Settings $settings -Slices $slices -Exporter $connection.Exporter
            $details['Collection'] = $collection
            if ($collection.Quarantined) { Write-DlpItem Warn ("{0} record(s) could not be read and were set aside (table 'quarantine'): {1}" -f $collection.Quarantined, (($collection.QuarantineReasons.GetEnumerator() | ForEach-Object { "$($_.Key) x$($_.Value)" }) -join ', ')) }
            if ($collection.Conflicts) { Write-DlpItem Info ("{0} event(s) had a sender, subject or recipient list different from an earlier event of the same message; the first values are kept." -f $collection.Conflicts) }
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
        $coverage = Get-DlpCoverage -Store $store -TargetId $targetId -StartMs $period.StartMs -EndMs $period.EndMs
        $reportClock = [Diagnostics.Stopwatch]::StartNew()
        $report = New-DlpReport -Store $store -TargetId $targetId -Settings $settings -Period $period -Coverage $coverage
        foreach ($file in $report.Result.Files) {
            Write-DlpItem Ok ("{0,-4} {1}   {2} messages {4} {3}" -f $file.Kind, (Split-Path $file.Path -Leaf), (Format-DlpNumber $file.Rows), (Format-DlpBytes $file.Bytes), [char]0x00B7) -Icon File
        }
        Write-DlpItem Info ("Written in {0}" -f (Format-DlpDuration $reportClock.Elapsed.TotalSeconds)) -Icon Clock
        $r = $report.Result
        if ($r.MessagesWithoutCount) { Write-DlpItem Info ("{0} message(s) without a reliable recipient count (column left empty)." -f (Format-DlpNumber $r.MessagesWithoutCount)) }
        $complete = $coverage.Percent -ge 100 -and -not $collectionFailed
        $summary = [ordered]@{
            'Period'   = @('Calendar', "$(Format-DlpRange $period.StartMs $period.EndMs $zone)  ($($settings.Report.TimeZone))")
            'Messages' = @('Mail', ('{0} unique messages {2} {1} senders' -f (Format-DlpNumber $r.Messages), (Format-DlpNumber $r.Senders), [char]0x00B7))
            'Coverage' = @($(if ($coverage.Percent -ge 100) { 'Chart' } else { 'Warn' }), $(if ($coverage.Percent -ge 100) { '100% of the period' } else { '{0}% of the period - missing: {1}' -f $coverage.Percent, ((@($coverage.Gaps) | Select-Object -First 3 | ForEach-Object { Format-DlpRange $_.Start $_.End $zone }) -join ', ') }))
            'Files'    = @('File', ('{0} file(s) {2} recipients {1}' -f $r.Files.Count, $(if ($settings.Report.IncludeRecipientDetails) { 'listed' } else { 'counted only' }), [char]0x00B7))
            'Folder'   = @('Folder', $report.Directory)
            'Duration' = @('Clock', (Format-DlpDuration $clock.Elapsed.TotalSeconds))
            'Log'      = @('Log', $logPath)
        }
        Write-DlpSummary -Title $(if ($complete) { 'Report ready' } else { 'Report written, but INCOMPLETE' }) -Values $summary -Status $(if ($complete) { 'Ok' } else { 'Warn' })
        $details['Report'] = [ordered]@{ Directory = $report.Directory; Messages = $r.Messages; Files = @($r.Files | ForEach-Object { $_.Path }); CoveragePercent = $coverage.Percent }
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
