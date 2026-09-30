#Requires -Version 7.4
<#
.SYNOPSIS
    Developer test: replays saved Activity Explorer records through the real collection,
    database and report code, without any connection to Microsoft 365.

.DESCRIPTION
    A fake exporter serves the records of a raw.ndjson file (one Activity Explorer record
    per line, as saved by the 2026 prototype runs) page by page, exactly like
    Export-ActivityExplorerData: same envelope fields (ResultData, LastPage, Watermark,
    TotalResultCount, ResultCode, ErrorData marker). Use it to measure performance or to
    check a change of the engine on real data.

.NOTES
    Author  : Nicolas Fabert
    Version : 2.1.1
    Not used in production.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$RawRecordsPath,
    [Parameter(Mandatory)][DateTimeOffset]$StartUtc,
    [Parameter(Mandatory)][DateTimeOffset]$EndUtc,
    [Parameter(Mandatory)][string]$RunDirectory,
    [string]$ConfigPath = (Join-Path $PSScriptRoot '..\config\PurviewDlpReport.config.psd1'),
    [int]$PageSize = 1000,
    [string]$SplitBy = 'Week',
    [int]$MaxRowsPerFile = 500000,
    [bool]$IncludeRecipientDetails = $true,
    [string]$ReferenceCsv
)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$total = [Diagnostics.Stopwatch]::StartNew()
if (Test-Path -LiteralPath $RunDirectory) { throw "RunDirectory already exists: $RunDirectory" }
[void][IO.Directory]::CreateDirectory($RunDirectory)
Import-Module (Join-Path $root 'PurviewDlpReport.psd1') -Force

# Test configuration: same values, but database, reports and logs inside the run directory.
$settings = Import-DlpConfiguration -Path $ConfigPath -Root $root
$settings.Storage.DatabasePath = Join-Path $RunDirectory 'data\PurviewDlpReport.sqlite'
$settings.Report.OutputPath = Join-Path $RunDirectory 'reports'
$settings.Report.SplitBy = $SplitBy
$settings.Report.MaxRowsPerFile = $MaxRowsPerFile
$settings.Report.IncludeRecipientDetails = $IncludeRecipientDetails
$settings.Collection.PageSize = $PageSize
$null = Start-DlpLog -Directory (Join-Path $RunDirectory 'logs')
Initialize-DlpEngine -Root $root

# Load the saved records, sorted by Happened.
$load = [Diagnostics.Stopwatch]::StartNew()
$times = [Collections.Generic.List[long]]::new()
$lines = [Collections.Generic.List[string]]::new()
if (-not ('ReplayLoader' -as [type])) {
    Add-Type -TypeDefinition @"
public static class ReplayLoader {
    public static void Load(string path, System.Collections.Generic.List<long> times, System.Collections.Generic.List<string> lines) {
        var rx = new System.Text.RegularExpressions.Regex("\"Happened\":\"([^\"]+)\"");
        foreach (var line in System.IO.File.ReadLines(path)) {
            if (line.Length == 0) continue;
            var m = rx.Match(line);
            System.DateTimeOffset t;
            if (!m.Success || !System.DateTimeOffset.TryParse(m.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.AssumeUniversal, out t)) continue;
            times.Add(t.ToUnixTimeMilliseconds()); lines.Add(line);
        }
    }
}
"@
}
[ReplayLoader]::Load($RawRecordsPath, $times, $lines)
$timeArray = $times.ToArray(); $lineArray = $lines.ToArray()
[Array]::Sort($timeArray, $lineArray)
Write-Host ("Loaded {0:N0} saved records in {1:0.0} s" -f $timeArray.Count, $load.Elapsed.TotalSeconds)

$counter = @{ Calls = 0 }
$exporter = {
    param([hashtable]$Query)
    $counter.Calls++
    $s = ([DateTimeOffset][DateTime]::SpecifyKind($Query.StartTime, 'Utc')).ToUnixTimeMilliseconds()
    $e = ([DateTimeOffset][DateTime]::SpecifyKind($Query.EndTime, 'Utc')).ToUnixTimeMilliseconds()
    $first = [Array]::BinarySearch($timeArray, $s); if ($first -lt 0) { $first = -bnot $first } else { while ($first -gt 0 -and $timeArray[$first - 1] -eq $s) { $first-- } }
    $last = [Array]::BinarySearch($timeArray, $e); if ($last -lt 0) { $last = -bnot $last } else { while ($last -gt 0 -and $timeArray[$last - 1] -eq $e) { $last-- } }
    $offset = if ($Query.ContainsKey('PageCookie')) { [int]$Query.PageCookie } else { 0 }
    $from = $first + $offset
    $count = [Math]::Max(0, [Math]::Min($Query.PageSize, $last - $from))
    $page = if ($count) { '[' + [string]::Join(',', $lineArray, $from, $count) + ']' } else { '[]' }
    $isLast = $from + $count -ge $last
    [pscustomobject]@{
        ResultData = $page; LastPage = $isLast; Watermark = $(if ($isLast) { '' } else { [string]($offset + $count) })
        TotalResultCount = $last - $first; RecordCount = $count; ResultCode = 'Success'
        ErrorData = 'Microsoft.Exchange.Hygiene.DataInsights.Common.DataInsightsErrorData'
    }
}.GetNewClosure()

$store = Open-DlpStore -Settings $settings
try {
    $targetId = $store.GetOrCreateTarget($settings.Target.PolicyId, $settings.Target.RuleId, $settings.Target.PolicyName, $settings.Target.RuleName)
    $runId = $store.StartRun('OfflineReplay', 'offline', [Environment]::MachineName, (Get-Module PurviewDlpReport).Version.ToString(), $null)
    $startMs = $StartUtc.ToUnixTimeMilliseconds(); $endMs = $EndUtc.ToUnixTimeMilliseconds()
    $plan = Get-DlpCollectionPlan -Store $store -TargetId $targetId -StartMs $startMs -EndMs $endMs -Settings $settings -NowMs ($endMs + 60000)
    Write-DlpStep 1 2 "Collecting (offline replay, $($plan.Slices.Count) windows)"
    $collection = Invoke-DlpCollection -Store $store -RunId $runId -TargetId $targetId -Settings $settings -Slices $plan.Slices -Exporter $exporter
    if (-not $collection.Completed) { throw "Collection failed: $($collection.Error)" }
    Write-DlpStep 2 2 'Writing the report'
    $period = [pscustomobject]@{ Range = 'Custom'; StartMs = $startMs; EndMs = $endMs; Label = '' }
    $coverage = Get-DlpCoverage -Store $store -TargetId $targetId -StartMs $startMs -EndMs $endMs
    $reportClock = [Diagnostics.Stopwatch]::StartNew()
    $report = New-DlpReport -Store $store -TargetId $targetId -Settings $settings -Period $period -Coverage $coverage
    $reportSeconds = $reportClock.Elapsed.TotalSeconds
    foreach ($f in $report.Result.Files) { Write-DlpItem Ok ("{0} {1} {2:N0} rows {3}" -f $f.Kind, (Split-Path $f.Path -Leaf), $f.Rows, (Format-DlpBytes $f.Bytes)) }
    $store.FinishRun($runId, 'Completed', $null)
    $stats = $store.GetStatistics($targetId)
} finally { $store.Dispose(); Stop-DlpLog }

$comparison = $null
if ($ReferenceCsv) {
    $reference = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    Import-Csv -LiteralPath $ReferenceCsv -Delimiter ';' | ForEach-Object { [void]$reference.Add($_.'Message ID') }
    $produced = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($f in $report.Result.Files | Where-Object Kind -eq 'CSV') { Import-Csv -LiteralPath $f.Path -Delimiter ';' | ForEach-Object { [void]$produced.Add($_.'Message ID') } }
    $onlyRef = @($reference | Where-Object { -not $produced.Contains($_) })
    $onlyNew = @($produced | Where-Object { -not $reference.Contains($_) })
    $comparison = [ordered]@{ Reference = $reference.Count; Produced = $produced.Count; OnlyReference = $onlyRef.Count; OnlyProduced = $onlyNew.Count; OnlyReferenceSample = @($onlyRef | Select-Object -First 5); OnlyProducedSample = @($onlyNew | Select-Object -First 5) }
}
$result = [ordered]@{
    RawRecords = $timeArray.Count; Calls = $counter.Calls; Collection = $collection
    ReportSeconds = $reportSeconds; Messages = $report.Result.Messages; Senders = $report.Result.Senders
    Files = @($report.Result.Files | ForEach-Object { [ordered]@{ Kind = $_.Kind; Path = $_.Path; Rows = $_.Rows; Bytes = $_.Bytes } })
    DatabaseBytes = $stats.FileBytes; DatabaseEvents = $stats.Events; DatabaseMessages = $stats.Messages
    TotalSeconds = $total.Elapsed.TotalSeconds; Comparison = $comparison
}
[IO.File]::WriteAllText((Join-Path $RunDirectory 'offline-replay-result.json'), ($result | ConvertTo-Json -Depth 8), [Text.UTF8Encoding]::new($false))
$result | ConvertTo-Json -Depth 4
