#
#  Purview DLP Report - module manifest
#  --------------------------------------------------------------------------
#  Author  : Nicolas Fabert
#  Version : see ModuleVersion
#
#  Loaded by Invoke-PurviewDlpReport.ps1 (Import-Module by path).
#
@{
    RootModule        = 'PurviewDlpReport.psm1'
    ModuleVersion     = '2.1.1'
    GUID              = 'de838866-574a-439e-b11d-08fd9239e1b2'
    Author            = 'Nicolas Fabert'
    Description       = 'Purview DLP Report: collects Exchange Online DLP rule matches from Activity Explorer into a local SQLite database and produces CSV/HTML reports (one row per Message ID).'
    PowerShellVersion = '7.4'

    # Functions called by Invoke-PurviewDlpReport.ps1 and by the tests. The other functions stay internal
    # to the module: add a function here only when the script or a test calls it.
    FunctionsToExport = @(
        'Import-DlpConfiguration', 'Initialize-DlpEngine', 'Open-DlpStore', 'Enter-DlpLock', 'Exit-DlpLock'
        'Start-DlpLog', 'Stop-DlpLog', 'Write-DlpLog'
        'Write-DlpBanner', 'Write-DlpStep', 'Write-DlpItem', 'Write-DlpSummary'
        'Format-DlpNumber', 'Format-DlpDuration', 'Format-DlpBytes', 'Format-DlpRange'
        'Get-DlpTimeZone', 'Resolve-DlpPeriod', 'New-DlpRangeList', 'Get-DlpCollectionPlan', 'Get-DlpCoverage', 'Get-DlpExpiringGaps'
        'Connect-DlpActivityExplorer', 'Disconnect-DlpActivityExplorer', 'Test-DlpPageEnvelope', 'Invoke-DlpCollection'
        'New-DlpReport', 'Show-DlpStatus', 'Invoke-DlpRetention'
    )
}