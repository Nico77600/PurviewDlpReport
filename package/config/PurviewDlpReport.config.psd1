#
#  Purview DLP Report - configuration file
#  --------------------------------------------------------------------------
#  Author  : Nicolas Fabert
#  Version : 2.1.1
#
#  This file is read by Invoke-PurviewDlpReport.ps1. It is a PowerShell data
#  file: text between quotes, $true / $false, numbers, and @( ) for lists.
#  Lines starting with # are comments.
#
#  Relative paths (.\data, .\reports ...) are relative to the tool folder.
#  Some values can be overridden for one execution on the command line
#  (see "Command-line overrides" in docs\PurviewDlpReport-Guide.md).
#
@{
    # ---------------------------------------------------------------------
    # Tenant
    # ---------------------------------------------------------------------
    Tenant = @{
        # Microsoft Entra tenant ID (GUID). Safety check in both modes: after sign-in,
        # the tool stops if it is connected to another tenant.
        TenantId     = ''
        # Initial domain (xxx.onmicrosoft.com). Certificate mode only:
        # Connect-IPPSSession -Organization requires the tenant name, not the GUID.
        Organization = ''
    }

    # ---------------------------------------------------------------------
    # DLP policy and rule to report on (Activity Explorer filters on them).
    #   PolicyId : Get-DlpCompliancePolicy -Identity '<name>' | Select-Object Name, ImmutableId, Guid
    #   RuleId   : Get-DlpComplianceRule   -Identity '<name>' | Select-Object Name, ImmutableId
    #   Use the values shown as PolicyId / RuleId in an Activity Explorer event
    #   of this rule (the rule value is the rule ImmutableId).
    # ---------------------------------------------------------------------
    Target = @{
        PolicyId   = ''
        RuleId     = ''
        PolicyName = ''    # optional, display only (console and report)
        RuleName   = ''    # optional, display only (console and report)
    }

    # ---------------------------------------------------------------------
    # Authentication to Security & Compliance PowerShell (Connect-IPPSSession).
    #   Interactive : an administrator signs in (browser window, MFA supported).
    #   Certificate : app-only, no human interaction (scheduled task).
    #                 Requires a dedicated app registration with the application permission
    #                 Exchange.ManageAsApp on 'Office 365 Exchange Online' (and 'Microsoft Exchange
    #                 Online Protection'), admin consent, and the Purview role 'Information
    #                 Protection Reader' through a role group. See the guide, chapter 7 and Annex F.
    # ---------------------------------------------------------------------
    Authentication = @{
        Mode                  = 'Interactive'      # Interactive | Certificate
        UserPrincipalName     = ''   # Interactive: expected account ('' = any account of the tenant)
        DisableWAM            = $true              # Interactive: use the browser instead of the Windows broker (WAM). Keep $true: WAM crashes PowerShell 7 (msalruntime.dll)
        AppId                 = ''                 # Certificate: application (client) ID
        CertificateThumbprint = ''                 # Certificate: thumbprint of the certificate in Cert:\CurrentUser\My or Cert:\LocalMachine\My
    }

    # ---------------------------------------------------------------------
    # Collection from Activity Explorer (Export-ActivityExplorerData).
    # ---------------------------------------------------------------------
    Collection = @{
        PageSize             = 1000   # records per call (1-5000). 1000 is validated in production-like volume.
        SliceHours           = 24     # maximum length of one request window (a window never crosses local midnight)
        SettlingHours        = 6      # the last N hours are re-collected at the next run (Activity Explorer can
                                      # show events 60-90 minutes after they happen, sometimes later)
        BackfillDays         = 30     # first run / Collect mode: how many days back to collect
        SourceRetentionDays  = 30     # Activity Explorer keeps 30 days: older periods cannot be collected any more
        RetentionWarningDays = 7      # warn when a missing period will become uncollectable within N days
        MaxSliceRestarts     = 2      # retries of one window after a recoverable error, before it is split in two
        MinimumSliceMinutes  = 15     # a window is never split below this length
        MaxPagesPerSlice     = 2000   # safety limit (2000 x 1000 = 2 million events in one window)
        CookieSafetySeconds  = 105    # the paging cookie expires after 120 s: restart the window before that
    }

    # ---------------------------------------------------------------------
    # Local database (SQLite). Keeps the history beyond the 30 days of
    # Activity Explorer. Budget about 1.5 to 2.5 GB per month at 30-50k
    # messages per day, recipients included (lab: 461 MB for 278,000 events),
    # that is about 9 to 15 GB with a retention of 180 days.
    # ---------------------------------------------------------------------
    Storage = @{
        DatabasePath           = '.\data\PurviewDlpReport.sqlite'
        RetentionDays          = 180   # events older than this are deleted at each collection (0 = keep everything)
        KeepQuarantinedRecords = $true # keep the raw JSON of records that could not be read (troubleshooting)
    }

    # ---------------------------------------------------------------------
    # Report files (CSV and HTML), written locally only.
    # ---------------------------------------------------------------------
    Report = @{
        DefaultRange            = 'Last7Days'      # used when -Range is not given: Last24Hours | Last7Days | Last30Days | PreviousMonth
        TimeZone                = 'Europe/Paris'   # time zone of dates, days, weeks and months in the report
        OutputPath              = '.\reports'      # one sub-folder per execution
        FilePrefix              = 'PurviewDLP'
        Formats                 = @('Csv', 'Html')
        IncludeRecipientDetails = $true            # $true: list of recipient addresses in CSV and HTML; $false: count only
        SplitBy                 = 'Week'           # used only above MaxRowsPerFile: Rows | Day | Week (Monday to Sunday)
        MaxRowsPerFile          = 500000           # no split below this number of messages; maximum 1,048,575 (Excel limit)
        CsvDelimiter            = ';'              # ';' opens directly in Excel with French regional settings
        Title                   = 'DLP messages with more than 25 recipients'
    }

    # ---------------------------------------------------------------------
    # Log files (one file per day, deleted after RetentionDays).
    # ---------------------------------------------------------------------
    Logging = @{
        Path          = '.\logs'
        RetentionDays = 30
    }
}
