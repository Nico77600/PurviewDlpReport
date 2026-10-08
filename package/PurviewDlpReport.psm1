#Requires -Version 7.4
<#
.SYNOPSIS
    Purview DLP Report - PowerShell module.

.DESCRIPTION
    Helper functions used by Invoke-PurviewDlpReport.ps1. The module is organised
    in regions, in the order of an execution:

        1. Console and log        Write-Dlp* functions (what the administrator sees)
        2. Configuration          Import-DlpConfiguration (reads and checks the .psd1 file)
        3. Engine                 Initialize-DlpEngine (SQLite + compiled C# engine)
        4. Periods and coverage   Resolve-DlpPeriod, Get-DlpCollectionPlan, Get-DlpCoverage
        5. Connection             Connect-DlpActivityExplorer / Disconnect-DlpActivityExplorer
        6. Collection             Invoke-DlpCollection (Activity Explorer -> SQLite)
        7. Report                 New-DlpReport (SQLite -> CSV / HTML)
        8. Status and maintenance Show-DlpStatus, Invoke-DlpRetention, Enter-DlpLock

    Performance-critical work (JSON parsing, database, file writing) is done by
    src\PurviewDlpReport.Engine.cs, compiled on first use.

.NOTES
    Author  : Nicolas Fabert
    Version : 2.1.1
    History : see CHANGELOG.md
#>
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:ToolVersion = '2.1.1'
$script:ToolRoot = $PSScriptRoot
# Prefix given to the Security & Compliance cmdlets imported by Connect-IPPSSession.
# It avoids any clash with another Exchange/Purview session opened in the same console.
$script:CommandPrefix = 'PdrAe'
$script:LogWriter = $null
$script:LogPath = $null
# ---------------------------------------------------------------------------------------------
# Console theme: colours and icons.
#   - Colours (ANSI) are disabled when the output is redirected (scheduled task, log capture)
#     or when the NO_COLOR environment variable is set; PDR_FORCE_COLOR=1 forces them.
#   - Icons: emoji in modern terminals (Windows Terminal, VS Code), simple symbols elsewhere.
#     Emoji are chosen among those always two columns wide (no variation selector), so frames
#     and columns stay aligned. The classic console (conhost) has no font fallback: a character
#     missing from its font is shown as an empty box. The symbols used outside the modern
#     terminals are therefore all present in Consolas and Lucida Console (code page 437
#     repertoire or Latin-1) - a test checks it. Force a style with the environment
#     variable PDR_ICONS = Emoji | Symbols | Ascii.
#   - Frames: rounded corners, present in Consolas (default font of the console). Lucida Console
#     and the raster font have no rounded corners: square corners are used with them (the
#     console font is read when the frame is drawn, see Get-DlpFrame).
# ---------------------------------------------------------------------------------------------
$script:C = @{ Reset = ''; Bold = ''; Dim = ''; Accent = ''; AccentBg = ''; Cyan = ''; Green = ''; Yellow = ''; Red = ''; White = '' }
if ($env:PDR_FORCE_COLOR -eq '1' -or (-not [Console]::IsOutputRedirected -and -not $env:NO_COLOR)) {
    $e = [char]27
    $script:C = @{
        Reset = "$e[0m"; Bold = "$e[1m"; Dim = "$e[90m"; White = "$e[97m"
        Accent = "$e[38;2;214;62;115m"; AccentBg = "$e[48;2;177;31;75m$e[97m"
        Cyan = "$e[38;2;97;214;214m"; Green = "$e[38;2;80;200;120m"; Yellow = "$e[38;2;240;200;90m"; Red = "$e[38;2;240;90;90m"
    }
}
$script:IconStyle = if ($env:PDR_ICONS -in 'Emoji', 'Symbols', 'Ascii') { $env:PDR_ICONS }
    elseif ([Console]::IsOutputRedirected) { 'Symbols' }
    elseif ($env:WT_SESSION -or $env:TERM_PROGRAM -eq 'vscode') { 'Emoji' }
    else { 'Symbols' }

function Get-DlpIconSet {
    <# Icons of one console style. Symbols: only characters of the classic console fonts. #>
    param([Parameter(Mandatory)][ValidateSet('Emoji', 'Symbols', 'Ascii')][string]$Style)
    $u = { param([int]$Code) [char]::ConvertFromUtf32($Code) }
    switch ($Style) {
        'Emoji' { return @{
                Logo = & $u 0x1F512; Ok = & $u 0x2705; Warn = (& $u 0x26A0) + [char]0xFE0F
                Fail = & $u 0x274C; Info = & $u 0x1F539; Skip = & $u 0x23E9
                Database = & $u 0x1F4BE; Plan = & $u 0x1F50E; Key = & $u 0x1F510
                Download = & $u 0x1F4E5; Report = & $u 0x1F4CA; Calendar = & $u 0x1F4C5
                File = & $u 0x1F4C4; Folder = & $u 0x1F4C1; Clock = & $u 0x23F3
                Mail = & $u 0x1F4E8; Target = & $u 0x1F3AF; Log = & $u 0x1F4DD
                Done = & $u 0x1F389; People = & $u 0x1F465; Chart = & $u 0x1F4C8
            } }
        'Symbols' { return @{
                Logo = & $u 0x2666; Ok = & $u 0x221A; Warn = & $u 0x25B2; Fail = & $u 0x00D7; Info = & $u 0x2022
                Skip = & $u 0x00BB; Database = & $u 0x25A0; Plan = & $u 0x25BA; Key = & $u 0x2194; Download = & $u 0x2193
                Report = & $u 0x2261; Calendar = & $u 0x263C; File = & $u 0x25AC; Folder = & $u 0x2302; Clock = & $u 0x25CB
                Mail = '@'; Target = & $u 0x25D9; Log = & $u 0x00B6; Done = & $u 0x221A; People = & $u 0x2192; Chart = & $u 0x2191
            } }
        default { return @{
                Logo = '*'; Ok = '+'; Warn = '!'; Fail = 'x'; Info = '-'; Skip = '>'; Database = '#'; Plan = '?'; Key = '@'; Download = 'v'; Report = '='; Calendar = ':'
                File = '-'; Folder = '>'; Clock = '~'; Mail = '@'; Target = 'o'; Log = '='; Done = '*'; People = '&'; Chart = '^'
            } }
    }
}

function Get-DlpFrameSet {
    <#
    .SYNOPSIS
        Frame characters. Rounded corners, except with the Ascii style and with the console fonts
        that have no rounded corners (Lucida Console, raster font 'Terminal'): square corners.
    #>
    param([Parameter(Mandatory)][ValidateSet('Emoji', 'Symbols', 'Ascii')][string]$Style, [AllowNull()][string]$FontName)
    if ($Style -eq 'Ascii') { return @{ TopLeft = [char]'+'; TopRight = [char]'+'; BottomLeft = [char]'+'; BottomRight = [char]'+'; Horizontal = [char]'-'; Vertical = [char]'|' } }
    if ($Style -eq 'Symbols' -and $FontName -in 'Lucida Console', 'Terminal') {
        return @{ TopLeft = [char]0x250C; TopRight = [char]0x2510; BottomLeft = [char]0x2514; BottomRight = [char]0x2518; Horizontal = [char]0x2500; Vertical = [char]0x2502 }
    }
    return @{ TopLeft = [char]0x256D; TopRight = [char]0x256E; BottomLeft = [char]0x2570; BottomRight = [char]0x256F; Horizontal = [char]0x2500; Vertical = [char]0x2502 }
}

function Get-DlpFrame {
    <# Frame characters for this console. In the classic console the font is read once, through the engine. #>
    if ($script:Frame) { return $script:Frame }
    $engine = [bool]('PurviewDlpReport.ConsoleFont' -as [type])
    $font = if ($script:IconStyle -eq 'Symbols' -and $engine) { [PurviewDlpReport.ConsoleFont]::FaceName() }
    $frame = Get-DlpFrameSet $script:IconStyle $font
    # Before the engine is loaded (error at start) the font is unknown: not kept, read again later.
    if ($script:IconStyle -ne 'Symbols' -or $engine) { $script:Frame = $frame }
    return $frame
}

$script:Icons = Get-DlpIconSet $script:IconStyle
$script:Frame = $null
# Emoji are two columns wide in the console; symbols are one: pad symbols so text stays aligned.
$script:IconPad = if ($script:IconStyle -eq 'Emoji') { ' ' } else { '  ' }

#region 1. Console and log ---------------------------------------------------------------

function Get-DlpIcon {
    param([Parameter(Mandatory)][string]$Name)
    return $script:Icons[$Name] + $script:IconPad
}

function Format-DlpNumber {
    param([Parameter(Mandatory)][AllowNull()]$Value)
    if ($null -eq $Value) { return '-' }
    return ([long]$Value).ToString('N0', [Globalization.CultureInfo]::GetCultureInfo('en-US'))
}

function Format-DlpDuration {
    param([Parameter(Mandatory)][double]$Seconds)
    # 0.0 (not 0): with an integer first argument PowerShell picks Math.Max(int, int) and drops the decimals.
    $t = [TimeSpan]::FromTicks([long]([Math]::Max(0.0, $Seconds) * 10000000))
    if ($t.TotalDays -ge 2) { return '{0} d {1:00} h' -f [int][Math]::Floor($t.TotalDays), $t.Hours }
    if ($t.TotalHours -ge 1) { return '{0} h {1:00} min' -f [int][Math]::Floor($t.TotalHours), $t.Minutes }
    if ($t.TotalMinutes -ge 1) { return '{0} min {1:00} s' -f $t.Minutes, $t.Seconds }
    return '{0:0.0} s' -f $t.TotalSeconds
}

function Format-DlpBytes {
    param([Parameter(Mandatory)][double]$Bytes)
    if ($Bytes -ge 1GB) { return '{0:0.00} GB' -f ($Bytes / 1GB) }
    if ($Bytes -ge 1MB) { return '{0:0.0} MB' -f ($Bytes / 1MB) }
    if ($Bytes -ge 1KB) { return '{0:0} KB' -f ($Bytes / 1KB) }
    return '{0:0} B' -f $Bytes
}

function Format-DlpLocalTime {
    param([Parameter(Mandatory)][long]$UnixMs, [Parameter(Mandatory)][TimeZoneInfo]$Zone, [string]$Format = 'yyyy-MM-dd HH:mm')
    return [TimeZoneInfo]::ConvertTime([DateTimeOffset]::FromUnixTimeMilliseconds($UnixMs), $Zone).ToString($Format, [Globalization.CultureInfo]::InvariantCulture)
}

function Format-DlpRange {
    param([long]$StartMs, [long]$EndMs, [TimeZoneInfo]$Zone)
    return '{0} {2} {1}' -f (Format-DlpLocalTime $StartMs $Zone), (Format-DlpLocalTime $EndMs $Zone), [char]0x2192
}

function Start-DlpLog {
    <# Opens (or continues) today's log file and deletes log files older than the retention. #>
    param([Parameter(Mandatory)][string]$Directory, [int]$RetentionDays = 30)
    [void][IO.Directory]::CreateDirectory($Directory)
    $script:LogPath = Join-Path $Directory ('PurviewDlpReport_{0:yyyyMMdd}.log' -f (Get-Date))
    $stream = [IO.FileStream]::new($script:LogPath, [IO.FileMode]::Append, [IO.FileAccess]::Write, [IO.FileShare]::ReadWrite)
    $script:LogWriter = [IO.StreamWriter]::new($stream, [Text.UTF8Encoding]::new($false))
    $script:LogWriter.AutoFlush = $true
    $limit = (Get-Date).AddDays(-$RetentionDays)
    Get-ChildItem -LiteralPath $Directory -Filter 'PurviewDlpReport_*.log' -File -ErrorAction SilentlyContinue |
        Where-Object LastWriteTime -lt $limit | Remove-Item -Force -ErrorAction SilentlyContinue
    return $script:LogPath
}

function Stop-DlpLog {
    if ($script:LogWriter) { $script:LogWriter.Dispose(); $script:LogWriter = $null }
}

function Write-DlpLog {
    <# Writes one line to the log file only (never to the console). The log never contains colours or icons. #>
    param([ValidateSet('INFO', 'OK', 'WARN', 'ERROR', 'STEP', 'DEBUG')][string]$Level = 'INFO', [Parameter(Mandatory)][AllowEmptyString()][string]$Message)
    if ($script:LogWriter) {
        $script:LogWriter.WriteLine(('{0:yyyy-MM-ddTHH:mm:ss.fffzzz} [{1,-5}] {2}' -f (Get-Date), $Level, $Message))
    }
}

function Write-DlpBanner {
    <#
    .SYNOPSIS
        Title card at the start of an execution:

          ╭──────────────────────────────────────────────────────────────────────╮
          │  🛡  Purview DLP Report                    v2.0.0 · Nicolas Fabert   │
          │     Exchange Online DLP matches · one row per Message ID             │
          ╰──────────────────────────────────────────────────────────────────────╯
             📅  Period     2026-09-22 19:58 → 2026-09-29 19:58 ...
    .PARAMETER Details
        Ordered list of rows: key = label, value = @(IconName, Text) or plain text.
    #>
    param([Parameter(Mandatory)][string]$Title, [string]$Subtitle, [System.Collections.Specialized.OrderedDictionary]$Details)
    $C = $script:C; $F = Get-DlpFrame; $width = 74
    $right = "v$($script:ToolVersion) $([char]0x00B7) Nicolas Fabert"
    $iconWidth = if ($script:IconStyle -eq 'Emoji') { 2 } else { 1 }
    $left = "  $($script:Icons.Logo)  $Title"
    $gap = [Math]::Max(1, $width - ($left.Length - $script:Icons.Logo.Length + $iconWidth) - $right.Length - 2)
    Write-Host ''
    Write-Host ("  {0}{1}{2}{3}{4}" -f $C.Accent, $F.TopLeft, [string]::new($F.Horizontal, $width), $F.TopRight, $C.Reset)
    Write-Host ("  {0}{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}" -f $C.Accent, $F.Vertical, $C.Reset, $C.Bold, $left, $C.Reset, [string]::new(' ', $gap), $C.Dim, $right, '  ', ($C.Accent + $F.Vertical), $C.Reset)
    if ($Subtitle) {
        $sub = "     $Subtitle"
        Write-Host ("  {0}{1}{2}{3}{4}{5}{0}{6}{2}" -f $C.Accent, $F.Vertical, $C.Reset, $C.Dim, $sub.PadRight($width), $C.Reset, $F.Vertical)
    }
    Write-Host ("  {0}{1}{2}{3}{4}" -f $C.Accent, $F.BottomLeft, [string]::new($F.Horizontal, $width), $F.BottomRight, $C.Reset)
    if ($Details) {
        foreach ($key in $Details.Keys) {
            $value = $Details[$key]
            $icon, $text = if ($value -is [array]) { (Get-DlpIcon $value[0]), $value[1] } else { '   ', $value }
            Write-Host ("     {0}{1}{2,-10}{3} {4}" -f $icon, $C.Dim, $key, $C.Reset, $text)
        }
    }
    Write-DlpLog 'STEP' "=== $Title v$($script:ToolVersion) ==="
    if ($Details) { foreach ($key in $Details.Keys) { $v = $Details[$key]; Write-DlpLog 'INFO' ("{0}: {1}" -f $key, $(if ($v -is [array]) { $v[1] } else { $v })) } }
}

function Write-DlpStep {
    <#
    .SYNOPSIS
        Step header with a coloured number pill and an icon, e.g.

          ─── 3/5 ─ 🔐  Connecting to Security & Compliance PowerShell
    #>
    param([Parameter(Mandatory)][int]$Number, [Parameter(Mandatory)][int]$Total, [Parameter(Mandatory)][string]$Title, [string]$Icon = 'Info')
    $C = $script:C
    Write-Host ''
    Write-Host ("  {0} {1}/{2} {3} {4}{5}{6}{3}" -f $C.AccentBg, $Number, $Total, $C.Reset, (Get-DlpIcon $Icon), $C.Bold, $Title)
    Write-DlpLog 'STEP' "[$Number/$Total] $Title"
}

function Write-DlpItem {
    <# One indented result line with a status icon, also written to the log. #>
    param([ValidateSet('Ok', 'Warn', 'Fail', 'Info', 'Skip')][string]$Status = 'Info', [Parameter(Mandatory)][AllowEmptyString()][string]$Text, [string]$Icon)
    $color = @{ Ok = $script:C.Green; Warn = $script:C.Yellow; Fail = $script:C.Red; Info = ''; Skip = $script:C.Dim }[$Status]
    $level = @{ Ok = 'OK'; Warn = 'WARN'; Fail = 'ERROR'; Info = 'INFO'; Skip = 'INFO' }[$Status]
    $symbol = Get-DlpIcon $(if ($Icon) { $Icon } else { $Status })
    $textColor = if ($Status -in 'Warn', 'Fail', 'Skip') { $color } else { '' }
    Write-Host ("      {0}{1}{2}{3}{4}{2}" -f $color, $symbol, $script:C.Reset, $textColor, $Text)
    Write-DlpLog $level $Text
}

function Write-DlpTableRow {
    <#
    .SYNOPSIS
        One aligned row of the collection table (one row per Activity Explorer window).
        Columns: status icon, window, events, new, pages, duration, rate.
    #>
    param([switch]$Header, [ValidateSet('Ok', 'Warn', 'Fail')][string]$Status = 'Ok', [string]$Window, $Events, $New, $Pages, [string]$Duration, [string]$Rate)
    $C = $script:C
    if ($Header) {
        Write-Host ("      {0}{1}{2,-35} {3,10} {4,10} {5,6}  {6,12}  {7,9}{8}" -f $C.Dim, ('  ' + $script:IconPad), 'Window', 'Events', 'New', 'Pages', 'Duration', 'Rate', $C.Reset)
        return
    }
    $color = @{ Ok = $C.Green; Warn = $C.Yellow; Fail = $C.Red }[$Status]
    $eventsText = Format-DlpNumber $Events
    $eventsColor = if ([long]$Events -gt 0) { $C.White + $C.Bold } else { $C.Dim }
    Write-Host ("      {0}{1}{2}{3,-35} {4}{5,10}{2} {6,10} {7,6}  {8,12}  {9}{10,9}{2}" -f $color, (Get-DlpIcon $Status), $C.Reset, $Window, $eventsColor, $eventsText, (Format-DlpNumber $New), $Pages, $Duration, $C.Dim, $Rate)
    Write-DlpLog $(if ($Status -eq 'Ok') { 'OK' } else { 'WARN' }) ("{0}  {1} events, {2} new, {3} pages, {4}, {5}" -f $Window, $eventsText, (Format-DlpNumber $New), $Pages, $Duration, $Rate)
}

function Write-DlpSummary {
    <#
    .SYNOPSIS
        Final summary card:

          ╭─ 🎉  Report ready ─────────────────────────────────────────────────╮
            📅  Period      ...
            ✉️  Messages    ...
          ╰────────────────────────────────────────────────────────────────────╯
    .PARAMETER Values
        Ordered list: key = label, value = @(IconName, Text) or plain text.
    #>
    param([Parameter(Mandatory)][string]$Title, [Parameter(Mandatory)][System.Collections.Specialized.OrderedDictionary]$Values, [ValidateSet('Ok', 'Warn', 'Fail')][string]$Status = 'Ok')
    $C = $script:C; $F = Get-DlpFrame; $width = 74
    $color = @{ Ok = $C.Green; Warn = $C.Yellow; Fail = $C.Red }[$Status]
    $icon = $script:Icons[@{ Ok = 'Done'; Warn = 'Warn'; Fail = 'Fail' }[$Status]]
    $iconWidth = if ($script:IconStyle -eq 'Emoji') { 2 } else { 1 }
    $head = " $icon  $Title "
    $rest = [Math]::Max(2, $width - 1 - ($head.Length - $icon.Length + $iconWidth))
    Write-Host ''
    Write-Host ("  {0}{1}{2}{3}{4}{0}{5}{6}{7}" -f $color, $F.TopLeft, $F.Horizontal, $C.Bold, $head, ($C.Reset + $color), ([string]::new($F.Horizontal, $rest) + $F.TopRight), $C.Reset)
    foreach ($key in $Values.Keys) {
        $value = $Values[$key]
        $rowIcon, $text = if ($value -is [array]) { (Get-DlpIcon $value[0]), $value[1] } else { '   ', $value }
        Write-Host ("    {0}{1}{2,-10}{3} {4}" -f $rowIcon, $C.Dim, $key, $C.Reset, $text)
        Write-DlpLog 'INFO' ("Summary - {0}: {1}" -f $key, $text)
    }
    Write-Host ("  {0}{1}{2}{3}{4}" -f $color, $F.BottomLeft, [string]::new($F.Horizontal, $width), $F.BottomRight, $C.Reset)
    Write-Host ''
}

#endregion
#region 2. Configuration ------------------------------------------------------------------

function Resolve-DlpPath {
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][string]$Root)
    $expanded = [Environment]::ExpandEnvironmentVariables($Path)
    if ([IO.Path]::IsPathRooted($expanded)) { return [IO.Path]::GetFullPath($expanded) }
    return [IO.Path]::GetFullPath((Join-Path $Root $expanded))
}

function Import-DlpConfiguration {
    <#
    .SYNOPSIS
        Reads the configuration file, checks every value and returns it with absolute paths.
        All problems are reported together so the administrator can fix them in one go.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path, [string]$Root = $script:ToolRoot)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Configuration file not found: $Path" }
    try { $config = Import-PowerShellDataFile -LiteralPath $Path }
    catch { throw "The configuration file is not valid PowerShell data ($Path): $($_.Exception.Message)" }

    $errors = [Collections.Generic.List[string]]::new()
    $guid = '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$'
    foreach ($section in 'Tenant', 'Target', 'Authentication', 'Collection', 'Storage', 'Report', 'Logging') {
        if (-not $config.ContainsKey($section) -or $config[$section] -isnot [hashtable]) { $errors.Add("Section '$section' is missing.") }
    }
    if ($errors.Count) { throw ("Invalid configuration ($Path):`n - " + ($errors -join "`n - ")) }

    function Get-Value([hashtable]$Section, [string]$SectionName, [string]$Key, $Default, [switch]$Required) {
        if ($Section.ContainsKey($Key) -and $null -ne $Section[$Key] -and "$($Section[$Key])" -ne '') { return $Section[$Key] }
        if ($Required) { $errors.Add("$SectionName.$Key is required.") }
        return $Default
    }
    function Test-Int($Value, [string]$Name, [long]$Min, [long]$Max) {
        $n = 0L
        if (-not [long]::TryParse("$Value", [ref]$n) -or $n -lt $Min -or $n -gt $Max) { $errors.Add("$Name must be a whole number between $Min and $Max (current value: '$Value').") }
        return [int]$n
    }
    function Test-Bool($Value, [string]$Name) {
        if ($Value -isnot [bool]) { $errors.Add("$Name must be `$true or `$false (current value: '$Value').") ; return $false }
        return $Value
    }

    $t = $config.Tenant; $g = $config.Target; $a = $config.Authentication; $c = $config.Collection; $s = $config.Storage; $r = $config.Report; $l = $config.Logging
    $settings = [ordered]@{
        ConfigPath = [IO.Path]::GetFullPath($Path)
        Tenant = [ordered]@{
            TenantId     = [string](Get-Value $t 'Tenant' 'TenantId' '' -Required)
            Organization = [string](Get-Value $t 'Tenant' 'Organization' '')
        }
        Target = [ordered]@{
            PolicyId   = [string](Get-Value $g 'Target' 'PolicyId' '' -Required)
            RuleId     = [string](Get-Value $g 'Target' 'RuleId' '' -Required)
            PolicyName = [string](Get-Value $g 'Target' 'PolicyName' '')
            RuleName   = [string](Get-Value $g 'Target' 'RuleName' '')
        }
        Authentication = [ordered]@{
            Mode                  = [string](Get-Value $a 'Authentication' 'Mode' 'Interactive')
            UserPrincipalName     = [string](Get-Value $a 'Authentication' 'UserPrincipalName' '')
            DisableWAM            = Get-Value $a 'Authentication' 'DisableWAM' $true
            AppId                 = [string](Get-Value $a 'Authentication' 'AppId' '')
            CertificateThumbprint = [string](Get-Value $a 'Authentication' 'CertificateThumbprint' '')
        }
        Collection = [ordered]@{
            PageSize             = Test-Int (Get-Value $c 'Collection' 'PageSize' 1000) 'Collection.PageSize' 1 5000
            SliceHours           = Test-Int (Get-Value $c 'Collection' 'SliceHours' 24) 'Collection.SliceHours' 1 24
            SettlingHours        = Test-Int (Get-Value $c 'Collection' 'SettlingHours' 6) 'Collection.SettlingHours' 0 72
            BackfillDays         = Test-Int (Get-Value $c 'Collection' 'BackfillDays' 30) 'Collection.BackfillDays' 1 30
            SourceRetentionDays  = Test-Int (Get-Value $c 'Collection' 'SourceRetentionDays' 30) 'Collection.SourceRetentionDays' 1 365
            RetentionWarningDays = Test-Int (Get-Value $c 'Collection' 'RetentionWarningDays' 7) 'Collection.RetentionWarningDays' 0 364
            MaxSliceRestarts     = Test-Int (Get-Value $c 'Collection' 'MaxSliceRestarts' 2) 'Collection.MaxSliceRestarts' 0 5
            MinimumSliceMinutes  = Test-Int (Get-Value $c 'Collection' 'MinimumSliceMinutes' 15) 'Collection.MinimumSliceMinutes' 1 720
            MaxPagesPerSlice     = Test-Int (Get-Value $c 'Collection' 'MaxPagesPerSlice' 2000) 'Collection.MaxPagesPerSlice' 1 100000
            CookieSafetySeconds  = Test-Int (Get-Value $c 'Collection' 'CookieSafetySeconds' 105) 'Collection.CookieSafetySeconds' 30 119
        }
        Storage = [ordered]@{
            DatabasePath           = Resolve-DlpPath ([string](Get-Value $s 'Storage' 'DatabasePath' '.\data\PurviewDlpReport.sqlite')) $Root
            RetentionDays          = Test-Int (Get-Value $s 'Storage' 'RetentionDays' 180) 'Storage.RetentionDays' 0 3650
            KeepQuarantinedRecords = Test-Bool (Get-Value $s 'Storage' 'KeepQuarantinedRecords' $true) 'Storage.KeepQuarantinedRecords'
        }
        Report = [ordered]@{
            DefaultRange            = [string](Get-Value $r 'Report' 'DefaultRange' 'Last7Days')
            TimeZone                = [string](Get-Value $r 'Report' 'TimeZone' 'Europe/Paris')
            OutputPath              = Resolve-DlpPath ([string](Get-Value $r 'Report' 'OutputPath' '.\reports')) $Root
            FilePrefix              = [string](Get-Value $r 'Report' 'FilePrefix' 'PurviewDLP')
            Formats                 = @(Get-Value $r 'Report' 'Formats' @('Csv', 'Html'))
            IncludeRecipientDetails = Test-Bool (Get-Value $r 'Report' 'IncludeRecipientDetails' $true) 'Report.IncludeRecipientDetails'
            SplitBy                 = [string](Get-Value $r 'Report' 'SplitBy' 'Week')
            MaxRowsPerFile          = Test-Int (Get-Value $r 'Report' 'MaxRowsPerFile' 500000) 'Report.MaxRowsPerFile' 1000 1048575
            CsvDelimiter            = [string](Get-Value $r 'Report' 'CsvDelimiter' ';')
            Title                   = [string](Get-Value $r 'Report' 'Title' 'DLP messages with more than 25 recipients')
            TemplatePath            = Join-Path $Root 'templates\Report.template.html'
        }
        Logging = [ordered]@{
            Path          = Resolve-DlpPath ([string](Get-Value $l 'Logging' 'Path' '.\logs')) $Root
            RetentionDays = Test-Int (Get-Value $l 'Logging' 'RetentionDays' 30) 'Logging.RetentionDays' 1 3650
        }
    }
    if ($settings.Tenant.TenantId -and $settings.Tenant.TenantId -notmatch $guid) { $errors.Add('Tenant.TenantId must be a GUID.') }
    if ($settings.Target.PolicyId -and $settings.Target.PolicyId -notmatch $guid) { $errors.Add('Target.PolicyId must be a GUID.') }
    if ($settings.Target.RuleId -and $settings.Target.RuleId -notmatch $guid) { $errors.Add('Target.RuleId must be a GUID.') }
    if ($settings.Authentication.Mode -notin 'Interactive', 'Certificate') { $errors.Add("Authentication.Mode must be 'Interactive' or 'Certificate'.") }
    [void](Test-Bool $settings.Authentication.DisableWAM 'Authentication.DisableWAM')
    if ($settings.Authentication.Mode -eq 'Certificate') {
        if ($settings.Authentication.AppId -notmatch $guid) { $errors.Add('Authentication.AppId must be the application (client) ID (GUID) in Certificate mode.') }
        if ($settings.Authentication.CertificateThumbprint -notmatch '^[0-9a-fA-F]{40}$') { $errors.Add('Authentication.CertificateThumbprint must be a 40-character thumbprint in Certificate mode.') }
        if ($settings.Tenant.Organization -notmatch '\.onmicrosoft\.com$') { $errors.Add('Tenant.Organization must be the initial domain (xxx.onmicrosoft.com) in Certificate mode.') }
    }
    if ($settings.Collection.RetentionWarningDays -ge $settings.Collection.SourceRetentionDays) { $errors.Add('Collection.RetentionWarningDays must be lower than Collection.SourceRetentionDays.') }
    if ($settings.Report.DefaultRange -notin 'Last24Hours', 'Last7Days', 'Last30Days', 'PreviousMonth') { $errors.Add('Report.DefaultRange must be Last24Hours, Last7Days, Last30Days or PreviousMonth.') }
    if ($settings.Report.SplitBy -notin 'Rows', 'Day', 'Week') { $errors.Add("Report.SplitBy must be 'Rows', 'Day' or 'Week'.") }
    if (-not $settings.Report.Formats.Count -or @($settings.Report.Formats | Where-Object { $_ -notin 'Csv', 'Html' }).Count) { $errors.Add("Report.Formats must contain 'Csv', 'Html' or both.") }
    if ($settings.Report.CsvDelimiter -notin ';', ',', "`t", '|') { $errors.Add("Report.CsvDelimiter must be ';', ',', '|' or a tab.") }
    if (-not $settings.Report.FilePrefix -or $settings.Report.FilePrefix.IndexOfAny([IO.Path]::GetInvalidFileNameChars()) -ge 0) { $errors.Add('Report.FilePrefix must be a valid file name part.') }
    try { $settings['Zone'] = Get-DlpTimeZone $settings.Report.TimeZone } catch { $errors.Add($_.Exception.Message) }
    if ($errors.Count) { throw ("Invalid configuration ($Path):`n - " + ($errors -join "`n - ")) }
    return $settings
}

#endregion

#region 3. Engine (SQLite + compiled C#) ------------------------------------------------------

function Initialize-DlpEngine {
    <#
    .SYNOPSIS
        Loads SQLite (lib\sqlite) and the C# engine. The engine is compiled from
        src\PurviewDlpReport.Engine.cs into bin\ the first time, and again only when
        the source file changes (the file name contains a hash of the source).
    #>
    [CmdletBinding()]
    param([string]$Root = $script:ToolRoot)
    if ('PurviewDlpReport.DlpStore' -as [type]) { return }
    $lib = Join-Path $Root 'lib\sqlite'
    $arch = if ([Runtime.InteropServices.RuntimeInformation]::ProcessArchitecture -eq 'Arm64') { 'win-arm64' } else { 'win-x64' }
    $native = Join-Path $lib "runtimes\$arch\e_sqlite3.dll"
    if (-not (Test-Path -LiteralPath $native)) { throw "SQLite native library not found: $native" }
    [void][Runtime.InteropServices.NativeLibrary]::Load($native)
    foreach ($name in 'SQLitePCLRaw.core', 'SQLitePCLRaw.provider.e_sqlite3', 'SQLitePCLRaw.batteries_v2', 'Microsoft.Data.Sqlite') {
        Add-Type -LiteralPath (Join-Path $lib "$name.dll")
    }
    [SQLitePCL.Batteries_V2]::Init()

    $source = Join-Path $Root 'src\PurviewDlpReport.Engine.cs'
    $hash = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash.Substring(0, 16)
    $bin = Join-Path $Root 'bin'
    $dll = Join-Path $bin "PurviewDlpReport.Engine.$hash.dll"
    if (-not (Test-Path -LiteralPath $dll)) {
        [void][IO.Directory]::CreateDirectory($bin)
        $references = @(
            (Join-Path $lib 'Microsoft.Data.Sqlite.dll'), 'System.IO.Compression', 'System.Text.Json', 'System.Text.Encodings.Web',
            'System.Data.Common', 'System.Linq', 'System.Collections', 'System.Text.RegularExpressions', 'System.Security.Cryptography',
            'System.Runtime', 'System.Memory', 'System.ComponentModel.Primitives', 'System.ComponentModel', 'System.Transactions.Local',
            'System.Text.Encoding.Extensions', 'System.Runtime.Extensions', 'System.IO', 'System.Runtime.InteropServices', 'System.Console', 'netstandard')
        $staging = Join-Path $bin ("compile-{0}.dll" -f [guid]::NewGuid().ToString('N'))
        Add-Type -LiteralPath $source -ReferencedAssemblies $references -OutputAssembly $staging -OutputType Library -IgnoreWarnings -WarningAction SilentlyContinue
        Move-Item -LiteralPath $staging -Destination $dll -Force
        # Older compiled versions are removed when they are not in use.
        Get-ChildItem -LiteralPath $bin -Filter 'PurviewDlpReport.Engine.*.dll' | Where-Object FullName -ne $dll |
            Remove-Item -Force -ErrorAction SilentlyContinue
    }
    if (-not ('PurviewDlpReport.DlpStore' -as [type])) { Add-Type -LiteralPath $dll }
}

function Open-DlpStore {
    param([Parameter(Mandatory)]$Settings, [switch]$ReadOnly)
    $path = $Settings.Storage.DatabasePath
    if ($ReadOnly -and -not (Test-Path -LiteralPath $path)) { throw "The database does not exist yet: $path. Run a collection first (-Mode Collect)." }
    $store = [PurviewDlpReport.DlpStore]::new($path, [bool]$ReadOnly, $script:ToolVersion)
    $store.KeepQuarantinedRaw = [bool]$Settings.Storage.KeepQuarantinedRecords
    return $store
}

#endregion

#region 4. Periods and coverage ------------------------------------------------------------

function Get-DlpTimeZone {
    <# Accepts an IANA name (Europe/Paris) or a Windows name (Romance Standard Time). #>
    param([Parameter(Mandatory)][string]$Id)
    try { return [TimeZoneInfo]::FindSystemTimeZoneById($Id) } catch { }
    $windowsId = $null
    if ([TimeZoneInfo]::TryConvertIanaIdToWindowsId($Id, [ref]$windowsId)) { try { return [TimeZoneInfo]::FindSystemTimeZoneById($windowsId) } catch { } }
    $known = @{ 'Europe/Paris' = 'Romance Standard Time'; 'Europe/Brussels' = 'Romance Standard Time'; 'Europe/London' = 'GMT Standard Time'; 'UTC' = 'UTC' }
    if ($known.ContainsKey($Id)) { return [TimeZoneInfo]::FindSystemTimeZoneById($known[$Id]) }
    throw "Unknown time zone '$Id' (Report.TimeZone). Use a name such as 'Europe/Paris' or 'Romance Standard Time'."
}

function ConvertTo-DlpUnixMs {
    <# Text date -> Unix ms. Without an explicit offset (Z, +02:00) the date is read in the report time zone. #>
    param([Parameter(Mandatory)][string]$Text, [Parameter(Mandatory)][TimeZoneInfo]$Zone)
    $culture = [Globalization.CultureInfo]::InvariantCulture
    if ($Text -match '(Z|[+-]\d{2}:?\d{2})$') {
        return [DateTimeOffset]::Parse($Text, $culture).ToUnixTimeMilliseconds()
    }
    $formats = [string[]]@('yyyy-MM-dd', 'yyyy-MM-dd HH:mm', 'yyyy-MM-ddTHH:mm', 'yyyy-MM-dd HH:mm:ss', 'yyyy-MM-ddTHH:mm:ss')
    $local = [datetime]::MinValue
    if (-not [datetime]::TryParseExact($Text, $formats, $culture, [Globalization.DateTimeStyles]::None, [ref]$local)) {
        throw "Invalid date '$Text'. Use yyyy-MM-dd, 'yyyy-MM-dd HH:mm' or an ISO 8601 value with an offset."
    }
    return [PurviewDlpReport.Coverage]::LocalToUnixMs($local, $Zone)
}

function Resolve-DlpPeriod {
    <#
    .SYNOPSIS
        Converts a range name into a [start, end) period in Unix milliseconds.
    .DESCRIPTION
        Last24Hours / Last7Days / Last30Days : rolling windows ending now.
        PreviousMonth : the previous calendar month in the report time zone.
        Month (-Month yyyy-MM), Day (-Date yyyy-MM-dd), Custom (-Start / -End).
        The end is never later than now.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateSet('Last24Hours', 'Last7Days', 'Last30Days', 'PreviousMonth', 'Month', 'Day', 'Custom')][string]$Range,
        [string]$Month, [string]$Date, [string]$Start, [string]$End,
        [Parameter(Mandatory)][TimeZoneInfo]$Zone,
        [DateTimeOffset]$Now = [DateTimeOffset]::UtcNow
    )
    $culture = [Globalization.CultureInfo]::InvariantCulture
    $nowMs = $Now.ToUnixTimeMilliseconds(); $nowMs -= $nowMs % 1000
    $day = 86400000L
    switch ($Range) {
        'Last24Hours' { $s = $nowMs - $day; $e = $nowMs }
        'Last7Days' { $s = $nowMs - 7 * $day; $e = $nowMs }
        'Last30Days' { $s = $nowMs - 30 * $day; $e = $nowMs }
        'PreviousMonth' {
            $local = [TimeZoneInfo]::ConvertTime($Now, $Zone).DateTime
            $first = [datetime]::new($local.Year, $local.Month, 1).AddMonths(-1)
            $s = [PurviewDlpReport.Coverage]::LocalToUnixMs($first, $Zone)
            $e = [PurviewDlpReport.Coverage]::LocalToUnixMs($first.AddMonths(1), $Zone)
        }
        'Month' {
            $first = [datetime]::MinValue
            if (-not $Month -or -not [datetime]::TryParseExact($Month, 'yyyy-MM', $culture, 'None', [ref]$first)) { throw "-Range Month requires -Month in the format yyyy-MM (for example -Month 2026-08)." }
            $s = [PurviewDlpReport.Coverage]::LocalToUnixMs($first, $Zone)
            $e = [PurviewDlpReport.Coverage]::LocalToUnixMs($first.AddMonths(1), $Zone)
        }
        'Day' {
            $d = [datetime]::MinValue
            if (-not $Date -or -not [datetime]::TryParseExact($Date, 'yyyy-MM-dd', $culture, 'None', [ref]$d)) { throw "-Range Day requires -Date in the format yyyy-MM-dd." }
            $s = [PurviewDlpReport.Coverage]::LocalToUnixMs($d, $Zone)
            $e = [PurviewDlpReport.Coverage]::LocalToUnixMs($d.AddDays(1), $Zone)
        }
        'Custom' {
            if (-not $Start -or -not $End) { throw '-Range Custom requires -Start and -End.' }
            $s = ConvertTo-DlpUnixMs $Start $Zone
            $e = ConvertTo-DlpUnixMs $End $Zone
        }
    }
    if ($e -gt $nowMs) { $e = $nowMs }
    if ($e -le $s) { throw "The requested period is empty or in the future ($Range)." }
    [pscustomobject]@{
        Range   = $Range
        StartMs = [long]$s
        EndMs   = [long]$e
        Label   = [PurviewDlpReport.ReportPlanner]::PeriodLabel($s, $e, $Zone)
    }
}

function Get-DlpRangeLength {
    <# Total length in milliseconds of a list of ranges (each with Start and End). #>
    param([AllowNull()][AllowEmptyCollection()]$Ranges)
    $sum = 0L
    foreach ($r in @($Ranges)) { if ($null -ne $r) { $sum += [long]$r.End - [long]$r.Start } }
    return $sum
}

function New-DlpRangeList {
    param([Parameter()][AllowEmptyCollection()][object[]]$Ranges)
    $list = [Collections.Generic.List[PurviewDlpReport.TimeRange]]::new()
    foreach ($r in $Ranges) { if ($null -ne $r) { $list.Add([PurviewDlpReport.TimeRange]::new([long]$r.Start, [long]$r.End)) } }
    return , $list
}

function Get-DlpCollectionPlan {
    <#
    .SYNOPSIS
        Decides what must be collected for a period.
    .DESCRIPTION
        - Ranges already collected and "settled" (collected at least SettlingHours after
          their end) are skipped.
        - Missing ranges younger than the Activity Explorer retention are collected,
          cut into slices that never cross local midnight.
        - Missing ranges older than the retention are reported as unrecoverable.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Store, [Parameter(Mandatory)][long]$TargetId,
        [Parameter(Mandatory)][long]$StartMs, [Parameter(Mandatory)][long]$EndMs,
        [Parameter(Mandatory)]$Settings, [long]$NowMs = [DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds()
    )
    $c = $Settings.Collection
    $settled = $Store.GetCoveredRanges($TargetId, [long]$c.SettlingHours * 3600000)
    $missing = [PurviewDlpReport.Coverage]::Gaps($settled, $StartMs, $EndMs)
    # One hour of margin: the oldest hour may already be purged by the time it is requested.
    $earliest = $NowMs - [long]$c.SourceRetentionDays * 86400000 + 3600000
    $collectable = [Collections.Generic.List[PurviewDlpReport.TimeRange]]::new()
    $lost = [Collections.Generic.List[PurviewDlpReport.TimeRange]]::new()
    foreach ($gap in $missing) {
        if ($gap.End -le $earliest) { $lost.Add($gap); continue }
        if ($gap.Start -lt $earliest) {
            $lost.Add([PurviewDlpReport.TimeRange]::new($gap.Start, $earliest))
            $collectable.Add([PurviewDlpReport.TimeRange]::new($earliest, $gap.End))
        } else { $collectable.Add($gap) }
    }
    $slices = [PurviewDlpReport.Coverage]::SplitIntoSlices($collectable, $Settings.Zone, $c.SliceHours)
    # Part of the collectable time never collected before (the rest is a refresh of recent data).
    $coveredAny = $Store.GetCoveredRanges($TargetId, 0)
    $neverCollected = 0L
    foreach ($range in $collectable) { $neverCollected += Get-DlpRangeLength ([PurviewDlpReport.Coverage]::Gaps($coveredAny, $range.Start, $range.End)) }
    [pscustomobject]@{
        Missing          = $missing
        Collectable      = $collectable
        Unrecoverable    = $lost
        Slices           = $slices
        CollectableMs    = Get-DlpRangeLength $collectable
        NewMs            = $neverCollected
        RefreshMs        = (Get-DlpRangeLength $collectable) - $neverCollected
        UnrecoverableMs  = Get-DlpRangeLength $lost
        EarliestSourceMs = $earliest
    }
}

function Get-DlpCoverage {
    <# Share of a period present in the database (any completed collection, settled or not). #>
    param([Parameter(Mandatory)]$Store, [Parameter(Mandatory)][long]$TargetId, [Parameter(Mandatory)][long]$StartMs, [Parameter(Mandatory)][long]$EndMs)
    $covered = $Store.GetCoveredRanges($TargetId, 0)
    $gaps = [PurviewDlpReport.Coverage]::Gaps($covered, $StartMs, $EndMs)
    $missingMs = Get-DlpRangeLength $gaps
    [pscustomobject]@{
        Percent   = if ($EndMs -gt $StartMs) { [Math]::Round(100.0 * ($EndMs - $StartMs - $missingMs) / ($EndMs - $StartMs), 2) } else { 100 }
        Gaps      = $gaps
        MissingMs = $missingMs
    }
}

function Get-DlpExpiringGaps {
    <# Missing ranges that Activity Explorer will no longer return within RetentionWarningDays. #>
    param([Parameter(Mandatory)]$Store, [Parameter(Mandatory)][long]$TargetId, [Parameter(Mandatory)]$Settings, [long]$NowMs = [DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds())
    $c = $Settings.Collection
    $stats = $Store.GetStatistics($TargetId)
    if ($null -eq $stats.FirstCoveredMs) { return @() }
    $windowStart = [Math]::Max([long]$stats.FirstCoveredMs, $NowMs - [long]$c.SourceRetentionDays * 86400000)
    $gaps = [PurviewDlpReport.Coverage]::Gaps($Store.GetCoveredRanges($TargetId, 0), $windowStart, $NowMs - [long]$c.SettlingHours * 3600000)
    $limit = $NowMs - [long]($c.SourceRetentionDays - $c.RetentionWarningDays) * 86400000
    foreach ($gap in $gaps) {
        if ($gap.Start -lt $limit) {
            [pscustomobject]@{ Start = $gap.Start; End = $gap.End; DaysLeft = [Math]::Max(0, [Math]::Floor(($gap.Start - ($NowMs - [long]$c.SourceRetentionDays * 86400000)) / 86400000.0)) }
        }
    }
}

#endregion

#region 5. Connection ------------------------------------------------------------------------

# ExchangeOnlineManagement versions that must not be used, by authentication mode (checked in the lab).
$script:ExoKnownIssues = @{
    '3.10.0' = @{ Modes = @('Certificate'); Issue = 'its certificate authentication fails with "Object reference not set to an instance of an object" (fixed in 3.10.1)' }
}

function Select-DlpExoModule {
    <#
    .SYNOPSIS
        Chooses the ExchangeOnlineManagement module to use: the version already loaded in the session,
        otherwise the most recent installed version, 3.9.0 or later, without a known issue for the mode.
        Returns the module and the versions skipped; throws a message with the command to run otherwise.
    #>
    param([AllowEmptyCollection()][object[]]$Available, [AllowNull()]$Loaded, [Parameter(Mandatory)][string]$Mode)
    $install = 'Install-Module ExchangeOnlineManagement -MinimumVersion 3.10.1 -Scope CurrentUser -Force'
    $problem = {
        param($Version)
        $known = $script:ExoKnownIssues[$Version.ToString()]
        if ($Version -lt [version]'3.9.0') { return 'it is older than 3.9.0' }
        if ($known -and $Mode -in $known.Modes) { return $known.Issue }
    }
    if ($Loaded) {
        $why = & $problem $Loaded.Version
        if ($why) { throw "ExchangeOnlineManagement $($Loaded.Version) is already loaded in this PowerShell session and cannot be used: $why. Open a new PowerShell window. If this version is loaded automatically, run: $install" }
        return [pscustomobject]@{ Module = $Loaded; Skipped = @() }
    }
    $sorted = @($Available | Where-Object { $_ } | Sort-Object Version -Descending)
    if (-not $sorted.Count) { throw "The ExchangeOnlineManagement module is not installed. Run: $install" }
    $skipped = [Collections.Generic.List[string]]::new()
    foreach ($m in $sorted) {
        $why = & $problem $m.Version
        if (-not $why) { return [pscustomobject]@{ Module = $m; Skipped = @($skipped) } }
        if (-not $skipped.Contains("$($m.Version): $why")) { $skipped.Add("$($m.Version): $why") }
    }
    throw "No usable ExchangeOnlineManagement version is installed (3.9.2 and 3.10.1 validated). $(@($skipped) -join '; '). Run: $install"
}

function Connect-DlpActivityExplorer {
    <#
    .SYNOPSIS
        Connects to Security & Compliance PowerShell and returns the function used to call
        Export-ActivityExplorerData. The tenant (and the account in interactive mode) is checked.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Settings)
    $auth = $Settings.Authentication
    $choice = Select-DlpExoModule -Available @(Get-Module ExchangeOnlineManagement -ListAvailable) -Loaded (Get-Module ExchangeOnlineManagement | Select-Object -First 1) -Mode $auth.Mode
    foreach ($s in $choice.Skipped) { Write-DlpItem Info "ExchangeOnlineManagement $s - skipped, $($choice.Module.Version) used." }
    $module = $choice.Module
    if (-not (Get-Module ExchangeOnlineManagement)) { Import-Module $module -Global -WarningAction SilentlyContinue }
    Write-DlpLog 'INFO' "ExchangeOnlineManagement $($module.Version) ($($module.ModuleBase))"

    $prefix = $script:CommandPrefix
    $existing = @(Get-ConnectionInformation -ErrorAction SilentlyContinue | Where-Object { $_.ModulePrefix -eq $prefix -and $_.State -eq 'Connected' })
    $connectedHere = $false
    if (-not $existing) {
        $parameters = @{ Prefix = $prefix; ShowBanner = $false; ErrorAction = 'Stop' }
        if ($auth.Mode -eq 'Certificate') {
            $parameters.AppId = $auth.AppId
            $parameters.CertificateThumbprint = $auth.CertificateThumbprint
            $parameters.Organization = $Settings.Tenant.Organization
        } else {
            if ($auth.UserPrincipalName) { $parameters.UserPrincipalName = $auth.UserPrincipalName }
            if ($auth.DisableWAM -and (Get-Command Connect-IPPSSession).Parameters.ContainsKey('DisableWAM')) { $parameters.DisableWAM = $true }
        }
        Connect-IPPSSession @parameters | Out-Null
        $connectedHere = $true
    }
    $connection = @(Get-ConnectionInformation | Where-Object { $_.ModulePrefix -eq $prefix -and $_.State -eq 'Connected' }) | Select-Object -Last 1
    if (-not $connection) { throw 'The Security & Compliance connection was not established.' }
    $account = if ($auth.Mode -eq 'Certificate') { "app $($auth.AppId)" } else { [string]$connection.UserPrincipalName }
    $problem = $null
    if ($Settings.Tenant.TenantId -and [string]$connection.TenantID -ne $Settings.Tenant.TenantId) { $problem = "Connected to tenant $($connection.TenantID), but the configuration expects $($Settings.Tenant.TenantId)." }
    elseif ($auth.Mode -eq 'Interactive' -and $auth.UserPrincipalName -and $account -ine $auth.UserPrincipalName) { $problem = "Signed in as $account, but the configuration expects $($auth.UserPrincipalName)." }
    if ($problem) {
        if ($connectedHere) { Disconnect-ExchangeOnline -ConnectionId $connection.ConnectionId -Confirm:$false -ErrorAction SilentlyContinue }
        throw $problem
    }
    $commandName = "Export-$($prefix)ActivityExplorerData"
    $command = Get-Command $commandName -ErrorAction SilentlyContinue
    if (-not $command) {
        throw "Export-ActivityExplorerData is not available for $account. The account (or application) needs a Purview role that can export Activity Explorer data - see the guide, chapter 4 (accounts) and Annex F (application)."
    }
    $exporter = { param([hashtable]$Query) & $command @Query -ErrorAction Stop }.GetNewClosure()
    [pscustomobject]@{
        Exporter      = $exporter
        Account       = $account
        TenantId      = [string]$connection.TenantID
        Mode          = $auth.Mode
        ConnectionId  = $connection.ConnectionId
        ConnectedHere = $connectedHere
    }
}

function Disconnect-DlpActivityExplorer {
    param($Connection)
    if ($Connection -and $Connection.ConnectedHere -and $Connection.ConnectionId) {
        Disconnect-ExchangeOnline -ConnectionId $Connection.ConnectionId -Confirm:$false -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null
    }
}

#endregion

#region 6. Collection ---------------------------------------------------------------------------

function Get-DlpProperty {
    param($Object, [string]$Name)
    if ($null -eq $Object) { return $null }
    if ($Object -is [Collections.IDictionary]) { if ($Object.Contains($Name)) { return $Object[$Name] }; return $null }
    $p = $Object.PSObject.Properties[$Name]
    if ($p) { return $p.Value }
    return $null
}

function Test-DlpPageEnvelope {
    <#
    .SYNOPSIS
        Checks one response of Export-ActivityExplorerData before its data is used.
    .NOTES
        Known false positive: the service fills ErrorData with the type name
        "Microsoft.Exchange.Hygiene.DataInsights.Common.DataInsightsErrorData" even when
        ResultCode is Success. That value is accepted only together with ResultCode = Success.
    #>
    param([Parameter(Mandatory)]$Response)
    $resultCode = [string](Get-DlpProperty $Response 'ResultCode')
    foreach ($field in 'Exception', 'ErrorData') {
        $value = Get-DlpProperty $Response $field
        if ($null -eq $value -or ($value -is [bool] -and -not $value)) { continue }
        $text = ([string]$value).Trim()
        if ($text -in '', '[]', '{}', 'null') { continue }
        if ($field -eq 'ErrorData' -and $text -eq 'Microsoft.Exchange.Hygiene.DataInsights.Common.DataInsightsErrorData' -and $resultCode -eq 'Success') { continue }
        throw "SourceEnvelopeError: $field = $text (ResultCode = $resultCode)."
    }
    $lastValue = Get-DlpProperty $Response 'LastPage'
    if ($lastValue -is [bool]) { $last = $lastValue }
    elseif ("$lastValue" -match '^(?i:true|false)$') { $last = [bool]::Parse("$lastValue") }
    else { throw 'MalformedPage: LastPage is missing or is not a boolean.' }
    $data = Get-DlpProperty $Response 'ResultData'
    # Only .Length is used on ResultData here: passing a large string to a .NET method is slow (AMSI scan).
    if ($data -isnot [string] -or $data.Length -eq 0) { throw 'MalformedPage: ResultData is missing (an empty page must be the JSON text []).' }
    $recordCount = $null
    $reported = Get-DlpProperty $Response 'RecordCount'
    if ($null -ne $reported) {
        $n = 0L
        if (-not [long]::TryParse("$reported", [ref]$n) -or $n -lt 0) { throw 'MalformedPage: RecordCount is not a number.' }
        $recordCount = $n
    }
    $total = $null
    $reportedTotal = Get-DlpProperty $Response 'TotalResultCount'
    if ($null -ne $reportedTotal) { $n = 0L; if ([long]::TryParse("$reportedTotal", [ref]$n)) { $total = $n } }
    [pscustomobject]@{
        LastPage         = $last
        Watermark        = [string](Get-DlpProperty $Response 'Watermark')
        ResultData       = $data
        RecordCount      = $recordCount
        TotalResultCount = $total
        ResultCode       = $resultCode
    }
}

function Invoke-DlpCollection {
    <#
    .SYNOPSIS
        Collects Activity Explorer events for a list of time slices and stores them in SQLite.
    .DESCRIPTION
        For each slice, pages are requested with the same filters (Exchange, DLPRuleMatch,
        policy, rule) until LastPage = true. Every page is stored immediately, so an
        interrupted execution never loses what was already received; the slice is marked
        "Completed" only when its last page has been stored.

        On a recoverable error the slice is retried from its beginning (the paging cookie
        cannot be reused after 120 s), then split in two halves. A permanent error
        (permissions, command not found) stops the collection.
    .OUTPUTS
        An object with the counters and the final state (Completed / Error).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Store,
        [Parameter(Mandatory)][long]$RunId,
        [Parameter(Mandatory)][long]$TargetId,
        [Parameter(Mandatory)]$Settings,
        [Parameter(Mandatory)][AllowEmptyCollection()]$Slices,
        [Parameter(Mandatory)][scriptblock]$Exporter
    )
    $c = $Settings.Collection
    $zone = $Settings.Zone
    $stats = [ordered]@{
        Completed = $false; Error = $null; Slices = @($Slices).Count; SlicesDone = 0; Pages = 0; Records = 0; NewEvents = 0
        Duplicates = 0; Quarantined = 0; Conflicts = 0; WithoutCount = 0; Restarts = 0; Subdivisions = 0
        SourceSeconds = 0.0; EngineSeconds = 0.0; Seconds = 0.0; ServiceAnnounced = 0; QuarantineReasons = @{}
    }
    $pending = [Collections.Generic.LinkedList[object]]::new()
    foreach ($slice in $Slices) { [void]$pending.AddLast([pscustomobject]@{ Start = [long]$slice.Start; End = [long]$slice.End; Restart = 0 }) }
    $totalSpan = [double](Get-DlpRangeLength $Slices)
    $doneSpan = 0.0
    $clock = [Diagnostics.Stopwatch]::StartNew()
    $pageData = [PurviewDlpReport.PageData]::new()
    $permanent = '(?i)(not recognized|CommandNotFound|ParameterBinding|access.{0,20}denied|forbidden|\b403\b|not authori[sz]ed|insufficient|does not have permission)'
    $index = 0
    if ($pending.Count) { Write-DlpTableRow -Header }
    try {
        while ($pending.Count) {
            $slice = $pending.First.Value; $pending.RemoveFirst()
            $index++
            $label = Format-DlpRange $slice.Start $slice.End $zone
            $intervalId = $Store.BeginInterval($RunId, $TargetId, $slice.Start, $slice.End)
            $pages = 0; $records = 0L; $new = 0L; $serviceTotal = $null; $cookie = $null
            $cookies = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
            $hashes = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
            $sinceLastPage = $null
            $sliceClock = [Diagnostics.Stopwatch]::StartNew()
            Write-DlpLog 'INFO' "Slice $label started (interval $intervalId, attempt $($slice.Restart + 1))"
            try {
                while ($true) {
                    if ($pages -ge $c.MaxPagesPerSlice) { throw "SlicePageBudgetExceeded: more than $($c.MaxPagesPerSlice) pages for one window." }
                    if ($sinceLastPage -and $sinceLastPage.Elapsed.TotalSeconds -ge $c.CookieSafetySeconds) { throw 'PageCookieExpiredLocalSafetyBudget: the next page could not be requested within the cookie lifetime.' }
                    $query = @{
                        StartTime    = [DateTimeOffset]::FromUnixTimeMilliseconds($slice.Start).UtcDateTime
                        EndTime      = [DateTimeOffset]::FromUnixTimeMilliseconds($slice.End).UtcDateTime
                        OutputFormat = 'Json'
                        PageSize     = $c.PageSize
                        Filter1      = @('Workload', 'Exchange')
                        Filter2      = @('Activity', 'DLPRuleMatch')
                        Filter3      = @('DLPPolicyId', $Settings.Target.PolicyId)
                        Filter4      = @('DLPPolicyRuleId', $Settings.Target.RuleId)
                    }
                    if ($cookie) { $query.PageCookie = $cookie }
                    $call = [Diagnostics.Stopwatch]::StartNew()
                    try { $responses = @(& $Exporter $query) }
                    finally { $stats.SourceSeconds += $call.Elapsed.TotalSeconds }
                    $sinceLastPage = [Diagnostics.Stopwatch]::StartNew()
                    if ($responses.Count -ne 1) { throw "MalformedPage: one response expected, $($responses.Count) received." }
                    $page = Test-DlpPageEnvelope -Response $responses[0]
                    $pageData.ResultData = $page.ResultData   # property assignment, see PageData in the engine
                    $result = $Store.IngestPage($intervalId, $TargetId, $Settings.Target.PolicyId, $Settings.Target.RuleId, $pageData)
                    $pages++
                    $stats.EngineSeconds += ($result.ParseMs + $result.NormalizeMs + $result.WriteMs + $result.CommitMs) / 1000
                    if ($result.Records -gt $c.PageSize) { throw 'MalformedPage: more records than the requested page size.' }
                    if ($null -ne $page.RecordCount -and $page.RecordCount -ne $result.Records) { throw "MalformedPage: RecordCount $($page.RecordCount) differs from the $($result.Records) records received." }
                    if ($result.Records -eq 0 -and -not $page.LastPage) { throw 'EmptyIntermediatePage: an empty page that is not the last one cannot be trusted.' }
                    if ($result.Records -gt 0 -and -not $hashes.Add($result.Sha256)) { throw 'RepeatedPage: the service returned the same page twice.' }
                    $records += $result.Records; $new += $result.NewEvents
                    $stats.Pages++; $stats.Records += $result.Records; $stats.NewEvents += $result.NewEvents; $stats.Duplicates += $result.Duplicates
                    $stats.Quarantined += $result.Quarantined; $stats.Conflicts += $result.Conflicts; $stats.WithoutCount += $result.WithoutCount
                    foreach ($reason in $result.QuarantineReasons.Keys) { $stats.QuarantineReasons[$reason] = [int]$stats.QuarantineReasons[$reason] + $result.QuarantineReasons[$reason] }
                    if ($null -eq $serviceTotal -and $null -ne $page.TotalResultCount) { $serviceTotal = [long]$page.TotalResultCount }

                    # Progress: share of the total time span, refined with the announced number of events.
                    $fraction = if ($serviceTotal -gt 0) { [Math]::Min(1.0, $records / $serviceTotal) } else { 0.0 }
                    $percent = if ($totalSpan -gt 0) { [Math]::Min(100.0, 100.0 * ($doneSpan + $fraction * ($slice.End - $slice.Start)) / $totalSpan) } else { 100 }
                    $rate = if ($clock.Elapsed.TotalSeconds -gt 0) { $stats.Records / $clock.Elapsed.TotalSeconds } else { 0 }
                    $eta = if ($percent -gt 1) { Format-DlpDuration ($clock.Elapsed.TotalSeconds * (100 - $percent) / $percent) } else { 'estimating...' }
                    Write-Progress -Id 1 -Activity 'Collecting Activity Explorer data' -PercentComplete $percent `
                        -Status ("Window {0}/{1}  {2}  |  {3} / {4} events  |  {5:0} events/s  |  remaining {6}" -f $index, ($index + $pending.Count), $label, (Format-DlpNumber $records), (Format-DlpNumber $serviceTotal), $rate, $eta)

                    if ($page.LastPage) { break }
                    if (-not $page.Watermark) { throw 'MissingCookie: LastPage is false but no Watermark was returned.' }
                    if (-not $cookies.Add($page.Watermark)) { throw 'RepeatedCookie: the paging cookie did not change.' }
                    $cookie = $page.Watermark
                }
                $Store.EndInterval($intervalId, $true, $pages, $records, $new, $serviceTotal, $null)
                $doneSpan += ($slice.End - $slice.Start)
                $stats.SlicesDone++
                if ($serviceTotal) { $stats.ServiceAnnounced += $serviceTotal }
                $sliceRate = if ($sliceClock.Elapsed.TotalSeconds -gt 0) { $records / $sliceClock.Elapsed.TotalSeconds } else { 0 }
                Write-DlpTableRow -Status Ok -Window $label -Events $records -New $new -Pages $pages -Duration (Format-DlpDuration $sliceClock.Elapsed.TotalSeconds) -Rate ("{0:0} ev/s" -f $sliceRate)
                if ($null -ne $serviceTotal -and $serviceTotal -ne $records) { Write-DlpLog 'INFO' "Slice $label : service announced $serviceTotal events, $records received (the announced total is known to be approximate)." }
            } catch {
                $message = $_.Exception.Message
                $Store.EndInterval($intervalId, $false, $pages, $records, $new, $serviceTotal, $message)
                Write-DlpLog 'WARN' "Slice $label failed after $pages page(s): $message"
                $length = $slice.End - $slice.Start
                if ($message -match $permanent) { throw "$label : $message" }
                if ($slice.Restart -lt $c.MaxSliceRestarts) {
                    $stats.Restarts++
                    $delay = if ($message -match '(?i)(429|throttl)') { 60 } else { 5 * [Math]::Pow(2, $slice.Restart) }
                    Write-DlpItem Warn ("{0}   interrupted ({1}) - retry {2}/{3} in {4} s" -f $label, (($message -split "`n")[0]), ($slice.Restart + 1), $c.MaxSliceRestarts, $delay)
                    Start-Sleep -Seconds $delay
                    [void]$pending.AddFirst([pscustomobject]@{ Start = $slice.Start; End = $slice.End; Restart = $slice.Restart + 1 })
                } elseif ($length -ge 2 * [long]$c.MinimumSliceMinutes * 60000) {
                    $stats.Subdivisions++
                    $middle = $slice.Start + [long][Math]::Floor($length / 2)
                    Write-DlpItem Warn ("{0}   still failing - split into two smaller windows" -f $label)
                    [void]$pending.AddFirst([pscustomobject]@{ Start = $middle; End = $slice.End; Restart = 0 })
                    [void]$pending.AddFirst([pscustomobject]@{ Start = $slice.Start; End = $middle; Restart = 0 })
                } else { throw "$label : $message" }
            }
        }
        $stats.Completed = $true
    } catch {
        $stats.Error = $_.Exception.Message
    } finally {
        Write-Progress -Id 1 -Activity 'Collecting Activity Explorer data' -Completed
        $stats.Seconds = $clock.Elapsed.TotalSeconds
    }
    return [pscustomobject]$stats
}

#endregion

#region 7. Report --------------------------------------------------------------------------------

function Move-DlpDirectory {
    <# Directory move with retries: antivirus or indexing can briefly lock files that were just written. #>
    param([Parameter(Mandatory)][string]$Source, [Parameter(Mandatory)][string]$Destination, [int]$Attempts = 10)
    for ($i = 1; $i -le $Attempts; $i++) {
        try { [IO.Directory]::Move($Source, $Destination); return }
        catch {
            $e = $_.Exception; while ($e.InnerException) { $e = $e.InnerException }
            $code = $e.HResult -band 0xFFFF
            $transient = $e -is [UnauthorizedAccessException] -or $code -in 5, 32, 33
            if (-not $transient -or $i -eq $Attempts) { throw }
            Start-Sleep -Milliseconds (250 * $i)
        }
    }
}

function New-DlpReport {
    <#
    .SYNOPSIS
        Writes the CSV and HTML files of a period from the database.
    .DESCRIPTION
        Files are written in a "<folder>.pending" directory, which is renamed only when
        every file is complete: a folder without ".pending" is always a finished report.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Store, [Parameter(Mandatory)][long]$TargetId, [Parameter(Mandatory)]$Settings,
        [Parameter(Mandatory)]$Period, [Parameter(Mandatory)]$Coverage
    )
    $r = $Settings.Report
    $zone = $Settings.Zone
    [void][IO.Directory]::CreateDirectory($r.OutputPath)
    $name = '{0:yyyy-MM-dd_HHmmss}_{1}' -f (Get-Date), $Period.Range
    $final = Join-Path $r.OutputPath $name
    for ($n = 2; (Test-Path -LiteralPath $final) -or (Test-Path -LiteralPath "$final.pending"); $n++) { $final = Join-Path $r.OutputPath "${name}_$n" }
    $staging = "$final.pending"
    $request = [PurviewDlpReport.ReportRequest]::new()
    $request.StartMs = $Period.StartMs
    $request.EndMs = $Period.EndMs
    $request.Zone = $zone
    $request.TimeZoneLabel = $r.TimeZone
    $request.SplitBy = $r.SplitBy
    $request.MaxRowsPerFile = $r.MaxRowsPerFile
    $request.IncludeRecipientDetails = [bool]$r.IncludeRecipientDetails
    $request.WriteCsv = 'Csv' -in $r.Formats
    $request.WriteHtml = 'Html' -in $r.Formats
    $request.CsvDelimiter = $r.CsvDelimiter
    $request.OutputDirectory = $staging
    $request.FilePrefix = $r.FilePrefix
    $request.HtmlTemplatePath = $r.TemplatePath
    $request.Title = $r.Title
    $request.PolicyName = $Settings.Target.PolicyName
    $request.RuleName = $Settings.Target.RuleName
    $request.ToolVersion = $script:ToolVersion
    $request.RangeName = $Period.Range
    $request.CoveragePercent = $Coverage.Percent
    if ($Coverage.Percent -lt 100) {
        $request.CoverageNote = 'no data collected for ' + ((@($Coverage.Gaps) | Select-Object -First 3 | ForEach-Object { Format-DlpRange $_.Start $_.End $zone }) -join ', ') + $(if (@($Coverage.Gaps).Count -gt 3) { ', ...' } else { '' })
    }
    $result = $Store.WriteReport($TargetId, $request)
    Move-DlpDirectory -Source $staging -Destination $final
    foreach ($file in $result.Files) { $file.Path = Join-Path $final (Split-Path $file.Path -Leaf) }
    [pscustomobject]@{ Directory = $final; Result = $result }
}

#endregion

#region 8. Status and maintenance ------------------------------------------------------------------

function Show-DlpStatus {
    <#
    .SYNOPSIS
        Prints the database content and, day by day, what has been collected:
        date, messages, a coverage bar and the state (complete, provisional, missing, lost).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Store, [Parameter(Mandatory)][long]$TargetId, [Parameter(Mandatory)]$Settings, [int]$Days = 35)
    $zone = $Settings.Zone; $cfg = $Settings.Collection; $K = $script:C; $dot = [char]0x00B7
    $now = [DateTimeOffset]::UtcNow
    $nowMs = $now.ToUnixTimeMilliseconds()
    $s = $Store.GetStatistics($TargetId)
    Write-DlpItem Info ('{0}  ({1})' -f $Store.DatabasePath, (Format-DlpBytes $s.FileBytes)) -Icon Database
    Write-DlpItem Info ('{0} events {4} {1} messages {4} {2} executions {4} {3} quarantined record(s)' -f (Format-DlpNumber $s.Events), (Format-DlpNumber $s.Messages), (Format-DlpNumber $s.Runs), (Format-DlpNumber $s.Quarantined), $dot) -Icon Chart
    if ($null -eq $s.FirstCoveredMs) { Write-DlpItem Warn 'Nothing has been collected yet. Run: .\Invoke-PurviewDlpReport.ps1 -Mode Collect'; return }
    Write-DlpItem Info ('History {0}   {1} last collection {2}' -f (Format-DlpRange $s.FirstCoveredMs $s.LastCoveredMs $zone), $dot, (Format-DlpLocalTime $s.LastSuccessfulCollectionMs $zone 'yyyy-MM-dd HH:mm:ss')) -Icon Calendar
    Write-Host ''
    $today = [TimeZoneInfo]::ConvertTime($now, $zone).DateTime.Date
    $first = $today.AddDays( - ($Days - 1))
    $firstCollected = [TimeZoneInfo]::ConvertTime([DateTimeOffset]::FromUnixTimeMilliseconds($s.FirstCoveredMs), $zone).DateTime.Date
    if ($firstCollected -gt $first) { $first = $firstCollected }
    $Days = [int]($today - $first).TotalDays + 1
    $rows = $Store.GetDailyStatistics($TargetId, $first, $Days, $zone, [long]$cfg.SettlingHours * 3600000)
    $sourceStart = $nowMs - [long]$cfg.SourceRetentionDays * 86400000
    $en = [Globalization.CultureInfo]::GetCultureInfo('en-US')
    Write-Host ('      {0}{1}{2,-15} {3,10}   {4,-16} {5}{6}' -f $K.Dim, ('  ' + $script:IconPad), 'Day', 'Messages', 'Collected', 'State', $K.Reset)
    foreach ($d in $rows) {
        if ($d.StartMs -ge $nowMs) { continue }
        $length = [Math]::Max(1, [Math]::Min($d.EndMs, $nowMs) - $d.StartMs)
        $percent = [int][Math]::Min(100, [Math]::Floor(100.0 * $d.CoveredMs / $length))
        # State, status icon and colour of the day.
        if ($d.LocalDate -eq $today) { $state = if ($percent -ge 100) { 'Today, provisional' } else { 'Today, collection in progress' }; $status = 'Info' }
        elseif ($percent -ge 100 -and $d.FinalCoveredMs -ge ($d.EndMs - $d.StartMs)) { $state = 'Complete'; $status = 'Ok' }
        elseif ($percent -ge 100) { $state = 'Complete, refreshed at next run'; $status = 'Ok' }
        elseif ($d.EndMs -le $sourceStart) { $state = if ($percent -gt 0) { 'Partial, rest no longer collectable' } else { 'Missing, no longer collectable' }; $status = 'Fail' }
        elseif ($d.StartMs -lt $sourceStart) { $state = "Partial, earlier hours beyond the $($cfg.SourceRetentionDays)-day retention"; $status = 'Info' }
        else {
            $left = [Math]::Max(0, [Math]::Floor(($d.StartMs - $sourceStart) / 86400000.0))
            $state = '{0}, collectable {1} more day(s)' -f $(if ($percent -gt 0) { 'Partial' } else { 'Missing' }), $left; $status = 'Warn'
        }
        $color = @{ Ok = $K.Green; Warn = $K.Yellow; Fail = $K.Red; Info = $K.Cyan }[$status]
        # Coverage bar: 12 cells.
        $filled = [int][Math]::Round(12 * $percent / 100.0)
        $bar = $color + [string]::new([char]0x2588, $filled) + $K.Dim + [string]::new([char]0x2591, 12 - $filled) + $K.Reset
        $day = '{0} {1}' -f $d.LocalDate.ToString('ddd', $en), $d.LocalDate.ToString('yyyy-MM-dd')
        $messages = if ($d.Messages) { $K.Bold + (Format-DlpNumber $d.Messages).PadLeft(10) + $K.Reset } else { $K.Dim + '0'.PadLeft(10) + $K.Reset }
        Write-Host ('      {0}{1}{2}{3,-15} {4}   {5} {6,4}%  {0}{7}{2}' -f $color, (Get-DlpIcon $status), $K.Reset, $day, $messages, $bar, $percent, $state)
        Write-DlpLog 'INFO' ("Status {0}: {1} messages, {2} events, {3}% collected, {4}" -f $d.LocalDate.ToString('yyyy-MM-dd'), $d.Messages, $d.Events, $percent, $state)
    }
}
function Invoke-DlpRetention {
    <# Deletes database content older than Storage.RetentionDays (0 = keep everything). #>
    param([Parameter(Mandatory)]$Store, [Parameter(Mandatory)]$Settings)
    if ($Settings.Storage.RetentionDays -le 0) { return $null }
    $cutoff = [DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds() - [long]$Settings.Storage.RetentionDays * 86400000
    return $Store.PurgeBefore($cutoff)
}

function Enter-DlpLock {
    <#
    .SYNOPSIS
        Prevents two executions from collecting at the same time (for example the scheduled
        task and an administrator). Returns the lock, to be released with Exit-DlpLock.
    #>
    param([Parameter(Mandatory)][string]$Path, [int]$TimeoutSeconds = 30)
    [void][IO.Directory]::CreateDirectory((Split-Path $Path -Parent))
    $deadline = [DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
    while ($true) {
        try {
            $stream = [IO.FileStream]::new($Path, [IO.FileMode]::OpenOrCreate, [IO.FileAccess]::ReadWrite, [IO.FileShare]::Read)
            $stream.SetLength(0)
            $bytes = [Text.Encoding]::UTF8.GetBytes(("{0} pid {1} since {2:o}" -f [Environment]::MachineName, $PID, (Get-Date)))
            $stream.Write($bytes, 0, $bytes.Length); $stream.Flush()
            return $stream
        } catch [IO.IOException] {
            if ([DateTime]::UtcNow -ge $deadline) {
                $owner = try { [IO.File]::ReadAllText($Path) } catch { 'unknown' }
                throw "Another execution is already collecting ($owner). Wait for it to finish, or use -NoCollect to build a report from the data already collected."
            }
            Start-Sleep -Seconds 2
        }
    }
}

function Exit-DlpLock {
    param($Lock)
    if ($Lock) { $Lock.Dispose() }
}

#endregion
