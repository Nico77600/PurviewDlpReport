#Requires -Version 7.4
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }
<#
    Purview DLP Report - automated tests (Pester 5 or later).
    Author  : Nicolas Fabert
    Version : 2.1.2

    Run:  Invoke-Pester -Path .\tests\PurviewDlpReport.Tests.ps1 -Output Detailed

    No connection to Microsoft 365 is made: Activity Explorer is replaced by a fake
    exporter that returns pages built from sample records.
#>

BeforeAll {
    $script:RepoRoot = Split-Path $PSScriptRoot -Parent
    $script:Root = Join-Path $script:RepoRoot 'package'
    Import-Module (Join-Path $script:Root 'PurviewDlpReport.psd1') -Force
    Initialize-DlpEngine -Root $script:Root
    # Fictitious identifiers: the tests never depend on a real tenant.
    $script:PolicyId = '11111111-2222-3333-4444-555555555555'
    $script:RuleId = '66666666-7777-8888-9999-000000000000'
    $script:Paris = Get-DlpTimeZone 'Europe/Paris'

    # Test configuration = the delivered configuration file with fictitious tenant values
    # (the delivered file may contain real values or empty ones).
    $script:ConfigDirectory = Join-Path $TestDrive 'configuration'
    [void][IO.Directory]::CreateDirectory($script:ConfigDirectory)
    $script:ConfigPath = Join-Path $script:ConfigDirectory 'test.config.psd1'
    $text = [IO.File]::ReadAllText((Join-Path $script:Root 'config\PurviewDlpReport.config.psd1'))
    $values = [ordered]@{
        TenantId = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'; Organization = 'contoso.onmicrosoft.com'
        PolicyId = $script:PolicyId; RuleId = $script:RuleId; PolicyName = 'Test policy'; RuleName = 'Test rule'
        UserPrincipalName = ''; AppId = ''; CertificateThumbprint = ''
    }
    foreach ($key in $values.Keys) {
        $pattern = "(?m)^(\s*$key\s*=\s*)'[^']*'"
        if ([regex]::Matches($text, $pattern).Count -ne 1) { throw "The key $key must appear once in the delivered configuration." }
        $text = [regex]::Replace($text, $pattern, "`${1}'$($values[$key])'")
    }
    [IO.File]::WriteAllText($script:ConfigPath, $text, [Text.UTF8Encoding]::new($true))

    function New-TestRecord {
        param(
            [string]$Id = [guid]::NewGuid().ToString(),
            [string]$MessageId = "<$([guid]::NewGuid().ToString('N'))@test.local>",
            [string]$Happened = '2026-09-28T10:00:00Z',
            [int]$Count = 30,
            [string]$Sender = 'alice@contoso.com',
            [string]$Subject = 'Quarterly update',
            [string[]]$Receivers = @('bob@contoso.com', 'carol@contoso.com'),
            [string]$RuleId = $script:RuleId,
            [switch]$EmailInfoAsString,
            [switch]$PolicyAsArray,
            [object]$ExplicitCount = $null
        )
        $email = [ordered]@{ Sender = $Sender; Receivers = $Receivers; Subject = $Subject; MessageID = $MessageId }
        if ($null -ne $ExplicitCount) { $email.RecipientCount = $ExplicitCount }
        $policy = [ordered]@{
            PolicyId = $script:PolicyId; PolicyName = 'Test policy'; RuleId = $RuleId; RuleName = 'Test rule'
            OtherConditions = @(@{ Condition = "RecipientCountOver $Count" }, @{ Condition = 'FromMemberOf grp@contoso.com' })
        }
        [ordered]@{
            RecordIdentity = $Id; Activity = 'DLP rule matched'; ActivityId = 'DLPRuleMatch'; Happened = $Happened; Workload = 'Exchange'
            PolicyMatchInfo = $(if ($PolicyAsArray) { , @($policy) } else { $policy })
            EmailInfo = $(if ($EmailInfoAsString) { $email | ConvertTo-Json -Compress -Depth 5 } else { $email })
        }
    }
    function ConvertTo-TestPage([object[]]$Records) {
        if (-not $Records -or $Records.Count -eq 0) { return '[]' }
        return ConvertTo-Json -InputObject @($Records) -Depth 10 -Compress
    }
    function Invoke-Normalize($Record, [string]$Rule = $script:RuleId) {
        $doc = [Text.Json.JsonDocument]::Parse(($Record | ConvertTo-Json -Depth 10 -Compress))
        try { return [PurviewDlpReport.EventNormalizer]::Normalize($doc.RootElement.Clone(), $script:PolicyId, $Rule) } finally { $doc.Dispose() }
    }
    function New-TestSettings([string]$Directory) {
        $settings = Import-DlpConfiguration -Path $script:ConfigPath -Root $script:Root
        $settings.Storage.DatabasePath = Join-Path $Directory 'test.sqlite'
        $settings.Report.OutputPath = Join-Path $Directory 'reports'
        $settings.Collection.PageSize = 2
        $settings.Collection.SettlingHours = 0
        return $settings
    }
    function New-TestPageData([string]$Json) { $p = [PurviewDlpReport.PageData]::new(); $p.ResultData = $Json; return $p }
    function Ms([string]$Utc) { return [DateTimeOffset]::Parse($Utc, [Globalization.CultureInfo]::InvariantCulture).ToUnixTimeMilliseconds() }
    function Read-TestHtml([string]$Path) {
        $html = Get-Content -LiteralPath $Path -Raw
        $meta = [regex]::Match($html, '<script id="report-meta" type="application/json">(.*?)</script>', 'Singleline').Groups[1].Value | ConvertFrom-Json
        $chunks = foreach ($match in [regex]::Matches($html, '<script type="application/x-dlp-chunk">([A-Za-z0-9+/=\r\n]+)</script>')) {
            $stream = [IO.MemoryStream]::new([Convert]::FromBase64String($match.Groups[1].Value))
            $gzip = [IO.Compression.GZipStream]::new($stream, [IO.Compression.CompressionMode]::Decompress)
            $reader = [IO.StreamReader]::new($gzip)
            try { $reader.ReadToEnd() | ConvertFrom-Json } finally { $reader.Dispose(); $stream.Dispose() }
        }
        [pscustomobject]@{ Html = $html; Meta = $meta; Chunks = @($chunks) }
    }
}

AfterAll {
    if ($script:ConfigDirectory -and (Test-Path -LiteralPath $script:ConfigDirectory)) { Remove-Item -LiteralPath $script:ConfigDirectory -Recurse -Force }
}

Describe 'Disclaimer synchronization' {
    BeforeAll {
        $script:DisclaimerPattern = '(?ms)^## Disclaimer\r?\n(?<Text>.*?)(?=^## |\z)'
        $sections = [regex]::Matches((Get-Content (Join-Path $script:RepoRoot 'README.md') -Raw), $script:DisclaimerPattern)
        if ($sections.Count -ne 1) { throw 'The root README must have exactly one Disclaimer section.' }
        $script:Disclaimer = ($sections[0].Groups['Text'].Value -replace "`r`n", "`n").Trim()
    }
    It 'ships the complete root disclaimer in the runtime README' {
        $sections = [regex]::Matches((Get-Content (Join-Path $script:Root 'README.md') -Raw), $script:DisclaimerPattern)
        $sections.Count | Should -Be 1
        ($sections[0].Groups['Text'].Value -replace "`r`n", "`n").Trim() | Should -BeExactly $script:Disclaimer
    }
    It 'includes the complete root disclaimer in the entry script help notes' {
        $errors = $null
        $ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $script:Root 'Invoke-PurviewDlpReport.ps1'), [ref]$null, [ref]$errors)
        $errors | Should -BeNullOrEmpty
        $help = $ast.GetHelpContent()
        $help | Should -Not -BeNullOrEmpty
        $notes = ($help.Notes -replace "`r`n", "`n" -replace '(?m)^[ \t]+', '').Trim()
        $start = $notes.IndexOf(($script:Disclaimer -split "`n")[0], [StringComparison]::Ordinal)
        $start | Should -BeGreaterOrEqual 0
        $notes.Substring($start) | Should -BeExactly $script:Disclaimer
    }
    It 'ships the exact runtime README and disclaimer through the existing package builder' {
        $destination = Join-Path $TestDrive 'disclaimer-package'
        Test-Path -LiteralPath $destination | Should -BeFalse
        & (Join-Path $script:RepoRoot 'tools\New-DlpPackage.ps1') -Destination $destination | Out-Null
        $readme = Join-Path $destination 'README.md'
        $readme | Should -Exist
        (Get-FileHash -LiteralPath $readme -Algorithm SHA256).Hash |
            Should -BeExactly (Get-FileHash -LiteralPath (Join-Path $script:Root 'README.md') -Algorithm SHA256).Hash
        $sections = [regex]::Matches((Get-Content -LiteralPath $readme -Raw), $script:DisclaimerPattern)
        $sections.Count | Should -Be 1
        ($sections[0].Groups['Text'].Value -replace "`r`n", "`n").Trim() | Should -BeExactly $script:Disclaimer
    }
}

Describe 'Configuration' {
    It 'loads the delivered configuration file (tenant values filled in)' {
        $s = Import-DlpConfiguration -Path $script:ConfigPath -Root $script:Root
        $s.Target.RuleId | Should -Be $script:RuleId
        [IO.Path]::IsPathRooted($s.Storage.DatabasePath) | Should -BeTrue
        $s.Zone.Id | Should -Not -BeNullOrEmpty
    }
    It 'reports every invalid value at once' {
        $path = Join-Path $TestDrive 'bad.psd1'
        $text = Get-Content $script:ConfigPath -Raw
        $text = $text.Replace('PageSize             = 1000', 'PageSize             = 9000').Replace("SplitBy                 = 'Week'", "SplitBy                 = 'Month'")
        Set-Content -LiteralPath $path -Value $text
        { Import-DlpConfiguration -Path $path -Root $script:Root } | Should -Throw -ExpectedMessage '*Collection.PageSize*Report.SplitBy*'
    }
    It 'requires the application settings in Certificate mode' {
        $path = Join-Path $TestDrive 'cert.psd1'
        (Get-Content $script:ConfigPath -Raw).Replace("Mode                  = 'Interactive'", "Mode                  = 'Certificate'") | Set-Content -LiteralPath $path
        { Import-DlpConfiguration -Path $path -Root $script:Root } | Should -Throw -ExpectedMessage '*AppId*CertificateThumbprint*'
    }
    It 'rejects source retention beyond the documented Activity Explorer window' {
        $path = Join-Path $TestDrive 'retention.psd1'
        (Get-Content $script:ConfigPath -Raw).Replace('SourceRetentionDays  = 30', 'SourceRetentionDays  = 31') | Set-Content -LiteralPath $path
        { Import-DlpConfiguration -Path $path -Root $script:Root } | Should -Throw -ExpectedMessage '*SourceRetentionDays*between 1 and 30*'
    }
    It 'rejects local retention shorter than the collection backfill' {
        $path = Join-Path $TestDrive 'storage-retention.psd1'
        (Get-Content $script:ConfigPath -Raw).Replace('RetentionDays          = 180', 'RetentionDays          = 20') | Set-Content -LiteralPath $path
        { Import-DlpConfiguration -Path $path -Root $script:Root } | Should -Throw -ExpectedMessage '*Storage.RetentionDays must be 0 or at least Collection.BackfillDays*'
    }
    It 'hides tenant information by default and keeps older configuration files compatible' {
        $settings = Import-DlpConfiguration -Path $script:ConfigPath -Root $script:Root
        $settings.Report.ShowTenantInformation | Should -BeFalse
        $settings.Report.CompanyName | Should -Be ''
        $settings.Report.EnvironmentName | Should -Be ''
        $path = Join-Path $TestDrive 'older.config.psd1'
        (Get-Content $script:ConfigPath -Raw) -replace '(?m)^\s*(CompanyName|ShowTenantInformation|EnvironmentName)\s*=.*\r?\n', '' | Set-Content $path
        (Import-DlpConfiguration -Path $path -Root $script:Root).Report.ShowTenantInformation | Should -BeFalse
    }
    It 'requires a boolean display switch and rejects configurable authors' {
        $path = Join-Path $TestDrive 'branding.config.psd1'
        $text = (Get-Content $script:ConfigPath -Raw).Replace('ShowTenantInformation   = $false', "ShowTenantInformation   = 'false'")
        $text.Replace("CompanyName             = ''", "CompanyName             = ''; Author = 'Other author'") | Set-Content $path
        { Import-DlpConfiguration -Path $path -Root $script:Root } | Should -Throw -ExpectedMessage '*Report.ShowTenantInformation*Report author is fixed*'
    }
}

Describe 'Console characters' {
    It 'uses only characters of the classic console fonts outside the emoji style' {
        # Repertoire of Consolas and Lucida Console (checked glyph by glyph): code page 437 and Latin-1.
        # The classic console has no font fallback: any other character is shown as an empty box.
        $safe = [Collections.Generic.HashSet[int]]::new()
        foreach ($c in (0x20..0x7E) + (0xA0..0xFF)) { [void]$safe.Add($c) }
        foreach ($c in '☺☻♥♦♣♠•◘○◙♂♀♪♫☼►◄↕‼¶§▬↨↑↓→←∟↔▲▼⌂₧ƒ⌐░▒▓│┤╡╢╖╕╣║╗╝╜╛┐└┴┬├─┼╞╟╚╔╩╦╠═╬╧╨╤╥╙╘╒╓╫╪┘┌█▄▌▐▀αΓπΣστΦΘΩδ∞φε∩≡≥≤⌠⌡≈∙√ⁿ■'.ToCharArray()) { [void]$safe.Add([int]$c) }
        $module = Get-Module PurviewDlpReport
        $used = [Collections.Generic.List[string]]::new()
        $sets = & $module { (Get-DlpIconSet 'Symbols'), (Get-DlpIconSet 'Ascii'), (Get-DlpFrameSet 'Symbols' 'Lucida Console'), (Get-DlpFrameSet 'Symbols' 'Terminal'), (Get-DlpFrameSet 'Ascii' $null) }
        foreach ($set in $sets) { foreach ($key in $set.Keys) { $used.Add("set.$key=$($set[$key])") } }
        # Rounded corners: only for the fonts that have them (Consolas: checked glyph by glyph).
        $rounded = & $module { (Get-DlpFrameSet 'Symbols' 'Consolas'), (Get-DlpFrameSet 'Symbols' $null) }
        foreach ($set in $rounded) { foreach ($key in $set.Keys) { if ([int]$set[$key] -notin 0x256D, 0x256E, 0x256F, 0x2570) { $used.Add("rounded.$key=$($set[$key])") } } }
        ($rounded[0].TopLeft, $sets[2].TopLeft) | Should -Be @([char]0x256D, [char]0x250C)
        # Characters written directly by the module and the script (the two style functions excluded).
        foreach ($file in 'PurviewDlpReport.psm1', 'Invoke-PurviewDlpReport.ps1') {
            $ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $script:Root $file), [ref]$null, [ref]$null)
            $text = $ast.Extent.Text
            $skip = $ast.FindAll({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -in 'Get-DlpIconSet', 'Get-DlpFrameSet' }, $true)
            foreach ($f in @($skip) | Sort-Object { $_.Extent.StartOffset } -Descending) { $text = $text.Remove($f.Extent.StartOffset, $f.Extent.EndOffset - $f.Extent.StartOffset) }
            foreach ($m in [regex]::Matches($text, '\[char\]0x([0-9A-Fa-f]{4})')) { $used.Add("${file}=$([char][Convert]::ToInt32($m.Groups[1].Value, 16))") }
        }
        $bad = @($used | Where-Object { $v = $_.Substring($_.IndexOf('=') + 1); @($v.ToCharArray() | Where-Object { -not $safe.Contains([int]$_) }).Count -gt 0 })
        $used.Count | Should -BeGreaterThan 40
        $bad | Should -BeNullOrEmpty
    }
}

Describe 'Entry script' {
    It 'chooses a usable ExchangeOnlineManagement version for the authentication mode' {
        $m = Get-Module PurviewDlpReport
        $v = { param([string]$Version) [pscustomobject]@{ Version = [version]$Version; ModuleBase = "C:\m\$Version" } }
        $both = @((& $v '3.10.0'), (& $v '3.9.2'), (& $v '3.8.0'))
        $cert = & $m { param($a) Select-DlpExoModule -Available $a -Loaded $null -Mode 'Certificate' } $both
        $cert.Module.Version | Should -Be ([version]'3.9.2')
        $cert.Skipped | Should -Match '^3\.10\.0: .*3\.10\.1'
        (& $m { param($a) Select-DlpExoModule -Available $a -Loaded $null -Mode 'Interactive' } $both).Module.Version | Should -Be ([version]'3.10.0')
        (& $m { param($a) Select-DlpExoModule -Available $a -Loaded $null -Mode 'Certificate' } @((& $v '3.10.1'), (& $v '3.10.0'))).Module.Version | Should -Be ([version]'3.10.1')
        { & $m { param($a) Select-DlpExoModule -Available $a -Loaded $null -Mode 'Certificate' } @((& $v '3.10.0')) } | Should -Throw -ExpectedMessage '*MinimumVersion 3.10.1*'
        { & $m { param($l) Select-DlpExoModule -Available @() -Loaded $l -Mode 'Certificate' } (& $v '3.10.0') } | Should -Throw -ExpectedMessage '*new PowerShell window*'
        { & $m { param($a) Select-DlpExoModule -Available $a -Loaded $null -Mode 'Interactive' } @((& $v '3.8.0')) } | Should -Throw -ExpectedMessage '*older than 3.9.0*'
        { & $m { Select-DlpExoModule -Available @() -Loaded $null -Mode 'Interactive' } } | Should -Throw -ExpectedMessage '*not installed*'
    }
    It 'Status on a new installation confirms the configuration and creates no database' {
        $dir = Join-Path $TestDrive 'fresh'
        $path = Join-Path $TestDrive 'fresh.psd1'
        $text = Get-Content $script:ConfigPath -Raw
        $text = $text.Replace("'.\data\PurviewDlpReport.sqlite'", "'$dir\data\PurviewDlpReport.sqlite'").Replace("Path          = '.\logs'", "Path          = '$dir\logs'")
        Set-Content -LiteralPath $path -Value $text
        $output = & pwsh -NoProfile -File (Join-Path $script:Root 'Invoke-PurviewDlpReport.ps1') -Mode Status -ConfigPath $path 2>&1 | Out-String
        $LASTEXITCODE | Should -Be 0
        $output | Should -Match 'Ready for the first collection'
        Test-Path -LiteralPath (Join-Path $dir 'data') | Should -BeFalse
    }
    It 'returns incomplete for a one-second gap even when coverage rounds to 100 percent' {
        $dir = Join-Path $TestDrive 'rounded-entry'
        $settings = New-TestSettings $dir
        $store = Open-DlpStore -Settings $settings
        try {
            $target = $store.GetOrCreateTarget($script:PolicyId, $script:RuleId, 'P', 'R')
            $run = $store.StartRun('Test', $null, $null, 'test', $null)
            $i = $store.BeginInterval($run, $target, (Ms '2026-09-28T00:00:00Z'), (Ms '2026-09-28T23:59:59Z'))
            $store.EndInterval($i, $true, 1, 0, 0, 0, $null)
            $store.FinishRun($run, 'Completed', '{}')
        } finally { $store.Dispose() }
        $path = Join-Path $TestDrive 'rounded-entry.psd1'
        $text = (Get-Content $script:ConfigPath -Raw).Replace("'.\data\PurviewDlpReport.sqlite'", "'$($settings.Storage.DatabasePath)'")
        $text = $text.Replace("OutputPath              = '.\reports'", "OutputPath              = '$($settings.Report.OutputPath)'").Replace("Path          = '.\logs'", "Path          = '$dir\logs'")
        Set-Content $path $text
        $output = & pwsh -NoProfile -File (Join-Path $script:Root 'Invoke-PurviewDlpReport.ps1') -Mode Report -Range Custom -Start '2026-09-28T00:00:00Z' -End '2026-09-29T00:00:00Z' -NoCollect -ConfigPath $path 2>&1 | Out-String
        $LASTEXITCODE | Should -Be 2
        $output | Should -Match 'Report written, but INCOMPLETE'
        $metadata = Get-ChildItem $settings.Report.OutputPath -Recurse -Filter report-metadata.json | Get-Content -Raw | ConvertFrom-Json
        $metadata.CoveragePercent | Should -Be 100
        $metadata.MissingMs | Should -Be 1000
        $metadata.CoverageComplete | Should -BeFalse
    }
}

Describe 'Periods' {
    BeforeAll { $script:Now = [DateTimeOffset]::Parse('2026-11-15T10:30:45.678Z') }
    It 'Last24Hours ends now, rounded to the second' {
        $p = Resolve-DlpPeriod -Range Last24Hours -Zone $script:Paris -Now $script:Now
        $p.EndMs | Should -Be (Ms '2026-11-15T10:30:45Z')
        ($p.EndMs - $p.StartMs) | Should -Be 86400000
    }
    It 'PreviousMonth follows the report time zone and summer time' {
        $p = Resolve-DlpPeriod -Range PreviousMonth -Zone $script:Paris -Now $script:Now
        $p.StartMs | Should -Be (Ms '2026-09-30T22:00:00Z')
        $p.EndMs | Should -Be (Ms '2026-10-31T23:00:00Z')
    }
    It 'Month and Day read calendar dates' {
        (Resolve-DlpPeriod -Range Month -Month '2026-08' -Zone $script:Paris -Now $script:Now).StartMs | Should -Be (Ms '2026-07-31T22:00:00Z')
        (Resolve-DlpPeriod -Range Day -Date '2026-11-02' -Zone $script:Paris -Now $script:Now).EndMs | Should -Be (Ms '2026-11-02T23:00:00Z')
    }
    It 'Custom dates without offset are local; the end is capped to now' {
        $p = Resolve-DlpPeriod -Range Custom -Start '2026-11-15 08:00' -End '2026-12-01' -Zone $script:Paris -Now $script:Now
        $p.StartMs | Should -Be (Ms '2026-11-15T07:00:00Z')
        $p.EndMs | Should -Be (Ms '2026-11-15T10:30:45Z')
    }
    It 'rejects a missing -Month' { { Resolve-DlpPeriod -Range Month -Zone $script:Paris -Now $script:Now } | Should -Throw '*yyyy-MM*' }
}

Describe 'Coverage arithmetic' {
    It 'finds the gaps of a period' {
        $covered = New-DlpRangeList @([pscustomobject]@{ Start = 10; End = 20 }, [pscustomobject]@{ Start = 15; End = 30 }, [pscustomobject]@{ Start = 40; End = 50 })
        $gaps = [PurviewDlpReport.Coverage]::Gaps($covered, 0, 60)
        @($gaps | ForEach-Object { "$($_.Start)-$($_.End)" }) | Should -Be @('0-10', '30-40', '50-60')
    }
    It 'cuts slices at local midnight' {
        $ranges = New-DlpRangeList @([pscustomobject]@{ Start = (Ms '2026-09-27T20:00:00Z'); End = (Ms '2026-09-29T01:00:00Z') })
        $slices = [PurviewDlpReport.Coverage]::SplitIntoSlices($ranges, $script:Paris, 24)
        $slices.Count | Should -Be 3
        $slices[0].End | Should -Be (Ms '2026-09-27T22:00:00Z')
        $slices[1].End | Should -Be (Ms '2026-09-28T22:00:00Z')
    }
}

Describe 'Normalization of Activity Explorer records' {
    It 'reads the report fields and the recipient count' {
        $n = Invoke-Normalize (New-TestRecord -Count 38 -Receivers @('a@x.com', 'b@x.com'))
        $n.Problem | Should -BeNullOrEmpty
        $n.Event.RecipientCount | Should -Be 38
        $n.Event.CountSource | Should -Be 'OtherConditions'
        $n.Event.Recipients | Should -Be 'a@x.com; b@x.com'
        $n.Event.HappenedMs | Should -Be (Ms '2026-09-28T10:00:00Z')
    }
    It 'accepts EmailInfo as a JSON string and PolicyMatchInfo as an array' {
        $n = Invoke-Normalize (New-TestRecord -EmailInfoAsString -PolicyAsArray -Count 27)
        $n.Problem | Should -BeNullOrEmpty
        $n.Event.RecipientCount | Should -Be 27
    }
    It 'quarantines a record of another rule' {
        (Invoke-Normalize (New-TestRecord -RuleId ([guid]::NewGuid()))).Problem | Should -Be 'TargetRuleNotInRecord'
    }
    It 'quarantines a record without MessageID' {
        $r = New-TestRecord; $r.EmailInfo.Remove('MessageID')
        (Invoke-Normalize $r).Problem | Should -Be 'MissingMessageID'
    }
    It 'never invents a count when sources disagree' {
        $n = Invoke-Normalize (New-TestRecord -Count 30 -ExplicitCount 31)
        $n.Event.RecipientCount | Should -BeNullOrEmpty
        $n.Event.CountSource | Should -Be 'ConflictingValues'
    }
    It 'accepts an explicit count equal to the condition value' {
        (Invoke-Normalize (New-TestRecord -Count 30 -ExplicitCount '30')).Event.CountSource | Should -Be 'RecipientCountField+OtherConditions'
    }
}

Describe 'Page envelope checks' {
    It 'accepts the known ErrorData marker only with ResultCode Success' {
        $ok = [pscustomobject]@{ ResultData = '[]'; LastPage = $true; ResultCode = 'Success'; ErrorData = 'Microsoft.Exchange.Hygiene.DataInsights.Common.DataInsightsErrorData' }
        (Test-DlpPageEnvelope $ok).LastPage | Should -BeTrue
        $ko = [pscustomobject]@{ ResultData = '[]'; LastPage = $true; ResultCode = 'Failure'; ErrorData = 'Microsoft.Exchange.Hygiene.DataInsights.Common.DataInsightsErrorData' }
        { Test-DlpPageEnvelope $ko } | Should -Throw '*SourceEnvelopeError*'
    }
    It 'rejects a page without LastPage or ResultData' {
        { Test-DlpPageEnvelope ([pscustomobject]@{ ResultData = '[]' }) } | Should -Throw '*LastPage*'
        { Test-DlpPageEnvelope ([pscustomobject]@{ LastPage = 'True' }) } | Should -Throw '*ResultData*'
    }
}

Describe 'Database' {
    BeforeEach {
        $script:Store = [PurviewDlpReport.DlpStore]::new((Join-Path $TestDrive ("db-{0}.sqlite" -f [guid]::NewGuid().ToString('N'))), $false, 'test')
        $script:Target = $script:Store.GetOrCreateTarget($script:PolicyId, $script:RuleId, 'P', 'R')
        $script:Run = $script:Store.StartRun('Test', $null, $null, 'test', $null)
        $script:Interval = $script:Store.BeginInterval($script:Run, $script:Target, 0, 1)
    }
    AfterEach { $script:Store.Dispose() }
    It 'stores a page once, even when received twice' {
        $page = New-TestPageData (ConvertTo-TestPage @((New-TestRecord), (New-TestRecord)))
        $first = $script:Store.IngestPage($script:Interval, $script:Target, $script:PolicyId, $script:RuleId, $page)
        $second = $script:Store.IngestPage($script:Interval, $script:Target, $script:PolicyId, $script:RuleId, $page)
        $first.NewEvents | Should -Be 2
        $second.Duplicates | Should -Be 2
        $script:Store.GetStatistics($script:Target).Events | Should -Be 2
    }
    It 'keeps one message for several events of the same MessageID and counts conflicts' {
        $page = New-TestPageData (ConvertTo-TestPage @((New-TestRecord -MessageId '<m1@x>'), (New-TestRecord -MessageId '<m1@x>' -Subject 'Other subject')))
        $r = $script:Store.IngestPage($script:Interval, $script:Target, $script:PolicyId, $script:RuleId, $page)
        $r.NewEvents | Should -Be 2
        $r.Conflicts | Should -Be 1
        $script:Store.GetStatistics($script:Target).Messages | Should -Be 1
    }
    It 'sets unreadable records aside in quarantine' {
        $bad = New-TestRecord; $bad.Workload = 'SharePoint'
        $r = $script:Store.IngestPage($script:Interval, $script:Target, $script:PolicyId, $script:RuleId, (New-TestPageData (ConvertTo-TestPage @($bad, (New-TestRecord)))))
        $r.Quarantined | Should -Be 1
        $r.QuarantineReasons['UnexpectedWorkload'] | Should -Be 1
        $script:Store.GetStatistics($script:Target).Quarantined | Should -Be 1
    }
    It 'rejects a page that is not a JSON array' {
        { $script:Store.IngestPage($script:Interval, $script:Target, $script:PolicyId, $script:RuleId, (New-TestPageData '{"a":1}')) } | Should -Throw '*MalformedPage*'
    }
    It 'returns only settled ranges when a settling time is given' {
        $i = $script:Store.BeginInterval($script:Run, $script:Target, 1000, 2000)
        $script:Store.EndInterval($i, $true, 1, 0, 0, $null, $null)
        $script:Store.GetCoveredRanges($script:Target, 0).Count | Should -Be 1
        $script:Store.GetCoveredRanges($script:Target, [long]::MaxValue / 2).Count | Should -Be 0
    }
    It 'keeps the old part of a recent range and re-collects only the last SettlingHours' {
        $now = [DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds()
        $i = $script:Store.BeginInterval($script:Run, $script:Target, $now - 24 * 3600000, $now - 3600000)
        $script:Store.EndInterval($i, $true, 1, 0, 0, $null, $null)
        $settled = $script:Store.GetCoveredRanges($script:Target, 6 * 3600000)
        $settled.Count | Should -Be 1
        $settled[0].Start | Should -Be ($now - 24 * 3600000)
        [Math]::Abs($settled[0].End - ($now - 6 * 3600000)) | Should -BeLessThan 60000
    }
    It 'reports current and settled coverage separately' {
        $now = [DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds()
        $i = $script:Store.BeginInterval($script:Run, $script:Target, $now - 24 * 3600000, $now - 3600000)
        $script:Store.EndInterval($i, $true, 1, 0, 0, $null, $null)
        $coverage = Get-DlpCoverage -Store $script:Store -TargetId $script:Target -StartMs ($now - 24 * 3600000) -EndMs ($now - 3600000) -SettlingMs (6 * 3600000)
        $coverage.Percent | Should -Be 100
        $coverage.FinalPercent | Should -BeLessThan 100
        $coverage.ProvisionalMs | Should -BeGreaterThan 0
    }
    It 'does not infer completeness from rounded coverage' {
        $start = Ms '2026-09-28T00:00:00Z'; $end = $start + 86400000
        $i = $script:Store.BeginInterval($script:Run, $script:Target, $start, $end - 1000)
        $script:Store.EndInterval($i, $true, 1, 0, 0, 0, $null)
        $coverage = Get-DlpCoverage -Store $script:Store -TargetId $script:Target -StartMs $start -EndMs $end
        $coverage.Percent | Should -Be 100
        $coverage.MissingMs | Should -Be 1000
        $coverage.Complete | Should -BeFalse
    }
    It 'binds authenticated tenants, flags legacy history and rejects another tenant' {
        $tenant = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'
        $script:Store.BindAuthenticatedTenant($tenant) | Should -BeTrue
        $script:Store.GetMetadata('legacy_tenant_history_unverified') | Should -Be 'true'
        $script:Store.BindAuthenticatedTenant($tenant.ToUpperInvariant()) | Should -BeFalse
        { $script:Store.ValidateTenant('bbbbbbbb-cccc-dddd-eeee-ffffffffffff') } | Should -Throw '*Database tenant mismatch*'
        { $script:Store.BindAuthenticatedTenant('bbbbbbbb-cccc-dddd-eeee-ffffffffffff') } | Should -Throw '*Database tenant mismatch*'
        $script:Store.GetMetadata('tenant_id') | Should -Be $tenant
        $settings = New-TestSettings $TestDrive
        $settings.Storage.DatabasePath = $script:Store.DatabasePath
        $settings.Tenant.TenantId = 'bbbbbbbb-cccc-dddd-eeee-ffffffffffff'
        { Open-DlpStore -Settings $settings -ReadOnly } | Should -Throw '*Database tenant mismatch*'
    }
    It 'binds an empty database without attributing legacy history' {
        $settings = New-TestSettings (Join-Path $TestDrive 'fresh-tenant')
        $store = Open-DlpStore -Settings $settings
        try {
            $store.BindAuthenticatedTenant($settings.Tenant.TenantId) | Should -BeFalse
            $store.GetMetadata('legacy_tenant_history_unverified') | Should -Be 'false'
        } finally { $store.Dispose() }
    }
    It 'classifies only exact source-second boundaries and counts stored versus duplicate observations' {
        $start = Ms '2026-09-28T00:00:00.851Z'; $end = Ms '2026-09-28T00:01:00Z'
        $i = $script:Store.BeginInterval($script:Run, $script:Target, $start, $end)
        $records = @(
            New-TestRecord -Happened '2026-09-28T00:00:00Z'
            New-TestRecord -Happened '2026-09-28T00:01:00Z'
            New-TestRecord -Happened '2026-09-28T00:00:00.500Z'
            New-TestRecord -Happened '2026-09-28T00:01:00.001Z'
            New-TestRecord -Happened '2026-09-28T00:01:01Z'
        )
        $page = New-TestPageData (ConvertTo-TestPage $records)
        $first = $script:Store.IngestPage($i, $script:Target, $script:PolicyId, $script:RuleId, $page)
        $first.OutOfRange | Should -Be 5
        $first.BoundaryEvents | Should -Be 2
        $first.BoundaryNewEvents | Should -Be 2
        $first.BoundaryDuplicates | Should -Be 0
        $first.UnexpectedOutOfRange | Should -Be 3
        $first.UnexpectedTimestampSamples[0].StartMs | Should -Be $start
        $second = $script:Store.IngestPage($i, $script:Target, $script:PolicyId, $script:RuleId, $page)
        $second.BoundaryDuplicates | Should -Be 2
        $second.BoundaryNewEvents | Should -Be 0
        $second.UnexpectedTimestampSamples[0].AlreadyPresent | Should -BeTrue
        $script:Store.GetStatistics($script:Target).Events | Should -Be 5
    }
    It 'describes Status as time coverage rather than collection progress' {
        $settings = New-TestSettings $TestDrive
        $settings.Collection.SettlingHours = 6
        $i = $script:Store.BeginInterval($script:Run, $script:Target, (Ms '2026-09-27T22:00:00Z'), (Ms '2026-09-28T16:00:00Z'))
        $script:Store.EndInterval($i, $true, 1, 0, 0, 0, $null)
        $output = Show-DlpStatus -Store $script:Store -TargetId $script:Target -Settings $settings -Days 1 -Now ([DateTimeOffset]'2026-09-28T18:00:00Z') 6>&1 | Out-String
        $output | Should -Match 'Read-only snapshot.*Europe/Paris'
        $output | Should -Match 'Today partially covered'
        $output | Should -Not -Match 'collection in progress'
        $output | Should -Match 'Not yet covered:.*2026-09-28 18:00.*20:00'
    }
    It 'flags normalized events outside their requested interval' {
        $start = Ms '2026-09-28T00:00:00Z'
        $end = Ms '2026-09-28T01:00:00Z'
        $i = $script:Store.BeginInterval($script:Run, $script:Target, $start, $end)
        $record = New-TestRecord -Happened '2026-09-28T02:00:00Z'
        $r = $script:Store.IngestPage($i, $script:Target, $script:PolicyId, $script:RuleId, (New-TestPageData (ConvertTo-TestPage @($record))))
        $r.OutOfRange | Should -Be 1
        $r.OutOfRangeHappenedMs.Count | Should -Be 1
        $r.OutOfRangeHappenedMs[0] | Should -Be (Ms '2026-09-28T02:00:00Z')
        $script:Store.GetStatistics($script:Target).Events | Should -Be 1
    }
    It 'purges events and coverage older than the retention' {
        $page = New-TestPageData (ConvertTo-TestPage @((New-TestRecord -Happened '2020-01-01T00:00:00Z'), (New-TestRecord -Happened '2026-09-28T00:00:00Z')))
        $null = $script:Store.IngestPage($script:Interval, $script:Target, $script:PolicyId, $script:RuleId, $page)
        $p = $script:Store.PurgeBefore((Ms '2025-01-01T00:00:00Z'))
        $p.Events | Should -Be 1
        $p.Messages | Should -Be 1
    }
}

Describe 'Collection (fake Activity Explorer)' {
    BeforeEach {
        $script:Dir = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        $script:Settings = New-TestSettings $script:Dir
        $script:Store = Open-DlpStore -Settings $script:Settings
        $script:Target = $script:Store.GetOrCreateTarget($script:PolicyId, $script:RuleId, 'P', 'R')
        $script:Run = $script:Store.StartRun('Test', $null, $null, 'test', $null)
        $script:Records = @(1..5 | ForEach-Object { New-TestRecord -Happened ('2026-09-28T0{0}:00:00Z' -f $_) })
        $script:Slices = New-DlpRangeList @([pscustomobject]@{ Start = (Ms '2026-09-28T00:00:00Z'); End = (Ms '2026-09-28T12:00:00Z') })
        $script:State = @{ Calls = 0; FailOnCall = 0; Permanent = $false; RepeatCookie = $false }
        $records = $script:Records; $state = $script:State
        $script:Exporter = {
            param([hashtable]$Query)
            $state.Calls++
            if ($state.FailOnCall -eq $state.Calls) { if ($state.Permanent) { throw 'Access denied: insufficient permissions' } else { throw 'The operation has timed out' } }
            $offset = if ($Query.ContainsKey('PageCookie')) { [int]$Query.PageCookie } else { 0 }
            $page = @($records | Select-Object -Skip $offset -First $Query.PageSize)
            $last = $offset + $page.Count -ge $records.Count
            $cookie = if ($state.RepeatCookie) { '2' } else { [string]($offset + $page.Count) }
            [pscustomobject]@{ ResultData = (ConvertTo-Json -InputObject $page -Depth 10 -Compress); LastPage = $last; Watermark = $(if ($last) { '' } else { $cookie }); TotalResultCount = $records.Count; ResultCode = 'Success' }
        }.GetNewClosure()
    }
    AfterEach { $script:Store.Dispose() }
    It 'pages until LastPage and marks the window as collected' {
        $r = Invoke-DlpCollection -Store $script:Store -RunId $script:Run -TargetId $script:Target -Settings $script:Settings -Slices $script:Slices -Exporter $script:Exporter 6>$null
        $r.Completed | Should -BeTrue
        $r.Pages | Should -Be 3
        $r.NewEvents | Should -Be 5
        $script:Store.GetCoveredRanges($script:Target, 0).Count | Should -Be 1
    }
    It 'restarts a window after a transient error' {
        $script:Settings.Collection.MaxSliceRestarts = 1
        $script:State.FailOnCall = 2
        Mock -ModuleName PurviewDlpReport Start-Sleep { }
        $r = Invoke-DlpCollection -Store $script:Store -RunId $script:Run -TargetId $script:Target -Settings $script:Settings -Slices $script:Slices -Exporter $script:Exporter 6>$null
        $r.Completed | Should -BeTrue
        $r.Restarts | Should -Be 1
        $r.NewEvents | Should -Be 5
    }
    It 'stops on a permanent error and keeps the window uncollected' {
        $script:State.FailOnCall = 1; $script:State.Permanent = $true
        $r = Invoke-DlpCollection -Store $script:Store -RunId $script:Run -TargetId $script:Target -Settings $script:Settings -Slices $script:Slices -Exporter $script:Exporter 6>$null
        $r.Completed | Should -BeFalse
        $r.Error | Should -BeLike '*Access denied*'
        $script:Store.GetCoveredRanges($script:Target, 0).Count | Should -Be 0
    }
    It 'splits a window that keeps failing' {
        $script:Settings.Collection.MaxSliceRestarts = 0
        $script:State.FailOnCall = 1
        $r = Invoke-DlpCollection -Store $script:Store -RunId $script:Run -TargetId $script:Target -Settings $script:Settings -Slices $script:Slices -Exporter $script:Exporter 6>$null
        $r.Completed | Should -BeTrue
        $r.Subdivisions | Should -Be 1
        (Get-DlpCoverage -Store $script:Store -TargetId $script:Target -StartMs (Ms '2026-09-28T00:00:00Z') -EndMs (Ms '2026-09-28T12:00:00Z')).Percent | Should -Be 100
    }
    It 'refuses a paging cookie that does not change' {
        $script:Settings.Collection.MaxSliceRestarts = 0; $script:Settings.Collection.MinimumSliceMinutes = 720
        $script:State.RepeatCookie = $true
        $r = Invoke-DlpCollection -Store $script:Store -RunId $script:Run -TargetId $script:Target -Settings $script:Settings -Slices $script:Slices -Exporter $script:Exporter 6>$null
        $r.Completed | Should -BeFalse
        $r.Error | Should -BeLike '*RepeatedCookie*'
    }
    It 'plans only what is missing and flags what is too old to collect' {
        $now = Ms '2026-09-29T12:00:00Z'
        $plan = Get-DlpCollectionPlan -Store $script:Store -TargetId $script:Target -StartMs (Ms '2026-08-01T00:00:00Z') -EndMs $now -Settings $script:Settings -NowMs $now
        $plan.UnrecoverableMs | Should -BeGreaterThan 0
        ($plan.Collectable | Select-Object -First 1).Start | Should -Be ($now - 30 * 86400000 + 3600000)
    }
    It 'aggregates boundary observations without treating them as unexpected timestamps' {
        $script:Settings.Collection.PageSize = 1000
        $records = @(New-TestRecord -Happened '2026-09-28T00:00:00Z')
        $json = ConvertTo-TestPage $records
        $exporter = { param($Query) [pscustomobject]@{ ResultData = $json; LastPage = $true; ResultCode = 'Success' } }.GetNewClosure()
        $slices = New-DlpRangeList @([pscustomobject]@{ Start = (Ms '2026-09-28T00:00:00.851Z'); End = (Ms '2026-09-28T00:01:00Z') })
        $first = Invoke-DlpCollection -Store $script:Store -RunId $script:Run -TargetId $script:Target -Settings $script:Settings -Slices $slices -Exporter $exporter 6>$null
        $second = Invoke-DlpCollection -Store $script:Store -RunId $script:Run -TargetId $script:Target -Settings $script:Settings -Slices $slices -Exporter $exporter 6>$null
        $first.Completed | Should -BeTrue -Because $first.Error
        $second.Completed | Should -BeTrue -Because $second.Error
        $first.BoundaryNewEvents | Should -Be 1
        $second.BoundaryDuplicates | Should -Be 1
        $second.UnexpectedOutOfRange | Should -Be 0
        $output = Write-DlpCollectionDiagnostics -Collection $second -Zone $script:Paris 6>&1 | Out-String
        $output | Should -Match '1 already in the database'
        $output | Should -Match 'none were discarded'
        $output | Should -Not -Match 'Unexpected timestamps'
    }
}

Describe 'Report files' {
    BeforeAll {
        $script:Dir = Join-Path $TestDrive 'report'
        $script:Settings = New-TestSettings $script:Dir
        $script:Store = Open-DlpStore -Settings $script:Settings
        $script:Target = $script:Store.GetOrCreateTarget($script:PolicyId, $script:RuleId, 'P', 'R')
        $run = $script:Store.StartRun('Test', $null, $null, 'test', $null)
        $interval = $script:Store.BeginInterval($run, $script:Target, (Ms '2026-09-01T00:00:00Z'), (Ms '2026-10-01T00:00:00Z'))
        # 3 messages in week 36 (Sept 1-6), 2 in week 37, one of them with a subject starting with '='.
        $records = @(
            New-TestRecord -Happened '2026-09-01T08:00:00Z' -MessageId '<a@x>'
            New-TestRecord -Happened '2026-09-01T09:00:00Z' -MessageId '<a@x>'
            New-TestRecord -Happened '2026-09-02T08:00:00Z' -MessageId '<b@x>' -Count 45
            New-TestRecord -Happened '2026-09-03T08:00:00Z' -MessageId '<c@x>' -Subject '=HYPERLINK("x")'
            New-TestRecord -Happened '2026-09-08T08:00:00Z' -MessageId '<d@x>' -Receivers @('one@x.com', 'two@x.com', 'three@x.com')
            New-TestRecord -Happened '2026-09-09T08:00:00Z' -MessageId '<e@x>' -Sender 'bob@contoso.com'
        )
        $null = $script:Store.IngestPage($interval, $script:Target, $script:PolicyId, $script:RuleId, (New-TestPageData (ConvertTo-TestPage $records)))
        $script:Store.EndInterval($interval, $true, 1, 6, 6, 6, $null)
        $script:Period = Resolve-DlpPeriod -Range Month -Month '2026-09' -Zone $script:Settings.Zone -Now ([DateTimeOffset]'2026-10-15T00:00:00Z')
        $script:Coverage = Get-DlpCoverage -Store $script:Store -TargetId $script:Target -StartMs $script:Period.StartMs -EndMs $script:Period.EndMs
        function Read-TestCsv($File) { Import-Csv -LiteralPath $File.Path -Delimiter ';' }
    }
    AfterAll { $script:Store.Dispose() }

    It 'writes one row per Message ID with the business columns' {
        $out = New-DlpReport -Store $script:Store -TargetId $script:Target -Settings $script:Settings -Period $script:Period -Coverage $script:Coverage
        $out.Result.Messages | Should -Be 5
        $csv = $out.Result.Files | Where-Object Kind -eq 'CSV'
        @($csv).Count | Should -Be 1
        (Split-Path $csv.Path -Leaf) | Should -Be 'PurviewDLP_2026-09-01_to_2026-09-30.csv'
        $rows = Read-TestCsv $csv
        @($rows[0].PSObject.Properties.Name) | Should -Be @('Detection time (Europe/Paris)', 'Sender', 'Recipients', 'Subject', 'Recipient count', 'Message ID')
        $rows[0].'Detection time (Europe/Paris)' | Should -Be '2026-09-01 10:00:00'
        ($rows | Where-Object 'Message ID' -eq '<c@x>').Subject | Should -Be "'=HYPERLINK(""x"")"
        ($rows | Where-Object 'Message ID' -eq '<d@x>').Recipients | Should -Be 'one@x.com; two@x.com; three@x.com'
        Test-Path -LiteralPath "$($out.Directory).pending" | Should -BeFalse
    }
    It 'omits the recipient addresses when IncludeRecipientDetails is false' {
        $script:Settings.Report.IncludeRecipientDetails = $false
        try {
            $out = New-DlpReport -Store $script:Store -TargetId $script:Target -Settings $script:Settings -Period $script:Period -Coverage $script:Coverage
            $rows = Read-TestCsv ($out.Result.Files | Where-Object Kind -eq 'CSV')
            @($rows[0].PSObject.Properties.Name) | Should -Not -Contain 'Recipients'
            $html = Get-Content -LiteralPath ($out.Result.Files | Where-Object Kind -eq 'HTML').Path -Raw
            $html | Should -Match '"includeRecipients":false'
        } finally { $script:Settings.Report.IncludeRecipientDetails = $true }
    }
    It 'splits by week above MaxRowsPerFile and names the files by period' {
        $script:Settings.Report.MaxRowsPerFile = 1000
        $request = [PurviewDlpReport.ReportRequest]::new()
        $request.StartMs = $script:Period.StartMs; $request.EndMs = $script:Period.EndMs; $request.Zone = $script:Settings.Zone
        $request.SplitBy = 'Week'; $request.MaxRowsPerFile = 2; $request.FilePrefix = 'X'
        $times = [Collections.Generic.List[long]]::new()
        foreach ($t in '2026-09-01T08:00:00Z', '2026-09-02T08:00:00Z', '2026-09-03T08:00:00Z', '2026-09-08T08:00:00Z') { $times.Add((Ms $t)) }
        $plan = [PurviewDlpReport.ReportPlanner]::Plan($times, $request)
        @($plan | ForEach-Object BaseName) | Should -Be @('X_2026-09-01_to_2026-09-06_part1of2', 'X_2026-09-01_to_2026-09-06_part2of2', 'X_2026-09-07_to_2026-09-13')
        @($plan | ForEach-Object RowCount) | Should -Be @(2, 1, 1)
    }
    It 'does not split below MaxRowsPerFile and writes a self-contained HTML file' {
        $out = New-DlpReport -Store $script:Store -TargetId $script:Target -Settings $script:Settings -Period $script:Period -Coverage $script:Coverage
        $html = Get-Content -LiteralPath ($out.Result.Files | Where-Object Kind -eq 'HTML').Path -Raw
        $html | Should -Not -Match '%%CHUNKS%%|%%META%%'
        $html | Should -Match 'application/x-dlp-chunk'
        $html | Should -Match '"rows":5'
    }
    It 'omits tenant and environment from HTML and metadata unless the administrator opts in' {
        $script:Settings.Report.EnvironmentName = 'Hidden environment marker'
        try {
            $out = New-DlpReport -Store $script:Store -TargetId $script:Target -Settings $script:Settings -Period $script:Period -Coverage $script:Coverage
            $data = Read-TestHtml ($out.Result.Files | Where-Object Kind -eq 'HTML').Path
            $data.Meta.PSObject.Properties.Name | Should -Not -Contain 'tenantId'
            $data.Html | Should -Not -Match 'Hidden environment marker|aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee|contoso.onmicrosoft.com'
            $manifest = Get-Content (Join-Path $out.Directory 'report-metadata.json') -Raw
            $manifest | Should -Not -Match 'TenantId|TenantOrganization|Hidden environment marker'
        } finally { $script:Settings.Report.EnvironmentName = '' }
    }
    It 'shows opted-in tenant and environment and separate company branding without changing CSV columns' {
        $script:Settings.Report.ShowTenantInformation = $true
        $script:Settings.Report.EnvironmentName = 'PFT'
        $script:Settings.Report.CompanyName = 'Example & Company <Ltd>'
        try {
            $null = $script:Store.BindAuthenticatedTenant($script:Settings.Tenant.TenantId)
            $out = New-DlpReport -Store $script:Store -TargetId $script:Target -Settings $script:Settings -Period $script:Period -Coverage $script:Coverage
            $data = Read-TestHtml ($out.Result.Files | Where-Object Kind -eq 'HTML').Path
            $data.Meta.tenantId | Should -Be $script:Settings.Tenant.TenantId
            $data.Meta.environmentName | Should -Be 'PFT'
            $data.Meta.companyName | Should -Be 'Example & Company <Ltd>'
            $data.Meta.author | Should -Be 'Nicolas Fabert'
            $manifest = Get-Content (Join-Path $out.Directory 'report-metadata.json') -Raw | ConvertFrom-Json
            $manifest.TenantIdentitySource | Should -Be 'Authenticated collection session'
            $manifest.LegacyTenantHistoryUnverified | Should -BeTrue
            $rows = Read-TestCsv ($out.Result.Files | Where-Object Kind -eq 'CSV')
            @($rows[0].PSObject.Properties.Name).Count | Should -Be 6
        } finally {
            $script:Settings.Report.ShowTenantInformation = $false
            $script:Settings.Report.EnvironmentName = ''; $script:Settings.Report.CompanyName = ''
        }
    }
    It 'shows company branding independently of tenant visibility' {
        $script:Settings.Report.CompanyName = 'Example Company'
        try {
            $out = New-DlpReport -Store $script:Store -TargetId $script:Target -Settings $script:Settings -Period $script:Period -Coverage $script:Coverage
            $meta = (Read-TestHtml ($out.Result.Files | Where-Object Kind -eq 'HTML').Path).Meta
            $meta.companyName | Should -Be 'Example Company'
            $meta.PSObject.Properties.Name | Should -Not -Contain 'tenantId'
        } finally { $script:Settings.Report.CompanyName = '' }
    }
    It 'keeps an exact missing gap and provisional notice despite rounded 100 percent' {
        $coverage = [pscustomobject]@{
            Percent = 100; MissingMs = 1; FinalPercent = 100; ProvisionalMs = 1
            Gaps = @([PurviewDlpReport.TimeRange]::new($script:Period.EndMs - 1, $script:Period.EndMs))
        }
        $out = New-DlpReport -Store $script:Store -TargetId $script:Target -Settings $script:Settings -Period $script:Period -Coverage $coverage
        $data = Read-TestHtml ($out.Result.Files | Where-Object Kind -eq 'HTML').Path
        $data.Meta.coverageComplete | Should -BeFalse
        $data.Meta.missingMs | Should -Be 1
        $data.Meta.coverageNote | Should -Match 'INCOMPLETE.*1 ms missing.*recent data is provisional'
    }
    It 'marks a failed refresh incomplete even when the requested period was already covered' {
        $coverage = [pscustomobject]@{ Percent = 100; MissingMs = 0; FinalPercent = 100; ProvisionalMs = 0; Gaps = @() }
        $out = New-DlpReport -Store $script:Store -TargetId $script:Target -Settings $script:Settings -Period $script:Period -Coverage $coverage -CollectionFailed
        $meta = (Read-TestHtml ($out.Result.Files | Where-Object Kind -eq 'HTML').Path).Meta
        $meta.coverageComplete | Should -BeFalse
        $meta.collectionSucceeded | Should -BeFalse
        $meta.coverageNote | Should -Match 'refresh did not finish'
    }
    It 'preserves UTC chronology and distinct offsets through autumn and spring daylight-saving changes' {
        $settings = New-TestSettings (Join-Path $TestDrive 'dst-report')
        $store = Open-DlpStore -Settings $settings
        try {
            $target = $store.GetOrCreateTarget($script:PolicyId, $script:RuleId, 'P', 'R')
            $run = $store.StartRun('Test', $null, $null, 'test', $null)
            $start = Ms '2026-03-01T00:00:00Z'; $end = Ms '2026-11-01T00:00:00Z'
            $i = $store.BeginInterval($run, $target, $start, $end)
            $times = @('2026-03-29T00:45:00Z', '2026-03-29T01:15:00Z', '2026-10-25T00:45:00.100Z', '2026-10-25T01:15:00.200Z')
            $records = @($times | ForEach-Object { New-TestRecord -Happened $_ })
            $null = $store.IngestPage($i, $target, $script:PolicyId, $script:RuleId, (New-TestPageData (ConvertTo-TestPage $records)))
            $store.EndInterval($i, $true, 1, 4, 4, 4, $null)
            $period = [pscustomobject]@{ Range = 'Custom'; StartMs = $start; EndMs = $end }
            $coverage = Get-DlpCoverage -Store $store -TargetId $target -StartMs $start -EndMs $end
            $out = New-DlpReport -Store $store -TargetId $target -Settings $settings -Period $period -Coverage $coverage
            $data = Read-TestHtml ($out.Result.Files | Where-Object Kind -eq 'HTML').Path
            $chunk = $data.Chunks[0]
            @($chunk.u | Sort-Object) | Should -Be @($times | ForEach-Object { Ms $_ })
            $autumn = @(0..3 | Where-Object { $chunk.u[$_] -ge (Ms '2026-10-25T00:00:00Z') } | Sort-Object { $chunk.u[$_] })
            @($autumn | ForEach-Object { $chunk.o[$_] }) | Should -Be @(120, 60)
            $chunk.t[$autumn[0]] | Should -BeGreaterThan $chunk.t[$autumn[1]]
            $data.Html | Should -Match 'U\[b\] - U\[a\]'
            $data.Html | Should -Match 'Detection time \(UTC\)'
            [PurviewDlpReport.ReportWriter]::FormatLocal((Ms $times[0]), $script:Paris) | Should -Be '2026-03-29 01:45:00'
            [PurviewDlpReport.ReportWriter]::FormatLocal((Ms $times[1]), $script:Paris) | Should -Be '2026-03-29 03:15:00'
        } finally { $store.Dispose() }
    }
    It 'protects CSV cells from formula injection' {
        [PurviewDlpReport.ReportWriter]::SafeCsv('=1+1', ';') | Should -Be "'=1+1"
        [PurviewDlpReport.ReportWriter]::SafeCsv('a;b', ';') | Should -Be '"a;b"'
    }
}

Describe 'Packaging' {
    It 'clears branding and disables tenant display without rejecting generic environment names in documentation' {
        $source = Join-Path $TestDrive 'package-source'
        [void][IO.Directory]::CreateDirectory($source)
        Copy-Item -LiteralPath $script:Root -Destination (Join-Path $source 'package') -Recurse
        Copy-Item -LiteralPath (Join-Path $script:RepoRoot 'tools') -Destination $source -Recurse
        $config = Join-Path $source 'package\config\PurviewDlpReport.config.psd1'
        [void][IO.Directory]::CreateDirectory((Split-Path $config -Parent))
        $text = (Get-Content $script:ConfigPath -Raw).Replace('contoso.onmicrosoft.com', 'packaging-fixture.onmicrosoft.com').Replace("CompanyName             = ''", "CompanyName             = 'Example Company'")
        $text = $text.Replace("EnvironmentName         = ''", "EnvironmentName         = 'PFT'").Replace('ShowTenantInformation   = $false', 'ShowTenantInformation   = $True')
        Set-Content -LiteralPath $config -Value $text
        $destination = Join-Path $TestDrive 'package-result'
        & (Join-Path $source 'tools\New-DlpPackage.ps1') -Destination $destination 6>$null
        $packaged = Import-PowerShellDataFile (Join-Path $destination 'config\PurviewDlpReport.config.psd1')
        $packaged.Report.CompanyName | Should -Be ''
        $packaged.Report.EnvironmentName | Should -Be ''
        $packaged.Report.ShowTenantInformation | Should -BeFalse
        $packaged.Tenant.TenantId | Should -Be ''
        Test-Path (Join-Path $destination 'data') | Should -BeFalse
    }
}
