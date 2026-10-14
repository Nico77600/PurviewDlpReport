#Requires -Version 7.4
# Real signature, timestamp, hashing and filesystem checks; no trust mocks.
BeforeAll {
    $script:Source = Join-Path (Split-Path $PSScriptRoot -Parent) 'package'
    $script:FixtureNumber = 0
    $script:Fixture = $null
    function Reset-SignedFixture {
        Remove-Module PurviewDlpReport -Force -ErrorAction SilentlyContinue
        # Loaded SQLite assemblies stay locked on Windows until the test process exits.
        $script:FixtureNumber++
        $script:Fixture = Join-Path $TestDrive ('signed-runtime-{0:D4}' -f $script:FixtureNumber)
        if (Test-Path -LiteralPath $script:Fixture) { throw "Unexpected existing test fixture: $script:Fixture" }
        Copy-Item -LiteralPath $script:Source -Destination $script:Fixture -Recurse
    }
    function Invoke-RealGuard {
        & (Join-Path $script:Fixture 'Test-PackageIntegrity.ps1') -Root $script:Fixture
    }
}
Describe 'Authenticated source and distribution integrity' {
    BeforeEach { Reset-SignedFixture }
    It 'accepts the real finalized source and a relocated byte-identical copy' {
        { & (Join-Path $script:Source 'Test-PackageIntegrity.ps1') -Root $script:Source } | Should -Not -Throw
        { Invoke-RealGuard } | Should -Not -Throw
        (Import-PowerShellDataFile (Join-Path $script:Fixture 'PackageIntegrity.psd1')).Files.Count | Should -Be 16
    }
    It 'requires valid timestamped owner signatures on all five PowerShell files' {
        foreach ($name in 'Invoke-PurviewDlpReport.ps1', 'PurviewDlpReport.psm1', 'PurviewDlpReport.psd1', 'Test-PackageIntegrity.ps1', 'PackageIntegrity.psd1') {
            $signature = Get-AuthenticodeSignature -LiteralPath (Join-Path $script:Fixture $name)
            [string]$signature.Status | Should -Be 'Valid'
            $signature.SignerCertificate.Thumbprint | Should -BeExactly '0F28E442889F17B5F4D251E8F96E7CC6E2C9C6D6'
            $signature.TimeStamperCertificate | Should -Not -BeNullOrEmpty
        }
    }
    It 'accepts editable configuration and ignored generated output but does not load caches' {
        Add-Content (Join-Path $script:Fixture 'config\PurviewDlpReport.config.psd1') '# Fictitious customization'
        [void][IO.Directory]::CreateDirectory((Join-Path $script:Fixture 'bin'))
        Set-Content (Join-Path $script:Fixture 'bin\PurviewDlpReport.Engine.invalid.dll') 'not an assembly'
        { Invoke-RealGuard } | Should -Not -Throw
        Import-Module (Join-Path $script:Fixture 'PurviewDlpReport.psd1') -Force
        { Initialize-DlpEngine -Root $script:Fixture } | Should -Not -Throw
        ('PurviewDlpReport.DlpStore' -as [type]).Assembly.Location | Should -BeNullOrEmpty
        @(Get-ChildItem (Join-Path $script:Fixture 'bin') -File).Count | Should -Be 1
    }
    It 'rejects mutations of every protected entry without executing the changed files' {
        $entries = (Import-PowerShellDataFile (Join-Path $script:Source 'PackageIntegrity.psd1')).Files
        foreach ($entry in $entries) {
            Reset-SignedFixture
            $path = Join-Path $script:Fixture $entry.Path
            if ([IO.Path]::GetExtension($path) -in '.ps1', '.psm1', '.psd1') {
                [IO.File]::WriteAllText($path, "# Evidence-only tamper`r`n" + [IO.File]::ReadAllText($path), [Text.UTF8Encoding]::new($true))
            } else { [IO.File]::AppendAllText($path, "`r`n# Evidence-only tamper") }
            if ($entry.Path -ceq 'Test-PackageIntegrity.ps1') {
                [string](Get-AuthenticodeSignature -LiteralPath (Join-Path $script:Fixture $entry.Path)).Status | Should -Be 'HashMismatch'
                { Import-Module (Join-Path $script:Fixture 'PurviewDlpReport.psd1') -Force -ErrorAction Stop } |
                    Should -Throw -ExpectedMessage '*verifier must have the trusted, timestamped owner signature*'
            } else {
                { Invoke-RealGuard } | Should -Throw -ExpectedMessage '*SHA256 mismatch*'
            }
        }
    }
    It 'rejects missing mandatory stable files' {
        foreach ($entry in (Import-PowerShellDataFile (Join-Path $script:Source 'PackageIntegrity.psd1')).Files) {
            Reset-SignedFixture
            Remove-Item -LiteralPath (Join-Path $script:Fixture $entry.Path)
            if ($entry.Path -ceq 'Test-PackageIntegrity.ps1') {
                { Import-Module (Join-Path $script:Fixture 'PurviewDlpReport.psd1') -Force -ErrorAction Stop } | Should -Throw
            } else { { Invoke-RealGuard } | Should -Throw }
        }
    }
    It 'rejects an absent or altered manifest before parsing it' {
        Remove-Item -LiteralPath (Join-Path $script:Fixture 'PackageIntegrity.psd1')
        { Invoke-RealGuard } | Should -Throw
        Reset-SignedFixture
        $path = Join-Path $script:Fixture 'PackageIntegrity.psd1'
        [IO.File]::WriteAllText($path, "# Changed`r`n" + [IO.File]::ReadAllText($path), [Text.UTF8Encoding]::new($true))
        { Invoke-RealGuard } | Should -Throw -ExpectedMessage '*expected trusted, timestamped owner signature*'
    }
    It 'rejects a valid owner-signed but wrong manifest schema' {
        Copy-Item -LiteralPath (Join-Path $script:Fixture 'PurviewDlpReport.psd1') -Destination (Join-Path $script:Fixture 'PackageIntegrity.psd1') -Force
        { Invoke-RealGuard } | Should -Throw -ExpectedMessage '*unsupported or incomplete manifest schema*'
    }
    It 'rejects a cryptographically valid manifest with its timestamp removed' {
        Add-Type -AssemblyName System.Security.Cryptography.Pkcs
        $path = Join-Path $script:Fixture 'PackageIntegrity.psd1'
        $encoding = [Text.UTF8Encoding]::new($false, $true)
        $text = $encoding.GetString([IO.File]::ReadAllBytes($path))
        $match = [regex]::Match($text, '(?ms)^# SIG # Begin signature block\r?\n(?<Data>.*?)^# SIG # End signature block\r?\n?\z')
        $match.Success | Should -BeTrue
        $cms = [Security.Cryptography.Pkcs.SignedCms]::new()
        $cms.Decode([Convert]::FromBase64String((($match.Groups['Data'].Value -replace '(?m)^#\s?', '') -replace '\s', '')))
        foreach ($attribute in @($cms.SignerInfos[0].UnsignedAttributes)) {
            foreach ($value in $attribute.Values) {
                $cms.SignerInfos[0].RemoveUnsignedAttribute([Security.Cryptography.AsnEncodedData]::new($attribute.Oid, $value.RawData))
            }
        }
        $base64 = [Convert]::ToBase64String($cms.Encode())
        $lines = for ($i = 0; $i -lt $base64.Length; $i += 64) { '# ' + $base64.Substring($i, [Math]::Min(64, $base64.Length - $i)) }
        [IO.File]::WriteAllBytes($path, $encoding.GetBytes($text.Substring(0, $match.Index) +
            "# SIG # Begin signature block`r`n" + ($lines -join "`r`n") + "`r`n# SIG # End signature block`r`n"))
        $signature = Get-AuthenticodeSignature -LiteralPath $path
        [string]$signature.Status | Should -Be 'Valid'
        $signature.TimeStamperCertificate | Should -BeNullOrEmpty
        { Invoke-RealGuard } | Should -Throw -ExpectedMessage '*expected trusted, timestamped owner signature*'
    }
    It 'rejects unknown code and documentation names rather than allowing directories or extensions' {
        foreach ($relative in 'lib\sqlite\extra.dll', 'src\extra.cs', 'docs\extra.ps1', 'docs\extra.html', 'docs\images\extra.png', 'config\extra.psd1') {
            Reset-SignedFixture
            Set-Content (Join-Path $script:Fixture $relative) 'not executable'
            { Invoke-RealGuard } | Should -Throw -ExpectedMessage '*unexpected stable file*'
        }
    }
    It 'allows only the declared repository-only documents without requiring them in a ZIP' {
        Remove-Item -LiteralPath (Join-Path $script:Fixture 'LICENSE')
        Remove-Item -LiteralPath (Join-Path $script:Fixture 'docs\PurviewDlpReport-Guide.md')
        Remove-Item -LiteralPath (Join-Path $script:Fixture 'docs\PurviewDlpReport-UserGuide.md')
        Remove-Item -LiteralPath (Join-Path $script:Fixture 'docs\images') -Recurse
        { Invoke-RealGuard } | Should -Not -Throw
        @(Get-ChildItem $script:Fixture -Recurse -File).Count | Should -Be 18
    }
    It 'rejects junctions in protected or source-document subtrees' {
        $outside = Join-Path $TestDrive 'junction-target'
        [void][IO.Directory]::CreateDirectory($outside)
        foreach ($relative in 'src', 'docs\images') {
            Reset-SignedFixture
            $link = Join-Path $script:Fixture $relative
            Remove-Item -LiteralPath $link -Recurse -Force
            New-Item -ItemType Junction -Path $link -Target $outside | Out-Null
            try { { Invoke-RealGuard } | Should -Throw }
            finally { (Get-Item -LiteralPath $link -Force).Delete() }
        }
    }
    It 'checks integrity again on warm initialization and refuses another engine root' {
        Import-Module (Join-Path $script:Fixture 'PurviewDlpReport.psd1') -Force
        Initialize-DlpEngine -Root $script:Fixture
        { Initialize-DlpEngine -Root $script:Source } | Should -Throw -ExpectedMessage '*engine root must be the imported package root*'
        Add-Content (Join-Path $script:Fixture 'src\PurviewDlpReport.Engine.cs') '// Changed after initialization'
        { Initialize-DlpEngine -Root $script:Fixture } | Should -Throw -ExpectedMessage '*SHA256 mismatch*'
    }
    It 'refuses an arbitrary external template before writing a report' {
        Import-Module (Join-Path $script:Fixture 'PurviewDlpReport.psd1') -Force
        $settings = @{ Report = @{ TemplatePath = Join-Path $TestDrive 'untrusted.html' } }
        { New-DlpReport -Store @{} -TargetId 1 -Settings $settings -Period @{} -Coverage @{} } |
            Should -Throw -ExpectedMessage '*report template must belong to the authenticated package*'
    }
}
AfterAll { Remove-Module PurviewDlpReport -Force -ErrorAction SilentlyContinue }
