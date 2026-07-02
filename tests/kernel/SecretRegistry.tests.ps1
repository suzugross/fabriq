# ========================================
# Pester v5 unit tests for the Secret Registry
# (Register-FabriqSecret / Get-FabriqMaskedText / Clear-FabriqSecrets)
# ========================================
# Functions: kernel/common.ps1 :: Secret Registry section + its three
#            sink integrations (Show-*, telemetry, execution history)
#            + registration chokepoints (Import-ModuleCsv ENC: decrypt,
#            Test-MasterPassphrase, Unprotect-PassphraseFromResume)
# Run     : powershell.exe -File ./dev/run_tests.ps1
#
# Pins the A7 guarantee: a value decrypted from an ENC: cell (or the
# master passphrase / host PIN) can NOT reach the delivered transcript,
# the telemetry corpus, or execution_history.csv in plaintext via the
# kernel sinks. This is the systemic fix for the leak class that pianist
# hit locally (masked in Wave 2A) - future modules get it for free.
# ========================================

BeforeAll {
    . "$PSScriptRoot\..\_helpers\test_state.ps1"
    $script:RepoRoot = Get-FabriqRepoRoot
    . (Join-Path $script:RepoRoot 'kernel\common.ps1')

    # Test-only encryptor: same independent spec mirror used by
    # UnprotectFabriqValue.tests.ps1.
    function Protect-TestValue {
        param(
            [Parameter(Mandatory)][string]$PlainText,
            [Parameter(Mandatory)][string]$Passphrase
        )
        $salt = [System.Text.Encoding]::UTF8.GetBytes('fabriq-fixed-salt-2024')
        $kdf = New-Object System.Security.Cryptography.Rfc2898DeriveBytes(
            $Passphrase, $salt, 100000,
            [System.Security.Cryptography.HashAlgorithmName]::SHA256
        )
        $key = $kdf.GetBytes(32)
        $iv  = $kdf.GetBytes(16)
        $kdf.Dispose()

        $aes = [System.Security.Cryptography.Aes]::Create()
        $aes.Mode = [System.Security.Cryptography.CipherMode]::CBC
        $aes.Padding = [System.Security.Cryptography.PaddingMode]::PKCS7
        $aes.Key = $key
        $aes.IV = $iv
        $enc = $aes.CreateEncryptor()
        $plainBytes = [System.Text.Encoding]::UTF8.GetBytes($PlainText)
        $cipherBytes = $enc.TransformFinalBlock($plainBytes, 0, $plainBytes.Length)
        $enc.Dispose(); $aes.Dispose()
        return 'ENC:' + [Convert]::ToBase64String($cipherBytes)
    }
}

Describe 'Secret Registry core' {

    BeforeEach { Clear-FabriqSecrets }
    AfterAll   { Clear-FabriqSecrets }

    It 'masks a registered secret with ***' {
        Register-FabriqSecret -Value 'Sup3rSecret!'
        (Get-FabriqMaskedText -Text 'pw=Sup3rSecret! ok') | Should -Be 'pw=*** ok'
    }

    It 'returns text unchanged when nothing is registered (fast path)' {
        (Get-FabriqMaskedText -Text 'plain text') | Should -Be 'plain text'
    }

    It 'applies longest-first so a substring secret never splits a longer one' {
        Register-FabriqSecret -Value 'abc'
        Register-FabriqSecret -Value 'abc-def-123'
        (Get-FabriqMaskedText -Text 'x abc-def-123 y abc z') | Should -Be 'x *** y *** z'
    }

    It 'rejects values shorter than 3 chars (no output shredding)' {
        Register-FabriqSecret -Value '1'
        Register-FabriqSecret -Value 'OK'
        (Get-FabriqMaskedText -Text 'result 1 is OK') | Should -Be 'result 1 is OK'
    }

    It 'rejects null / whitespace without throwing' {
        { Register-FabriqSecret -Value $null } | Should -Not -Throw
        { Register-FabriqSecret -Value '   ' } | Should -Not -Throw
        (Get-FabriqMaskedText -Text 'unchanged') | Should -Be 'unchanged'
    }

    It 'Clear-FabriqSecrets empties the registry' {
        Register-FabriqSecret -Value 'Sup3rSecret!'
        Clear-FabriqSecrets
        (Get-FabriqMaskedText -Text 'pw=Sup3rSecret!') | Should -Be 'pw=Sup3rSecret!'
    }
}

Describe 'Sink integrations' {

    BeforeEach { Clear-FabriqSecrets }
    AfterAll   { Clear-FabriqSecrets }

    It 'Show-* masks registered secrets on the console (transcript surface)' {
        Register-FabriqSecret -Value 'LeakyValue99'
        $out = ((Show-Info -Message 'value is LeakyValue99') *>&1 | Out-String)
        $out | Should -Not -Match 'LeakyValue99'
        $out | Should -Match '\*\*\*'
    }

    It 'New-TelemetryRedactMap hard-redacts registered secrets (no hash)' {
        Register-FabriqSecret -Value 'TelemetrySecretX'
        $map = New-TelemetryRedactMap
        $map['TelemetrySecretX'] | Should -Be '[REDACTED]'
    }

    It 'Write-ExecutionHistory masks secrets in the message column' {
        Register-FabriqSecret -Value 'HistSecret77'
        $histDir = Join-Path $TestDrive 'hist'
        New-Item -ItemType Directory -Path $histDir -Force | Out-Null
        $origPath = $script:HistoryPath
        try {
            $script:HistoryPath = Join-Path $histDir 'execution_history.csv'
            $null = Write-ExecutionHistory -ModuleName 'rig' -Category 'Test' `
                -Status 'Success' -Message 'pw=HistSecret77 applied'
            $csv = Get-Content -Path $script:HistoryPath -Raw
            $csv | Should -Not -Match 'HistSecret77'
            $csv | Should -Match '\*\*\*'
        }
        finally {
            $script:HistoryPath = $origPath
        }
    }
}

Describe 'Registration chokepoints' {

    BeforeEach { Clear-FabriqSecrets }
    AfterAll   {
        Clear-FabriqSecrets
        $global:FabriqMasterPassphrase = $null
    }

    It 'Import-ModuleCsv auto-registers decrypted ENC: values' {
        $global:FabriqMasterPassphrase = 'rig-master-key'
        $enc = Protect-TestValue -PlainText 'CsvSecret42!' -Passphrase 'rig-master-key'
        $csvPath = Join-Path $TestDrive 'secret_list.csv'
        @(
            '"Enabled","Name","Password"'
            "`"1`",`"row1`",`"$enc`""
        ) | Set-Content -Path $csvPath -Encoding Ascii

        $items = Import-ModuleCsv -Path $csvPath -FilterEnabled
        $items[0].Password | Should -Be 'CsvSecret42!'
        (Get-FabriqMaskedText -Text 'x CsvSecret42! y') | Should -Be 'x *** y'
    }

    It 'Test-MasterPassphrase registers the passphrase on success only' {
        $token = Protect-TestValue -PlainText 'surkitinisme' -Passphrase 'good-pass-123'
        $tokenPath = Join-Path $TestDrive 'verify.txt'
        Set-Content -Path $tokenPath -Value $token -Encoding Ascii

        (Test-MasterPassphrase -Passphrase 'wrong-pass-456' -VerifyTokenPath $tokenPath) | Should -BeFalse
        (Get-FabriqMaskedText -Text 'wrong-pass-456') | Should -Be 'wrong-pass-456'

        (Test-MasterPassphrase -Passphrase 'good-pass-123' -VerifyTokenPath $tokenPath) | Should -BeTrue
        (Get-FabriqMaskedText -Text 'p=good-pass-123') | Should -Be 'p=***'
    }

    It 'Unprotect-PassphraseFromResume registers the restored secret' {
        $protected = Protect-PassphraseForResume -Passphrase 'ResumePin987'
        Clear-FabriqSecrets
        (Unprotect-PassphraseFromResume -ProtectedBase64 $protected) | Should -Be 'ResumePin987'
        (Get-FabriqMaskedText -Text 'pin ResumePin987') | Should -Be 'pin ***'
    }

    It 'Reset-FabriqState clears the registry (next-customer isolation, AST check)' {
        # Running Reset-FabriqState in a test would restart transcripts and
        # delete state files; assert structurally instead: its body must
        # call Clear-FabriqSecrets.
        $tokens = $null; $errors = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile(
            (Join-Path $script:RepoRoot 'kernel\common.ps1'), [ref]$tokens, [ref]$errors)
        $reset = $ast.FindAll({
            param($n)
            $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
            $n.Name -eq 'Reset-FabriqState'
        }, $true) | Select-Object -First 1
        $reset | Should -Not -BeNullOrEmpty
        $calls = $reset.FindAll({
            param($n)
            $n -is [System.Management.Automation.Language.CommandAst] -and
            $n.GetCommandName() -eq 'Clear-FabriqSecrets'
        }, $true)
        @($calls).Count | Should -BeGreaterOrEqual 1
    }
}
