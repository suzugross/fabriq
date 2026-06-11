# ========================================
# Pester v5 unit tests for ENC: decryption (Unprotect-FabriqValue /
# Test-MasterPassphrase)
# ========================================
# Functions: kernel/common.ps1 :: Unprotect-FabriqValue / Test-MasterPassphrase
# Run     : powershell.exe -File ./dev/run_tests.ps1
#
# Pins the ENC: cipher spec (TM t-0024 (2)) - a silent drift here bricks
# every ENC: credential in customer hostlists/CSVs:
#   Key derivation : PBKDF2-HMAC-SHA256, 100,000 iterations,
#                    fixed salt "fabriq-fixed-salt-2024"
#   Key / IV       : 32 + 16 bytes drawn from the SAME KDF stream
#   Cipher         : AES-256-CBC, PKCS7 padding
#   Encoding       : UTF-8 plaintext, "ENC:" + Base64 ciphertext
#
# The encryptor below is a deliberate INDEPENDENT reimplementation of
# that spec (mirroring the C# CryptoPoC) - if common.ps1's decryptor
# drifts from the spec in any parameter, the roundtrips fail loudly.
# Japanese plaintext is built from codepoints so this source stays ASCII.
# ========================================

BeforeAll {
    . "$PSScriptRoot\..\_helpers\test_state.ps1"
    $script:RepoRoot = Get-FabriqRepoRoot
    . (Join-Path $script:RepoRoot 'kernel\common.ps1')

    # Test-only encryptor: independent spec mirror (see header).
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

    # Robust wrong-key assertion: AES-CBC with a wrong key either throws
    # a padding exception or (rarely, ~1/256) yields garbage with valid
    # padding - both count as "did not reveal the plaintext".
    function Test-DecryptionDenied {
        param([string]$Encrypted, [string]$Passphrase, [string]$RealPlain)
        try {
            $out = Unprotect-FabriqValue -EncryptedValue $Encrypted -Passphrase $Passphrase
            return ($out -ne $RealPlain)
        }
        catch { return $true }
    }
}

Describe 'Unprotect-FabriqValue' {

    Context 'spec roundtrip (independent encryptor -> kernel decryptor)' {

        It 'decrypts ASCII plaintext' {
            $enc = Protect-TestValue -PlainText 'P@ssw0rd-123' -Passphrase 'master-key'
            Unprotect-FabriqValue -EncryptedValue $enc -Passphrase 'master-key' |
                Should -Be 'P@ssw0rd-123'
        }

        It 'decrypts multibyte (Japanese) plaintext via UTF-8' {
            # Codepoint-built so this test source stays ASCII
            $jp = [string]::Join('', @([char]0x30D1, [char]0x30B9, [char]0x30EF, [char]0x30FC, [char]0x30C9))  # katakana "password"
            $enc = Protect-TestValue -PlainText $jp -Passphrase 'master-key'
            Unprotect-FabriqValue -EncryptedValue $enc -Passphrase 'master-key' |
                Should -Be $jp
        }

        It 'decrypts long plaintext spanning many AES blocks' {
            $long = ('fabriq-' * 64)   # 448 chars
            $enc = Protect-TestValue -PlainText $long -Passphrase 'master-key'
            Unprotect-FabriqValue -EncryptedValue $enc -Passphrase 'master-key' |
                Should -Be $long
        }
    }

    Context 'routing and failure modes' {

        It 'returns non-ENC: values unchanged (plaintext passthrough)' {
            Unprotect-FabriqValue -EncryptedValue 'plain-value' -Passphrase 'whatever' |
                Should -Be 'plain-value'
        }

        It 'does not reveal the plaintext under a wrong passphrase' {
            $enc = Protect-TestValue -PlainText 'top-secret' -Passphrase 'right-pass'
            Test-DecryptionDenied -Encrypted $enc -Passphrase 'wrong-pass' -RealPlain 'top-secret' |
                Should -BeTrue
        }

        It 'throws on malformed Base64 after the ENC: prefix' {
            { Unprotect-FabriqValue -EncryptedValue 'ENC:not-base64!!' -Passphrase 'x' } |
                Should -Throw
        }
    }
}

Describe 'Test-MasterPassphrase' {

    BeforeEach {
        # Verification token = ENC: of the fixed verify plaintext,
        # exactly what Fabriq Studio generates.
        $script:tokenPath = Join-Path $env:TEMP `
            ("fabriq-verify-token-{0}.txt" -f ([guid]::NewGuid().ToString('N')))
        Mock Show-Warning { }
    }

    AfterEach {
        if ($script:tokenPath -and (Test-Path $script:tokenPath)) {
            Remove-Item $script:tokenPath -Force -ErrorAction SilentlyContinue
        }
    }

    It 'accepts the correct passphrase against a Studio-style token' {
        Set-Content -Path $script:tokenPath -Value (Protect-TestValue -PlainText 'surkitinisme' -Passphrase 'site-master') -Encoding Ascii
        Test-MasterPassphrase -Passphrase 'site-master' -VerifyTokenPath $script:tokenPath |
            Should -BeTrue
    }

    It 'rejects a wrong passphrase' {
        Set-Content -Path $script:tokenPath -Value (Protect-TestValue -PlainText 'surkitinisme' -Passphrase 'site-master') -Encoding Ascii
        Test-MasterPassphrase -Passphrase 'typo-master' -VerifyTokenPath $script:tokenPath |
            Should -BeFalse
    }

    It 'rejects a token file without the ENC: prefix (invalid token)' {
        Set-Content -Path $script:tokenPath -Value 'surkitinisme' -Encoding Ascii
        Test-MasterPassphrase -Passphrase 'site-master' -VerifyTokenPath $script:tokenPath |
            Should -BeFalse
        Should -Invoke Show-Warning -Times 1 -Exactly
    }
}
