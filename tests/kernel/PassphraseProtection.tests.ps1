# ========================================
# Pester v5 unit tests for Protect-/Unprotect-PassphraseForResume
# ========================================
# Function: kernel/common.ps1 :: Protect-PassphraseForResume /
#                                Unprotect-PassphraseFromResume
# Run    : powershell.exe -File ./dev/run_tests.ps1
#
# Pins the DPAPI (LocalMachine) round-trip that resume_state.json relies
# on across __RESTART__ reboots (TM t-0007 (1)). A silent change to the
# encryption format - scope, encoding, base64 framing - would strand
# every interrupted kitting run at resume time with an undecryptable
# passphrase, observable only after a mid-profile reboot in the field.
#
# DPAPI LocalMachine round-trips only on the same machine, which is
# exactly the production contract (the PC that saved the resume state
# is the PC that resumes), so these tests run unmocked.
#
# NOTE: source stays pure ASCII (CLAUDE.md English-only rule); the
# multibyte passphrase case builds its string from [char] code points
# at runtime instead of embedding Japanese literals.
# ========================================

BeforeAll {
    . "$PSScriptRoot\..\_helpers\test_state.ps1"
    $script:RepoRoot = Get-FabriqRepoRoot
    . (Join-Path $script:RepoRoot 'kernel\common.ps1')
}

Describe 'Protect-PassphraseForResume / Unprotect-PassphraseFromResume' {

    Context 'Round-trip integrity' {

        It 'round-trips an ASCII passphrase' {
            $p = 'kitting-master-2026'
            $enc = Protect-PassphraseForResume -Passphrase $p
            Unprotect-PassphraseFromResume -ProtectedBase64 $enc | Should -Be $p
        }

        It 'round-trips a passphrase with spaces, quotes and symbols' {
            $p = 'pass "with" sp~aces & $ymbols\!'
            $enc = Protect-PassphraseForResume -Passphrase $p
            Unprotect-PassphraseFromResume -ProtectedBase64 $enc | Should -Be $p
        }

        It 'round-trips a multibyte (UTF-8) passphrase built from code points' {
            # "fabriq aikotoba" in katakana/hiragana via code points - keeps
            # this source file ASCII-only while exercising the UTF-8 byte path.
            $p = -join @([char]0x30D5, [char]0x30A1, [char]0x30D6, [char]0x308A, [char]0x304F)
            $enc = Protect-PassphraseForResume -Passphrase $p
            Unprotect-PassphraseFromResume -ProtectedBase64 $enc | Should -Be $p
        }

        It 'round-trips a long passphrase (256 chars)' {
            $p = ('Ab3$' * 64)
            $p.Length | Should -Be 256
            $enc = Protect-PassphraseForResume -Passphrase $p
            Unprotect-PassphraseFromResume -ProtectedBase64 $enc | Should -Be $p
        }
    }

    Context 'Output format contract' {

        It 'returns valid Base64 that is not the plaintext' {
            $p = 'plain-visibility-check'
            $enc = Protect-PassphraseForResume -Passphrase $p
            $enc | Should -Not -Be $p
            { [Convert]::FromBase64String($enc) } | Should -Not -Throw
        }
    }

    Context 'Failure paths' {

        It 'throws on a tampered ciphertext blob (DPAPI integrity)' {
            $enc = Protect-PassphraseForResume -Passphrase 'tamper-target'
            $bytes = [Convert]::FromBase64String($enc)
            $mid = [int]($bytes.Length / 2)
            $bytes[$mid] = $bytes[$mid] -bxor 0xFF
            $corrupted = [Convert]::ToBase64String($bytes)
            { Unprotect-PassphraseFromResume -ProtectedBase64 $corrupted } | Should -Throw
        }

        It 'throws on non-Base64 garbage input' {
            { Unprotect-PassphraseFromResume -ProtectedBase64 'not base64 at all!!' } | Should -Throw
        }

        It 'rejects an empty passphrase (Mandatory binding)' {
            { Protect-PassphraseForResume -Passphrase '' } | Should -Throw
        }
    }
}
