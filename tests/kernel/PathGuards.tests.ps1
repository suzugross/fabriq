# ========================================
# Pester v5 unit tests for Destructive Path Guards
# ========================================
# Functions: kernel/common.ps1 :: Test-FabriqSafePathComponent /
#            Test-FabriqProtectedPath
# Run     : powershell.exe -File ./dev/run_tests.ps1
#
# Pins the CLAUDE.md section-8 guard contract (TM t-0009):
#   - component validation rejects everything that lets a CSV value
#     escape its base directory (the C:\Users Join-Path collapse class)
#   - protected-path verdicts: protected roots, parents of protected
#     roots, sub-3-segment paths, unresolvable paths
#   - wildcard-leaf handling: PS 5.1 GetFullPath throws on * / ?, and
#     the shipped history_destroyer CSV relies on wildcard targets, so
#     the parent directory is validated instead. A regression here
#     either bricks every shipped DeletePath row (guard too strict) or
#     reopens the recursive-delete hole (guard too loose).
# Pure functions - no mocks, no machine state touched.
# ========================================

BeforeAll {
    . "$PSScriptRoot\..\_helpers\test_state.ps1"
    $script:RepoRoot = Get-FabriqRepoRoot
    . (Join-Path $script:RepoRoot 'kernel\common.ps1')
}

Describe 'Test-FabriqSafePathComponent' {

    Context 'legitimate components pass byte-identical (no transform)' {
        It 'accepts plain user names (shipped profile_list.csv values)' {
            Test-FabriqSafePathComponent -Value 'test' | Should -BeTrue
            Test-FabriqSafePathComponent -Value 'defaultuser0' | Should -BeTrue
        }
        It 'accepts names with inner spaces and dots (driver model names)' {
            Test-FabriqSafePathComponent -Value 'X1 Carbon' | Should -BeTrue
            Test-FabriqSafePathComponent -Value 'ThinkCentre.M75q' | Should -BeTrue
        }
    }

    Context 'escape vectors are rejected' {
        It 'rejects empty and whitespace-only (Join-Path collapses to the base itself)' {
            Test-FabriqSafePathComponent -Value '' | Should -BeFalse
            Test-FabriqSafePathComponent -Value '   ' | Should -BeFalse
        }
        It 'rejects . and .. (traversal)' {
            Test-FabriqSafePathComponent -Value '.' | Should -BeFalse
            Test-FabriqSafePathComponent -Value '..' | Should -BeFalse
        }
        It 'rejects path separators' {
            Test-FabriqSafePathComponent -Value 'a\b' | Should -BeFalse
            Test-FabriqSafePathComponent -Value 'a/b' | Should -BeFalse
            Test-FabriqSafePathComponent -Value '..\..\x' | Should -BeFalse
        }
        It 'rejects wildcards and other invalid filename chars' {
            Test-FabriqSafePathComponent -Value 'a*' | Should -BeFalse
            Test-FabriqSafePathComponent -Value 'a?b' | Should -BeFalse
            Test-FabriqSafePathComponent -Value 'a:b' | Should -BeFalse
            Test-FabriqSafePathComponent -Value 'a"b' | Should -BeFalse
        }
        It 'rejects trailing dot / space (silently trimmed by Win32 resolution)' {
            Test-FabriqSafePathComponent -Value 'name.' | Should -BeFalse
            Test-FabriqSafePathComponent -Value 'name ' | Should -BeFalse
        }
    }
}

Describe 'Test-FabriqProtectedPath' {

    Context 'blocked targets' {
        It 'blocks empty input' {
            $v = Test-FabriqProtectedPath -Path ''
            $v.IsSafe | Should -BeFalse
        }
        It 'blocks protected roots exactly (trailing backslash tolerated)' {
            (Test-FabriqProtectedPath -Path 'C:\Users').IsSafe | Should -BeFalse
            (Test-FabriqProtectedPath -Path 'C:\Users\').IsSafe | Should -BeFalse
            (Test-FabriqProtectedPath -Path 'C:\Windows\System32').IsSafe | Should -BeFalse
            (Test-FabriqProtectedPath -Path 'C:\').IsSafe | Should -BeFalse
        }
        It 'blocks the current user profile root' {
            (Test-FabriqProtectedPath -Path $env:USERPROFILE).IsSafe | Should -BeFalse
        }
        It 'blocks the fabriq root itself' {
            (Test-FabriqProtectedPath -Path $script:RepoRoot).IsSafe | Should -BeFalse
        }
        It 'blocks paths shallower than 3 segments (the C:\Users collapse class)' {
            $v = Test-FabriqProtectedPath -Path 'C:\Win'
            $v.IsSafe | Should -BeFalse
            $v.Reason | Should -Match 'shallow'
        }
        It 'blocks a wildcard whose parent is a protected root (C:\Users\*)' {
            $v = Test-FabriqProtectedPath -Path 'C:\Users\*'
            $v.IsSafe | Should -BeFalse
            $v.Reason | Should -Match 'Protected'
        }
        It 'blocks a wildcard whose parent is too shallow (C:\x\*)' {
            (Test-FabriqProtectedPath -Path 'C:\x\*').IsSafe | Should -BeFalse
        }
        It 'blocks non-leaf wildcards as unresolvable (fail-closed)' {
            $v = Test-FabriqProtectedPath -Path 'C:\Users\*\AppData\Local\Temp'
            $v.IsSafe | Should -BeFalse
            $v.Reason | Should -Match 'Unresolvable'
        }
    }

    Context 'device / UNC / extended-length namespace bypass (TM t-0041)' {
        # The protected-root list is drive-letter form (C:\...), but
        # [IO.Path]::GetFullPath preserves \\?\ / \\.\ / UNC prefixes, so
        # these alternate spellings resolve to a protected root yet evade
        # the list. Must be blocked fail-closed (verified bypass on PS 5.1).
        It 'blocks extended-length \\?\ pointing at a protected root' {
            (Test-FabriqProtectedPath -Path '\\?\C:\Windows').IsSafe | Should -BeFalse
            (Test-FabriqProtectedPath -Path '\\?\C:\Users\Public').IsSafe | Should -BeFalse
        }
        It 'blocks device namespace \\.\ pointing at a protected root' {
            (Test-FabriqProtectedPath -Path '\\.\C:\Windows').IsSafe | Should -BeFalse
        }
        It 'blocks \\?\UNC\ form' {
            (Test-FabriqProtectedPath -Path '\\?\UNC\localhost\c$\Windows').IsSafe | Should -BeFalse
        }
        It 'blocks UNC administrative shares (\\host\c$)' {
            (Test-FabriqProtectedPath -Path '\\localhost\c$\Windows').IsSafe | Should -BeFalse
            (Test-FabriqProtectedPath -Path '\\127.0.0.1\c$\Windows').IsSafe | Should -BeFalse
        }
        It 'blocks forward-slash and mixed-separator variants (normalize-first)' {
            (Test-FabriqProtectedPath -Path '//?/C:/Windows').IsSafe | Should -BeFalse
            (Test-FabriqProtectedPath -Path '//./C:/Windows').IsSafe | Should -BeFalse
            (Test-FabriqProtectedPath -Path '//localhost/c$/Windows').IsSafe | Should -BeFalse
            (Test-FabriqProtectedPath -Path '\\?/C:/Windows').IsSafe | Should -BeFalse
        }
        It 'reports a namespace-specific block reason (not a generic pass)' {
            (Test-FabriqProtectedPath -Path '\\?\C:\Windows').Reason | Should -Match 'namespace'
            (Test-FabriqProtectedPath -Path '\\localhost\c$\Windows').Reason | Should -Match 'UNC'
        }
        It 'still allows a legit deep drive-letter target (no false positive)' {
            (Test-FabriqProtectedPath -Path 'C:\Users\fabriq-test-no-such-user\AppData\Local\Temp\junk').IsSafe | Should -BeTrue
        }
    }

    Context 'legitimate targets pass (shipped CSV regression pins)' {
        It 'allows a sibling user profile folder (profile_delete purpose)' {
            (Test-FabriqProtectedPath -Path 'C:\Users\fabriq-test-no-such-user').IsSafe | Should -BeTrue
        }
        It 'allows deep wildcard targets (history_destroyer shipped rows)' {
            $recent = Join-Path $env:APPDATA 'Microsoft\Windows\Recent\*'
            (Test-FabriqProtectedPath -Path $recent).IsSafe | Should -BeTrue
            $temp = Join-Path $env:TEMP '*'
            (Test-FabriqProtectedPath -Path $temp).IsSafe | Should -BeTrue
        }
        It 'allows 3-segment system cache dirs used by shipped rows (C:\Windows\Temp\*, Prefetch\*)' {
            (Test-FabriqProtectedPath -Path "$env:windir\Temp\*").IsSafe | Should -BeTrue
            (Test-FabriqProtectedPath -Path "$env:windir\Prefetch\*").IsSafe | Should -BeTrue
        }
        It 'allows wildcard FILE patterns in a deep dir (thumbcache_*.db)' {
            $tc = Join-Path $env:LOCALAPPDATA 'Microsoft\Windows\Explorer\thumbcache_*.db'
            (Test-FabriqProtectedPath -Path $tc).IsSafe | Should -BeTrue
        }
        It 'returns an empty Reason and a populated NormalizedPath when safe' {
            $v = Test-FabriqProtectedPath -Path 'C:\Users\fabriq-test-no-such-user'
            $v.Reason | Should -Be ''
            $v.NormalizedPath | Should -Be 'c:\users\fabriq-test-no-such-user'
        }
    }
}
