# ========================================
# Pester v5 unit tests for New-BatchResult
# ========================================
# Function: kernel/common.ps1 :: New-BatchResult
# Run    : powershell.exe -File ./dev/run_tests.ps1
#
# Pins the Status auto-determination matrix and Verified pass-through
# behavior. New-BatchResult underpins almost every standard module's
# return path; a regression here silently changes execution_history
# Status across the entire fleet.
# ========================================

BeforeAll {
    . "$PSScriptRoot\..\_helpers\test_state.ps1"
    $script:RepoRoot = Get-FabriqRepoRoot
    . (Join-Path $script:RepoRoot 'kernel\common.ps1')
}

Describe 'New-BatchResult' {

    Context 'Status auto-determination matrix' {

        It 'returns Success when only Success>0' {
            $r = New-BatchResult -Success 3 -Skip 0 -Fail 0
            $r.Status | Should -Be 'Success'
        }

        It 'returns Success when Success>0 and Skip>0 (no failures)' {
            $r = New-BatchResult -Success 3 -Skip 2 -Fail 0
            $r.Status | Should -Be 'Success'
        }

        It 'returns Partial when Success>0 and Fail>0' {
            $r = New-BatchResult -Success 3 -Skip 0 -Fail 1
            $r.Status | Should -Be 'Partial'
        }

        It 'returns Partial when Success>0 and Skip>0 and Fail>0' {
            $r = New-BatchResult -Success 3 -Skip 2 -Fail 1
            $r.Status | Should -Be 'Partial'
        }

        It 'returns Skipped when only Skip>0' {
            $r = New-BatchResult -Success 0 -Skip 3 -Fail 0
            $r.Status | Should -Be 'Skipped'
        }

        It 'returns Error when only Fail>0' {
            $r = New-BatchResult -Success 0 -Skip 0 -Fail 2
            $r.Status | Should -Be 'Error'
        }

        It 'returns Error when Fail>0 and Skip>0 (no successes)' {
            $r = New-BatchResult -Success 0 -Skip 1 -Fail 2
            $r.Status | Should -Be 'Error'
        }

        It 'returns Success on the all-zero edge case (default branch)' {
            # Documents current behavior: empty result defaults to Success.
            # Modules typically guard against this earlier with explicit
            # Skipped returns, so this branch should rarely fire in practice.
            $r = New-BatchResult -Success 0 -Skip 0 -Fail 0
            $r.Status | Should -Be 'Success'
        }
    }

    Context 'Verified pass-through' {

        It 'omits Verified ($null) when -Verified is not supplied' {
            $r = New-BatchResult -Success 1
            $r.Verified | Should -BeNullOrEmpty
        }

        It 'propagates -Verified $true' {
            $r = New-BatchResult -Success 1 -Verified $true
            $r.Verified | Should -BeTrue
        }

        It 'propagates -Verified $false' {
            $r = New-BatchResult -Success 1 -Verified $false
            $r.Verified | Should -BeFalse
        }

        It 'propagates -Verified $false even on a Partial Status' {
            $r = New-BatchResult -Success 1 -Fail 1 -Verified $false
            $r.Status   | Should -Be 'Partial'
            $r.Verified | Should -BeFalse
        }
    }

    Context 'Message format' {

        It 'emits "Success: N, Skip: M, Fail: K" by default' {
            $r = New-BatchResult -Success 3 -Skip 2 -Fail 1
            $r.Message | Should -Be 'Success: 3, Skip: 2, Fail: 1'
        }

        It 'appends MessageSuffix with a leading space' {
            $r = New-BatchResult -Success 1 -MessageSuffix '(restart pending)'
            $r.Message | Should -Be 'Success: 1, Skip: 0, Fail: 0 (restart pending)'
        }

        It 'omits suffix when MessageSuffix is empty string (default)' {
            $r = New-BatchResult -Success 1
            $r.Message | Should -Be 'Success: 1, Skip: 0, Fail: 0'
        }
    }

    Context 'ModuleResult shape contract (KERNEL_API.md §5)' {

        It 'sets _IsModuleResult marker to $true' {
            $r = New-BatchResult -Success 1
            $r._IsModuleResult | Should -BeTrue
        }

        It 'populates Timestamp with a DateTime' {
            $r = New-BatchResult -Success 1
            $r.Timestamp | Should -BeOfType [datetime]
        }

        It 'leaves Details as the New-ModuleResult default (empty array)' {
            $r = New-BatchResult -Success 1
            ,$r.Details | Should -BeOfType [System.Object[]]
            $r.Details.Count | Should -Be 0
        }
    }
}
