# ========================================
# Pester v5 unit tests for Invoke-SafeCommand
# ========================================
# Function: kernel/common.ps1 :: Invoke-SafeCommand
# Run    : pwsh ./dev/run_tests.ps1
#
# Pins the fail-closed result-contract behavior (TM t-0005): a module
# script that completes WITHOUT returning a ModuleResult is recorded as
# Error (not Success), while every legitimate return path - pipeline
# capture, $global:_LastModuleResult fallback, module-reported status,
# thrown exception - keeps its existing semantics. A regression back to
# fail-open would let a broken module sail through as a silent false
# Success in execution history and the HTML checklist.
#
# Invoke-SafeCommandAsync shares the same extraction/fail-closed shape
# but spawns a real runspace (CWD-dependent, slow); it is covered by
# the manual async verification documented in CHANGELOG / t-0005 notes,
# not by this file.
# ========================================

BeforeAll {
    . "$PSScriptRoot\..\_helpers\test_state.ps1"
    $script:RepoRoot = Get-FabriqRepoRoot
    . (Join-Path $script:RepoRoot 'kernel\common.ps1')
}

Describe 'Invoke-SafeCommand' {

    Context 'Fail-closed: no ModuleResult returned (contract violation)' {

        It 'records Error when the scriptblock returns nothing' {
            $r = Invoke-SafeCommand -OperationName 'NoResult' -ScriptBlock { } -ContinueOnError
            $r.Status  | Should -Be 'Error'
            $r.Success | Should -BeFalse
            $r.Message | Should -Match 'No ModuleResult returned'
        }

        It 'records Error when the scriptblock returns a plain hashtable (windows_update shape)' {
            $r = Invoke-SafeCommand -OperationName 'HashtableResult' -ScriptBlock {
                @{ Status = 'Success'; RebootRequired = $false }
            } -ContinueOnError
            $r.Status  | Should -Be 'Error'
            $r.Success | Should -BeFalse
            $r.Message | Should -Match 'No ModuleResult returned'
        }

        It 'records Error when the scriptblock returns a PSCustomObject without the result marker' {
            $r = Invoke-SafeCommand -OperationName 'UnmarkedObject' -ScriptBlock {
                [PSCustomObject]@{ Foo = 1 }
            } -ContinueOnError
            $r.Status  | Should -Be 'Error'
            $r.Success | Should -BeFalse
        }

        It 'leaves Verified $null on the contract-violation path' {
            $r = Invoke-SafeCommand -OperationName 'NoResultVerified' -ScriptBlock { } -ContinueOnError
            $r.Verified | Should -Be $null
        }
    }

    Context 'ModuleResult passthrough (regression pin)' {

        It 'passes through a Success ModuleResult from the pipeline' {
            $r = Invoke-SafeCommand -OperationName 'PipeSuccess' -ScriptBlock {
                New-ModuleResult -Status 'Success' -Message 'done'
            } -ContinueOnError
            $r.Status  | Should -Be 'Success'
            $r.Success | Should -BeTrue
            $r.Message | Should -Be 'done'
        }

        It 'passes through an Error ModuleResult reported by the module' {
            $r = Invoke-SafeCommand -OperationName 'PipeError' -ScriptBlock {
                New-ModuleResult -Status 'Error' -Message 'module says no'
            } -ContinueOnError
            $r.Status  | Should -Be 'Error'
            $r.Success | Should -BeFalse
            $r.Message | Should -Be 'module says no'
        }

        It 'passes through Cancelled without converting it to Error' {
            $r = Invoke-SafeCommand -OperationName 'PipeCancelled' -ScriptBlock {
                New-ModuleResult -Status 'Cancelled' -Message 'user said N'
            } -ContinueOnError
            $r.Status  | Should -Be 'Cancelled'
            $r.Success | Should -BeFalse
        }

        It 'passes through Skipped without converting it to Error' {
            $r = Invoke-SafeCommand -OperationName 'PipeSkipped' -ScriptBlock {
                New-ModuleResult -Status 'Skipped' -Message 'nothing to do'
            } -ContinueOnError
            $r.Status  | Should -Be 'Skipped'
            $r.Success | Should -BeFalse
        }

        It 'passes through the Verified flag' {
            $r = Invoke-SafeCommand -OperationName 'PipeVerified' -ScriptBlock {
                New-ModuleResult -Status 'Success' -Message 'v' -Verified $true
            } -ContinueOnError
            $r.Verified | Should -BeTrue
        }

        It 'finds the ModuleResult among extra pipeline noise' {
            $r = Invoke-SafeCommand -OperationName 'PipeNoise' -ScriptBlock {
                'stray string output'
                42
                New-ModuleResult -Status 'Success' -Message 'found me'
            } -ContinueOnError
            $r.Status  | Should -Be 'Success'
            $r.Message | Should -Be 'found me'
        }
    }

    Context 'Global fallback recovery (regression pin)' {

        It 'recovers a ModuleResult captured into a variable (empty pipeline)' {
            $r = Invoke-SafeCommand -OperationName 'GlobalFallback' -ScriptBlock {
                $captured = New-ModuleResult -Status 'Partial' -Message 'via-global'
                return
            } -ContinueOnError
            $r.Status  | Should -Be 'Partial'
            $r.Message | Should -Be 'via-global'
        }
    }

    Context 'Exception path (regression pin)' {

        It 'records Error with the exception message when the scriptblock throws' {
            $r = Invoke-SafeCommand -OperationName 'Throws' -ScriptBlock {
                throw 'boom'
            } -ContinueOnError
            $r.Status  | Should -Be 'Error'
            $r.Success | Should -BeFalse
            $r.Message | Should -Match 'boom'
        }
    }
}
