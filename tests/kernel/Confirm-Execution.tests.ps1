# ========================================
# Pester v5 unit tests for Confirm-Execution / Wait-KeyPress /
# Confirm-ModuleExecution
# ========================================
# Functions: kernel/common.ps1 :: Confirm-Execution / Wait-KeyPress /
#            Confirm-ModuleExecution
# Run    : pwsh ./dev/run_tests.ps1
#
# Pins the AutoPilot / AutoConfirm short-circuit contract (KERNEL_API.md
# §1.4 / §2). Both globals were introduced in different kernel cycles
# (AutoPilot in 2.x; AutoConfirmMode in 3.1.5 for FlexProfile single-row
# [Run] buttons) and the dual short-circuit on every Y/N prompt and
# Press-Enter wait is the linchpin that separates "unattended profile run"
# from "operator clicks one row to retry it." A regression here either
# breaks unattended runs (operator dialog appears mid-AutoPilot) or breaks
# Flex single-execution (Y/N prompt appears mid-[Run] click), neither
# observable in static analysis.
# ========================================

BeforeAll {
    . "$PSScriptRoot\..\_helpers\test_state.ps1"
    $script:RepoRoot = Get-FabriqRepoRoot
    . (Join-Path $script:RepoRoot 'kernel\common.ps1')
}

Describe 'Confirm-Execution' {

    BeforeEach {
        # Reset both short-circuit globals between cases so the previous
        # test's state never bleeds into the current one.
        $global:AutoPilotMode    = $false
        $global:AutoConfirmMode  = $false
    }

    Context 'AutoPilot short-circuit' {

        It 'returns $true without prompting when AutoPilotMode is $true' {
            $global:AutoPilotMode = $true
            Mock Read-Host { throw 'Read-Host should not be called under AutoPilot' }
            (Confirm-Execution -Message 'proceed?') | Should -BeTrue
            Should -Invoke Read-Host -Times 0 -Exactly
        }
    }

    Context 'AutoConfirm short-circuit (FlexProfile single-execution)' {

        It 'returns $true without prompting when only AutoConfirmMode is $true' {
            $global:AutoConfirmMode = $true
            Mock Read-Host { throw 'Read-Host should not be called under AutoConfirm' }
            (Confirm-Execution -Message 'proceed?') | Should -BeTrue
            Should -Invoke Read-Host -Times 0 -Exactly
        }

        It 'still short-circuits when both AutoPilot and AutoConfirm are $true' {
            # The two flags are not mutually exclusive at runtime; observable
            # behavior must remain "auto-Y" regardless of which one is set.
            $global:AutoPilotMode   = $true
            $global:AutoConfirmMode = $true
            Mock Read-Host { throw 'Read-Host should not be called when either short-circuit is active' }
            (Confirm-Execution -Message 'proceed?') | Should -BeTrue
        }
    }

    Context 'Interactive prompt (Read-Host loop)' {

        It "returns `$true on uppercase 'Y'" {
            Mock Read-Host { 'Y' }
            (Confirm-Execution) | Should -BeTrue
        }

        It "returns `$true on lowercase 'y'" {
            Mock Read-Host { 'y' }
            (Confirm-Execution) | Should -BeTrue
        }

        It "returns `$false on uppercase 'N'" {
            Mock Read-Host { 'N' }
            (Confirm-Execution) | Should -BeFalse
        }

        It "returns `$false on lowercase 'n'" {
            Mock Read-Host { 'n' }
            (Confirm-Execution) | Should -BeFalse
        }

        It 'loops on invalid input until a Y/N is supplied' {
            # First two reads are garbage, third is 'Y'; function should
            # loop past the invalid responses and eventually return $true.
            $script:_replyIdx = 0
            $script:_replies  = @('what','','Y')
            Mock Read-Host {
                $reply = $script:_replies[$script:_replyIdx]
                $script:_replyIdx++
                return $reply
            }
            (Confirm-Execution) | Should -BeTrue
            Should -Invoke Read-Host -Times 3 -Exactly
        }
    }
}

Describe 'Wait-KeyPress' {

    BeforeEach {
        $global:AutoPilotMode    = $false
        $global:AutoConfirmMode  = $false
    }

    It 'skips Read-Host when AutoPilotMode is $true' {
        $global:AutoPilotMode = $true
        Mock Read-Host { throw 'Read-Host should not be called under AutoPilot' }
        Wait-KeyPress -Message 'Press Enter'
        Should -Invoke Read-Host -Times 0 -Exactly
    }

    It 'skips Read-Host when AutoConfirmMode is $true' {
        $global:AutoConfirmMode = $true
        Mock Read-Host { throw 'Read-Host should not be called under AutoConfirm' }
        Wait-KeyPress -Message 'Press Enter'
        Should -Invoke Read-Host -Times 0 -Exactly
    }

    It 'invokes Read-Host once when both short-circuit flags are $false' {
        Mock Read-Host { '' }
        Wait-KeyPress -Message 'Press Enter'
        Should -Invoke Read-Host -Times 1 -Exactly
    }
}

Describe 'Confirm-ModuleExecution' {

    BeforeEach {
        $global:AutoPilotMode    = $false
        $global:AutoConfirmMode  = $false
        $global:_LastModuleResult = $null
    }

    It 'returns $null (proceed signal) under AutoPilotMode' {
        $global:AutoPilotMode = $true
        Mock Read-Host { throw 'Read-Host should not be called under AutoPilot' }
        $r = Confirm-ModuleExecution -Message 'run module?'
        $null -eq $r | Should -BeTrue
    }

    It "returns `$null when the operator answers 'Y'" {
        Mock Read-Host { 'Y' }
        $r = Confirm-ModuleExecution -Message 'run module?'
        $null -eq $r | Should -BeTrue
    }

    It "returns a Cancelled ModuleResult when the operator answers 'N'" {
        Mock Read-Host { 'N' }
        $r = Confirm-ModuleExecution -Message 'run module?'
        $r                | Should -Not -BeNullOrEmpty
        $r._IsModuleResult | Should -BeTrue
        $r.Status         | Should -Be 'Cancelled'
        $r.Message        | Should -Be 'User canceled'
    }
}
