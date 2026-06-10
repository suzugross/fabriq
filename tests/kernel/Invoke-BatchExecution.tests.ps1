# ========================================
# Pester v5 unit tests for Invoke-BatchExecution
# ========================================
# Function: kernel/main.ps1 :: Invoke-BatchExecution
# Run    : powershell.exe -File ./dev/run_tests.ps1
#
# Pins the batch-orchestration branches that a field kitting run depends
# on (TM t-0007 (3)):
#   - __RESTART__ chain: Save-ResumeState -> Register-FabriqRunOnce,
#     including the failure chain (RunOnce fails -> Remove-ResumeState ->
#     record Error -> KEEP GOING) and the success early-exit (resume
#     state survives, Invoke-CountdownRestart fires, finalize skipped).
#   - AutoPilot ErrorMode: retry (bounded by $script:AutoPilotMaxRetry),
#     skip (record and continue), empty (dialog fallback).
#   - Attended error notification without AutoPilot.
#   - FinalizeOnComplete routing to Complete-ProfileExecution.
#   - finally contract: $script:LastBatchResults is published and
#     AutoPilot globals are reset even when the loop throws mid-batch.
#
# The function is extracted from main.ps1 via AST
# (Get-FabriqMainFunctionScriptBlock) - main.ps1 itself is not
# dot-sourceable (top-level startup side effects).
#
# SAFETY: Invoke-BatchExecution's dependency surface includes functions
# that reboot the machine, write HKLM, kill explorer.exe and take real
# screenshots. Two layers of protection:
#   1. BeforeEach mocks every machine-affecting dependency.
#   2. BeforeAll defines THROWING STUBS underneath the mocks, so if a
#      future edit drops a mock, the call fails the test loudly instead
#      of acting on the dev machine. Do not remove the stubs.
# Real (unmocked) on purpose: Save-ResumeState / Remove-ResumeState /
# Load path - they only touch the temp resume file injected via
# Set-FabriqTestState, and the on-disk round trip IS the contract under
# test for the __RESTART__ chain.
# ========================================

BeforeAll {
    . "$PSScriptRoot\..\_helpers\test_state.ps1"
    $script:RepoRoot = Get-FabriqRepoRoot
    . (Join-Path $script:RepoRoot 'kernel\common.ps1')

    # Inject Invoke-BatchExecution into the test scope (AST extraction).
    $sb = Get-FabriqMainFunctionScriptBlock -Name 'Invoke-BatchExecution'
    . $sb

    # ---- SAFETY STUBS (see header) ------------------------------------
    function Invoke-CountdownRestart   { throw 'SAFETY STUB: Invoke-CountdownRestart called without a Pester mock' }
    function Register-FabriqRunOnce    { throw 'SAFETY STUB: Register-FabriqRunOnce called without a Pester mock' }
    function Capture-ScreenEvidence    { throw 'SAFETY STUB: Capture-ScreenEvidence called without a Pester mock' }
    function Show-AutoPilotErrorDialog { throw 'SAFETY STUB: Show-AutoPilotErrorDialog (blocking dialog) called without a Pester mock' }
    function Stop-Process              { throw 'SAFETY STUB: Stop-Process called without a Pester mock' }
    function Start-Process             { throw 'SAFETY STUB: Start-Process called without a Pester mock' }
    function Restart-Computer          { throw 'SAFETY STUB: Restart-Computer called without a Pester mock' }
    # --------------------------------------------------------------------

    # Module fixture builder. Mirrors the attribute set Resolve-ProfileModules
    # attaches to profile rows (_IsRestart / _IsAsync / _ErrorMode etc.).
    # _IsAsync stays $false so the dispatch short-circuits to the sync path
    # without consulting async_config.json.
    function New-BatchModule {
        param(
            [Parameter(Mandatory)][int]$Order,
            [string]$MenuName = '',
            [switch]$Restart,
            [string]$ErrorMode = ''
        )
        if (-not $MenuName) {
            $MenuName = if ($Restart) { '[RESTART]' } else { "Module$Order" }
        }
        [PSCustomObject]@{
            Order          = $Order
            MenuName       = $MenuName
            Category       = 'Test'
            Script         = "C:\fake\module$Order.ps1"
            _IsRestart     = [bool]$Restart
            _IsReexplorer  = $false
            _IsAsync       = $false
            _ErrorMode     = $ErrorMode
            _AutoLogonUser = $null
            _Segment       = $null
        }
    }

    # Result object shaped like Invoke-SafeCommand's return value.
    function New-FakeRunResult {
        param([string]$Status = 'Success', [string]$Message = 'ok', $Verified = $null)
        [PSCustomObject]@{
            Operation = 'fake'
            Success   = ($Status -eq 'Success')
            Status    = $Status
            Message   = $Message
            Duration  = [TimeSpan]::Zero
            Error     = $null
            Verified  = $Verified
        }
    }
}

Describe 'Invoke-BatchExecution' {

    BeforeEach {
        # Isolated resume-state file per test (real Save/Remove run against it).
        # StatusFilePath is redirected too: the real Clear-ExecutionResults
        # calls Write-StatusFile, which would otherwise drop a runtime
        # status.json into the repo's kernel/json/ during test runs.
        $script:tmpResume = Join-Path $env:TEMP `
            ("fabriq-batch-resume-{0}.json" -f ([guid]::NewGuid().ToString('N')))
        $script:tmpStatus = Join-Path $env:TEMP `
            ("fabriq-batch-status-{0}.json" -f ([guid]::NewGuid().ToString('N')))
        Set-FabriqTestState -ResumeStatePath $script:tmpResume -SessionID 'batch-test' `
            -StatusFilePath $script:tmpStatus

        # Script/global state the function (and real Save-ResumeState) reads.
        $script:AutoPilotMaxRetry      = 2
        $script:LastBatchResults       = 'sentinel-not-yet-published'
        $global:AutoPilotMode          = $false
        $global:AutoPilotWaitSec       = 0
        $global:FabriqMasterPassphrase = $null
        $global:FabriqEvidenceBasePath = 'C:\test\evidence'

        # Unconditional interception of every machine-affecting or
        # prompting dependency (argument checks live in Should -Invoke).
        Mock Show-BatchConfirmation    { $true }
        Mock Restore-ExecutionHistory  { }
        Mock Write-KernelTelemetryEvent { }
        Mock Invoke-SafeCommand        { New-FakeRunResult -Status 'Success' }
        Mock Invoke-SafeCommandAsync   { New-FakeRunResult -Status 'Success' }
        Mock Register-FabriqRunOnce    { $true }
        Mock Invoke-CountdownRestart   { }
        Mock Capture-ScreenEvidence    { }
        Mock Add-ExecutionResult       { }
        Mock Write-ExecutionHistory    { }
        Mock Invoke-ErrorNotification  { }
        Mock Show-AutoPilotErrorDialog { 'Skip' }
        Mock Complete-ProfileExecution { }
        Mock Show-ExecutionSummary     { }
        Mock Start-Sleep               { }
    }

    AfterEach {
        if ($script:tmpResume -and (Test-Path $script:tmpResume)) {
            Remove-Item $script:tmpResume -Force -ErrorAction SilentlyContinue
        }
        if ($script:tmpStatus) {
            Remove-Item $script:tmpStatus -Force -ErrorAction SilentlyContinue
            Remove-Item "$($script:tmpStatus).tmp" -Force -ErrorAction SilentlyContinue
        }
        $global:AutoPilotMode    = $false
        $global:AutoPilotWaitSec = 3
    }

    Context 'Cancel path' {

        It 'runs nothing when the operator declines the confirmation, but still publishes LastBatchResults' {
            Mock Show-BatchConfirmation { $false }
            $mods = @((New-BatchModule -Order 10), (New-BatchModule -Order 20))

            Invoke-BatchExecution -SelectedModules $mods

            Should -Invoke Invoke-SafeCommand -Exactly -Times 0
            ($null -eq $script:LastBatchResults) | Should -BeFalse
            @($script:LastBatchResults).Count | Should -Be 0
        }
    }

    Context '__RESTART__ marker' {

        It 'on RunOnce success: saves resume state, fires the countdown restart, and exits before later modules and finalize' {
            $mods = @(
                (New-BatchModule -Order 10),
                (New-BatchModule -Order 20 -Restart),
                (New-BatchModule -Order 30)
            )

            Invoke-BatchExecution -SelectedModules $mods `
                -ProfilePath 'C:\profiles\test.csv' -ProfileName 'TestProfile'

            # Only the pre-marker module ran; the marker exited the batch.
            Should -Invoke Invoke-SafeCommand -Exactly -Times 1
            Should -Invoke Invoke-CountdownRestart -Exactly -Times 1
            Should -Invoke Complete-ProfileExecution -Exactly -Times 0

            # Resume state survives for the post-reboot leg and points
            # past the marker's Order.
            Test-Path $script:tmpResume | Should -BeTrue
            $state = Get-Content $script:tmpResume -Raw | ConvertFrom-Json
            $state.ResumeAfterOrder | Should -Be 20
        }

        It 'on RunOnce failure: removes the resume state, records the Error, skips the restart and keeps executing' {
            Mock Register-FabriqRunOnce { $false }
            $mods = @(
                (New-BatchModule -Order 10),
                (New-BatchModule -Order 20 -Restart),
                (New-BatchModule -Order 30)
            )

            Invoke-BatchExecution -SelectedModules $mods `
                -ProfilePath 'C:\profiles\test.csv' -ProfileName 'TestProfile'

            # No reboot; the failure is recorded; the batch continued to
            # the module after the marker.
            Should -Invoke Invoke-CountdownRestart -Exactly -Times 0
            Should -Invoke Add-ExecutionResult -Exactly -Times 1 -ParameterFilter {
                $Operation -eq '[RESTART]' -and $Status -eq 'Error'
            }
            Should -Invoke Invoke-SafeCommand -Exactly -Times 2

            # No stale resume file is left behind (failure-path Remove +
            # natural-completion Remove both target the temp file).
            Test-Path $script:tmpResume | Should -BeFalse
        }

        It 'outside a profile run (no ProfilePath): silently skips the marker and never saves resume state' {
            $mods = @(
                (New-BatchModule -Order 10),
                (New-BatchModule -Order 20 -Restart),
                (New-BatchModule -Order 30)
            )

            Invoke-BatchExecution -SelectedModules $mods

            Should -Invoke Invoke-SafeCommand -Exactly -Times 2
            Should -Invoke Invoke-CountdownRestart -Exactly -Times 0
            Should -Invoke Register-FabriqRunOnce -Exactly -Times 0
            Test-Path $script:tmpResume | Should -BeFalse
        }
    }

    Context 'AutoPilot ErrorMode' {

        It 'retry: re-runs a failing module up to AutoPilotMaxRetry, then records the Error once' {
            Mock Invoke-SafeCommand { New-FakeRunResult -Status 'Error' -Message 'always fails' }
            $mods = @((New-BatchModule -Order 10 -ErrorMode 'retry'))

            Invoke-BatchExecution -SelectedModules $mods -AutoPilot

            # 1 initial attempt + 2 retries (AutoPilotMaxRetry seeded to 2)
            Should -Invoke Invoke-SafeCommand -Exactly -Times 3
            Should -Invoke Add-ExecutionResult -Exactly -Times 1 -ParameterFilter {
                $Operation -eq 'Module10' -and $Status -eq 'Error'
            }
        }

        It 'skip: records the Error and continues to the next module without retrying' {
            $script:safeCmdCalls = 0
            Mock Invoke-SafeCommand {
                $script:safeCmdCalls++
                if ($script:safeCmdCalls -eq 1) { New-FakeRunResult -Status 'Error' -Message 'first fails' }
                else { New-FakeRunResult -Status 'Success' }
            }
            $mods = @(
                (New-BatchModule -Order 10 -ErrorMode 'skip'),
                (New-BatchModule -Order 20)
            )

            Invoke-BatchExecution -SelectedModules $mods -AutoPilot

            Should -Invoke Invoke-SafeCommand -Exactly -Times 2
            Should -Invoke Add-ExecutionResult -Exactly -Times 1 -ParameterFilter {
                $Operation -eq 'Module10' -and $Status -eq 'Error'
            }
            Should -Invoke Add-ExecutionResult -Exactly -Times 1 -ParameterFilter {
                $Operation -eq 'Module20' -and $Status -eq 'Success'
            }
        }

        It 'empty ErrorMode: falls back to the operator dialog (mocked Skip = no retry)' {
            Mock Invoke-SafeCommand { New-FakeRunResult -Status 'Error' -Message 'fails' }
            $mods = @((New-BatchModule -Order 10))

            Invoke-BatchExecution -SelectedModules $mods -AutoPilot

            Should -Invoke Show-AutoPilotErrorDialog -Exactly -Times 1
            Should -Invoke Invoke-SafeCommand -Exactly -Times 1
        }
    }

    Context 'Attended (non-AutoPilot) error handling' {

        It 'fires the error notification once and does not retry or open the AutoPilot dialog' {
            Mock Invoke-SafeCommand { New-FakeRunResult -Status 'Error' -Message 'fails' }
            $mods = @((New-BatchModule -Order 10))

            Invoke-BatchExecution -SelectedModules $mods

            Should -Invoke Invoke-ErrorNotification -Exactly -Times 1
            Should -Invoke Show-AutoPilotErrorDialog -Exactly -Times 0
            Should -Invoke Invoke-SafeCommand -Exactly -Times 1
        }
    }

    Context 'Finalize routing' {

        It 'natural completion of a profile run fires Complete-ProfileExecution in Auto mode' {
            $mods = @((New-BatchModule -Order 10), (New-BatchModule -Order 20))

            Invoke-BatchExecution -SelectedModules $mods `
                -ProfilePath 'C:\profiles\test.csv' -ProfileName 'TestProfile'

            Should -Invoke Complete-ProfileExecution -Exactly -Times 1 -ParameterFilter {
                $ProfileName -eq 'TestProfile' -and $Mode -eq 'Auto'
            }
            @($script:LastBatchResults).Count | Should -Be 2
            @($script:LastBatchResults)[0].Status | Should -Be 'Success'
            @($script:LastBatchResults)[1].Status | Should -Be 'Success'
        }

        It '-FinalizeOnComplete:$false suppresses Complete-ProfileExecution (FlexProfile contract)' {
            $mods = @((New-BatchModule -Order 10))

            Invoke-BatchExecution -SelectedModules $mods `
                -ProfilePath 'C:\profiles\test.csv' -ProfileName 'TestProfile' `
                -FinalizeOnComplete:$false

            Should -Invoke Complete-ProfileExecution -Exactly -Times 0
        }
    }

    Context 'finally contract (mid-batch throw)' {

        It 'publishes the partial LastBatchResults snapshot and resets AutoPilot globals when the loop throws' {
            $script:safeCmdCalls = 0
            Mock Invoke-SafeCommand {
                $script:safeCmdCalls++
                if ($script:safeCmdCalls -ge 2) { throw 'harness exploded' }
                New-FakeRunResult -Status 'Success'
            }
            $mods = @((New-BatchModule -Order 10), (New-BatchModule -Order 20))

            { Invoke-BatchExecution -SelectedModules $mods -AutoPilot } |
                Should -Throw '*harness exploded*'

            # Module 1 completed before the throw; the snapshot reflects it.
            @($script:LastBatchResults).Count | Should -Be 1
            @($script:LastBatchResults)[0].MenuName | Should -Be 'Module10'
            # AutoPilot flag set by -AutoPilot must be reset by finally.
            $global:AutoPilotMode | Should -BeFalse
        }
    }
}
