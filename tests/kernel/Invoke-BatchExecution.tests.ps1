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
#   - Profile identity (TM t-0111): results are recorded with the profile
#     name, and the __GATE__ seed only trusts this profile's history rows.
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
#
# A second Describe at the end of the file replays the t-0111 field
# report end to end (profile A batch -> profile B dashboard state and
# checklist) with the history functions real against a temp file.
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
            [switch]$Gate,
            [string]$ErrorMode = ''
        )
        if (-not $MenuName) {
            $MenuName = if ($Restart) { '[RESTART]' } elseif ($Gate) { '[GATE]' } else { "Module$Order" }
        }
        [PSCustomObject]@{
            Order          = $Order
            MenuName       = $MenuName
            Category       = 'Test'
            Script         = "C:\fake\module$Order.ps1"
            _IsRestart     = [bool]$Restart
            _IsGate        = [bool]$Gate
            _IsReexplorer  = $false
            _IsAsync       = $false
            _ErrorMode     = $ErrorMode
            _AutoLogonUser = $null
            _Segment       = $null
        }
    }

    # One execution_history.csv row as Import-ExecutionHistory returns it
    # (all cells are strings), in the session Set-FabriqTestState selects.
    function New-HistoryRow {
        param([string]$ModuleName, [string]$Status, [int]$Order, [string]$ProfileName)
        [PSCustomObject]@{
            ModuleName = $ModuleName; Status = $Status; Verified = ''
            Order = "$Order"; Profile = $ProfileName; SessionID = 'batch-test'
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

        It 'on RunOnce failure: removes the resume state, records the Error, aborts the restart and STOPS (fail-closed) - post-reboot modules do not run' {
            Mock Register-FabriqRunOnce { $false }
            $mods = @(
                (New-BatchModule -Order 10),
                (New-BatchModule -Order 20 -Restart),
                (New-BatchModule -Order 30)
            )

            Invoke-BatchExecution -SelectedModules $mods `
                -ProfilePath 'C:\profiles\test.csv' -ProfileName 'TestProfile'

            # No reboot; the failure is recorded; the batch STOPS at the marker
            # (fail-closed, TM t-0045) so the post-reboot module (Order 30) does
            # NOT run on the un-rebooted/stale system state.
            Should -Invoke Invoke-CountdownRestart -Exactly -Times 0
            Should -Invoke Add-ExecutionResult -Exactly -Times 1 -ParameterFilter {
                $Operation -eq '[RESTART]' -and $Status -eq 'Error'
            }
            # Only the pre-marker module (Order 10) ran; Order 30 was skipped.
            Should -Invoke Invoke-SafeCommand -Exactly -Times 1

            # Unlike the RunOnce-success path (which returns before finalize so
            # the post-reboot resume can finalize later), the abort falls through
            # to natural completion and finalizes once with the Error recorded.
            Should -Invoke Complete-ProfileExecution -Exactly -Times 1

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

    Context 'Profile identity of recorded results (TM t-0111)' {
        # Order is local to a profile. Every result recorded for a Profile
        # row carries the profile name, so that the dashboard / checklist /
        # gate of ANOTHER profile opened in the same session cannot mistake
        # it for their own same-Order row.

        It 'records module results, history rows and LastBatchResults with the profile name' {
            $mods = @((New-BatchModule -Order 10), (New-BatchModule -Order 20))

            Invoke-BatchExecution -SelectedModules $mods `
                -ProfilePath 'C:\profiles\test.csv' -ProfileName 'TestProfile' -FinalizeOnComplete:$false

            Should -Invoke Add-ExecutionResult -Exactly -Times 2 -ParameterFilter { $ProfileName -eq 'TestProfile' }
            Should -Invoke Write-ExecutionHistory -Exactly -Times 2 -ParameterFilter { $ProfileName -eq 'TestProfile' }
            @($script:LastBatchResults).Count | Should -Be 2
            @($script:LastBatchResults | Where-Object { $_.ProfileName -eq 'TestProfile' }).Count | Should -Be 2
        }

        It 'records the [RESTART] marker with the profile name' {
            $mods = @((New-BatchModule -Order 10), (New-BatchModule -Order 20 -Restart))

            Invoke-BatchExecution -SelectedModules $mods `
                -ProfilePath 'C:\profiles\test.csv' -ProfileName 'TestProfile'

            Should -Invoke Write-ExecutionHistory -Exactly -Times 1 -ParameterFilter {
                $ModuleName -eq '[RESTART]' -and $Order -eq 20 -and $ProfileName -eq 'TestProfile'
            }
        }

        It 'records a menu-mode batch (no profile) with an EMPTY profile name' {
            $mods = @((New-BatchModule -Order 10))

            Invoke-BatchExecution -SelectedModules $mods

            Should -Invoke Add-ExecutionResult -Exactly -Times 1 -ParameterFilter { $ProfileName -eq '' }
            Should -Invoke Write-ExecutionHistory -Exactly -Times 1 -ParameterFilter { $ProfileName -eq '' }
            @($script:LastBatchResults)[0].ProfileName | Should -Be ''
        }
    }

    Context '__GATE__ seed from session history is bound to the profile (TM t-0111)' {
        # Profile under test: Module10, [GATE] at 20, Module30. Only Module30
        # is selected (sparse Flex run), so whether it runs depends solely on
        # what the history says about Order 10.

        BeforeEach {
            $script:savedKanri = $env:SELECTED_KANRI_NO
            $env:SELECTED_KANRI_NO = 'K-1'
            $script:gateProfile = @(
                (New-BatchModule -Order 10),
                (New-BatchModule -Order 20 -Gate),
                (New-BatchModule -Order 30)
            )
            $script:fakeHistory = @()
            # Import-ExecutionHistory returns newest first.
            Mock Import-ExecutionHistory { $script:fakeHistory }
        }

        AfterEach {
            $env:SELECTED_KANRI_NO = $script:savedKanri
        }

        It 'blocks past the gate on THIS profile''s own Error (baseline)' {
            $script:fakeHistory = @((New-HistoryRow -ModuleName 'Module10' -Status 'Error' -Order 10 -ProfileName 'TestProfile'))

            Invoke-BatchExecution -SelectedModules @($script:gateProfile[2]) -FullProfileModules $script:gateProfile `
                -ProfilePath 'C:\profiles\test.csv' -ProfileName 'TestProfile' -FinalizeOnComplete:$false

            Should -Invoke Invoke-SafeCommand -Exactly -Times 0
        }

        It 'is not blocked by another profile''s Error at the same Order and MenuName' {
            $script:fakeHistory = @((New-HistoryRow -ModuleName 'Module10' -Status 'Error' -Order 10 -ProfileName 'Other'))

            Invoke-BatchExecution -SelectedModules @($script:gateProfile[2]) -FullProfileModules $script:gateProfile `
                -ProfilePath 'C:\profiles\test.csv' -ProfileName 'TestProfile' -FinalizeOnComplete:$false

            Should -Invoke Invoke-SafeCommand -Exactly -Times 1
        }

        It 'stays blocked when another profile''s NEWER Success sits on top of this profile''s Error' {
            # The seed takes the first row per Order (= latest). An unscoped
            # seed would take the other profile's Success and open the gate.
            $script:fakeHistory = @(
                (New-HistoryRow -ModuleName 'Module10' -Status 'Success' -Order 10 -ProfileName 'Other'),
                (New-HistoryRow -ModuleName 'Module10' -Status 'Error'   -Order 10 -ProfileName 'TestProfile')
            )

            Invoke-BatchExecution -SelectedModules @($script:gateProfile[2]) -FullProfileModules $script:gateProfile `
                -ProfilePath 'C:\profiles\test.csv' -ProfileName 'TestProfile' -FinalizeOnComplete:$false

            Should -Invoke Invoke-SafeCommand -Exactly -Times 0
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

    Context 'Profile data overlay context lifetime (dev/PROFILE_DATA_OVERLAY_PLAN.md)' {

        BeforeEach {
            # profiles/<name>.csv + sibling profiles/<name>/ data folder
            # GetFullPath: $env:TEMP may be an 8.3 short name that the kernel
            # expands when it normalizes the profile path.
            $script:ovRoot    = [System.IO.Path]::GetFullPath((Join-Path $env:TEMP ("fabriq-pdf-batch-{0}" -f ([guid]::NewGuid().ToString('N')))))
            $script:ovProfile = Join-Path $script:ovRoot 'Master_OV.csv'
            $script:ovFolder  = Join-Path $script:ovRoot 'Master_OV'
            $null = New-Item -ItemType Directory -Path $script:ovFolder -Force
            [System.IO.File]::WriteAllLines($script:ovProfile, @('Order,ScriptPath,Enabled'), [System.Text.Encoding]::ASCII)
            Clear-FabriqProfileDataContext
            $script:seenDataDir = 'not-captured'
        }

        AfterEach {
            Clear-FabriqProfileDataContext
            # Destructive-path guard (CLAUDE.md section 8): only the temp
            # folder this context created.
            $tempRoot = [System.IO.Path]::GetFullPath($env:TEMP)
            if ($script:ovRoot -and $script:ovRoot.StartsWith($tempRoot, [System.StringComparison]::OrdinalIgnoreCase) `
                -and $script:ovRoot -like '*fabriq-pdf-batch-*' -and (Test-Path $script:ovRoot)) {
                Remove-Item $script:ovRoot -Recurse -Force -ErrorAction SilentlyContinue
            }
        }

        It 'exposes FABRIQ_PROFILE_DATA_DIR to modules during the batch and clears it afterwards' {
            Mock Invoke-SafeCommand {
                $script:seenDataDir = "$env:FABRIQ_PROFILE_DATA_DIR"
                New-FakeRunResult -Status 'Success'
            }
            $mods = @((New-BatchModule -Order 10))

            Invoke-BatchExecution -SelectedModules $mods `
                -ProfilePath $script:ovProfile -ProfileName 'Master_OV' -FinalizeOnComplete:$false

            $script:seenDataDir | Should -Be $script:ovFolder
            [string]::IsNullOrEmpty($env:FABRIQ_PROFILE_DATA_DIR) | Should -BeTrue
        }

        It 'sets no context for a profile without a data folder (backward compatibility)' {
            Remove-Item $script:ovFolder -Recurse -Force
            Mock Invoke-SafeCommand {
                $script:seenDataDir = "$env:FABRIQ_PROFILE_DATA_DIR"
                New-FakeRunResult -Status 'Success'
            }
            $mods = @((New-BatchModule -Order 10))

            Invoke-BatchExecution -SelectedModules $mods `
                -ProfilePath $script:ovProfile -ProfileName 'Master_OV' -FinalizeOnComplete:$false

            $script:seenDataDir | Should -Be ''
        }

        It 'sets no context for a menu-mode batch without a ProfilePath' {
            Mock Invoke-SafeCommand {
                $script:seenDataDir = "$env:FABRIQ_PROFILE_DATA_DIR"
                New-FakeRunResult -Status 'Success'
            }
            $mods = @((New-BatchModule -Order 10))

            Invoke-BatchExecution -SelectedModules $mods

            $script:seenDataDir | Should -Be ''
        }

        It 'clears the context even when the loop throws mid-batch' {
            Mock Invoke-SafeCommand {
                $script:seenDataDir = "$env:FABRIQ_PROFILE_DATA_DIR"
                throw 'harness exploded'
            }
            $mods = @((New-BatchModule -Order 10))

            { Invoke-BatchExecution -SelectedModules $mods `
                -ProfilePath $script:ovProfile -ProfileName 'Master_OV' -FinalizeOnComplete:$false } |
                Should -Throw '*harness exploded*'

            $script:seenDataDir | Should -Be $script:ovFolder
            [string]::IsNullOrEmpty($env:FABRIQ_PROFILE_DATA_DIR) | Should -BeTrue
        }

        It 'records the data folder in resume_state.json at a __RESTART__ marker' {
            $mods = @((New-BatchModule -Order 10), (New-BatchModule -Order 20 -Restart), (New-BatchModule -Order 30))

            Invoke-BatchExecution -SelectedModules $mods `
                -ProfilePath $script:ovProfile -ProfileName 'Master_OV' -AutoPilot

            (Test-Path $script:tmpResume) | Should -BeTrue
            $saved = Get-Content $script:tmpResume -Raw | ConvertFrom-Json
            $saved.ProfileDataDir | Should -Be $script:ovFolder
        }
    }
}

# ========================================
# End to end: two profiles in one session (TM t-0111)
# ========================================
# The field report, replayed through the real write and read paths: run
# one module of profile A via Invoke-BatchExecution, then look at profile
# B the way the FlexProfile dashboard and the HTML checklist do. Unlike
# the Describe above, Add-ExecutionResult / Write-ExecutionHistory /
# Restore-ExecutionHistory / Import-ExecutionHistory are REAL here (temp
# history file), so this also pins that the profile name the batch writes
# is the key the readers match on. Machine-affecting dependencies stay
# mocked, on top of the throwing safety stubs from BeforeAll.
# ========================================
Describe 'Profile isolation end to end (batch -> dashboard state / checklist)' {

    BeforeAll {
        . (Join-Path $script:RepoRoot 'apps\fabriq_operator\lib\flex_dashboard.ps1')

        $script:e2eProfileA = @(
            (New-BatchModule -Order 10 -MenuName 'Hostname Change'),
            (New-BatchModule -Order 20 -MenuName 'Network Config')
        )
        $script:e2eProfileB = @(
            (New-BatchModule -Order 10 -MenuName 'App Install'),
            (New-BatchModule -Order 20 -MenuName 'Printer Install')
        )

        # What the dashboard's per-row [Run] does (Invoke-FlexProfileLoop
        # "RunSingle"): one module, Flex mode, finalize left to [Complete].
        function Invoke-FlexSingleRun {
            param([array]$ProfileRows, [string]$ProfileName, [int]$Order)
            $target = @($ProfileRows | Where-Object { $_.Order -eq $Order })[0]
            Invoke-BatchExecution -SelectedModules @($target) -FullProfileModules $ProfileRows `
                -ProfilePath "C:\profiles\$ProfileName.csv" -ProfileName $ProfileName `
                -FinalizeOnComplete:$false -ExecutionMode 'Flex' -SelectedOrders @($Order)
        }

        function Get-ChecklistHtml {
            param([array]$ProfileRows, [string]$ProfileName)
            $path = Export-HtmlChecklist -ProfileName $ProfileName -ProfilePath '' `
                -DefinedModules $ProfileRows -ExecutionResults $script:ExecutionResults
            return (Get-Content $path -Raw)
        }
    }

    BeforeEach {
        $script:e2eRoot = [System.IO.Path]::GetFullPath((Join-Path $env:TEMP ("fabriq-batch-e2e-{0}" -f ([guid]::NewGuid().ToString('N')))))
        $null = New-Item -ItemType Directory -Path $script:e2eRoot -Force
        Set-FabriqTestState -ResumeStatePath (Join-Path $script:e2eRoot 'resume_state.json') `
            -HistoryPath (Join-Path $script:e2eRoot 'history.csv') `
            -StatusFilePath (Join-Path $script:e2eRoot 'status.json') `
            -SessionID 'e2e-session'

        $script:ExecutionResults       = @()
        $script:LastBatchResults       = @()
        $script:SessionInfo            = $null
        $script:AutoPilotMaxRetry      = 2
        $global:AutoPilotMode          = $false
        $global:AutoPilotWaitSec       = 0
        $global:FabriqMasterPassphrase = $null
        $global:FabriqEvidenceBasePath = Join-Path $script:e2eRoot 'evidence'
        $global:FabriqUniqueId         = 'TESTUID'
        $script:savedKanri = $env:SELECTED_KANRI_NO
        $env:SELECTED_KANRI_NO = 'K-1'

        $script:e2eNextResult = { New-FakeRunResult -Status 'Success' -Message 'renamed to PC-001' -Verified $true }

        Mock Show-BatchConfirmation    { $true }
        Mock Write-KernelTelemetryEvent { }
        Mock Invoke-SafeCommand        { & $script:e2eNextResult }
        Mock Invoke-SafeCommandAsync   { & $script:e2eNextResult }
        Mock Register-FabriqRunOnce    { $true }
        Mock Invoke-CountdownRestart   { }
        Mock Capture-ScreenEvidence    { }
        Mock Invoke-ErrorNotification  { }
        Mock Show-AutoPilotErrorDialog { 'Skip' }
        Mock Complete-ProfileExecution { }
        Mock Show-ExecutionSummary     { }
        Mock Start-Sleep               { }
        # Hardware / WMI surfaces read by Export-HtmlChecklist
        Mock Get-CurrentPCInfo {
            @{ ComputerName = 'TEST-PC'; EthernetIP = ''; EthernetSubnet = ''; EthernetGateway = ''
               WifiIP = ''; WifiSubnet = ''; WifiGateway = ''; DNS = @(); Printers = @() }
        }
        Mock Get-WmiObject { }
        Mock Get-BitLockerVolume { }
    }

    AfterEach {
        $env:SELECTED_KANRI_NO = $script:savedKanri
        $global:FabriqEvidenceBasePath = $null
        $global:AutoPilotMode    = $false
        $global:AutoPilotWaitSec = 3
        $script:ExecutionResults = @()
        $script:LastBatchResults = @()
        # Destructive-path guard (CLAUDE.md section 8): only the temp
        # folder this test created.
        $tempRoot = [System.IO.Path]::GetFullPath($env:TEMP)
        if ($script:e2eRoot -and $script:e2eRoot.StartsWith($tempRoot, [System.StringComparison]::OrdinalIgnoreCase) `
            -and $script:e2eRoot -like '*fabriq-batch-e2e-*' -and (Test-Path $script:e2eRoot)) {
            Remove-Item $script:e2eRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    It 'a module run in profile A leaves every row of profile B Pending on the dashboard' {
        Invoke-FlexSingleRun -ProfileRows $script:e2eProfileA -ProfileName 'A' -Order 10

        $mapB = Get-FlexModuleStateMap -Rows $script:e2eProfileB -ProfileName 'B' -LastBatchResults $script:LastBatchResults
        $mapA = Get-FlexModuleStateMap -Rows $script:e2eProfileA -ProfileName 'A' -LastBatchResults $script:LastBatchResults

        $mapB[10].Status   | Should -Be 'Pending'
        $mapB[20].Status   | Should -Be 'Pending'
        # ...while profile A itself shows what just ran
        $mapA[10].Status   | Should -Be 'Success'
        $mapA[10].Verified | Should -BeTrue
        $mapA[10].Message  | Should -Be 'renamed to PC-001'
        $mapA[20].Status   | Should -Be 'Pending'
    }

    It 'the dashboard state survives without the in-process overlay (history file alone)' {
        Invoke-FlexSingleRun -ProfileRows $script:e2eProfileA -ProfileName 'A' -Order 10

        $mapB = Get-FlexModuleStateMap -Rows $script:e2eProfileB -ProfileName 'B'
        $mapA = Get-FlexModuleStateMap -Rows $script:e2eProfileA -ProfileName 'A'

        $mapB[10].Status | Should -Be 'Pending'
        $mapA[10].Status | Should -Be 'Success'
    }

    It 'the checklist of profile B does not report the module run in profile A' {
        Invoke-FlexSingleRun -ProfileRows $script:e2eProfileA -ProfileName 'A' -Order 10

        # [Complete] pressed on profile B right away, then on profile A
        $htmlB = Get-ChecklistHtml -ProfileRows $script:e2eProfileB -ProfileName 'B'
        $htmlA = Get-ChecklistHtml -ProfileRows $script:e2eProfileA -ProfileName 'A'

        $htmlB | Should -Not -Match 'renamed to PC-001'
        $htmlB | Should -Match 'chip-ok">OK 0<'
        $htmlB | Should -Match 'chip-notrun">Not Run 2<'
        $htmlA | Should -Match 'renamed to PC-001'
        $htmlA | Should -Match 'chip-ok">OK 1<'
        $htmlA | Should -Match 'chip-notrun">Not Run 1<'
    }

    It 'running profile B afterwards changes B only; A keeps its result (history reloaded at batch start)' {
        Invoke-FlexSingleRun -ProfileRows $script:e2eProfileA -ProfileName 'A' -Order 10
        $script:e2eNextResult = { New-FakeRunResult -Status 'Error' -Message 'installer failed' }
        Invoke-FlexSingleRun -ProfileRows $script:e2eProfileB -ProfileName 'B' -Order 10

        $mapA = Get-FlexModuleStateMap -Rows $script:e2eProfileA -ProfileName 'A' -LastBatchResults $script:LastBatchResults
        $mapB = Get-FlexModuleStateMap -Rows $script:e2eProfileB -ProfileName 'B' -LastBatchResults $script:LastBatchResults
        $htmlA = Get-ChecklistHtml -ProfileRows $script:e2eProfileA -ProfileName 'A'
        $htmlB = Get-ChecklistHtml -ProfileRows $script:e2eProfileB -ProfileName 'B'

        $mapA[10].Status  | Should -Be 'Success'
        $mapB[10].Status  | Should -Be 'Error'
        $mapB[10].Message | Should -Be 'installer failed'

        $htmlA | Should -Match 'chip-ok">OK 1<'
        $htmlA | Should -Match 'chip-ng">NG 0<'
        $htmlA | Should -Not -Match 'installer failed'
        $htmlB | Should -Match 'chip-ng">NG 1<'
        $htmlB | Should -Match 'chip-ok">OK 0<'
        $htmlB | Should -Not -Match 'renamed to PC-001'
    }
}
