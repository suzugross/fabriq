# ========================================
# Pester v5 unit tests for Save-ResumeState / Load-ResumeState
# ========================================
# Function: kernel/common.ps1 :: Save-ResumeState / Load-ResumeState
# Run    : powershell.exe -File ./dev/run_tests.ps1
#
# Pins the on-disk JSON shape contract (FlexProfile schemaVersion=1
# vs schemaVersion=2) and the Save -> reboot -> Load round-trip
# integrity that __RESTART__ depends on. This is internal-API-tier
# coverage (KERNEL_API.md §6) and exists primarily for regression
# detection: a silent change to resume_state.json layout would break
# every cross-restart profile in the field.
#
# Coverage targets:
#   - kernel 3.1.0 : FlexProfile resume v2 schema introduced
#   - kernel 3.1.x : Linear/Flex Save coexistence (byte-compat for v1)
# ========================================

BeforeAll {
    . "$PSScriptRoot\..\_helpers\test_state.ps1"
    $script:RepoRoot = Get-FabriqRepoRoot
    . (Join-Path $script:RepoRoot 'kernel\common.ps1')

    # Suppress telemetry write attempts (best-effort path in production;
    # would try to write to logs/telemetry/... in tests).
    Mock Write-KernelTelemetryEvent { }
}

Describe 'Save-ResumeState / Load-ResumeState' {

    BeforeEach {
        $script:tmpStatePath = Join-Path $env:TEMP `
            ("fabriq-resume-{0}.json" -f ([guid]::NewGuid().ToString('N')))
        Set-FabriqTestState -ResumeStatePath $script:tmpStatePath -SessionID 'test-session-001'

        # Reset globals Save-ResumeState reads.
        $global:AutoPilotMode          = $false
        $global:AutoPilotWaitSec       = 3
        $global:FabriqMasterPassphrase = $null
        $global:FabriqEvidenceBasePath = 'C:\test\evidence'
    }

    AfterEach {
        if (Test-Path $script:tmpStatePath) {
            Remove-Item $script:tmpStatePath -Force -ErrorAction SilentlyContinue
        }
    }

    Context 'Linear (v1) Save format' {

        It 'omits schemaVersion field for byte-compat with pre-FlexProfile v1' {
            Save-ResumeState -ProfilePath 'C:\profiles\foo.csv' `
                             -ProfileName 'foo' `
                             -ResumeAfterOrder 30 `
                             -CompletedModules @()
            $loaded = Get-Content $script:tmpStatePath -Raw | ConvertFrom-Json
            $loaded.PSObject.Properties.Name | Should -Not -Contain 'schemaVersion'
            $loaded.PSObject.Properties.Name | Should -Not -Contain 'ExecutionMode'
            $loaded.PSObject.Properties.Name | Should -Not -Contain 'SelectedOrders'
            $loaded.PSObject.Properties.Name | Should -Not -Contain 'ModuleStates'
        }

        It 'includes the core resume fields' {
            Save-ResumeState -ProfilePath 'C:\profiles\foo.csv' `
                             -ProfileName 'foo' `
                             -ResumeAfterOrder 30 `
                             -CompletedModules @()
            $loaded = Get-Content $script:tmpStatePath -Raw | ConvertFrom-Json
            $loaded.ProfilePath      | Should -Be 'C:\profiles\foo.csv'
            $loaded.ProfileName      | Should -Be 'foo'
            $loaded.ResumeAfterOrder | Should -Be 30
            $loaded.SessionID        | Should -Be 'test-session-001'
            $loaded.AutoPilot        | Should -BeFalse
            $loaded.AutoPilotWaitSec | Should -Be 3
            $loaded.EvidenceBasePath | Should -Be 'C:\test\evidence'
        }

        It 'serializes CompletedModules as Order/MenuName/Status triples' {
            $completed = @(
                [PSCustomObject]@{ Order = 10; MenuName = 'Hostname';    Status = 'Success' }
                [PSCustomObject]@{ Order = 20; MenuName = 'IP Address';  Status = 'Skipped' }
                [PSCustomObject]@{ Order = 30; MenuName = 'Reg HKLM';    Status = 'Partial' }
            )
            Save-ResumeState -ProfilePath 'foo' -ProfileName 'foo' -ResumeAfterOrder 40 -CompletedModules $completed
            $loaded = Get-Content $script:tmpStatePath -Raw | ConvertFrom-Json
            @($loaded.CompletedModules).Count    | Should -Be 3
            $loaded.CompletedModules[0].Order    | Should -Be 10
            $loaded.CompletedModules[0].MenuName | Should -Be 'Hostname'
            $loaded.CompletedModules[0].Status   | Should -Be 'Success'
            $loaded.CompletedModules[2].Status   | Should -Be 'Partial'
        }

        It 'serializes ProfileStartTime in ISO 8601 round-trip format' {
            $start = Get-Date '2026-05-10T13:28:39+09:00'
            Save-ResumeState -ProfilePath 'foo' -ProfileName 'foo' -ResumeAfterOrder 0 `
                             -CompletedModules @() -ProfileStartTime $start
            $loaded = Get-Content $script:tmpStatePath -Raw | ConvertFrom-Json
            # "o" round-trip format: yyyy-MM-ddTHH:mm:ss.fffffffzzz
            $loaded.ProfileStartTime | Should -Match '^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{7}'
        }
    }

    Context 'Flex (v2) Save format' {

        It 'emits schemaVersion=2 when -ExecutionMode Flex' {
            Save-ResumeState -ProfilePath 'foo' -ProfileName 'foo' -ResumeAfterOrder -1 `
                             -CompletedModules @() -ExecutionMode 'Flex'
            $loaded = Get-Content $script:tmpStatePath -Raw | ConvertFrom-Json
            $loaded.schemaVersion | Should -Be 2
            $loaded.ExecutionMode | Should -Be 'Flex'
        }

        It 'serializes SelectedOrders array' {
            Save-ResumeState -ProfilePath 'foo' -ProfileName 'foo' -ResumeAfterOrder -1 `
                             -CompletedModules @() -ExecutionMode 'Flex' `
                             -SelectedOrders @(10, 30, 50)
            $loaded = Get-Content $script:tmpStatePath -Raw | ConvertFrom-Json
            @($loaded.SelectedOrders).Count | Should -Be 3
            @($loaded.SelectedOrders)[0]    | Should -Be 10
            @($loaded.SelectedOrders)[2]    | Should -Be 50
        }

        It 'serializes ModuleStates hashtable as JSON object' {
            $states = @{
                '10' = @{ Status = 'Success'; Verified = $true  }
                '20' = @{ Status = 'Error';   Verified = $false }
            }
            Save-ResumeState -ProfilePath 'foo' -ProfileName 'foo' -ResumeAfterOrder -1 `
                             -CompletedModules @() -ExecutionMode 'Flex' `
                             -ModuleStates $states
            $loaded = Get-Content $script:tmpStatePath -Raw | ConvertFrom-Json
            $loaded.ModuleStates.'10'.Status | Should -Be 'Success'
            $loaded.ModuleStates.'10'.Verified | Should -BeTrue
            $loaded.ModuleStates.'20'.Status | Should -Be 'Error'
            $loaded.ModuleStates.'20'.Verified | Should -BeFalse
        }

        It 'still includes all v1 core fields alongside v2 additions' {
            Save-ResumeState -ProfilePath 'C:\profiles\flex.csv' -ProfileName 'flex' `
                             -ResumeAfterOrder -1 -CompletedModules @() `
                             -ExecutionMode 'Flex'
            $loaded = Get-Content $script:tmpStatePath -Raw | ConvertFrom-Json
            $loaded.ProfilePath      | Should -Be 'C:\profiles\flex.csv'
            $loaded.SessionID        | Should -Be 'test-session-001'
            $loaded.HostEnvironment  | Should -Not -BeNullOrEmpty
        }
    }

    Context 'Save -> Load round-trip' {

        It 'round-trips a Linear (v1) state' {
            Save-ResumeState -ProfilePath 'C:\profiles\foo.csv' -ProfileName 'foo' `
                             -ResumeAfterOrder 50 -CompletedModules @(
                                 [PSCustomObject]@{ Order = 10; MenuName = 'A'; Status = 'Success' }
                             )
            $loaded = Load-ResumeState
            $loaded                        | Should -Not -BeNullOrEmpty
            $loaded.ProfilePath            | Should -Be 'C:\profiles\foo.csv'
            $loaded.ProfileName            | Should -Be 'foo'
            $loaded.ResumeAfterOrder       | Should -Be 50
            $loaded.SessionID              | Should -Be 'test-session-001'
            @($loaded.CompletedModules).Count | Should -Be 1
            $loaded.CompletedModules[0].MenuName | Should -Be 'A'
        }

        It 'round-trips a Flex (v2) state' {
            $states = @{
                '10' = @{ Status = 'Success'; Verified = $true }
            }
            Save-ResumeState -ProfilePath 'C:\profiles\flex.csv' -ProfileName 'flex' `
                             -ResumeAfterOrder -1 -CompletedModules @() `
                             -ExecutionMode 'Flex' `
                             -SelectedOrders @(10, 20) `
                             -ModuleStates $states
            $loaded = Load-ResumeState
            $loaded                        | Should -Not -BeNullOrEmpty
            $loaded.schemaVersion          | Should -Be 2
            $loaded.ExecutionMode          | Should -Be 'Flex'
            @($loaded.SelectedOrders).Count | Should -Be 2
            $loaded.ModuleStates.'10'.Status | Should -Be 'Success'
        }

        It 'preserves HostEnvironment snapshot through round-trip' {
            # Set a recognizable env var the Save snapshot is known to capture.
            $orig = [Environment]::GetEnvironmentVariable('SELECTED_NEW_PCNAME')
            try {
                [Environment]::SetEnvironmentVariable('SELECTED_NEW_PCNAME', 'TEST-PC-RT')
                Save-ResumeState -ProfilePath 'foo' -ProfileName 'foo' -ResumeAfterOrder 0 `
                                 -CompletedModules @()
                $loaded = Load-ResumeState
                $loaded.HostEnvironment.SELECTED_NEW_PCNAME | Should -Be 'TEST-PC-RT'
            } finally {
                [Environment]::SetEnvironmentVariable('SELECTED_NEW_PCNAME', $orig)
            }
        }

        It 'preserves Printer slot env vars (1..10 x NAME/DRIVER/PORT) in HostEnvironment' {
            Save-ResumeState -ProfilePath 'foo' -ProfileName 'foo' -ResumeAfterOrder 0 `
                             -CompletedModules @()
            $loaded = Load-ResumeState
            # Values may be null on the test machine; structural check.
            $printerKeys = $loaded.HostEnvironment.PSObject.Properties.Name |
                Where-Object { $_ -like 'SELECTED_PRINTER_*' }
            $printerKeys.Count | Should -Be 30   # 10 slots * 3 suffixes
        }
    }

    Context 'Load boundary conditions' {

        It 'returns $null when the resume_state file does not exist' {
            $loaded = Load-ResumeState   # tmpStatePath has been set but no file written
            $loaded | Should -BeNullOrEmpty
        }

        It 'returns $null when the resume_state file is malformed JSON' {
            'this is not { valid json [' | Out-File -FilePath $script:tmpStatePath -Encoding UTF8 -Force
            $loaded = Load-ResumeState
            $loaded | Should -BeNullOrEmpty
        }

        It 'reads pre-FlexProfile v1 file (no schemaVersion field) without error' {
            # Hand-craft a v1 file (no schemaVersion / ExecutionMode fields)
            # to verify forward-compat reading path used by sites still on
            # pre-3.1.0 resume state files captured before kernel upgrade.
            $v1 = @{
                ProfilePath      = 'foo'
                ProfileName      = 'foo'
                AutoPilot        = $true
                AutoPilotWaitSec = 5
                SessionID        = 'legacy-session'
                ResumeAfterOrder = 20
                CompletedModules = @()
                HostEnvironment  = @{}
                EvidenceBasePath = ''
                ProfileStartTime = (Get-Date).ToString('o')
            }
            $v1 | ConvertTo-Json -Depth 5 | Out-File -FilePath $script:tmpStatePath -Encoding UTF8 -Force
            $loaded = Load-ResumeState
            $loaded                  | Should -Not -BeNullOrEmpty
            $loaded.SessionID        | Should -Be 'legacy-session'
            $loaded.ResumeAfterOrder | Should -Be 20
            $loaded.PSObject.Properties.Name | Should -Not -Contain 'schemaVersion'
        }
    }
}
