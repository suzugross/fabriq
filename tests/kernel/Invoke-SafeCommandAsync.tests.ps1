# ========================================
# Pester v5 integration tests for Invoke-SafeCommandAsync
# ========================================
# Function: kernel/common.ps1 :: Invoke-SafeCommandAsync
# Run    : powershell.exe -File ./dev/run_tests.ps1
#
# Pins the ASYNC execution path (TM t-0024 (1)) - the DEFAULT path for
# every module run since DefaultAsync=true shipped (kernel 3.3.0), yet
# previously covered only by throwaway manual verification (t-0005).
#
# INTEGRATION STYLE, BY NECESSITY: the child runspace has its own
# session state, so Pester mocks cannot reach inside it. Each test runs
# a REAL runspace executing a stub module .ps1 written by the test -
# which is also the point: runspace creation, global injection, the
# monitor loop and forced stop ARE the contract under test.
#
# CWD CONTAINMENT: the function resolves common.ps1 / async_config.json
# / the skip flag / telemetry paths relative to the CURRENT DIRECTORY,
# so every test runs inside a temp fabriq-like tree (real common.ps1
# copied in, async_config.json test-controlled). No repo side effects.
#
# The Skip-interrupt test uses a self-skip trick: the caller thread is
# blocked inside Invoke-SafeCommandAsync's monitor loop, so the CHILD
# stub itself writes the skip flag before sleeping - deterministic,
# no second thread needed.
# ========================================

BeforeAll {
    . "$PSScriptRoot\..\_helpers\test_state.ps1"
    $script:RepoRoot = Get-FabriqRepoRoot
    . (Join-Path $script:RepoRoot 'kernel\common.ps1')

    # Writes a stub module under the temp tree; returns the relative path
    # (the function Resolve-Path's it against the temp CWD, exactly like
    # real callers resolve module paths against the fabriq root).
    function New-AsyncStubModule {
        param(
            [Parameter(Mandatory)][string]$Name,
            [Parameter(Mandatory)][string[]]$BodyLines
        )
        $dir = Join-Path (Get-Location) 'modules\teststubs'
        if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        Set-Content -Path (Join-Path $dir $Name) -Value ($BodyLines -join "`r`n") -Encoding Ascii -Force
        return ".\modules\teststubs\$Name"
    }
}

Describe 'Invoke-SafeCommandAsync' {

    BeforeEach {
        # Temp fabriq-like tree (the function resolves everything from CWD)
        $script:tmpRoot = Join-Path $env:TEMP `
            ("fabriq-async-test-{0}" -f ([guid]::NewGuid().ToString('N')))
        New-Item -ItemType Directory -Path (Join-Path $script:tmpRoot 'kernel\json') -Force | Out-Null
        Copy-Item (Join-Path $script:RepoRoot 'kernel\common.ps1') (Join-Path $script:tmpRoot 'kernel\common.ps1')

        # Test-controlled async config: fast polling, no default timeout.
        # Forward slashes avoid JSON backslash escaping; the function
        # normalizes the path via Join-Path + GetFullPath anyway.
        Set-Content -Path (Join-Path $script:tmpRoot 'kernel\json\async_config.json') -Encoding Ascii -Force -Value @'
{ "Enabled": true, "DefaultAsync": true, "DefaultTimeoutSec": 0, "PollIntervalMs": 100, "SkipFlagPath": "./kernel/json/skip_request.flag" }
'@
        $script:skipFlag = Join-Path $script:tmpRoot 'kernel\json\skip_request.flag'

        Push-Location $script:tmpRoot

        $global:FabriqMasterPassphrase = $null
        $global:AutoPilotMode = $false
    }

    AfterEach {
        Pop-Location
        if ($script:tmpRoot -and (Test-Path $script:tmpRoot)) {
            Remove-Item $script:tmpRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
        $global:FabriqMasterPassphrase = $null
        $global:AutoPilotMode = $false
    }

    Context 'result contract' {

        It 'passes a pipeline ModuleResult through (Status / Success / Verified / Message)' {
            $stub = New-AsyncStubModule -Name 'ok.ps1' -BodyLines @(
                'return (New-ModuleResult -Status "Success" -Message "async ok" -Verified $true)'
            )

            $result = Invoke-SafeCommandAsync -ScriptPath $stub -OperationName 'AsyncOk'

            $result.Status | Should -Be 'Success'
            $result.Success | Should -BeTrue
            $result.Verified | Should -BeTrue
            $result.Message | Should -Be 'async ok'
        }

        It 'recovers a result via the $global:_LastModuleResult fallback across the runspace boundary' {
            $stub = New-AsyncStubModule -Name 'lastresult.ps1' -BodyLines @(
                '$null = New-ModuleResult -Status "Partial" -Message "via global"'
            )

            $result = Invoke-SafeCommandAsync -ScriptPath $stub -OperationName 'AsyncLastResult'

            $result.Status | Should -Be 'Partial'
            $result.Success | Should -BeFalse
            $result.Message | Should -Be 'via global'
        }

        It 'fail-closed: a module returning no ModuleResult is recorded as Error (contract violation)' {
            $stub = New-AsyncStubModule -Name 'noresult.ps1' -BodyLines @(
                'Write-Host "stub ran but returned nothing"'
            )

            $result = Invoke-SafeCommandAsync -ScriptPath $stub -OperationName 'AsyncNoResult'

            $result.Status | Should -Be 'Error'
            $result.Success | Should -BeFalse
            $result.Message | Should -Match 'No ModuleResult returned'
        }

        It 'returns Error without starting a runspace when the script path does not exist' {
            $result = Invoke-SafeCommandAsync -ScriptPath '.\modules\teststubs\no-such.ps1' -OperationName 'AsyncMissing'

            $result.Status | Should -Be 'Error'
            $result.Message | Should -Match 'Script not found'
        }
    }

    Context 'global injection' {

        It 'injects parent $global:* values into the child runspace' {
            $global:FabriqMasterPassphrase = 'test-pass-123'
            $global:AutoPilotMode = $true
            $stub = New-AsyncStubModule -Name 'globals.ps1' -BodyLines @(
                'return (New-ModuleResult -Status "Success" -Message "pp=$global:FabriqMasterPassphrase ap=$global:AutoPilotMode")'
            )

            $result = Invoke-SafeCommandAsync -ScriptPath $stub -OperationName 'AsyncGlobals'

            $result.Message | Should -Be 'pp=test-pass-123 ap=True'
        }
    }

    Context 'interrupts' {

        It 'Skip flag stops the runspace and reports Error (self-skip: the child writes the flag)' {
            $stub = New-AsyncStubModule -Name 'selfskip.ps1' -BodyLines @(
                'Set-Content -Path ".\kernel\json\skip_request.flag" -Value "skip" -Force',
                'Start-Sleep -Seconds 30',
                'return (New-ModuleResult -Status "Success" -Message "must never reach this")'
            )

            $result = Invoke-SafeCommandAsync -ScriptPath $stub -OperationName 'AsyncSkip'

            $result.Status | Should -Be 'Error'
            $result.Message | Should -Match 'skipped by operator'
            # Forced stop, not the 30s sleep running out
            $result.Duration.TotalSeconds | Should -BeLessThan 15
        }

        It 'Timeout stops the runspace and reports Error with the configured limit' {
            $stub = New-AsyncStubModule -Name 'slow.ps1' -BodyLines @(
                'Start-Sleep -Seconds 30',
                'return (New-ModuleResult -Status "Success" -Message "must never reach this")'
            )

            $result = Invoke-SafeCommandAsync -ScriptPath $stub -OperationName 'AsyncTimeout' -TimeoutSec 2

            $result.Status | Should -Be 'Error'
            $result.Message | Should -Match 'exceeded timeout'
            $result.Duration.TotalSeconds | Should -BeLessThan 15
        }

        It 'clears a STALE skip flag before starting (no spurious interrupt)' {
            Set-Content -Path $script:skipFlag -Value 'stale' -Force
            $stub = New-AsyncStubModule -Name 'quick.ps1' -BodyLines @(
                'return (New-ModuleResult -Status "Success" -Message "ran to completion")'
            )

            $result = Invoke-SafeCommandAsync -ScriptPath $stub -OperationName 'AsyncStaleFlag'

            $result.Status | Should -Be 'Success'
            $result.Message | Should -Be 'ran to completion'
            Test-Path $script:skipFlag | Should -BeFalse
        }
    }
}
