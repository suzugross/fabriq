# ========================================
# Pester v5 unit tests for Invoke-WindowsUpdateLoop (AutoLogon cleanup)
# ========================================
# Function: kernel/main.ps1 :: Invoke-WindowsUpdateLoop
# Run    : powershell.exe -File ./dev/run_tests.ps1
#
# Pins the AutoLogon credential-cleanup contract (TM t-0011): the WU
# reboot loop writes a PLAINTEXT DefaultPassword + AutoLogonCount into
# HKLM Winlogon before each reboot, and every exit path that owns those
# credentials must clear them:
#   - operator cancel on a RESUMED loop (LoopCount > 0)  -> Clear fires
#   - operator cancel on a FRESH first pass              -> registry
#     untouched (WU wrote nothing; autologon_config may own the values)
#   - Register-FabriqRunOnce failure right after Set     -> Clear fires
#   - RunOnce success (happy path)                       -> no Clear,
#     countdown restart fires, wu_state.json survives for the next leg
#
# The function is extracted from main.ps1 via AST
# (Get-FabriqMainFunctionScriptBlock). It resolves its module paths from
# the CURRENT DIRECTORY, so each test runs inside a temp tree containing
# a stub modules\standard\windows_update\windows_update.ps1 whose return
# value drives the scenario (the stub is invoked as a file - it cannot
# be Pester-mocked).
#
# SAFETY: Set-/Clear-WindowsUpdateAutoLogon write HKLM Winlogon and
# Invoke-CountdownRestart reboots the machine. Throwing stubs sit under
# the Pester mocks so a dropped mock fails loudly instead of acting on
# the dev machine. Do not remove the stubs.
# ========================================

BeforeAll {
    . "$PSScriptRoot\..\_helpers\test_state.ps1"
    $script:RepoRoot = Get-FabriqRepoRoot
    . (Join-Path $script:RepoRoot 'kernel\common.ps1')

    # Inject the function under test (AST extraction; main.ps1 itself is
    # not dot-sourceable).
    $sb = Get-FabriqMainFunctionScriptBlock -Name 'Invoke-WindowsUpdateLoop'
    . $sb

    # ---- SAFETY STUBS (see header) ------------------------------------
    function Set-WindowsUpdateAutoLogon   { throw 'SAFETY STUB: Set-WindowsUpdateAutoLogon (writes HKLM) called without a Pester mock' }
    function Clear-WindowsUpdateAutoLogon { throw 'SAFETY STUB: Clear-WindowsUpdateAutoLogon (writes HKLM) called without a Pester mock' }
    function Register-FabriqRunOnce       { throw 'SAFETY STUB: Register-FabriqRunOnce called without a Pester mock' }
    function Invoke-CountdownRestart      { throw 'SAFETY STUB: Invoke-CountdownRestart called without a Pester mock' }
    function Show-WindowsUpdateSummary    { throw 'SAFETY STUB: Show-WindowsUpdateSummary called without a Pester mock' }
    # --------------------------------------------------------------------

    # Writes the stub WU module whose PSCustomObject return drives the
    # scenario. $BodyLines are appended after the param() block.
    function Set-WuStubScript {
        param([Parameter(Mandatory)][string[]]$BodyLines)
        $lines = @('param($SkipKBs, [switch]$AutoConfirm)') + $BodyLines
        Set-Content -Path $script:wuScriptPath -Value ($lines -join "`r`n") -Encoding Ascii -Force
    }

    function Set-WuStateFile {
        param([int]$LoopCount = 1, [bool]$AutoLogon = $true)
        @{
            LoopCount    = $LoopCount
            MaxLoops     = 5
            InstalledKBs = @()
            FailedKBs    = @()
            StartTime    = (Get-Date).ToString('o')
            RebootSec    = 15
            AutoLogon    = $AutoLogon
        } | ConvertTo-Json -Depth 5 | Out-File -FilePath $script:wuStatePath -Encoding UTF8 -Force
    }
}

Describe 'Invoke-WindowsUpdateLoop - AutoLogon credential cleanup' {

    BeforeEach {
        # Temp fabriq-like tree; the function resolves all paths from CWD.
        $script:tmpRoot = Join-Path $env:TEMP `
            ("fabriq-wu-test-{0}" -f ([guid]::NewGuid().ToString('N')))
        $wuDir = Join-Path $script:tmpRoot 'modules\standard\windows_update'
        New-Item -ItemType Directory -Path $wuDir -Force | Out-Null
        $script:wuScriptPath = Join-Path $wuDir 'windows_update.ps1'
        $script:wuStatePath  = Join-Path $wuDir 'wu_state.json'
        Push-Location $script:tmpRoot

        Mock Set-WindowsUpdateAutoLogon   { }
        Mock Clear-WindowsUpdateAutoLogon { }
        Mock Register-FabriqRunOnce       { $true }
        Mock Invoke-CountdownRestart      { }
        Mock Show-WindowsUpdateSummary    { }
        Mock Add-ExecutionResult          { }
        Mock Write-ExecutionHistory       { }
    }

    AfterEach {
        Pop-Location
        if ($script:tmpRoot -and (Test-Path $script:tmpRoot)) {
            Remove-Item $script:tmpRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    Context 'operator cancel' {

        It 'fresh first pass (LoopCount=0): leaves the registry untouched and removes nothing it did not write' {
            Set-WuStubScript -BodyLines @(
                'return [PSCustomObject]@{ Status = "Cancelled" }'
            )

            Invoke-WindowsUpdateLoop

            # WU never configured AutoLogon on this pass - clearing here
            # could clobber values owned by autologon_config.
            Should -Invoke Clear-WindowsUpdateAutoLogon -Exactly -Times 0
            Should -Invoke Set-WindowsUpdateAutoLogon -Exactly -Times 0
            Test-Path $script:wuStatePath | Should -BeFalse
        }

        It 'resumed loop (LoopCount=1, AutoLogon=true): clears the credentials WU set before the previous reboot' {
            Set-WuStateFile -LoopCount 1 -AutoLogon $true
            Set-WuStubScript -BodyLines @(
                'return [PSCustomObject]@{ Status = "Cancelled" }'
            )

            Invoke-WindowsUpdateLoop

            Should -Invoke Clear-WindowsUpdateAutoLogon -Exactly -Times 1
            Test-Path $script:wuStatePath | Should -BeFalse
        }

        It 'resumed loop with AutoLogon=false: nothing to clear' {
            Set-WuStateFile -LoopCount 1 -AutoLogon $false
            Set-WuStubScript -BodyLines @(
                'return [PSCustomObject]@{ Status = "Cancelled" }'
            )

            Invoke-WindowsUpdateLoop

            Should -Invoke Clear-WindowsUpdateAutoLogon -Exactly -Times 0
        }
    }

    Context 'reboot leg' {

        It 'RunOnce failure right after Set: clears the just-written credentials and aborts without restarting' {
            Mock Register-FabriqRunOnce { $false }
            Set-WuStubScript -BodyLines @(
                'return [PSCustomObject]@{',
                '    Status = "Success"; RebootRequired = $true',
                '    InstalledCount = 1; FailedCount = 0',
                '    InstalledKBs = @([PSCustomObject]@{ KB = "KB1"; Title = "T" })',
                '    FailedKBs = @()',
                '}'
            )

            Invoke-WindowsUpdateLoop

            Should -Invoke Set-WindowsUpdateAutoLogon -Exactly -Times 1
            Should -Invoke Clear-WindowsUpdateAutoLogon -Exactly -Times 1
            Should -Invoke Invoke-CountdownRestart -Exactly -Times 0
            Test-Path $script:wuStatePath | Should -BeFalse
        }

        It 'RunOnce success (happy path): no Clear, countdown restart fires, wu_state.json survives for the next leg' {
            Set-WuStubScript -BodyLines @(
                'return [PSCustomObject]@{',
                '    Status = "Success"; RebootRequired = $true',
                '    InstalledCount = 1; FailedCount = 0',
                '    InstalledKBs = @([PSCustomObject]@{ KB = "KB1"; Title = "T" })',
                '    FailedKBs = @()',
                '}'
            )

            Invoke-WindowsUpdateLoop

            Should -Invoke Set-WindowsUpdateAutoLogon -Exactly -Times 1
            Should -Invoke Clear-WindowsUpdateAutoLogon -Exactly -Times 0
            Should -Invoke Invoke-CountdownRestart -Exactly -Times 1
            Test-Path $script:wuStatePath | Should -BeTrue
        }
    }
}
