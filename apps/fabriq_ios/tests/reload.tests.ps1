# Pester tests for the `reload` flow.
# Run: Invoke-Pester apps\fabriq_ios\tests\reload.tests.ps1
#
# A real reboot cannot be exercised, so every external effect is mocked:
# Read-Host (the confirmation), Register-FabriqIosRunOnce (relaunch arming),
# Show-FabriqIosReloadTheatre (the farewell sequence), and crucially
# Restart-Computer. The mocks are established in BeforeEach so the real
# Restart-Computer can never fire during a test run. The assertions pin the
# decision contract: when (and only when) does the reboot happen.

BeforeAll {
    . "$PSScriptRoot\..\lib\commands\reload.ps1"
    # reload.ps1's fail-closed branch calls Show-Error (normally from
    # kernel\common.ps1). Provide a no-op stub so the branch is harmless;
    # the RunOnce body that also uses it is fully mocked away below.
    function Show-Error { param([string]$Message) }
}

Describe 'Invoke-FabriqIosReload' {
    BeforeEach {
        # Restart-Computer is mocked FIRST and unconditionally so a real
        # reboot is impossible regardless of which branch a test takes.
        Mock Restart-Computer { }
        Mock Register-FabriqIosRunOnce { $true }
        Mock Show-FabriqIosReloadTheatre { }
    }

    Context 'confirmation gate' {
        It 'aborts on "n": neither arms nor reboots' {
            Mock Read-Host { 'n' }
            Invoke-FabriqIosReload -State @{ Mode = 'PrivilegedExec' }
            Should -Invoke Register-FabriqIosRunOnce -Times 0 -Exactly
            Should -Invoke Restart-Computer -Times 0 -Exactly
        }

        It 'aborts on unrelated input' {
            Mock Read-Host { 'maybe' }
            Invoke-FabriqIosReload -State @{ Mode = 'PrivilegedExec' }
            Should -Invoke Restart-Computer -Times 0 -Exactly
        }

        It 'proceeds and reboots on empty input (Enter)' {
            Mock Read-Host { '' }
            Invoke-FabriqIosReload -State @{ Mode = 'PrivilegedExec' }
            Should -Invoke Restart-Computer -Times 1 -Exactly
        }

        It 'proceeds and reboots on "y"' {
            Mock Read-Host { 'y' }
            Invoke-FabriqIosReload -State @{ Mode = 'PrivilegedExec' }
            Should -Invoke Restart-Computer -Times 1 -Exactly
        }
    }

    Context 'fail-closed when the relaunch cannot be armed' {
        It 'does not reboot if RunOnce registration fails' {
            Mock Read-Host { 'y' }
            Mock Register-FabriqIosRunOnce { $false }
            Invoke-FabriqIosReload -State @{ Mode = 'PrivilegedExec' }
            Should -Invoke Register-FabriqIosRunOnce -Times 1 -Exactly
            Should -Invoke Restart-Computer -Times 0 -Exactly
        }
    }
}
