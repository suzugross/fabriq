# ========================================
# Pester v5 unit tests for Register-FabriqRunOnce
# ========================================
# Function: kernel/common.ps1 :: Register-FabriqRunOnce
# Run    : powershell.exe -File ./dev/run_tests.ps1
#
# Pins the RunOnce registration that every __RESTART__ marker and the
# Windows Update reboot loop depend on (TM t-0007 (2)). A regression in
# the value name, the quoted command line, or the $false-on-failure
# contract would either drop the auto-resume after reboot (operator
# stranded at a login screen) or let Invoke-BatchExecution proceed to
# reboot without a registered resume (kitting run lost).
#
# SAFETY: New-Item / New-ItemProperty are mocked UNCONDITIONALLY (no
# ParameterFilter on the Mock itself) so no test path can ever touch the
# real HKLM RunOnce key - a conditional mock would fall through to the
# real cmdlet on a filter mismatch and schedule Fabriq.exe to launch on
# the dev machine's next boot. Argument assertions are done on the
# Should -Invoke side instead.
# ========================================

BeforeAll {
    . "$PSScriptRoot\..\_helpers\test_state.ps1"
    $script:RepoRoot = Get-FabriqRepoRoot
    . (Join-Path $script:RepoRoot 'kernel\common.ps1')

    $script:RunOncePath = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce'
}

Describe 'Register-FabriqRunOnce' {

    BeforeEach {
        # Unconditional registry-write interception (see SAFETY note above).
        Mock New-Item {}
        Mock New-ItemProperty {}
        # Default environment: Fabriq.exe present, RunOnce key present.
        # Individual tests override with more-recently-defined mocks.
        Mock Test-Path { $true }
    }

    Context 'Success path' {

        It 'registers FabriqAutoStart under the canonical RunOnce key and returns $true' {
            $result = Register-FabriqRunOnce
            $result | Should -BeTrue
            Should -Invoke New-ItemProperty -Exactly -Times 1 -ParameterFilter {
                $Path -eq $script:RunOncePath -and
                $Name -eq 'FabriqAutoStart' -and
                $PropertyType -eq 'String'
            }
        }

        It 'writes the exe path as a quoted command line (space-safe)' {
            $null = Register-FabriqRunOnce
            Should -Invoke New-ItemProperty -Exactly -Times 1 -ParameterFilter {
                $Value -match '^".+\\Fabriq\.exe"$'
            }
        }

        It 'does not create the RunOnce key when it already exists' {
            $null = Register-FabriqRunOnce
            Should -Invoke New-Item -Exactly -Times 0
        }
    }

    Context 'RunOnce key missing' {

        It 'creates the key first, then registers the value' {
            Mock Test-Path { $false } -ParameterFilter { $Path -eq $script:RunOncePath }
            $result = Register-FabriqRunOnce
            $result | Should -BeTrue
            Should -Invoke New-Item -Exactly -Times 1 -ParameterFilter {
                $Path -eq $script:RunOncePath
            }
            Should -Invoke New-ItemProperty -Exactly -Times 1
        }
    }

    Context 'Failure paths (fail-closed contract for __RESTART__)' {

        It 'returns $false and never touches the registry when Fabriq.exe is missing' {
            Mock Test-Path { $false } -ParameterFilter { $Path -like '*Fabriq.exe' }
            $result = Register-FabriqRunOnce
            $result | Should -BeFalse
            Should -Invoke New-ItemProperty -Exactly -Times 0
            Should -Invoke New-Item -Exactly -Times 0
        }

        It 'returns $false when the registry write throws (access denied etc.)' {
            Mock New-ItemProperty { throw 'Requested registry access is not allowed.' }
            $result = Register-FabriqRunOnce
            $result | Should -BeFalse
        }
    }
}
