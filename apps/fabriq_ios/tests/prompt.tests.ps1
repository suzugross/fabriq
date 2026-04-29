# Pester tests for the prompt builder.
# Run: Invoke-Pester apps\fabriq_ios\tests\prompt.tests.ps1

BeforeAll {
    . "$PSScriptRoot\..\lib\shell_state.ps1"
    . "$PSScriptRoot\..\lib\prompt.ps1"
}

Describe 'Get-FabriqIosPrompt' {
    It 'returns fabriq> for UserExec' {
        $s = New-ShellState
        Get-FabriqIosPrompt -State $s | Should -Be 'fabriq>'
    }

    It 'returns fabriq# for PrivilegedExec' {
        $s = New-ShellState
        $s.Mode = 'PrivilegedExec'
        Get-FabriqIosPrompt -State $s | Should -Be 'fabriq#'
    }

    It 'returns fabriq(config)# for GlobalConfig' {
        $s = New-ShellState
        $s.Mode = 'GlobalConfig'
        Get-FabriqIosPrompt -State $s | Should -Be 'fabriq(config)#'
    }

    It 'returns fabriq(config-if)# for InterfaceConfig' {
        $s = New-ShellState
        $s.Mode = 'InterfaceConfig'
        Get-FabriqIosPrompt -State $s | Should -Be 'fabriq(config-if)#'
    }
}
