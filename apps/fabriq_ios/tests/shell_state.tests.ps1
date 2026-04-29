# Pester tests for the shell state model and mode transitions.
# Run: Invoke-Pester apps\fabriq_ios\tests\shell_state.tests.ps1

BeforeAll {
    . "$PSScriptRoot\..\lib\shell_state.ps1"
}

Describe 'New-ShellState' {
    It 'starts in UserExec mode' {
        (New-ShellState).Mode | Should -Be 'UserExec'
    }

    It 'has the expected schema fields' {
        $s = New-ShellState
        $s.ContainsKey('Mode')             | Should -BeTrue
        $s.ContainsKey('CurrentInterface') | Should -BeTrue
        $s.ContainsKey('Passphrase')       | Should -BeTrue
        $s.ContainsKey('SelectedHost')     | Should -BeTrue
        $s.ContainsKey('ShouldExit')       | Should -BeTrue
        $s.ContainsKey('ProcessId')        | Should -BeTrue
    }

    It 'starts with no current interface and no passphrase' {
        $s = New-ShellState
        $s.CurrentInterface | Should -BeNullOrEmpty
        $s.Passphrase       | Should -BeNullOrEmpty
        $s.ShouldExit       | Should -BeFalse
    }
}

Describe 'Set-ShellMode' {
    It 'allows UserExec -> PrivilegedExec' {
        $s = New-ShellState
        Set-ShellMode -State $s -NewMode 'PrivilegedExec'
        $s.Mode | Should -Be 'PrivilegedExec'
    }

    It 'allows PrivilegedExec -> GlobalConfig' {
        $s = New-ShellState
        $s.Mode = 'PrivilegedExec'
        Set-ShellMode -State $s -NewMode 'GlobalConfig'
        $s.Mode | Should -Be 'GlobalConfig'
    }

    It 'allows InterfaceConfig -> PrivilegedExec (end)' {
        $s = New-ShellState
        $s.Mode = 'InterfaceConfig'
        Set-ShellMode -State $s -NewMode 'PrivilegedExec'
        $s.Mode | Should -Be 'PrivilegedExec'
    }

    It 'rejects UserExec -> GlobalConfig directly' {
        $s = New-ShellState
        { Set-ShellMode -State $s -NewMode 'GlobalConfig' } | Should -Throw
    }

    It 'rejects unknown mode names' {
        $s = New-ShellState
        { Set-ShellMode -State $s -NewMode 'NotAMode' } | Should -Throw
    }
}
