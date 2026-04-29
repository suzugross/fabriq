# Pester tests for the completion engine (PSReadLine-independent).
# Run: Invoke-Pester apps\fabriq_ios\tests\completer.tests.ps1

BeforeAll {
    # Establish $script:FabriqRoot before dot-sourcing libs that reference it.
    $script:FabriqIosRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
    $script:FabriqRoot    = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..')).Path

    . (Join-Path $script:FabriqRoot 'kernel\common.ps1')
    . (Join-Path $script:FabriqIosRoot 'lib\shell_state.ps1')
    . (Join-Path $script:FabriqIosRoot 'lib\parser.ps1')
    . (Join-Path $script:FabriqIosRoot 'lib\dispatch.ps1')
    . (Join-Path $script:FabriqIosRoot 'lib\commands\hostname.ps1')
    . (Join-Path $script:FabriqIosRoot 'lib\commands\interface.ps1')
    . (Join-Path $script:FabriqIosRoot 'lib\commands\ip_address.ps1')
    . (Join-Path $script:FabriqIosRoot 'lib\completer.ps1')
}

Describe 'Get-FabriqIosCompletion - top-level commands' {
    It 'lists UserExec command names when the line is empty' {
        $state = New-ShellState
        $r = Get-FabriqIosCompletion -Line '' -Position 0 -Mode 'UserExec' -State $state
        $r | Should -Contain 'enable'
        $r | Should -Contain 'show'
        $r | Should -Contain 'exit'
    }

    It 'filters UserExec commands by prefix' {
        $state = New-ShellState
        $r = Get-FabriqIosCompletion -Line 'e' -Position 1 -Mode 'UserExec' -State $state
        $r | Should -Contain 'enable'
        $r | Should -Contain 'exit'
        $r | Should -Not -Contain 'show'
    }

    It 'returns sub-vocabulary for show in UserExec' {
        $state = New-ShellState
        $r = Get-FabriqIosCompletion -Line 'show ' -Position 5 -Mode 'UserExec' -State $state
        $r | Should -Contain 'version'
        $r | Should -Contain 'manifesto'
        $r | Should -Not -Contain 'running-config'
    }

    It 'returns extended sub-vocabulary for show in PrivilegedExec' {
        $state = New-ShellState
        $r = Get-FabriqIosCompletion -Line 'show ' -Position 5 -Mode 'PrivilegedExec' -State $state
        $r | Should -Contain 'running-config'
        $r | Should -Contain 'profiles'
    }

    It 'completes show subcommand prefix' {
        $state = New-ShellState
        $r = Get-FabriqIosCompletion -Line 'show ru' -Position 7 -Mode 'PrivilegedExec' -State $state
        $r | Should -Contain 'running-config'
        $r.Count | Should -Be 1
    }
}

Describe 'Get-FabriqIosCompletion - dynamic sources' {
    It 'completes hostname candidates from hostlist' {
        $state = New-ShellState
        $r = Get-FabriqIosCompletion -Line 'host ' -Position 5 -Mode 'GlobalConfig' -State $state
        # Hostlist on this repo has at least NEW-PC-01.
        $r | Should -Contain 'NEW-PC-01'
    }

    It 'completes ip address candidates with from-hostlist literal' {
        $state = New-ShellState
        $r = Get-FabriqIosCompletion -Line 'ip address ' -Position 11 -Mode 'InterfaceConfig' -State $state
        $r | Should -Contain 'from-hostlist'
    }

    It 'returns empty for unknown leading command' {
        $state = New-ShellState
        $r = Get-FabriqIosCompletion -Line 'nonsense ' -Position 9 -Mode 'UserExec' -State $state
        $r.Count | Should -Be 0
    }
}

Describe 'Get-CommonPrefix' {
    It 'finds shared prefix' {
        Get-CommonPrefix -Strings @('hello','help','helpful') | Should -Be 'hel'
    }

    It 'returns empty when no shared prefix' {
        Get-CommonPrefix -Strings @('foo','bar') | Should -Be ''
    }

    It 'returns the single string for one input' {
        Get-CommonPrefix -Strings @('alone') | Should -Be 'alone'
    }
}
