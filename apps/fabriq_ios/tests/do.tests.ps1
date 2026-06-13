# Pester tests for the `do <EXEC command>` config-mode helper.
# Run: Invoke-Pester apps\fabriq_ios\tests\do.tests.ps1
#
# These exercise Invoke-FabriqIosDoCommand's gate logic in isolation:
# the downstream Invoke-PrivilegedExecCommand is replaced with a
# recording stub so the tests assert WHAT would be dispatched (and that
# the shell mode / context is never mutated) without running real
# show / reload side effects.

BeforeAll {
    . "$PSScriptRoot\..\lib\parser.ps1"
    . "$PSScriptRoot\..\lib\commands\do.ps1"

    # Recording stub: shadows the real dispatcher. Captures every
    # dispatch so tests can assert the resolved command/args, and never
    # mutates $State - mirroring show/reload which also leave mode alone.
    function Invoke-PrivilegedExecCommand {
        param([hashtable]$Resolved, [hashtable]$State)
        $script:DoDispatchLog += , $Resolved
    }
}

Describe 'Invoke-FabriqIosDoCommand' {
    # Pester 5 disallows BeforeEach at the container root, so the
    # per-test log reset lives inside this Describe.
    BeforeEach {
        $script:DoDispatchLog = @()
    }

    Context 'whitelist accepts' {
        It "dispatches 'do sh run' as show running-config" {
            $state = @{ Mode = 'GlobalConfig' }
            Invoke-FabriqIosDoCommand -ExecTokens @('sh', 'run') -State $state
            $script:DoDispatchLog.Count | Should -Be 1
            $script:DoDispatchLog[0].Command | Should -Be 'show'
            $script:DoDispatchLog[0].Args[0] | Should -Be 'running-config'
        }

        It "dispatches 'do show version'" {
            $state = @{ Mode = 'InterfaceConfig' }
            Invoke-FabriqIosDoCommand -ExecTokens @('show', 'version') -State $state
            $script:DoDispatchLog.Count | Should -Be 1
            $script:DoDispatchLog[0].Command | Should -Be 'show'
        }

        It "dispatches 'do reload'" {
            $state = @{ Mode = 'ModuleConfig' }
            Invoke-FabriqIosDoCommand -ExecTokens @('reload') -State $state
            $script:DoDispatchLog.Count | Should -Be 1
            $script:DoDispatchLog[0].Command | Should -Be 'reload'
        }

        It "dispatches 'do ping 8.8.8.8' (read-only EXEC)" {
            $state = @{ Mode = 'GlobalConfig' }
            Invoke-FabriqIosDoCommand -ExecTokens @('ping', '8.8.8.8') -State $state
            $script:DoDispatchLog.Count | Should -Be 1
            $script:DoDispatchLog[0].Command | Should -Be 'ping'
            $script:DoDispatchLog[0].Args[0] | Should -Be '8.8.8.8'
        }

        It "dispatches 'do trace 8.8.8.8' (abbreviation -> traceroute)" {
            $state = @{ Mode = 'GlobalConfig' }
            Invoke-FabriqIosDoCommand -ExecTokens @('trace', '8.8.8.8') -State $state
            $script:DoDispatchLog.Count | Should -Be 1
            $script:DoDispatchLog[0].Command | Should -Be 'traceroute'
        }
    }

    Context 'whitelist rejects mode/lifecycle mutators' {
        It "rejects 'do exit' (would tear down the shell)" {
            $state = @{ Mode = 'GlobalConfig' }
            Invoke-FabriqIosDoCommand -ExecTokens @('exit') -State $state
            $script:DoDispatchLog.Count | Should -Be 0
        }

        It "rejects 'do conf t' (would change mode)" {
            $state = @{ Mode = 'GlobalConfig' }
            Invoke-FabriqIosDoCommand -ExecTokens @('conf', 't') -State $state
            $script:DoDispatchLog.Count | Should -Be 0
        }

        It "rejects 'do disable' (would deprivilege)" {
            $state = @{ Mode = 'GlobalConfig' }
            Invoke-FabriqIosDoCommand -ExecTokens @('disable') -State $state
            $script:DoDispatchLog.Count | Should -Be 0
        }
    }

    Context 'malformed input stays in mode' {
        It 'reports incomplete on empty exec tokens and does not dispatch' {
            $state = @{ Mode = 'GlobalConfig' }
            Invoke-FabriqIosDoCommand -ExecTokens @() -State $state
            $script:DoDispatchLog.Count | Should -Be 0
        }

        It 'reports error on an unknown exec command and does not dispatch' {
            $state = @{ Mode = 'GlobalConfig' }
            Invoke-FabriqIosDoCommand -ExecTokens @('bogus') -State $state
            $script:DoDispatchLog.Count | Should -Be 0
        }
    }

    Context 'context is preserved' {
        It 'never mutates the shell mode or interface context' {
            $state = @{ Mode = 'InterfaceConfig'; CurrentInterface = 'Ethernet0' }
            Invoke-FabriqIosDoCommand -ExecTokens @('show', 'version') -State $state
            $state.Mode             | Should -Be 'InterfaceConfig'
            $state.CurrentInterface | Should -Be 'Ethernet0'
        }

        It 'never mutates the module-config context' {
            $state = @{ Mode = 'ModuleConfig'; ConfigModuleName = 'acl_config'; CurrentCategoryId = 'settings' }
            Invoke-FabriqIosDoCommand -ExecTokens @('sh', 'ru') -State $state
            $state.Mode              | Should -Be 'ModuleConfig'
            $state.ConfigModuleName  | Should -Be 'acl_config'
            $state.CurrentCategoryId | Should -Be 'settings'
        }
    }
}

Describe 'design invariant - do is absent from the resolver vocabulary' {
    It 'is not a vocabulary entry in any configuration mode' {
        foreach ($mode in @('GlobalConfig', 'InterfaceConfig', 'ModuleConfig')) {
            (Get-FabriqIosCommandVocabulary -Mode $mode) | Should -Not -Contain 'do'
        }
    }
}
