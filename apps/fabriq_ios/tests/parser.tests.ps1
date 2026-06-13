# Pester tests for the command parser.
# Run: Invoke-Pester apps\fabriq_ios\tests\parser.tests.ps1

BeforeAll {
    . "$PSScriptRoot\..\lib\parser.ps1"
}

Describe 'ConvertTo-FabriqIosTokens' {
    It 'returns empty for empty input' {
        ConvertTo-FabriqIosTokens '' | Should -BeNullOrEmpty
    }

    It 'returns empty for whitespace-only input' {
        ConvertTo-FabriqIosTokens '   ' | Should -BeNullOrEmpty
    }

    It 'splits whitespace-separated tokens' {
        $r = ConvertTo-FabriqIosTokens 'show running-config'
        $r.Count | Should -Be 2
        $r[0]   | Should -Be 'show'
        $r[1]   | Should -Be 'running-config'
    }

    It 'preserves quoted arguments as a single token' {
        $r = ConvertTo-FabriqIosTokens 'interface "Ethernet 0"'
        $r.Count | Should -Be 2
        $r[0]   | Should -Be 'interface'
        $r[1]   | Should -Be 'Ethernet 0'
    }

    It 'collapses consecutive whitespace' {
        $r = ConvertTo-FabriqIosTokens '  show    version  '
        $r.Count | Should -Be 2
    }

    It 'preserves a Japanese interface alias verbatim' {
        $r = ConvertTo-FabriqIosTokens 'interface イーサネット'
        $r.Count | Should -Be 2
        $r[1]   | Should -Be 'イーサネット'
    }

    It "splits a trailing newline so a bare keyword resolves (PSReadLine 'end' continuation)" {
        # PSReadLine returns "end`n" because 'end' is an incomplete
        # PowerShell statement; the tokenizer must still yield just 'end'.
        $r = ConvertTo-FabriqIosTokens "end`n"
        $r.Count | Should -Be 1
        $r[0]    | Should -Be 'end'
    }

    It 'splits on embedded CRLF' {
        $r = ConvertTo-FabriqIosTokens "end`r`nenable"
        $r.Count | Should -Be 2
        $r[0]    | Should -Be 'end'
        $r[1]    | Should -Be 'enable'
    }

    It 'collapses consecutive newlines (no empty tokens)' {
        $r = ConvertTo-FabriqIosTokens "end`n`n"
        $r.Count | Should -Be 1
        $r[0]    | Should -Be 'end'
    }
}

Describe 'Expand-FabriqIosAbbreviation' {
    It "expands 'conf t' to 'configure terminal' in PrivilegedExec" {
        $r = Expand-FabriqIosAbbreviation -Tokens @('conf','t') -Mode 'PrivilegedExec'
        $r[0] | Should -Be 'configure'
        $r[1] | Should -Be 'terminal'
    }

    It "expands 'sh ru' to 'show running-config' in PrivilegedExec" {
        $r = Expand-FabriqIosAbbreviation -Tokens @('sh','ru') -Mode 'PrivilegedExec'
        $r[0] | Should -Be 'show'
        $r[1] | Should -Be 'running-config'
    }

    It "treats exact match as winner against longer prefix ('host' vs 'hosts')" {
        $r = Expand-FabriqIosAbbreviation -Tokens @('show','host') -Mode 'PrivilegedExec'
        $r[1] | Should -Be 'host'
    }

    It 'rejects ambiguous prefixes' {
        # In UserExec, 'e' matches both 'enable' and 'exit'.
        { Expand-FabriqIosAbbreviation -Tokens @('e') -Mode 'UserExec' } | Should -Throw
    }

    It 'returns dynamic args verbatim after expansion' {
        $r = Expand-FabriqIosAbbreviation -Tokens @('host','NEW-PC-01') -Mode 'GlobalConfig'
        $r[0] | Should -Be 'hostname'
        $r[1] | Should -Be 'NEW-PC-01'
    }
}

Describe 'Resolve-FabriqIosCommand' {
    It 'returns the canonical command record on a hit' {
        $r = Resolve-FabriqIosCommand -Tokens @('en') -Mode 'UserExec'
        $r.Command | Should -Be 'enable'
        $r.Error   | Should -BeNullOrEmpty
    }

    It 'returns Error for unknown command' {
        $r = Resolve-FabriqIosCommand -Tokens @('nonsense') -Mode 'UserExec'
        $r.Command | Should -BeNullOrEmpty
        $r.Error   | Should -Not -BeNullOrEmpty
    }

    It 'returns Error for ambiguous prefix' {
        $r = Resolve-FabriqIosCommand -Tokens @('e') -Mode 'UserExec'
        $r.Command | Should -BeNullOrEmpty
        $r.Error   | Should -Match 'Ambiguous'
    }

    It 'preserves Args including dynamic value' {
        $r = Resolve-FabriqIosCommand -Tokens @('host','NEW-PC-01') -Mode 'GlobalConfig'
        $r.Command    | Should -Be 'hostname'
        $r.Args.Count | Should -Be 1
        $r.Args[0]    | Should -Be 'NEW-PC-01'
    }
}
