# ========================================
# Pester v5 unit tests for domain_join's pure logic
# ========================================
# Functions: modules/standard/domain_join/domain_join.ps1 ::
#            Get-PendingComputerName / Write-JoinDiagnostics (status decoding)
# Run     : powershell.exe -File ./dev/run_tests.ps1
#
# The module body runs on import (there is no -LoadOnly switch on a shipped
# module and adding one would change its contract), so the function definitions
# are lifted out of the file with the AST and evaluated on their own. Nothing
# outside the function bodies is executed.
#
# What is covered is the staged-rename detection that decides whether the join
# runs under the old or the new computer name. Getting that wrong produces a
# machine whose AD account name stops matching after the reboot, which breaks
# the secure channel before the PC is delivered.
# ========================================

BeforeAll {
    . "$PSScriptRoot\..\..\_helpers\test_state.ps1"
    $script:RepoRoot = Get-FabriqRepoRoot
    $script:ModuleScript = Join-Path $script:RepoRoot 'modules\standard\domain_join\domain_join.ps1'

    $tokens = $null
    $parseErrors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile(
        $script:ModuleScript, [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors -and $parseErrors.Count -gt 0) {
        throw "domain_join.ps1 failed to parse: $($parseErrors[0].Message)"
    }

    $script:Functions = $ast.FindAll(
        { param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] },
        $false)
    foreach ($fn in $script:Functions) {
        . ([scriptblock]::Create($fn.Extent.Text))
    }

    $script:ModuleText = Get-Content -LiteralPath $script:ModuleScript -Raw
}

Describe 'Get-PendingComputerName' {

    It 'does not throw on a machine with no staged rename' {
        { Get-PendingComputerName } | Should -Not -Throw
    }

    It 'returns either $null or a non-empty name' {
        $result = Get-PendingComputerName
        if ($null -ne $result) { $result | Should -Not -BeNullOrEmpty }
    }

    It 'reads the pending name from ComputerName, not ActiveComputerName' {
        # The whole point is the DIFFERENCE between the two keys. A version
        # that reads only one of them would compile and silently never detect
        # a staged rename, so pin both reads.
        $body = ($script:Functions | Where-Object { $_.Name -eq 'Get-PendingComputerName' }).Extent.Text
        $body | Should -Match 'Control\\ComputerName'
        $body | Should -Match 'ActiveComputerName'
    }
}

Describe 'Staged-rename join options' {

    It 'passes AccountCreate together with JoinWithNewName' {
        # -Options REPLACES the cmdlet default (AccountCreate), it does not add
        # to it. JoinWithNewName on its own would drop account creation and
        # make every first-time join fail, so the pair must stay together.
        $script:ModuleText | Should -Match "@\('AccountCreate',\s*'JoinWithNewName'\)"
    }
}

Describe 'netsetup.log status decoding' {

    It 'decodes every status code measured against a real DC' {
        # 0x2 (missing / non-OU path) and 0x8b0 (name exists, and with 'ou' set
        # it exists in a different OU) were both measured on Windows Server
        # 2025; 0xaac and 0x5 are the two re-use gates.
        foreach ($code in @('0x2', '0x8b0', '0xaac', '0x5', '0x525')) {
            $script:ModuleText | Should -Match ([regex]::Escape("'$code'"))
        }
    }

    It 'does not match on the localized exception text anywhere' {
        # The Win32 tail of an Add-Computer error is localized, so any
        # substring match against it stops firing on a Japanese Windows -
        # exactly the machines this module runs on.
        $script:ModuleText | Should -Not -Match '\$errorMsg\s+-(?:like|match)'
    }
}

Describe 'OU placement is reported, never verified' {

    It 'does not build a directory query' {
        # Measured 2026-09-13: the join API itself refuses a mismatch (0x8b0),
        # so a read-back could only ever return "matches". If someone adds one
        # later, this test should make them justify it first.
        $script:ModuleText | Should -Not -Match 'DirectorySearcher'
        $script:ModuleText | Should -Not -Match 'LDAP://'
    }

    It 'still reports the requested OU to the operator' {
        $script:ModuleText | Should -Match 'Requested OU'
    }
}
