# ========================================
# Pester v5 source-contract test (TM t-0043)
# ========================================
# fabriq_ios reads commands via [PSConsoleReadLine]::ReadLine, whose
# default HistorySaveStyle (SaveIncrementally) persists every accepted
# line - including `set <col> <secret>` - to the shared, per-user
# ConsoleHost_history.txt in plaintext. Initialize-FabriqIos must disable
# that persistence (HistorySaveStyle SaveNothing) so secrets typed in the
# REPL never leak to a later PowerShell session.
#
# This is a SOURCE contract (AST parse, no execution) because the main
# fabriq_ios.ps1 self-spawns a subprocess when dot-sourced, so the
# function cannot be invoked directly under Pester. Pinning the option
# call guards against a silent regression of the security fix.
# Run: powershell.exe -File ./dev/run_tests.ps1
# ========================================

BeforeAll {
    $script:IosScript = (Resolve-Path (Join-Path $PSScriptRoot '..\fabriq_ios.ps1')).Path
    $tokens = $null
    $errors = $null
    $script:IosAst = [System.Management.Automation.Language.Parser]::ParseFile(
        $script:IosScript, [ref]$tokens, [ref]$errors)
    $script:IosParseErrors = $errors
}

Describe 'fabriq_ios history privacy (TM t-0043)' {

    It 'fabriq_ios.ps1 parses without errors' {
        @($script:IosParseErrors).Count | Should -Be 0
    }

    It 'Initialize-FabriqIos disables PSReadLine history file persistence' {
        $fn = $script:IosAst.FindAll({
            param($n)
            $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
            $n.Name -eq 'Initialize-FabriqIos'
        }, $true) | Select-Object -First 1

        $fn | Should -Not -BeNullOrEmpty

        # Find Set-PSReadLineOption calls inside the function and require one
        # that turns history persistence off (HistorySaveStyle SaveNothing).
        $calls = $fn.Body.FindAll({
            param($n)
            $n -is [System.Management.Automation.Language.CommandAst] -and
            $n.GetCommandName() -eq 'Set-PSReadLineOption'
        }, $true)

        $hasSaveNothing = @($calls | Where-Object {
            $t = $_.Extent.Text
            $t -match 'HistorySaveStyle' -and $t -match 'SaveNothing'
        }).Count -ge 1

        $hasSaveNothing | Should -BeTrue
    }
}
