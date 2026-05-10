# ========================================
# Fabriq Test Helpers - Path / State
# ========================================
# Path resolution and (future) script-scope state setup helpers shared
# across kernel unit tests. Production code is never touched by tests;
# any state that production normally sets via Initialize-Session is
# expected to be configured here when Phase 2 (Save-/Load-ResumeState
# tests) lands.
# ========================================

function Get-FabriqRepoRoot {
    # tests/_helpers/test_state.ps1 -> repo root is two levels up.
    return (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
}

function Get-FabriqMainFunctionScriptBlock {
    # Extracts a single named function from kernel/main.ps1 as a script
    # block, without executing main.ps1's top-level startup code (which
    # would call Enable-SleepSuppression / Set-ConsoleSize / load
    # fabriq_operator GUI / exit 1 if no GUI - none of which is sane in
    # a test runner).
    #
    # Approach: parse main.ps1 with the PowerShell language Parser into
    # an AST, find the requested FunctionDefinitionAst by name, and
    # return [scriptblock]::Create on its source extent. Callers in
    # BeforeAll then dot-source the returned scriptblock to define the
    # function in the test scope:
    #
    #   $sb = Get-FabriqMainFunctionScriptBlock -Name 'Set-SelectedHostEnvironment'
    #   . $sb
    #
    # Throws if main.ps1 fails to parse or the named function is absent
    # (regression signal: rename / removal of the production function
    # the test was pinning).
    param(
        [Parameter(Mandatory)][string]$Name
    )
    $mainPath = Join-Path (Get-FabriqRepoRoot) 'kernel\main.ps1'
    if (-not (Test-Path $mainPath)) {
        throw "kernel/main.ps1 not found at $mainPath"
    }

    $tokens = $null
    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile(
        $mainPath, [ref]$tokens, [ref]$errors
    )
    if ($errors -and $errors.Count -gt 0) {
        $msg = ($errors | ForEach-Object { $_.Message }) -join '; '
        throw "Parse errors in main.ps1: $msg"
    }

    $fn = $ast.FindAll({
        param($node)
        $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
        $node.Name -eq $Name
    }, $true) | Select-Object -First 1

    if (-not $fn) {
        throw "Function '$Name' not found in kernel/main.ps1"
    }

    return [scriptblock]::Create($fn.Extent.Text)
}

function Set-FabriqTestState {
    # Sets script-scope state that production normally configures via
    # Initialize-Session in kernel/main.ps1. Tests dot-source common.ps1
    # without going through main.ps1, so any function that reads
    # $script:ResumeStatePath / $script:SessionID / $script:HistoryPath
    # would otherwise see the production defaults (relative paths
    # rooted at the test runner's CWD, which would clobber real state).
    #
    # Both this helper and common.ps1 are dot-sourced into the same
    # Pester scope, so $script: in either resolves to the same backing
    # variable.
    param(
        [string]$ResumeStatePath,
        [string]$SessionID = 'fabriq-test',
        [string]$HistoryPath,
        [string]$StatusFilePath
    )
    if ($PSBoundParameters.ContainsKey('ResumeStatePath')) {
        $script:ResumeStatePath = $ResumeStatePath
    }
    $script:SessionID = $SessionID
    if ($PSBoundParameters.ContainsKey('HistoryPath')) {
        $script:HistoryPath = $HistoryPath
    }
    if ($PSBoundParameters.ContainsKey('StatusFilePath')) {
        $script:StatusFilePath = $StatusFilePath
    }
}
