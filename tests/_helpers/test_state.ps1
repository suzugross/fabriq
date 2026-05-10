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
