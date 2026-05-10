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
