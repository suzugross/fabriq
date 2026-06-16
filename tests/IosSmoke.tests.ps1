# ========================================
# Pester v5 wrapper for the fabriq_ios phase smoke suites
# ========================================
# Run    : powershell.exe -File ./dev/run_tests.ps1
#
# apps/fabriq_ios/tests/_phase*_smoke.ps1 are standalone integration smoke
# scripts written with a self-contained Check / "exit $Fail" harness. They
# are NOT named *.tests.ps1, so run_tests.ps1's Pester discovery never picks
# them up and their ~339 assertions only run when launched by hand.
#
# This wrapper folds them into the suite without rewriting them: each smoke
# script is executed as an isolated child powershell.exe process and its
# exit code (0 = all Checks passed) is asserted. The Check harness is left
# untouched, so the smoke scripts remain runnable on their own too.
# ========================================

$script:RepoRoot    = Split-Path -Parent $PSScriptRoot
$script:IosTestsDir = Join-Path $script:RepoRoot 'apps\fabriq_ios\tests'

# Enumerate at DISCOVERY time so -ForEach generates one It per smoke script.
$smokeFiles = @(Get-ChildItem -Path $script:IosTestsDir -Filter '_phase*_smoke.ps1' -ErrorAction SilentlyContinue |
    Sort-Object Name)

# Regression guard: a vanished smoke suite must fail loudly, not silently
# shrink coverage to zero It blocks.
if ($smokeFiles.Count -eq 0) {
    throw "No _phase*_smoke.ps1 found under $script:IosTestsDir"
}

$smokeCases = @($smokeFiles | ForEach-Object { @{ Name = $_.Name; Path = $_.FullName } })

Describe 'fabriq_ios phase smoke suites (wrapped)' {
    It 'exits 0 (all Checks pass): <Name>' -ForEach $smokeCases {
        $output = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $Path 2>&1
        $code = $LASTEXITCODE
        if ($code -ne 0) {
            # Surface the child output so a failing smoke run is diagnosable
            # from the Pester log instead of an opaque non-zero exit.
            Write-Host "--- $Name (exit $code) ---" -ForegroundColor Red
            Write-Host ($output | Out-String)
        }
        $code | Should -Be 0
    }
}
