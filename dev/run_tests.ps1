# ========================================
# Fabriq Test Runner (Pester v5+)
# ========================================
# Discovers and runs all *.tests.ps1 under tests/ and apps/fabriq_ios/tests/.
#
# Prerequisite: Pester v5+ must be installed.
#   Install-Module -Name Pester -MinimumVersion 5.0.0 -Scope CurrentUser -Force -SkipPublisherCheck
#
# The Pester module that ships with Windows is v3.4.0, which is signed by
# Microsoft and cannot be auto-upgraded silently; the -SkipPublisherCheck
# flag is required because v5 is signed by the Pester team.
#
# Usage:
#   pwsh ./dev/run_tests.ps1     (preferred, PowerShell 7+)
#   powershell -File ./dev/run_tests.ps1
# Exit code: 0 = all pass, 1 = any failure or environment problem.
# ========================================

$ErrorActionPreference = 'Stop'

$pesterV5 = Get-Module -Name Pester -ListAvailable |
    Where-Object { $_.Version.Major -ge 5 } |
    Sort-Object Version -Descending |
    Select-Object -First 1

if (-not $pesterV5) {
    Write-Host '[ERROR] Pester v5+ is required but not installed.' -ForegroundColor Red
    Write-Host ''
    Write-Host 'Install via:' -ForegroundColor Yellow
    Write-Host '  Install-Module -Name Pester -MinimumVersion 5.0.0 -Scope CurrentUser -Force -SkipPublisherCheck' -ForegroundColor White
    Write-Host ''
    Write-Host 'Currently installed Pester versions:' -ForegroundColor Yellow
    $installed = Get-Module -Name Pester -ListAvailable | Select-Object Version, Path
    if ($installed) {
        $installed | Format-Table -AutoSize | Out-String | Write-Host
    } else {
        Write-Host '  (none)' -ForegroundColor DarkGray
    }
    exit 1
}

Import-Module Pester -MinimumVersion 5.0.0 -Force

$repoRoot = Split-Path -Parent $PSScriptRoot
$testPaths = @(
    (Join-Path $repoRoot 'tests'),
    (Join-Path $repoRoot 'apps\fabriq_ios\tests')
) | Where-Object { Test-Path $_ }

if (-not $testPaths) {
    Write-Host "[ERROR] No test directories found under $repoRoot" -ForegroundColor Red
    exit 1
}

Write-Host '========================================' -ForegroundColor Cyan
Write-Host " Fabriq Test Runner (Pester $($pesterV5.Version))" -ForegroundColor Cyan
Write-Host '========================================' -ForegroundColor Cyan
foreach ($p in $testPaths) {
    Write-Host "  Path: $p" -ForegroundColor Gray
}
Write-Host ''

$config = New-PesterConfiguration
$config.Run.Path        = $testPaths
$config.Run.Exit        = $true
$config.Output.Verbosity = 'Detailed'

Invoke-Pester -Configuration $config
