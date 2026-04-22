# ========================================
# Version Consistency Checker
# ========================================
# Verifies that version strings across the project match kernel/KERNEL_VERSION.
#
# Targets:
#   - kernel/KERNEL_VERSION (source of truth, X.Y.Z)
#   - README.md L1          "# Fabriq ver{X.Y}"
#   - kernel/common.ps1 L2  "Common Function Library v{X.Y.Z}"   (full)
#   - kernel/main.ps1   L3  "Fabriq ver{X.Y}"                    (major.minor)
#
# Usage:
#   pwsh ./dev/check_version.ps1
#   (exits 0 on success, 1 on any mismatch)
# ========================================

$ErrorActionPreference = "Stop"

# Resolve project root as the parent of this script's directory
$projectRoot = Split-Path -Parent $PSScriptRoot
if (-not $projectRoot) { $projectRoot = (Resolve-Path "$PSScriptRoot\..").Path }

$versionFile = Join-Path $projectRoot "kernel\KERNEL_VERSION"
if (-not (Test-Path $versionFile)) {
    Write-Host "[ERROR] kernel/KERNEL_VERSION file not found: $versionFile" -ForegroundColor Red
    exit 1
}

# Parse KERNEL_VERSION file
$rawVersion = (Get-Content $versionFile -Raw).Trim()
if ($rawVersion -notmatch '^\d+\.\d+\.\d+$') {
    Write-Host "[ERROR] KERNEL_VERSION must be SemVer X.Y.Z, got: '$rawVersion'" -ForegroundColor Red
    exit 1
}

$parts       = $rawVersion.Split('.')
$majorMinor  = "$($parts[0]).$($parts[1])"
$fullVersion = $rawVersion

Write-Host "========================================"
Write-Host " Fabriq Kernel Version Consistency Check"
Write-Host "========================================"
Write-Host ""
Write-Host "KERNEL_VERSION : $fullVersion"
Write-Host "Expected X.Y   : $majorMinor"
Write-Host ""

$mismatches = @()

function Test-VersionLine {
    param(
        [string]$Label,
        [string]$FilePath,
        [int]$LineNumber,
        [string]$Pattern,
        [string]$Expected
    )
    if (-not (Test-Path $FilePath)) {
        Write-Host "[SKIP] $Label : file not found ($FilePath)" -ForegroundColor Yellow
        return
    }
    $lines = Get-Content $FilePath
    if ($LineNumber -gt $lines.Count) {
        Write-Host "[FAIL] $Label : line $LineNumber does not exist" -ForegroundColor Red
        $script:mismatches += $Label
        return
    }
    $line = $lines[$LineNumber - 1]
    if ($line -match $Pattern) {
        $actual = $Matches[1]
        if ($actual -eq $Expected) {
            Write-Host "[ OK ] $Label : $actual" -ForegroundColor Green
        } else {
            Write-Host "[FAIL] $Label : expected '$Expected', got '$actual'" -ForegroundColor Red
            Write-Host "       line $LineNumber : $line" -ForegroundColor DarkGray
            $script:mismatches += $Label
        }
    } else {
        Write-Host "[FAIL] $Label : version pattern not found on line $LineNumber" -ForegroundColor Red
        Write-Host "       line $LineNumber : $line" -ForegroundColor DarkGray
        $script:mismatches += $Label
    }
}

# README.md L1 : "# Fabriq ver{X.Y}"
Test-VersionLine `
    -Label "README.md L1" `
    -FilePath (Join-Path $projectRoot "README.md") `
    -LineNumber 1 `
    -Pattern '(?i)Fabriq\s+ver(\d+\.\d+)' `
    -Expected $majorMinor

# kernel/common.ps1 L2 : "Common Function Library v{X.Y.Z}"  (full version)
Test-VersionLine `
    -Label "kernel/common.ps1 L2" `
    -FilePath (Join-Path $projectRoot "kernel\common.ps1") `
    -LineNumber 2 `
    -Pattern 'Common Function Library\s+v(\d+\.\d+\.\d+)' `
    -Expected $fullVersion

# kernel/main.ps1 L3 : "Fabriq ver{X.Y}"
Test-VersionLine `
    -Label "kernel/main.ps1 L3" `
    -FilePath (Join-Path $projectRoot "kernel\main.ps1") `
    -LineNumber 3 `
    -Pattern '(?i)Fabriq\s+ver(\d+\.\d+)' `
    -Expected $majorMinor

Write-Host ""
if ($mismatches.Count -eq 0) {
    Write-Host "All version strings are consistent (X.Y = $majorMinor, full = $fullVersion)." -ForegroundColor Green
    exit 0
} else {
    Write-Host "Mismatches found in:" -ForegroundColor Red
    foreach ($m in $mismatches) { Write-Host "  - $m" -ForegroundColor Red }
    Write-Host ""
    Write-Host "Fix the files above to match KERNEL_VERSION ($fullVersion), or update KERNEL_VERSION." -ForegroundColor Yellow
    exit 1
}
