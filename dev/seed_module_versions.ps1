# ========================================
# Fabriq Module VERSION / REQUIRES_KERNEL Baseline Seeder
# ========================================
# Idempotent one-time (or repeatable, safe) tool that seeds a baseline
# VERSION file and a REQUIRES_KERNEL file to every module under
# modules/standard/ and modules/extended/ that is missing one or both.
#
# Rationale:
# The lazy-seed policy (CLAUDE.md rule H) originally said "stamp 1.0.0
# the first time Claude touches a module", assuming VERSION files would
# gradually proliferate over time. In practice, that left 70+ modules
# without VERSION files, which breaks the fabriq_studio "update from
# template" feature (both-missing -> SKIP, but contents actually differ
# between old deployments and current template).
#
# Solution: batch-seed all untracked modules NOW to VERSION=1.0.0 and
# REQUIRES_KERNEL=2.0.0. After this runs:
#   - Every module has a VERSION file
#   - template vs target comparison reliably detects missing-target case
#     (Case 4 in KERNEL_API.md § 9.4: template has VERSION, target does
#     not -> UPDATE with lazy seed)
#   - Modules that already have a VERSION file (e.g. evidence_config =
#     1.1.1) are NOT touched. Their history is preserved.
#
# Baseline choice:
#   VERSION = 1.0.0 (matches CLAUDE.md rule H: "initial touch stamps 1.0.0")
#   REQUIRES_KERNEL = 2.0.0 (matches KERNEL_API.md § 8 baseline)
#
# Usage:
#   powershell.exe -File .\dev\seed_module_versions.ps1
#   powershell.exe -File .\dev\seed_module_versions.ps1 -DryRun
#
# Parameters:
#   -DryRun : list what would be created without writing any files
# ========================================

param(
    [switch]$DryRun
)

$ErrorActionPreference = "Stop"

# Repo root = parent of dev/
$repoRoot = Split-Path -Parent $PSScriptRoot

# Sanity check
$kernelVersionFile = Join-Path $repoRoot "kernel\KERNEL_VERSION"
if (-not (Test-Path $kernelVersionFile)) {
    Write-Host "[ERROR] Not a fabriq root (missing kernel\KERNEL_VERSION): $repoRoot" -ForegroundColor Red
    exit 1
}

$BASELINE_VERSION         = "1.0.0"
$BASELINE_REQUIRES_KERNEL = "2.0.0"

$modulesRoot = Join-Path $repoRoot "modules"
if (-not (Test-Path $modulesRoot)) {
    Write-Host "[ERROR] modules/ directory not found: $modulesRoot" -ForegroundColor Red
    exit 1
}

Write-Host "Repo root : $repoRoot" -ForegroundColor Cyan
Write-Host "Baseline  : VERSION=$BASELINE_VERSION, REQUIRES_KERNEL=$BASELINE_REQUIRES_KERNEL" -ForegroundColor Cyan
if ($DryRun) {
    Write-Host "Mode      : DRY RUN (no files will be written)" -ForegroundColor Yellow
}
Write-Host ""

# Enumerate module directories under modules/{standard,extended}/*
$moduleTypes = @("standard", "extended")
$allModules = @()
foreach ($t in $moduleTypes) {
    $typeDir = Join-Path $modulesRoot $t
    if (-not (Test-Path $typeDir)) { continue }
    Get-ChildItem -Path $typeDir -Directory -ErrorAction SilentlyContinue | ForEach-Object {
        $allModules += [PSCustomObject]@{
            Type = $t
            Name = $_.Name
            Path = $_.FullName
        }
    }
}

if ($allModules.Count -eq 0) {
    Write-Host "[WARN] No modules found under modules/standard/ or modules/extended/" -ForegroundColor Yellow
    exit 0
}

Write-Host "Scanning $($allModules.Count) modules..." -ForegroundColor Cyan
Write-Host ""

# Write a single-line file with trailing newline (matches existing VERSION
# files produced by lazy-seed operations in the past).
function Write-SingleLineFile {
    param(
        [string]$Path,
        [string]$Content
    )
    # Use [System.IO.File] to control the exact bytes written (avoids
    # PS Out-File inserting BOM on UTF-8). These files are ASCII-safe.
    $bytes = [System.Text.Encoding]::UTF8.GetBytes("$Content`n")
    [System.IO.File]::WriteAllBytes($Path, $bytes)
}

$stats = @{
    VersionCreated         = 0
    VersionAlreadyPresent  = 0
    RequiresCreated        = 0
    RequiresAlreadyPresent = 0
}

$actions = @()

foreach ($m in $allModules) {
    $versionFile  = Join-Path $m.Path "VERSION"
    $requiresFile = Join-Path $m.Path "REQUIRES_KERNEL"
    $shortLabel   = "modules/$($m.Type)/$($m.Name)"

    # VERSION
    if (Test-Path $versionFile) {
        $existing = (Get-Content $versionFile -Raw -ErrorAction SilentlyContinue).Trim()
        $stats.VersionAlreadyPresent++
        $actions += [PSCustomObject]@{
            Module = $shortLabel
            File   = "VERSION"
            Action = "keep"
            Value  = $existing
        }
    }
    else {
        $actions += [PSCustomObject]@{
            Module = $shortLabel
            File   = "VERSION"
            Action = "create"
            Value  = $BASELINE_VERSION
        }
        if (-not $DryRun) {
            Write-SingleLineFile -Path $versionFile -Content $BASELINE_VERSION
        }
        $stats.VersionCreated++
    }

    # REQUIRES_KERNEL
    if (Test-Path $requiresFile) {
        $existing = (Get-Content $requiresFile -Raw -ErrorAction SilentlyContinue).Trim()
        $stats.RequiresAlreadyPresent++
        $actions += [PSCustomObject]@{
            Module = $shortLabel
            File   = "REQUIRES_KERNEL"
            Action = "keep"
            Value  = $existing
        }
    }
    else {
        $actions += [PSCustomObject]@{
            Module = $shortLabel
            File   = "REQUIRES_KERNEL"
            Action = "create"
            Value  = $BASELINE_REQUIRES_KERNEL
        }
        if (-not $DryRun) {
            Write-SingleLineFile -Path $requiresFile -Content $BASELINE_REQUIRES_KERNEL
        }
        $stats.RequiresCreated++
    }
}

# Show detailed actions (create / keep)
$actionsByKind = $actions | Group-Object Action
foreach ($grp in $actionsByKind | Sort-Object Name) {
    Write-Host ("---- Action: {0} ({1}) ----" -f $grp.Name, $grp.Count) -ForegroundColor Cyan
    foreach ($a in ($grp.Group | Sort-Object Module, File)) {
        $color = if ($a.Action -eq 'create') { 'Green' } else { 'DarkGray' }
        Write-Host ("  {0,-12} {1,-50} {2}" -f $a.File, $a.Module, $a.Value) -ForegroundColor $color
    }
    Write-Host ""
}

# Summary
Write-Host "========================================" -ForegroundColor Cyan
Write-Host " Summary" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ("VERSION          : created={0}, already_present={1}" -f $stats.VersionCreated, $stats.VersionAlreadyPresent) -ForegroundColor White
Write-Host ("REQUIRES_KERNEL  : created={0}, already_present={1}" -f $stats.RequiresCreated, $stats.RequiresAlreadyPresent) -ForegroundColor White
if ($DryRun) {
    Write-Host ""
    Write-Host "DRY RUN complete. Re-run without -DryRun to apply." -ForegroundColor Yellow
}
else {
    Write-Host ""
    Write-Host "Seeding complete." -ForegroundColor Green
}
