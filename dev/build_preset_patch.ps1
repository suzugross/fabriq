# ========================================
# Fabriq Preset Patch Builder
# ========================================
# Builds a NARROW patch folder containing ONLY the per-module preset.csv
# files (UI dropdown definitions) mirrored from the fabriq source tree,
# preserving the modules/<type>/<name>/ structure. Drops a self-contained
# Apply-Presets.ps1 + README into the folder so the result can be carried
# to a deployed fabriq and overlaid without touching any other file.
#
# preset.csv is a FRAMEWORK artifact (whitelisted in
# dev/framework_overlay_rules.json), NOT site-specific kitting data, so it
# is safe to overwrite on the target. This builder is the "targeted patch"
# counterpart to build_framework_patch.ps1 (full-tree patch), scoped to a
# single artifact family.
#
# Re-run this builder any time to refresh the folder to the latest source
# ("latest version" patch).
#
# Usage:
#   powershell.exe -File .\dev\build_preset_patch.ps1
#   powershell.exe -File .\dev\build_preset_patch.ps1 -OutDir D:\share\patches
#   powershell.exe -File .\dev\build_preset_patch.ps1 -PatchName my-preset-patch
#
# Parameters:
#   -OutDir    : parent directory where the patch folder is created
#                (default: $env:USERPROFILE\Desktop)
#   -PatchName : patch folder name
#                (default: fabriq_preset_patch_{yyyy-MM-dd}_kernel-v{KERNEL_VERSION})
# ========================================

param(
    [string]$OutDir = (Join-Path $env:USERPROFILE "Desktop"),
    [string]$PatchName = $null
)

$ErrorActionPreference = "Stop"

# ---- Resolve source = repo root (parent of dev/ where this script lives) ----
$src = Split-Path -Parent $PSScriptRoot

# ---- Sanity check: this must be a fabriq root ----
$kernelVersionFile = Join-Path $src "kernel\KERNEL_VERSION"
if (-not (Test-Path $kernelVersionFile)) {
    Write-Host "[ERROR] Not a fabriq root (missing kernel\KERNEL_VERSION): $src" -ForegroundColor Red
    exit 1
}
$kernelVersion = (Get-Content $kernelVersionFile -Raw).Trim()

$modulesSrc = Join-Path $src "modules"
if (-not (Test-Path $modulesSrc)) {
    Write-Host "[ERROR] Source modules directory not found: $modulesSrc" -ForegroundColor Red
    exit 1
}

# ---- Auto-generate patch name if not specified ----
if ([string]::IsNullOrWhiteSpace($PatchName)) {
    $dateStr = Get-Date -Format "yyyy-MM-dd"
    $PatchName = "fabriq_preset_patch_${dateStr}_kernel-v${kernelVersion}"
}

# ---- Destructive path guard (CLAUDE.md section 8) ----
# $dst is Remove-Item -Recurse'd below. PatchName must therefore be a
# single safe path component: "." resolves $dst to $OutDir itself
# (default: the user's Desktop) and "..\name" escapes it - either way
# the wrong tree gets deleted. Same rules as build_framework_patch.ps1.
if ($PatchName -eq '.' -or $PatchName -eq '..' -or
    $PatchName.IndexOfAny([System.IO.Path]::GetInvalidFileNameChars()) -ge 0 -or
    $PatchName -ne $PatchName.TrimEnd('.', ' ')) {
    Write-Host "[ERROR] Invalid -PatchName '$PatchName'." -ForegroundColor Red
    Write-Host "        Must be a single folder name: no path separators, no '.' / '..'," -ForegroundColor Red
    Write-Host "        no wildcard or other invalid filename chars, no trailing dot/space." -ForegroundColor Red
    exit 1
}

# ---- Validate output directory ----
if (-not (Test-Path $OutDir)) {
    Write-Host "[ERROR] Output directory does not exist: $OutDir" -ForegroundColor Red
    Write-Host "Create it first, or pass an existing path via -OutDir." -ForegroundColor Red
    exit 1
}

$dst = Join-Path $OutDir $PatchName

Write-Host "Source : $src" -ForegroundColor Cyan
Write-Host "Dest   : $dst" -ForegroundColor Cyan
Write-Host "Kernel : $kernelVersion" -ForegroundColor Cyan
Write-Host "Scope  : modules/**/preset.csv only" -ForegroundColor Cyan
Write-Host ""

if (Test-Path $dst) {
    # Containment assert: never delete anything outside $OutDir
    # (belt-and-suspenders behind the PatchName guard above).
    $outFull = [System.IO.Path]::GetFullPath($OutDir).TrimEnd('\')
    $dstFull = [System.IO.Path]::GetFullPath($dst).TrimEnd('\')
    if (-not $dstFull.StartsWith($outFull + '\')) {
        Write-Host "[ERROR] Destination escapes OutDir, refusing to delete: $dstFull" -ForegroundColor Red
        exit 1
    }
    Write-Host "Removing existing patch folder..." -ForegroundColor Yellow
    Remove-Item -Path $dst -Recurse -Force
}
$null = New-Item -Path $dst -ItemType Directory -Force

# ========================================
# Step 1: Mirror modules/**/preset.csv (structure preserved)
# ========================================
# robocopy copies preset.csv recursively under modules/, preserving the
# modules/<type>/<name>/ directory layout. Binary copy preserves the
# UTF-8 BOM that these CSVs require (CLAUDE.md section 7).
Write-Host "[1/3] Mirroring modules/**/preset.csv..." -ForegroundColor Cyan
$dstModules = Join-Path $dst "modules"
$robocopyArgs = @($modulesSrc, $dstModules, "preset.csv", "/S", "/NFL", "/NDL", "/NJH", "/NJS", "/NC", "/NS", "/NP")
$null = & robocopy @robocopyArgs
# robocopy exit codes: 0-7 = success (various), 8+ = error
if ($LASTEXITCODE -ge 8) {
    Write-Host "[ERROR] robocopy failed with exit code $LASTEXITCODE" -ForegroundColor Red
    exit 1
}
$presetFiles = @(Get-ChildItem $dstModules -Recurse -Filter "preset.csv" -ErrorAction SilentlyContinue)
Write-Host "  preset.csv files mirrored: $($presetFiles.Count)" -ForegroundColor DarkGray

if ($presetFiles.Count -eq 0) {
    Write-Host "[ERROR] No preset.csv files were mirrored. Aborting." -ForegroundColor Red
    exit 1
}

# ========================================
# Step 2: Write Apply-Presets.ps1 (run on the target)
# ========================================
Write-Host "[2/3] Writing Apply-Presets.ps1..." -ForegroundColor Cyan
$applyScript = @'
# ========================================
# Fabriq Preset Patch - Applier
# ========================================
# Overlays the preset.csv files bundled in this folder onto a deployed
# fabriq, in place, one module at a time. Only modules that ALREADY exist
# on the target are touched - this patch never creates a new module
# directory (a preset.csv without its module scripts is useless).
#
# preset.csv is a framework UI artifact, not site-specific kitting data,
# so overwriting it is safe. No other file on the target is modified.
#
# Usage (run on the target machine):
#   powershell.exe -File .\Apply-Presets.ps1 -TargetRoot E:\fabriq
#   powershell.exe -File .\Apply-Presets.ps1 -TargetRoot E:\fabriq -DryRun
#
# Parameters:
#   -TargetRoot : path to the deployed fabriq root (the folder that
#                 contains kernel\ and modules\). REQUIRED.
#   -DryRun     : report what would change without copying anything.
# ========================================

param(
    [Parameter(Mandatory = $true)]
    [string]$TargetRoot,
    [switch]$DryRun
)

$ErrorActionPreference = "Stop"

# ---- Patch source = this folder ----
$patchRoot = $PSScriptRoot
$patchModules = Join-Path $patchRoot "modules"
if (-not (Test-Path $patchModules)) {
    Write-Host "[ERROR] This patch folder has no modules\ subtree: $patchModules" -ForegroundColor Red
    exit 1
}

# ---- Validate target is a fabriq root ----
if (-not (Test-Path $TargetRoot)) {
    Write-Host "[ERROR] TargetRoot does not exist: $TargetRoot" -ForegroundColor Red
    exit 1
}
$targetKernelVersion = Join-Path $TargetRoot "kernel\KERNEL_VERSION"
if (-not (Test-Path $targetKernelVersion)) {
    Write-Host "[ERROR] TargetRoot is not a fabriq root (missing kernel\KERNEL_VERSION): $TargetRoot" -ForegroundColor Red
    Write-Host "        Pass the deployed fabriq folder, e.g. -TargetRoot E:\fabriq" -ForegroundColor Red
    exit 1
}
$targetModules = Join-Path $TargetRoot "modules"
if (-not (Test-Path $targetModules)) {
    Write-Host "[ERROR] TargetRoot has no modules\ directory: $targetModules" -ForegroundColor Red
    exit 1
}

Write-Host "Patch  : $patchRoot" -ForegroundColor Cyan
Write-Host "Target : $TargetRoot (kernel $((Get-Content $targetKernelVersion -Raw).Trim()))" -ForegroundColor Cyan
if ($DryRun) { Write-Host "Mode   : DRY RUN (no files will be written)" -ForegroundColor Yellow }
Write-Host ""

$presetFiles = @(Get-ChildItem $patchModules -Recurse -Filter "preset.csv" -ErrorAction SilentlyContinue)
if ($presetFiles.Count -eq 0) {
    Write-Host "[ERROR] No preset.csv files found in this patch folder." -ForegroundColor Red
    exit 1
}

$applied = 0
$unchanged = 0
$skippedMissing = 0
$failed = 0

foreach ($f in $presetFiles) {
    # Relative path like modules\standard\app_config\preset.csv
    $rel = $f.FullName.Substring($patchModules.Length).TrimStart('\')
    $relFull = Join-Path "modules" $rel
    $targetFile = Join-Path $targetModules $rel
    $targetDir = Split-Path -Parent $targetFile

    if (-not (Test-Path $targetDir)) {
        Write-Host "  [SKIP] $relFull (module not on target)" -ForegroundColor DarkYellow
        $skippedMissing++
        continue
    }

    # Skip identical content to keep the change set honest.
    if (Test-Path $targetFile) {
        $srcHash = (Get-FileHash -Path $f.FullName -Algorithm SHA256).Hash
        $dstHash = (Get-FileHash -Path $targetFile -Algorithm SHA256).Hash
        if ($srcHash -eq $dstHash) {
            Write-Host "  [SAME] $relFull" -ForegroundColor DarkGray
            $unchanged++
            continue
        }
    }

    if ($DryRun) {
        Write-Host "  [WOULD] $relFull" -ForegroundColor Cyan
        $applied++
        continue
    }

    try {
        Copy-Item -Path $f.FullName -Destination $targetFile -Force
        Write-Host "  [OK]   $relFull" -ForegroundColor Green
        $applied++
    }
    catch {
        Write-Host "  [FAIL] $relFull : $_" -ForegroundColor Red
        $failed++
    }
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Green
$verb = if ($DryRun) { "Would apply" } else { "Applied" }
Write-Host (" {0}            : {1}" -f $verb, $applied) -ForegroundColor White
Write-Host (" Unchanged (same)     : {0}" -f $unchanged) -ForegroundColor White
Write-Host (" Skipped (no module)  : {0}" -f $skippedMissing) -ForegroundColor White
Write-Host (" Failed               : {0}" -f $failed) -ForegroundColor White
Write-Host "========================================" -ForegroundColor Green

if ($failed -gt 0) { exit 1 }
exit 0
'@
$applyScript | Out-File -FilePath (Join-Path $dst "Apply-Presets.ps1") -Encoding ASCII -Force

# ========================================
# Step 3: Write README
# ========================================
Write-Host "[3/3] Writing PRESET_PATCH_README.md..." -ForegroundColor Cyan
$generatedAt = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$readme = @"
# Fabriq Preset Patch ($PatchName)

Generated : $generatedAt
Source    : $src
Kernel    : $kernelVersion
Scope     : modules/**/preset.csv ONLY ($($presetFiles.Count) files)

## What this is

A narrow patch that refreshes every module's ``preset.csv`` (the UI
dropdown definitions) on a deployed fabriq to the latest source version.
``preset.csv`` is a framework artifact, not site-specific kitting data, so
overwriting it is safe. NO other file is touched.

## How to apply

1. Copy this whole folder to the target machine (or a share it can reach).
2. From inside this folder, run:

   ``powershell.exe -File .\Apply-Presets.ps1 -TargetRoot E:\fabriq``

   (replace ``E:\fabriq`` with the deployed fabriq root - the folder that
   contains ``kernel\`` and ``modules\``).

3. To preview without writing, add ``-DryRun``:

   ``powershell.exe -File .\Apply-Presets.ps1 -TargetRoot E:\fabriq -DryRun``

## Behaviour

- Only modules that already exist on the target are updated. A module
  present here but absent on the target is reported as ``[SKIP]`` (a
  preset.csv without its module scripts is useless).
- Files whose content already matches the target are reported ``[SAME]``
  and not rewritten.
- Exit code is 0 on success, 1 if any copy failed.

## Refreshing this patch

Re-run ``dev\build_preset_patch.ps1`` on the source machine to regenerate
this folder from the latest source.
"@
$readme | Out-File -FilePath (Join-Path $dst "PRESET_PATCH_README.md") -Encoding UTF8 -Force

# ========================================
# Final summary
# ========================================
Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host " Preset patch built: $dst" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
$allFiles = @(Get-ChildItem $dst -Recurse -File)
$totalSize = ($allFiles | Measure-Object -Property Length -Sum).Sum
Write-Host ("preset.csv files : {0}" -f $presetFiles.Count) -ForegroundColor White
Write-Host ("Total files      : {0}" -f $allFiles.Count) -ForegroundColor White
Write-Host ("Size             : {0:N1} KB" -f ($totalSize / 1KB)) -ForegroundColor White
Write-Host ""
Write-Host "Apply on target with:" -ForegroundColor Cyan
Write-Host "  powershell.exe -File .\Apply-Presets.ps1 -TargetRoot E:\fabriq" -ForegroundColor White
Write-Host ""
