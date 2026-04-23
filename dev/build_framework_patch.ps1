# ========================================
# Fabriq Framework Patch Builder
# ========================================
# Mirrors the fabriq source tree into an output directory, excluding
# site-specific configuration CSVs and runtime artifacts. Produces a
# full-distribution patch that can be overlaid onto a deployed fabriq
# instance without clobbering the operator's kitting data.
#
# This is the FRAMEWORK patch builder (full tree, site-CSV stripped).
# For small targeted patches (hand-picked changed files only), mirror
# the directory structure manually on the destination instead.
#
# Usage:
#   powershell.exe -File .\dev\build_framework_patch.ps1
#   powershell.exe -File .\dev\build_framework_patch.ps1 -OutDir D:\share\patches
#   powershell.exe -File .\dev\build_framework_patch.ps1 -PatchName my-patch
#   powershell.exe -File .\dev\build_framework_patch.ps1 -Purpose "release candidate 1"
#
# Parameters:
#   -OutDir    : parent directory where the patch folder is created
#                (default: $env:USERPROFILE\Desktop)
#   -PatchName : patch folder name
#                (default: fabriq_patch_{yyyy-MM-dd}_kernel-v{KERNEL_VERSION})
#   -Purpose   : optional free-text purpose string injected into
#                PATCH_README.md (default: omitted)
# ========================================

param(
    [string]$OutDir = (Join-Path $env:USERPROFILE "Desktop"),
    [string]$PatchName = $null,
    [string]$Purpose = $null
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

# ---- Auto-generate patch name if not specified ----
if ([string]::IsNullOrWhiteSpace($PatchName)) {
    $dateStr = Get-Date -Format "yyyy-MM-dd"
    $PatchName = "fabriq_patch_${dateStr}_kernel-v${kernelVersion}"
}

# ---- Validate / create output directory ----
if (-not (Test-Path $OutDir)) {
    Write-Host "[ERROR] Output directory does not exist: $OutDir" -ForegroundColor Red
    Write-Host "Create it first, or pass an existing path via -OutDir." -ForegroundColor Red
    exit 1
}

$dst = Join-Path $OutDir $PatchName

Write-Host "Source : $src" -ForegroundColor Cyan
Write-Host "Dest   : $dst" -ForegroundColor Cyan
Write-Host "Kernel : $kernelVersion" -ForegroundColor Cyan
Write-Host ""

if (Test-Path $dst) {
    Write-Host "Removing existing patch folder..." -ForegroundColor Yellow
    Remove-Item -Path $dst -Recurse -Force
}
$null = New-Item -Path $dst -ItemType Directory -Force

# ========================================
# Step 1: Mirror source tree
# ========================================
# Excludes dev/runtime directories but keeps everything else. Site-specific
# and runtime files are stripped in Steps 2 and 3 below.
Write-Host "[1/4] Mirroring source tree (excluding .git / .claude / evidence / logs)..." -ForegroundColor Cyan
$robocopyArgs = @(
    $src, $dst, "/E",
    "/XD", ".git", ".claude", "evidence", "logs",
    "/NFL", "/NDL", "/NJH", "/NJS", "/NC", "/NS", "/NP"
)
$null = & robocopy @robocopyArgs
# robocopy exit codes: 0-7 = success (various), 8+ = error
if ($LASTEXITCODE -ge 8) {
    Write-Host "[ERROR] robocopy failed with exit code $LASTEXITCODE" -ForegroundColor Red
    exit 1
}

# ========================================
# Step 2: Strip kernel-level site files and runtime artifacts
# ========================================
Write-Host "[2/4] Stripping kernel-level site-specific files and runtime artifacts..." -ForegroundColor Cyan
$removeList = @(
    ".gitignore",
    "kernel\csv\hostlist.csv",
    "kernel\csv\workers.csv",
    "kernel\csv\log_destinations.csv",
    "kernel\json\art_pulse.txt",
    "kernel\json\resume_state.json",
    "kernel\json\session.json",
    "kernel\json\skip_request.flag",
    "kernel\json\status.json",
    "kernel\txt\passphrase_verify.txt",
    "kernel\txt\silence.flag",
    "profiles\Custom Plan.csv"
)
$removedCount = 0
foreach ($rel in $removeList) {
    $p = Join-Path $dst $rel
    if (Test-Path $p) {
        Remove-Item $p -Force -ErrorAction SilentlyContinue
        Write-Host "  - removed $rel" -ForegroundColor DarkGray
        $removedCount++
    }
}

# ========================================
# Step 3: Strip all module CSVs except module.csv and preset.csv
# ========================================
# Strict rule: any CSV under modules/ is treated as site-specific
# configuration (e.g. *_list.csv, office_key.csv, license_key.csv,
# domain.csv, builtin_admin.csv). Only module.csv (menu metadata)
# and preset.csv (UI dropdown definitions) are framework artifacts.
Write-Host "[3/4] Stripping module-level CSVs (keeping only module.csv / preset.csv)..." -ForegroundColor Cyan
$keepModuleCsvs = @("module.csv", "preset.csv")
$moduleCsvs = @(Get-ChildItem (Join-Path $dst "modules") -Recurse -Filter "*.csv" -ErrorAction SilentlyContinue)
$moduleStripped = 0
$moduleKept = 0
foreach ($f in $moduleCsvs) {
    if ($f.Name -in $keepModuleCsvs) {
        $moduleKept++
        continue
    }
    $relPath = $f.FullName.Substring($dst.Length + 1)
    Remove-Item $f.FullName -Force -ErrorAction SilentlyContinue
    Write-Host "  - removed $relPath" -ForegroundColor DarkGray
    $moduleStripped++
}
Write-Host "  (kept: $moduleKept files; stripped: $moduleStripped files)" -ForegroundColor DarkCyan

# ========================================
# Step 4: Write PATCH_README.md
# ========================================
Write-Host "[4/4] Writing PATCH_README.md..." -ForegroundColor Cyan
$generatedAt = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

$purposeBlock = ""
if (-not [string]::IsNullOrWhiteSpace($Purpose)) {
    $purposeBlock = @"

## Patch purpose (this build)

$Purpose

"@
}

$manifest = @"
# Fabriq Framework Patch ($PatchName)

Generated : $generatedAt
Source    : $src
Kernel    : $kernelVersion
$purposeBlock
## Purpose

Apply framework updates to an existing deployed fabriq without touching
the site-specific kitting configuration. Overlay this folder on top of
the deployed fabriq directory to upgrade code, docs, and framework
definitions while preserving the operator's CSV data.

## Included

- All PowerShell sources (``kernel/``, ``apps/``, ``commands/``, ``modules/``)
- Framework metadata (module-level): ``module.csv``, ``preset.csv`` only
- Framework metadata (kernel-level): ``categories.csv``, ``manifesto.csv``
- Profile templates under ``profiles/``:
  - ``Master_*.csv`` (framework master profiles)
  - ``_test_harness.csv``, ``_test_harness_async.csv`` (test profiles)
  - ``sysprep.csv`` (system profile)
  - ``easy_template/`` subfolder
- Documentation: ``README.md``, ``CHANGELOG.md``, ``CLAUDE.md``, ``LICENSE``
- Kernel API artifacts: ``kernel/KERNEL_VERSION``, ``kernel/KERNEL_API.md``
- Framework config: ``kernel/json/async_config.json``
- ``dev/`` tooling (template, check_version, build_framework_patch,
  launcher, ico, etc.)
- ``Fabriq.exe``, ``Deploy.bat``

## Excluded (kept on target system, not overwritten)

### Site-specific kitting data (kernel level)
- ``kernel/csv/hostlist.csv`` -- target PC list
- ``kernel/csv/workers.csv`` -- worker list
- ``kernel/csv/log_destinations.csv`` -- log upload destinations

### Site-specific kitting data (module level) -- strict rule
All CSVs under ``modules/`` EXCEPT ``module.csv`` and ``preset.csv``.
Examples:
- ``*_list.csv`` family (autologon_list, reg_hklm_list, app_list, ...)
- ``office_key.csv`` / ``license_key.csv`` -- product keys
- ``domain.csv`` -- domain join credentials
- ``builtin_admin.csv`` -- built-in admin credentials
- ``target_apps.csv``, ``driver.csv``, ``printer_delete.csv``, ``restore_dest.csv``
- Any future non-standard-named module CSV

### Site-customizable profile
- ``profiles/Custom Plan.csv``

### Runtime artifacts (auto-generated at runtime)
- ``kernel/json/art_pulse.txt``
- ``kernel/json/resume_state.json``
- ``kernel/json/session.json``
- ``kernel/json/skip_request.flag``
- ``kernel/json/status.json``
- ``kernel/txt/passphrase_verify.txt`` (Fabriq Studio-generated; site-specific)
- ``kernel/txt/silence.flag``

### Development-only
- ``.git/``, ``.claude/``, ``.gitignore``
- ``evidence/``, ``logs/`` (runtime output)

## Known caveats

1. **Master_*.csv may have been site-customized**: Framework master
   profiles are included in the patch. If the target has customized
   these, back up before overlaying or manually merge after.

2. **New modules from this patch**: If this patch introduces a wholly
   new module, the target will receive ``module.csv`` + ``preset.csv`` +
   the ``.ps1`` script(s) but NOT a starter ``_list.csv``. Operator must
   create one (copy from patch source if needed, or see the module's
   ``Guide.txt``).

3. **Module CSV schema changes**: If a module's ``.ps1`` update requires
   a new column in ``_list.csv`` that the target does not have, the
   module will fail at runtime. Schema changes are called out in
   ``CHANGELOG.md`` under ``Changed`` with migration notes.

## How to apply

1. Stop any running Fabriq instance on the target machine
2. Back up the target fabriq folder (recommended)
3. Copy the CONTENTS of this patch folder over the deployed fabriq folder
   (overwrite when prompted). Example:
   ``robocopy .\$PatchName E:\fabriq /E``
4. Verify with ``powershell.exe -File .\dev\check_version.ps1`` on the
   target (expect KERNEL_VERSION=$kernelVersion and all X.Y synced)
5. Launch ``Fabriq.exe`` and confirm normal startup

## Rollback

Restore the backup from step 2. Overlay is additive/overwrite; there is
no destructive removal of target-side files outside what this patch
contains.
"@

$manifest | Out-File -FilePath (Join-Path $dst "PATCH_README.md") -Encoding UTF8 -Force

# ========================================
# Final summary
# ========================================
Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host " Patch built: $dst" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green

$allFiles = @(Get-ChildItem $dst -Recurse -File)
$totalSize = ($allFiles | Measure-Object -Property Length -Sum).Sum
Write-Host ("Files              : {0}" -f $allFiles.Count) -ForegroundColor White
Write-Host ("Size               : {0:N1} MB" -f ($totalSize / 1MB)) -ForegroundColor White
Write-Host ("Kernel-level strip : {0}" -f $removedCount) -ForegroundColor White
Write-Host ("Module CSVs kept   : {0} (module.csv + preset.csv)" -f $moduleKept) -ForegroundColor White
Write-Host ("Module CSVs strip  : {0}" -f $moduleStripped) -ForegroundColor White
Write-Host ""
