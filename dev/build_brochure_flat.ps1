# ========================================
# Build Fabriq Brochure (Flat Layout)
# ========================================
# Copies brochure source materials into a single flat directory,
# prefixing each filename with its category to avoid collisions
# and preserve provenance.
#
# Source : E:\tmp\fabriq_brochure_materials\99_old\
#            apps/  contracts/  kernel/  modules/  profiles/
# Target : <Desktop>\fabriq_brochure_flat\        (default)
#
# Output naming: <category>__<original_name>.md
#   e.g.  kernel/01_overview.md  ->  kernel__01_overview.md
#
# Usage:
#   pwsh .\dev\build_brochure_flat.ps1
#   pwsh .\dev\build_brochure_flat.ps1 -DryRun
#   pwsh .\dev\build_brochure_flat.ps1 -SourceRoot "E:\other\source"
#
# Safety policy:
#   Target MUST be inside the user's Desktop AND its leaf name MUST
#   contain "brochure". Anything else is rejected. Cleanup of an
#   existing target only removes files matching the layout pattern
#   (`<category>__<name>.md` and `INDEX.md`); no recursive removal.
# ========================================

[CmdletBinding()]
param(
    [string]$SourceRoot = 'E:\tmp\fabriq_brochure_materials\99_old',
    [string]$TargetDir,
    [switch]$DryRun
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ----------------------------------------
# Resolve target (default: <Desktop>\fabriq_brochure_flat)
# ----------------------------------------
$desktop = [Environment]::GetFolderPath('Desktop')
if ([string]::IsNullOrWhiteSpace($desktop)) {
    throw "Cannot resolve Desktop folder via [Environment]::GetFolderPath('Desktop')."
}
$desktopFull = [System.IO.Path]::GetFullPath($desktop).TrimEnd('\')

if ([string]::IsNullOrWhiteSpace($TargetDir)) {
    $TargetDir = Join-Path $desktopFull 'fabriq_brochure_flat'
}
$targetFull = [System.IO.Path]::GetFullPath($TargetDir).TrimEnd('\')

# ----------------------------------------
# SAFETY GUARDS (multi-layer; called again before any deletion)
# ----------------------------------------
function Assert-SafeTarget {
    param(
        [Parameter(Mandatory)] [string]$Path,
        [Parameter(Mandatory)] [string]$DesktopRoot
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        throw "[SAFETY] Target path is empty."
    }
    if ($Path -match '\.\.[\\/]') {
        throw "[SAFETY] Target contains path traversal: $Path"
    }
    if ($Path -ieq $DesktopRoot) {
        throw "[SAFETY] Refusing to operate on Desktop root itself: $Path"
    }
    if (-not $Path.StartsWith($DesktopRoot + '\', [StringComparison]::OrdinalIgnoreCase)) {
        throw "[SAFETY] Target must be inside Desktop. Got: $Path (Desktop: $DesktopRoot)"
    }
    $relative = $Path.Substring($DesktopRoot.Length).TrimStart('\')
    if ([string]::IsNullOrWhiteSpace($relative)) {
        throw "[SAFETY] Target resolves to Desktop root after trim: $Path"
    }
    $leaf = Split-Path -Leaf $Path
    if ($leaf -notmatch 'brochure') {
        throw "[SAFETY] Target leaf must contain 'brochure'. Got: $leaf"
    }
}

Assert-SafeTarget -Path $targetFull -DesktopRoot $desktopFull

# ----------------------------------------
# Validate source
# ----------------------------------------
if (-not (Test-Path -LiteralPath $SourceRoot -PathType Container)) {
    throw "Source directory not found: $SourceRoot"
}
$sourceFull = [System.IO.Path]::GetFullPath($SourceRoot).TrimEnd('\')

Write-Host "========================================"
Write-Host " Fabriq Brochure Flat Builder"
Write-Host "========================================"
Write-Host "Source : $sourceFull"
Write-Host "Target : $targetFull"
Write-Host "DryRun : $DryRun"
Write-Host ""

# ----------------------------------------
# Prepare target directory
# ----------------------------------------
if (Test-Path -LiteralPath $targetFull) {
    Write-Host "[INFO] Target exists. Cleaning known files only (pattern: *__*.md and INDEX.md)."

    # Defense in depth: re-validate before any deletion
    Assert-SafeTarget -Path $targetFull -DesktopRoot $desktopFull

    $existing = Get-ChildItem -LiteralPath $targetFull -File -ErrorAction Stop |
        Where-Object {
            $_.Name -match '^[a-z0-9_]+__.+\.md$' -or $_.Name -ieq 'INDEX.md'
        }

    foreach ($f in $existing) {
        # Belt-and-suspenders: confirm each file is still inside target
        $ff = [System.IO.Path]::GetFullPath($f.FullName)
        if (-not $ff.StartsWith($targetFull + '\', [StringComparison]::OrdinalIgnoreCase)) {
            throw "[SAFETY] File outside target detected: $ff"
        }
        if ($DryRun) {
            Write-Host "  [DRY] would remove: $($f.Name)"
        } else {
            Remove-Item -LiteralPath $ff -Force -ErrorAction Stop
        }
    }
    Write-Host "[INFO] Cleaned $($existing.Count) old files."
} else {
    if ($DryRun) {
        Write-Host "[DRY] would create directory: $targetFull"
    } else {
        New-Item -ItemType Directory -Path $targetFull -Force | Out-Null
        Write-Host "[INFO] Created target directory: $targetFull"
    }
}

# ----------------------------------------
# Enumerate source markdown files
# ----------------------------------------
$mdFiles = Get-ChildItem -LiteralPath $sourceFull -Recurse -File -Filter '*.md' |
    Sort-Object FullName

if ($mdFiles.Count -eq 0) {
    throw "No .md files found under: $sourceFull"
}

# ----------------------------------------
# Copy with category prefix; collect manifest
# ----------------------------------------
$copied     = 0
$skipped    = 0
$collisions = New-Object System.Collections.Generic.List[string]
$manifest   = New-Object System.Collections.Generic.List[psobject]

foreach ($f in $mdFiles) {
    $relative = $f.FullName.Substring($sourceFull.Length).TrimStart('\')
    $parts    = $relative -split '[\\/]'

    if ($parts.Count -lt 2) {
        Write-Warning "[SKIP] File at source root (no category): $relative"
        $skipped++
        continue
    }

    $category = $parts[0].ToLowerInvariant()
    if ($category -notmatch '^[a-z0-9_]+$') {
        throw "Unexpected category name (not [a-z0-9_]+): $category"
    }

    $newName = "${category}__$($f.Name)"
    $dest    = Join-Path $targetFull $newName

    if (Test-Path -LiteralPath $dest) {
        $collisions.Add($newName)
    }

    if ($DryRun) {
        Write-Host "  [DRY] would copy: $relative -> $newName"
    } else {
        Copy-Item -LiteralPath $f.FullName -Destination $dest -Force
    }

    $manifest.Add([pscustomobject]@{
        Category = $category
        Source   = $relative
        Flat     = $newName
        Bytes    = $f.Length
    })
    $copied++
}

if ($collisions.Count -gt 0) {
    Write-Warning "Filename collisions overwritten: $($collisions -join ', ')"
}

# ----------------------------------------
# Build INDEX.md (auto-generated, English only)
# ----------------------------------------
$indexLines = New-Object System.Collections.Generic.List[string]
$null = $indexLines.Add("# Fabriq Brochure Materials (Flat Layout)")
$null = $indexLines.Add("")
$null = $indexLines.Add("Auto-generated by ``dev/build_brochure_flat.ps1`` on $(Get-Date -Format 'yyyy-MM-dd HH:mm').")
$null = $indexLines.Add("")
$null = $indexLines.Add("Source : ``$sourceFull``  ")
$null = $indexLines.Add("Files  : $copied")
$null = $indexLines.Add("")
$null = $indexLines.Add("Each filename encodes its original category as ``<category>__<original_name>.md``.")
$null = $indexLines.Add("Upload this entire folder to a chat-based Claude/AI to use as brochure source material.")
$null = $indexLines.Add("")
$null = $indexLines.Add("---")
$null = $indexLines.Add("")

$grouped = $manifest | Group-Object Category | Sort-Object Name
foreach ($g in $grouped) {
    $null = $indexLines.Add("## $($g.Name)  ($($g.Count) files)")
    $null = $indexLines.Add("")
    foreach ($item in ($g.Group | Sort-Object Flat)) {
        $null = $indexLines.Add("- ``$($item.Flat)``  <-  $($item.Source)")
    }
    $null = $indexLines.Add("")
}

$indexPath = Join-Path $targetFull 'INDEX.md'
if ($DryRun) {
    Write-Host "[DRY] would write INDEX.md ($($indexLines.Count) lines)"
} else {
    # Plain UTF-8 (no BOM) - INDEX.md is English only
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllLines($indexPath, $indexLines, $utf8NoBom)
}

# ----------------------------------------
# Summary
# ----------------------------------------
Write-Host ""
Write-Host "========================================"
Write-Host "[OK] Done."
Write-Host "  Copied   : $copied"
Write-Host "  Skipped  : $skipped"
Write-Host "  Output   : $targetFull"
if (-not $DryRun) {
    Write-Host "  INDEX    : $indexPath"
}
Write-Host "========================================"
