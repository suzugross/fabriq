# ========================================
# ACL Hybrid Restore Script
# ========================================
# Restores directory ACLs using a hybrid approach:
# 1. Full restore from icacls /restore (bulk ACLs)
# 2. Individual overrides for non-inherited subdirectories (shallow-to-deep)
#
# [NOTES]
# - Administrator privileges required
# - Requires prior backup from ACL Backup module
# - Restore order: full tree first, then individual overrides (shallow to deep)
# ========================================

Write-Host ""
Show-Separator
Write-Host "ACL Hybrid Restore" -ForegroundColor Cyan
Show-Separator
Write-Host ""


# ========================================
# Local helper functions
# ========================================
function Get-PathHash {
    param([string]$RelativePath)
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($RelativePath.ToLower())
    $sha = [System.Security.Cryptography.SHA256]::Create()
    $hash = $sha.ComputeHash($bytes)
    return ($hash[0..7] | ForEach-Object { $_.ToString("x2") }) -join ""
}

function Get-SafeDirectoryName {
    param([string]$Path)
    $leaf = Split-Path $Path -Leaf
    return ($leaf -replace '[\\/:*?"<>|]', '_')
}


# ========================================
# Step 1: CSV reading
# ========================================
$csvPath = Join-Path $PSScriptRoot "acl_list.csv"

$enabledItems = Import-ModuleCsv -Path $csvPath -FilterEnabled `
    -RequiredColumns @("Enabled", "Id", "TargetPath")

if ($null -eq $enabledItems) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load acl_list.csv")
}
if ($enabledItems.Count -eq 0) {
    return (New-ModuleResult -Status "Skipped" -Message "No enabled entries")
}

# Expand environment variables in TargetPath
foreach ($item in $enabledItems) {
    $item.TargetPath = Expand-UserEnvironmentVariables $item.TargetPath
}


# ========================================
# Step 2: Pre-flight check
# ========================================
if (-not (Test-AdminPrivilege)) {
    Show-Error "Administrator privileges are required."
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Administrator privileges required")
}

$icaclsCmd = Get-Command "icacls.exe" -ErrorAction SilentlyContinue
if (-not $icaclsCmd) {
    Show-Error "icacls.exe not found on this system."
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "icacls.exe not found")
}

$backupBaseDir = Join-Path $PSScriptRoot "backup"
if (-not (Test-Path $backupBaseDir)) {
    Show-Error "Backup directory not found: $backupBaseDir"
    Show-Info "Run 'ACL Backup' first to create a backup."
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Backup directory not found. Run ACL Backup first.")
}


# ========================================
# Step 3: Pre-execution display (dry-run)
# ========================================
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "ACL Restore Targets" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

# Pre-load manifest data for display
$manifestData = @{}

foreach ($item in $enabledItems) {
    $displayName = if ($item.Description) { $item.Description } else { $item.TargetPath }
    $safeName = Get-SafeDirectoryName $item.TargetPath
    $backupDir = Join-Path $backupBaseDir "$($item.Id)_$safeName"
    $fullBackupFile = Join-Path $backupDir "_full_acl.txt"
    $manifestPath = Join-Path $backupDir "_manifest.csv"

    # Target path check
    $targetExists = Test-Path $item.TargetPath
    # Backup check
    $backupExists = Test-Path $fullBackupFile

    if (-not $targetExists) {
        Write-Host "  [TARGET NOT FOUND] $displayName" -ForegroundColor Red
        Write-Host "    Path: $($item.TargetPath)" -ForegroundColor DarkGray
        Write-Host ""
        continue
    }

    if (-not $backupExists) {
        Write-Host "  [NO BACKUP] $displayName" -ForegroundColor Red
        Write-Host "    Path: $($item.TargetPath)" -ForegroundColor DarkGray
        Write-Host "    Expected: $fullBackupFile" -ForegroundColor DarkGray
        Write-Host ""
        continue
    }

    # Load manifest
    $manifest = @()
    if (Test-Path $manifestPath) {
        $rawManifest = @(Import-Csv -Path $manifestPath -Encoding UTF8 -ErrorAction SilentlyContinue)
        # Filter out empty rows (header-only manifests)
        $manifest = @($rawManifest | Where-Object { -not [string]::IsNullOrWhiteSpace($_.RelativePath) })
    }
    $manifestData[$item.Id] = $manifest

    Write-Host "  [RESTORE] $displayName" -ForegroundColor Yellow
    Write-Host "    Path: $($item.TargetPath)" -ForegroundColor DarkGray
    Write-Host "    Backup: $backupDir" -ForegroundColor DarkGray
    if ($manifest.Count -gt 0) {
        Write-Host "    Restore: full tree + $($manifest.Count) individual overrides" -ForegroundColor DarkGray
    }
    else {
        Write-Host "    Restore: full tree only (no individual overrides)" -ForegroundColor DarkGray
    }
    Write-Host ""
}

Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

Show-Warning "This will OVERWRITE current ACL settings on the target directories."
Write-Host ""


# ========================================
# Step 4: Confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Restore ACL settings from backup?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""


# ========================================
# Step 5: Execution loop
# ========================================
$successCount = 0
$skipCount    = 0
$failCount    = 0

foreach ($item in $enabledItems) {
    $displayName = if ($item.Description) { $item.Description } else { $item.TargetPath }
    $safeName = Get-SafeDirectoryName $item.TargetPath
    $backupDir = Join-Path $backupBaseDir "$($item.Id)_$safeName"

    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "Restoring: $displayName" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor White

    # ----------------------------------------
    # Pre-check: target path and backup existence
    # ----------------------------------------
    if (-not (Test-Path $item.TargetPath)) {
        Show-Skip "Target path not found: $($item.TargetPath)"
        Write-Host ""
        $skipCount++
        continue
    }

    $fullBackupFile = Join-Path $backupDir "_full_acl.txt"
    if (-not (Test-Path $fullBackupFile)) {
        Show-Skip "Backup not found: $fullBackupFile"
        Write-Host ""
        $skipCount++
        continue
    }

    # ----------------------------------------
    # Main processing
    # ----------------------------------------
    try {
        # Phase A: Full tree restore
        # icacls /restore requires the parent directory of the original save target
        $parentPath = Split-Path $item.TargetPath -Parent
        Show-Info "Phase 1: Full tree restore..."
        Show-Info "  icacls `"$parentPath`" /restore `"$fullBackupFile`" /C"

        $proc = Start-Process -FilePath "icacls.exe" `
            -ArgumentList "`"$parentPath`" /restore `"$fullBackupFile`" /C" `
            -Wait -NoNewWindow -PassThru -ErrorAction Stop

        if ($proc.ExitCode -ne 0) {
            Show-Error "Full restore failed (ExitCode=$($proc.ExitCode))"
            $failCount++
            Write-Host ""
            continue
        }
        Show-Success "Full tree restore completed"

        # Phase B: Individual overrides (shallow-to-deep order)
        $manifest = $manifestData[$item.Id]

        if ($null -eq $manifest -or $manifest.Count -eq 0) {
            # Check if manifest file exists but was empty
            $manifestPath = Join-Path $backupDir "_manifest.csv"
            if (-not (Test-Path $manifestPath)) {
                Show-Warning "Manifest not found. Skipping individual overrides."
            }
            else {
                Show-Info "No individual overrides needed"
            }
        }
        else {
            # Sort by depth ascending (shallow first, deep last)
            $sortedManifest = @($manifest | Sort-Object { [int]$_.Depth })

            Show-Info "Phase 2: Restoring $($sortedManifest.Count) individual non-inherited overrides..."
            $current = 0
            $indivSuccess = 0
            $indivFail = 0

            foreach ($entry in $sortedManifest) {
                $current++
                $indivBackupFile = Join-Path $backupDir "individual\$($entry.BackupFile)"

                if (-not (Test-Path $indivBackupFile)) {
                    Show-Warning "  [$current/$($sortedManifest.Count)] Backup file not found: $($entry.BackupFile)"
                    $indivFail++
                    continue
                }

                # Restore parent = parent of the non-inherited subdirectory
                $subdirFullPath = Join-Path $item.TargetPath $entry.RelativePath
                $subdirParent = Split-Path $subdirFullPath -Parent

                Write-Host "  [$current/$($sortedManifest.Count)] $($entry.RelativePath)" -ForegroundColor DarkGray

                $indivProc = Start-Process -FilePath "icacls.exe" `
                    -ArgumentList "`"$subdirParent`" /restore `"$indivBackupFile`" /C" `
                    -Wait -NoNewWindow -PassThru -ErrorAction Stop

                if ($indivProc.ExitCode -ne 0) {
                    Show-Warning "  Individual restore failed (ExitCode=$($indivProc.ExitCode)): $($entry.RelativePath)"
                    $indivFail++
                }
                else {
                    $indivSuccess++
                }
            }

            Show-Info "Individual restores: $indivSuccess succeeded, $indivFail failed"
        }

        Show-Success "Completed: $displayName"
        $successCount++
    }
    catch {
        Show-Error "Failed: $displayName : $_"
        $failCount++
    }

    Write-Host ""
}


# ========================================
# Step 6: Result
# ========================================
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount `
    -Title "ACL Restore Results")
