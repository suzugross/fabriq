# ========================================
# Userdata Restore
# ========================================
# [PURPOSE]
# Replay a userdata_backup snapshot back onto this PC. Reads
# manifest.json (schemaVersion=1, fabriq-userdata-backup) and
# robocopies each entry's data/ tree back to its resolvedPath.
#
# [NOTES]
# - Requires administrator privileges.
# - Hostlist-driven: source PC name resolved from
#   $env:SELECTED_OLD_PCNAME by default. CSV columns
#   SourcePcName / BackupTimestamp in userdata_backup_config.csv
#   are explicit overrides.
# - osArch / osVersion are NOT enforced (file contents are not
#   architecture-specific; we just record for diagnostics).
# - Per-entry OnConflict honored: skip / overwrite / rename.
# - Same engine (robocopy) as backup. /B backup mode + /COPYALL
#   when entry was captured with IncludeAcl=1.
# ========================================

Write-Host ""
Show-Separator
Write-Host "Userdata Restore" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# ========================================
# Step 1: Load Config CSV (restore-side overrides)
# ========================================
$csvPath = Join-Path $PSScriptRoot "userdata_backup_config.csv"

$cfgItems = Import-ModuleCsv -Path $csvPath -FilterEnabled `
    -RequiredColumns @("Enabled", "SourcePcName", "BackupTimestamp")

if ($null -eq $cfgItems) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load userdata_backup_config.csv")
}
if (@($cfgItems).Count -eq 0) {
    return (New-ModuleResult -Status "Skipped" -Message "Restore disabled in config CSV")
}

$cfg = @($cfgItems)[0]
$cfgSourcePcName    = if ([string]::IsNullOrWhiteSpace($cfg.SourcePcName))    { $null } else { $cfg.SourcePcName.Trim() }
$cfgBackupTimestamp = if ([string]::IsNullOrWhiteSpace($cfg.BackupTimestamp)) { $null } else { $cfg.BackupTimestamp.Trim() }


# ========================================
# Step 2: Prerequisite Checks
# ========================================
if (-not (Test-AdminPrivilege)) {
    Show-Error "Administrator privileges are required"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Administrator privileges required")
}

$robocopyExe = Get-Command robocopy.exe -ErrorAction SilentlyContinue
if ($null -eq $robocopyExe) {
    Show-Error "robocopy.exe not found in PATH"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "robocopy.exe not found")
}


# ========================================
# Step 3a: Locate Source Backup
# ========================================
# Resolution priority:
#   PcName    : 1. CSV SourcePcName  2. $env:SELECTED_OLD_PCNAME  3. error
#   Timestamp : 1. CSV BackupTimestamp  2. latest manifest.collectedAt
$backupRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot "backup"))

if (-not (Test-Path $backupRoot)) {
    Show-Error "Backup root not found: $backupRoot"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Backup root not found")
}

Show-Info "Backup root: $backupRoot"

$pcNameSource = $null
$resolvedPcName = $null
if (-not [string]::IsNullOrWhiteSpace($cfgSourcePcName)) {
    $resolvedPcName = $cfgSourcePcName
    $pcNameSource   = "CSV SourcePcName"
}
elseif (-not [string]::IsNullOrWhiteSpace($env:SELECTED_OLD_PCNAME)) {
    $resolvedPcName = $env:SELECTED_OLD_PCNAME.Trim()
    $pcNameSource   = "hostlist OldPCname"
}
else {
    Show-Error "No source PC name available."
    Show-Error "  Either select a host (so SELECTED_OLD_PCNAME is set) or"
    Show-Error "  set SourcePcName in userdata_backup_config.csv."
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "No source PC name (host not selected, CSV empty)")
}

Show-Info "Source PC name : $resolvedPcName (from $pcNameSource)"

$pcDir = Get-ChildItem -Path $backupRoot -Directory -ErrorAction SilentlyContinue |
         Where-Object { $_.Name -ieq $resolvedPcName } |
         Select-Object -First 1

if ($null -eq $pcDir) {
    Show-Error "No backup subfolder for PC '$resolvedPcName' under: $backupRoot"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "No backup folder for source PC '$resolvedPcName'")
}

$candidates = @()
foreach ($tsDir in @(Get-ChildItem -Path $pcDir.FullName -Directory -ErrorAction SilentlyContinue)) {
    if (-not [string]::IsNullOrWhiteSpace($cfgBackupTimestamp) -and ($tsDir.Name -ine $cfgBackupTimestamp)) { continue }
    $mfPath = Join-Path $tsDir.FullName "manifest.json"
    if (-not (Test-Path $mfPath)) { continue }
    $candidates += [PSCustomObject]@{
        Timestamp = $tsDir.Name
        Path      = $tsDir.FullName
        Manifest  = $mfPath
    }
}

if ($candidates.Count -eq 0) {
    Show-Error "No backup folders with manifest.json under: $($pcDir.FullName)"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "No backups under PC '$resolvedPcName'")
}

$timestampSource = if (-not [string]::IsNullOrWhiteSpace($cfgBackupTimestamp)) { "CSV BackupTimestamp" } else { "latest manifest.collectedAt" }
$chosen = $null
$chosenAt = [DateTime]::MinValue
foreach ($c in $candidates) {
    try {
        $m = Get-Content -Path $c.Manifest -Raw | ConvertFrom-Json
        $at = [DateTime]::MinValue
        if ($m.collectedAt) { $null = [DateTime]::TryParse($m.collectedAt, [ref]$at) }
        if ($at -eq [DateTime]::MinValue) { $at = (Get-Item $c.Path).LastWriteTime }
        if ($at -gt $chosenAt) { $chosen = $c; $chosenAt = $at }
    }
    catch { }
}

if ($null -eq $chosen) {
    Show-Error "No readable manifest.json among $($candidates.Count) candidate(s)"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "No readable manifest.json")
}

$backupDir = $chosen.Path
Show-Success "Selected backup: $resolvedPcName / $($chosen.Timestamp)"
Show-Info "  PcName from   : $pcNameSource"
Show-Info "  Timestamp from: $timestampSource"
Show-Info "  Path: $backupDir"
Write-Host ""


# ========================================
# Step 3b: Load & Validate Manifest
# ========================================
$manifest = $null
try {
    $manifest = Get-Content -Path $chosen.Manifest -Raw | ConvertFrom-Json
}
catch {
    Show-Error "Failed to parse manifest.json: $_"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "manifest.json parse error")
}

if ($manifest.manifestType -ne "fabriq-userdata-backup") {
    Show-Error "Unexpected manifestType: '$($manifest.manifestType)' (expected 'fabriq-userdata-backup')"
    return (New-ModuleResult -Status "Error" -Message "Wrong manifestType")
}
if ([int]$manifest.schemaVersion -ne 1) {
    Show-Error "Unsupported schemaVersion: $($manifest.schemaVersion) (this module handles schemaVersion=1)"
    return (New-ModuleResult -Status "Error" -Message "Unsupported schemaVersion")
}


# ========================================
# Step 3c: Build Restore Plan
# ========================================
$allEntries = @($manifest.items.entries)
$plannedEntries = @($allEntries | Where-Object { $_.status -ne 'Skipped' -and -not [string]::IsNullOrWhiteSpace($_.backupSubpath) })

Write-Host "========================================" -ForegroundColor Yellow
Write-Host "Restore Plan" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "  Source PC          : $($manifest.computerName) ($($manifest.collectedAt))" -ForegroundColor White
Write-Host "  Target PC          : $env:COMPUTERNAME" -ForegroundColor White
Write-Host "  Backup osArch      : $($manifest.osArch)" -ForegroundColor DarkGray
Write-Host "  Backup osVersion   : $($manifest.osVersion)" -ForegroundColor DarkGray
Write-Host "  Entries to restore : $($plannedEntries.Count) (of $($allEntries.Count) in manifest)" -ForegroundColor White
Write-Host "  Resolved via       : PcName=$pcNameSource, Timestamp=$timestampSource" -ForegroundColor DarkGray
Write-Host ""
foreach ($pe in $plannedEntries) {
    $kind = if ($pe.isDirectory) { "[DIR ]" } else { "[FILE]" }
    $aclTag = if ($pe.includeAcl) { " /ACL" } else { "" }
    Write-Host ("  [{0}] {1} {2} -> {3}{4}" -f $pe.id, $kind, $pe.backupSubpath, $pe.resolvedPath, $aclTag) -ForegroundColor White
    Write-Host ("         onConflict={0} files={1} bytes={2:N0}" -f $pe.onConflict, $pe.fileCount, $pe.byteCount) -ForegroundColor DarkGray
}
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

if ($plannedEntries.Count -eq 0) {
    Show-Info "No restorable entries in manifest"
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "No restorable entries")
}


# ========================================
# Step 4: Confirm
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Proceed with userdata restore?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""


# ========================================
# Step 5: Execute Restore
# ========================================
$warnings = @()
$successCount = 0
$skipCount    = 0
$failCount    = 0

foreach ($pe in $plannedEntries) {
    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "[$($pe.id)] $($pe.resolvedPath)" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor White

    $srcDataDir = Join-Path $backupDir $pe.backupSubpath
    if (-not (Test-Path $srcDataDir)) {
        Show-Error "  Backup data folder missing: $srcDataDir"
        $warnings += "Missing data folder for entry $($pe.id)"
        $failCount++
        Write-Host ""
        continue
    }

    $targetPath = $pe.resolvedPath
    $isDir      = [bool]$pe.isDirectory
    $onConflict = if ([string]::IsNullOrWhiteSpace($pe.onConflict)) { "skip" } else { $pe.onConflict.ToLower() }
    $includeAcl = [bool]$pe.includeAcl

    # Conflict handling
    $proceed = $true
    if ($isDir) {
        if (Test-Path -LiteralPath $targetPath -PathType Container) {
            switch ($onConflict) {
                "skip" {
                    Show-Skip "  Target directory exists (OnConflict=skip): $targetPath"
                    $skipCount++
                    $proceed = $false
                }
                "overwrite" {
                    Show-Info "  Target directory exists, will overwrite (OnConflict=overwrite)"
                }
                "rename" {
                    $renamed = "$targetPath.bak_$(Get-Date -Format yyyyMMdd_HHmmss)"
                    try {
                        Rename-Item -LiteralPath $targetPath -NewName (Split-Path $renamed -Leaf) -ErrorAction Stop
                        Show-Info "  Renamed existing target -> $(Split-Path $renamed -Leaf)"
                    }
                    catch {
                        Show-Error "  Rename failed: $($_.Exception.Message)"
                        $warnings += "Rename failed for $targetPath"
                        $failCount++
                        $proceed = $false
                    }
                }
                default {
                    Show-Error "  Unknown OnConflict value '$onConflict' - skipping"
                    $warnings += "Unknown OnConflict for entry $($pe.id)"
                    $failCount++
                    $proceed = $false
                }
            }
        }
        elseif (Test-Path -LiteralPath $targetPath -PathType Leaf) {
            Show-Error "  Target path is a file but backup expected a directory: $targetPath"
            $warnings += "Type mismatch (file vs dir) for entry $($pe.id)"
            $failCount++
            $proceed = $false
        }
    }
    else {
        if (Test-Path -LiteralPath $targetPath -PathType Leaf) {
            switch ($onConflict) {
                "skip" {
                    Show-Skip "  Target file exists (OnConflict=skip): $targetPath"
                    $skipCount++
                    $proceed = $false
                }
                "overwrite" {
                    Show-Info "  Target file exists, will overwrite (OnConflict=overwrite)"
                }
                "rename" {
                    $renamed = "$targetPath.bak_$(Get-Date -Format yyyyMMdd_HHmmss)"
                    try {
                        Rename-Item -LiteralPath $targetPath -NewName (Split-Path $renamed -Leaf) -ErrorAction Stop
                        Show-Info "  Renamed existing target -> $(Split-Path $renamed -Leaf)"
                    }
                    catch {
                        Show-Error "  Rename failed: $($_.Exception.Message)"
                        $warnings += "Rename failed for $targetPath"
                        $failCount++
                        $proceed = $false
                    }
                }
                default {
                    Show-Error "  Unknown OnConflict value '$onConflict' - skipping"
                    $warnings += "Unknown OnConflict for entry $($pe.id)"
                    $failCount++
                    $proceed = $false
                }
            }
        }
    }

    if (-not $proceed) { Write-Host ""; continue }

    # Ensure parent dir exists for files / target dir exists for dirs
    if ($isDir) {
        if (-not (Test-Path -LiteralPath $targetPath)) {
            try {
                $null = New-Item -ItemType Directory -Path $targetPath -Force -ErrorAction Stop
            }
            catch {
                Show-Error "  Failed to create target dir: $($_.Exception.Message)"
                $warnings += "Target dir creation failed: $targetPath"
                $failCount++
                Write-Host ""
                continue
            }
        }
    }
    else {
        $parent = Split-Path -Path $targetPath -Parent
        if (-not (Test-Path -LiteralPath $parent)) {
            try {
                $null = New-Item -ItemType Directory -Path $parent -Force -ErrorAction Stop
            }
            catch {
                Show-Error "  Failed to create target parent dir: $($_.Exception.Message)"
                $warnings += "Target parent creation failed: $parent"
                $failCount++
                Write-Host ""
                continue
            }
        }
    }

    $copyFlag = if ($includeAcl) { '/COPYALL' } else { '/COPY:DAT' }

    if ($isDir) {
        # Mirror backup data tree into target. Use /E (incl. empty subdirs).
        $rcArgs = @($srcDataDir, $targetPath, '/E', $copyFlag, '/B', '/R:1', '/W:1', '/NP')
    }
    else {
        $fileName = Split-Path -Path $targetPath -Leaf
        $targetDir = Split-Path -Path $targetPath -Parent
        $rcArgs = @($srcDataDir, $targetDir, $fileName, $copyFlag, '/B', '/R:1', '/W:1', '/NP', '/NDL', '/NS', '/NC')
    }

    Show-Info "  robocopy $($rcArgs -join ' ')"
    & robocopy.exe @rcArgs 2>&1 | Out-Null
    $rcExit = $LASTEXITCODE

    if ($rcExit -ge 8) {
        Show-Error "  robocopy failures (exit $rcExit)"
        $warnings += "robocopy exit $rcExit for entry $($pe.id)"
        $failCount++
    }
    elseif ($rcExit -ge 4) {
        Show-Warning "  robocopy mismatches (exit $rcExit)"
        $warnings += "robocopy mismatch (exit $rcExit) for entry $($pe.id)"
        $successCount++
    }
    else {
        Show-Success "  restored (robocopy exit $rcExit)"
        $successCount++
    }

    Write-Host ""
}


# ========================================
# Step 5.5: Post-Apply Verification
# ========================================
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Post-Apply Verification" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

$verifyFail = 0
foreach ($pe in $plannedEntries) {
    if (-not (Test-Path -LiteralPath $pe.resolvedPath)) {
        Write-Host "  [VERIFY FAILED] $($pe.resolvedPath) - not present after restore" -ForegroundColor Red
        $verifyFail++
        continue
    }
    Write-Host "  [VERIFIED] $($pe.resolvedPath)" -ForegroundColor Green
}

Write-Host ""
$verified = ($verifyFail -eq 0 -and $failCount -eq 0)


# ========================================
# Step 6: Result Summary
# ========================================
Show-Separator
Write-Host "Userdata Restore Results" -ForegroundColor Cyan
Show-Separator
Write-Host "  Source       : $($manifest.computerName) / $($chosen.Timestamp)" -ForegroundColor White
Write-Host "  Entries      : $successCount success / $skipCount skip / $failCount fail" -ForegroundColor White
if ($warnings.Count -gt 0) {
    Write-Host "  Warnings     : $($warnings.Count)" -ForegroundColor Yellow
}
if ($verified) {
    Write-Host "  Verified     : PASS" -ForegroundColor Green
} else {
    Write-Host "  Verified     : FAIL" -ForegroundColor Red
}
Show-Separator
Write-Host ""

return (New-BatchResult `
    -Success $successCount `
    -Skip $skipCount `
    -Fail $failCount `
    -Title "Userdata Restore Results" `
    -Verified $verified)
