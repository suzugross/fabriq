# ========================================
# BitLocker Await Script
# ========================================
# Wait for BitLocker volumes that are currently encrypting / decrypting
# to reach a steady state.
#
# [NOTES]
# - Requires administrator privileges
# - Watches only the drives listed in bitlocker_list.csv (TargetDrive column)
# - Idempotency: returns Skipped when no drives are encrypting / decrypting
# - EncryptionInProgress: waits for 100% -> FullyEncrypted
# - DecryptionInProgress: waits for 0% -> FullyDecrypted
# - Polling interval: 30 seconds
# ========================================

Write-Host ""
Show-Separator
Write-Host "BitLocker Await" -ForegroundColor Cyan
Show-Separator
Write-Host ""


# ========================================
# Step 1: Load CSV
# ========================================
$csvPath = Join-Path $PSScriptRoot "bitlocker_list.csv"

$enabledItems = Import-ModuleCsv -Path $csvPath -FilterEnabled `
    -RequiredColumns @("Enabled", "TargetDrive")

if ($null -eq $enabledItems) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load bitlocker_list.csv")
}
if ($enabledItems.Count -eq 0) {
    return (New-ModuleResult -Status "Skipped" -Message "No enabled entries")
}


# ========================================
# Step 3: Dry-run summary before execution
# ========================================
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "BitLocker Encryption Status" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

$hasAwaitTarget = $false

foreach ($item in $enabledItems) {
    $driveLetter = $item.TargetDrive
    $displayName = if ($item.Description) { "$driveLetter $($item.Description)" } else { $driveLetter }

    # Verify the drive exists
    if (-not (Test-Path "${driveLetter}\")) {
        Write-Host "  [NOT FOUND] $displayName" -ForegroundColor DarkGray
        Write-Host "    Drive does not exist" -ForegroundColor DarkGray
        Write-Host ""
        continue
    }

    # Read BitLocker status
    $blVolume = Get-BitLockerVolume -MountPoint $driveLetter -ErrorAction SilentlyContinue
    if ($null -eq $blVolume) {
        Write-Host "  [NOT FOUND] $displayName" -ForegroundColor DarkGray
        Write-Host "    Unable to get BitLocker status" -ForegroundColor DarkGray
        Write-Host ""
        continue
    }

    $volumeStatus     = $blVolume.VolumeStatus
    $protectionStatus = $blVolume.ProtectionStatus
    $encryptPercent   = $blVolume.EncryptionPercentage

    # Show a status-specific marker
    if ($volumeStatus -eq "FullyEncrypted") {
        Write-Host "  [COMPLETE] $displayName" -ForegroundColor DarkGray
        Write-Host "    Status: $protectionStatus ($volumeStatus)" -ForegroundColor DarkGray
    }
    elseif ($volumeStatus -eq "EncryptionInProgress") {
        Write-Host "  [ENCRYPTING] $displayName" -ForegroundColor Yellow
        Write-Host "    Status: $protectionStatus ($volumeStatus) - ${encryptPercent}% encrypted" -ForegroundColor Yellow
        $hasAwaitTarget = $true
    }
    elseif ($volumeStatus -eq "DecryptionInProgress") {
        Write-Host "  [DECRYPTING] $displayName" -ForegroundColor Yellow
        Write-Host "    Status: $protectionStatus ($volumeStatus) - ${encryptPercent}% remaining" -ForegroundColor Yellow
        $hasAwaitTarget = $true
    }
    elseif ($volumeStatus -eq "FullyDecrypted") {
        Write-Host "  [DECRYPTED] $displayName" -ForegroundColor DarkGray
        Write-Host "    Status: $protectionStatus ($volumeStatus)" -ForegroundColor DarkGray
    }
    else {
        Write-Host "  [OTHER] $displayName" -ForegroundColor DarkGray
        Write-Host "    Status: $protectionStatus ($volumeStatus)" -ForegroundColor DarkGray
    }
    Write-Host ""
}

Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

if (-not $hasAwaitTarget) {
    Show-Skip "No drives are currently encrypting or decrypting"
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "No drives are currently encrypting or decrypting")
}


# ========================================
# Step 4: User confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Wait for encryption/decryption to complete?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""


# ========================================
# Step 5: Wait loop
# ========================================
# Stale timeout: progress is tracked per drive. If a drive shows
# zero movement for 30 minutes it is treated as stalled and dropped
# from the watch list. Any progress resets the stale counter.
# ========================================
$pollIntervalSec  = 30
$staleTimeoutSec  = 1800   # 30 minutes

# Per-drive progress tracking table
$driveTracker = @{}
foreach ($item in $enabledItems) {
    $dl = $item.TargetDrive
    $blVol = Get-BitLockerVolume -MountPoint $dl -ErrorAction SilentlyContinue
    if ($blVol -and ($blVol.VolumeStatus -eq "EncryptionInProgress" -or $blVol.VolumeStatus -eq "DecryptionInProgress")) {
        $driveTracker[$dl] = @{
            LastPercent  = $blVol.EncryptionPercentage
            StaleElapsed = 0
            TimedOut     = $false
            Mode         = $blVol.VolumeStatus  # EncryptionInProgress or DecryptionInProgress
        }
    }
}

Show-Info "Monitoring progress (polling every ${pollIntervalSec}s, stale timeout ${staleTimeoutSec}s)..."
Write-Host ""

while ($true) {
    $pendingDrives = @()

    foreach ($item in $enabledItems) {
        $driveLetter = $item.TargetDrive

        # Skip drives that already timed out
        if ($driveTracker.ContainsKey($driveLetter) -and $driveTracker[$driveLetter].TimedOut) { continue }

        if (-not (Test-Path "${driveLetter}\")) { continue }

        $blVolume = Get-BitLockerVolume -MountPoint $driveLetter -ErrorAction SilentlyContinue
        if ($null -eq $blVolume) { continue }

        $isEncrypting  = $blVolume.VolumeStatus -eq "EncryptionInProgress"
        $isDecrypting  = $blVolume.VolumeStatus -eq "DecryptionInProgress"

        if ($isEncrypting -or $isDecrypting) {
            $currentPercent = $blVolume.EncryptionPercentage

            # Progress check
            if ($driveTracker.ContainsKey($driveLetter)) {
                $tracker = $driveTracker[$driveLetter]

                # Encrypting: a higher percent means progress; Decrypting: a lower percent means progress
                $hasProgress = if ($isEncrypting) { $currentPercent -gt $tracker.LastPercent } else { $currentPercent -lt $tracker.LastPercent }

                if ($hasProgress) {
                    # Progress observed -> reset the stale counter
                    $tracker.LastPercent  = $currentPercent
                    $tracker.StaleElapsed = 0
                }
                else {
                    # No progress -> accumulate the stale counter
                    $tracker.StaleElapsed += $pollIntervalSec
                    if ($tracker.StaleElapsed -ge $staleTimeoutSec) {
                        $tracker.TimedOut = $true
                        $displayName = if ($item.Description) { "$driveLetter $($item.Description)" } else { $driveLetter }
                        Show-Error "Stale timeout: $displayName (no progress for $($staleTimeoutSec / 60) min at ${currentPercent}%)"
                        continue
                    }
                }
            }

            $modeLabel = if ($isEncrypting) { "Encrypting" } else { "Decrypting" }
            $pendingDrives += @{
                Drive   = $driveLetter
                Percent = $currentPercent
                Label   = $modeLabel
            }
        }
    }

    if ($pendingDrives.Count -eq 0) {
        break
    }

    # Progress line
    $timestamp = Get-Date -Format "HH:mm:ss"
    $progressParts = @()
    foreach ($pd in $pendingDrives) {
        $progressParts += "$($pd.Drive) $($pd.Label) $($pd.Percent)%"
    }
    $progressText = $progressParts -join " / "
    Write-Host "  [$timestamp] $progressText" -ForegroundColor DarkGray

    Start-Sleep -Seconds $pollIntervalSec
}

Write-Host ""


# ========================================
# Step 6: Aggregate and return result
# ========================================
# Final check: confirm each drive reached a terminal state
$successCount = 0
$skipCount    = 0
$failCount    = 0

foreach ($item in $enabledItems) {
    $driveLetter = $item.TargetDrive
    $displayName = if ($item.Description) { "$driveLetter $($item.Description)" } else { $driveLetter }

    if (-not (Test-Path "${driveLetter}\")) {
        $skipCount++
        continue
    }

    $blVolume = Get-BitLockerVolume -MountPoint $driveLetter -ErrorAction SilentlyContinue
    if ($null -eq $blVolume) {
        $skipCount++
        continue
    }

    if ($blVolume.VolumeStatus -eq "FullyEncrypted") {
        Show-Success "Encryption complete: $displayName"
        $successCount++
    }
    elseif ($blVolume.VolumeStatus -eq "FullyDecrypted") {
        # Success if this drive was actively waited on; otherwise Skip
        if ($driveTracker.ContainsKey($driveLetter)) {
            Show-Success "Decryption complete: $displayName"
            $successCount++
        }
        else {
            $skipCount++
        }
    }
    else {
        Show-Error "Unexpected status on ${displayName}: $($blVolume.VolumeStatus)"
        $failCount++
    }
}

Write-Host ""
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount `
    -Title "BitLocker Await Results")
