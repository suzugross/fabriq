# ========================================
# BitLocker Configuration Script
# ========================================

# ========================================
# Helper: Build Evidence Content
# ========================================
function New-BitLockerEvidence {
    param(
        [object]$Volume,
        [object]$RecoveryKey,
        [string]$PCName
    )

    # Key Protectors summary
    $protectorLines = @()
    foreach ($kp in $Volume.KeyProtector) {
        $line = "  Type: $($kp.KeyProtectorType)  ID: $($kp.KeyProtectorId)"
        if ($kp.KeyProtectorType -eq "RecoveryPassword") {
            $line += "  Password: $($kp.RecoveryPassword)"
        }
        $protectorLines += $line
    }
    $protectorSection = $protectorLines -join "`r`n"

    $content = @(
        "========================================"
        "BitLocker Evidence Report"
        "========================================"
        ""
        "[General]"
        "Date:                    $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        "Computer:                $PCName"
        ""
        "[Volume]"
        "Mount Point:             $($Volume.MountPoint)"
        "Volume Type:             $($Volume.VolumeType)"
        "Volume Status:           $($Volume.VolumeStatus)"
        "Encryption Percentage:   $($Volume.EncryptionPercentage)%"
        "Encryption Method:       $($Volume.EncryptionMethod)"
        "Protection Status:       $($Volume.ProtectionStatus)"
        "Lock Status:             $($Volume.LockStatus)"
        "Auto Unlock Enabled:     $($Volume.AutoUnlockEnabled)"
        "Auto Unlock Key Stored:  $($Volume.AutoUnlockKeyStored)"
        ""
        "[Recovery Key]"
        "Identifier:              $($RecoveryKey.KeyProtectorId)"
        "Recovery Password:       $($RecoveryKey.RecoveryPassword)"
        ""
        "[All Key Protectors]"
        $protectorSection
        ""
        "========================================"
    ) -join "`r`n"

    return $content
}

Show-Info "Executing BitLocker configuration..."
Write-Host ""

# ========================================
# Pre-check: TPM Status
# ========================================
Show-Info "Checking TPM status..."

try {
    $tpm = Get-Tpm -ErrorAction Stop
}
catch {
    Show-Error "Failed to get TPM status: $_"
    return (New-ModuleResult -Status "Error" -Message "Failed to get TPM status")
}

if (-not $tpm.TpmPresent) {
    Show-Error "TPM is not present on this system"
    return (New-ModuleResult -Status "Error" -Message "TPM not present")
}

if (-not $tpm.TpmReady) {
    Show-Warning "TPM is present but not ready"
}
else {
    Show-Success "TPM is present and ready"
}

Write-Host ""

# ========================================
# Load CSV
# ========================================
$csvPath = Join-Path $PSScriptRoot "bitlocker_list.csv"

$driveList = Import-ModuleCsv -Path $csvPath -FilterEnabled
if ($null -eq $driveList) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load bitlocker_list.csv")
}
if ($driveList.Count -eq 0) {
    return (New-ModuleResult -Status "Skipped" -Message "No enabled entries")
}

Write-Host ""

# ========================================
# Drive Existence Check & Current Status
# ========================================
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host "BitLocker Configuration List" -ForegroundColor Cyan
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""

$validDrives = @()

foreach ($drive in $driveList) {
    $driveLetter = $drive.TargetDrive

    # Check drive exists
    if (-not (Test-Path "${driveLetter}\")) {
        Show-Warning "Drive $driveLetter does not exist (will be skipped)"
        Write-Host ""
        continue
    }

    # Get current BitLocker status
    $blVolume = Get-BitLockerVolume -MountPoint $driveLetter -ErrorAction SilentlyContinue
    $currentStatus = if ($blVolume) { $blVolume.ProtectionStatus } else { "Unknown" }
    $currentEncrypt = if ($blVolume) { $blVolume.VolumeStatus } else { "Unknown" }

    $skipHwText = if ($drive.SkipHardwareTest -eq "TRUE") { "Yes" } else { "No" }
    $usedOnlyText = if ($drive.UsedSpaceOnly -eq "TRUE") { "Yes" } else { "No" }
    $autoUnlockText = if ($drive.AutoUnlock -eq "TRUE") { "Yes" } else { "No" }

    Write-Host "  [$driveLetter] $($drive.Description)" -ForegroundColor Yellow
    Write-Host "    Current Status:      $currentStatus ($currentEncrypt)" -ForegroundColor Gray
    Write-Host "    Encryption Method:   $($drive.EncryptionMethod)" -ForegroundColor White
    Write-Host "    Used Space Only:     $usedOnlyText" -ForegroundColor White
    Write-Host "    Skip HW Test:        $skipHwText" -ForegroundColor White
    Write-Host "    Auto Unlock:         $autoUnlockText" -ForegroundColor White
    Write-Host ""

    $validDrives += $drive
}

if ($validDrives.Count -eq 0) {
    Show-Info "No valid drives to configure"
    return (New-ModuleResult -Status "Skipped" -Message "No valid drives found")
}

Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""

# ========================================
# Confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Enable BitLocker on the above drives?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""

# ========================================
# Evidence Directory Setup
# ========================================
$pcName = if (-not [string]::IsNullOrEmpty($env:SELECTED_NEW_PCNAME)) {
    $env:SELECTED_NEW_PCNAME
} else {
    $env:COMPUTERNAME
}
$dateStr = Get-Date -Format "yyyy_MM_dd"
$evidenceDir = Join-Path $PSScriptRoot "..\..\..\evidence\bitlocker\${dateStr}_${pcName}"

if (-not (Test-Path $evidenceDir)) {
    $null = New-Item -ItemType Directory -Path $evidenceDir -Force
}

# ========================================
# BitLocker Encryption Process
# ========================================
$successCount = 0
$skipCount = 0
$failCount = 0

foreach ($drive in $validDrives) {
    $driveLetter = $drive.TargetDrive
    $driveLabel = $driveLetter -replace ':', ''

    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "[$driveLetter] $($drive.Description)" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor White

    # --- Check if already encrypted ---
    $blVolume = Get-BitLockerVolume -MountPoint $driveLetter -ErrorAction SilentlyContinue
    if ($blVolume -and $blVolume.ProtectionStatus -eq "On") {
        Show-Skip "$driveLetter is already encrypted (ProtectionStatus: On)"

        # Still save recovery key for already-encrypted drives
        $existingKey = $blVolume.KeyProtector | Where-Object { $_.KeyProtectorType -eq "RecoveryPassword" } | Select-Object -First 1
        if ($null -ne $existingKey) {
            $evidencePath = Join-Path $evidenceDir "${pcName}_${driveLabel}.txt"
            $evidenceContent = New-BitLockerEvidence -Volume $blVolume -RecoveryKey $existingKey -PCName $pcName
            $evidenceContent | Out-File -FilePath $evidencePath -Encoding UTF8 -Force
            Show-Info "Recovery key saved: $evidencePath"
        }

        $skipCount++
        Write-Host ""
        continue
    }

    # --- Enable BitLocker ---
    try {
        Show-Info "Enabling BitLocker on $driveLetter..."

        $blParams = @{
            MountPoint                = $driveLetter
            EncryptionMethod          = $drive.EncryptionMethod
            RecoveryPasswordProtector  = $true
            ErrorAction               = "Stop"
        }

        if ($drive.UsedSpaceOnly -eq "TRUE") {
            $blParams.UsedSpaceOnly = $true
        }

        if ($drive.SkipHardwareTest -eq "TRUE") {
            $blParams.SkipHardwareTest = $true
        }

        $null = Enable-BitLocker @blParams
        Show-Success "BitLocker enabled on $driveLetter"
    }
    catch {
        Show-Error "Failed to enable BitLocker on ${driveLetter}: $_"
        $failCount++
        Write-Host ""
        continue
    }

    # --- Retrieve Recovery Key ---
    Show-Info "Retrieving recovery key..."

    $blVolume = Get-BitLockerVolume -MountPoint $driveLetter -ErrorAction SilentlyContinue
    if ($null -eq $blVolume) {
        Show-Error "Failed to retrieve BitLocker volume info for $driveLetter"
        $failCount++
        Write-Host ""
        continue
    }

    $recoveryKey = $blVolume.KeyProtector | Where-Object { $_.KeyProtectorType -eq "RecoveryPassword" } | Select-Object -First 1

    if ($null -eq $recoveryKey -or [string]::IsNullOrWhiteSpace($recoveryKey.RecoveryPassword)) {
        Show-Error "Recovery key is empty for $driveLetter"
        $failCount++
        Write-Host ""
        continue
    }

    Show-Success "Recovery key retrieved: $($recoveryKey.KeyProtectorId)"

    # --- Save Evidence ---
    Show-Info "Saving recovery key to evidence..."

    $evidencePath = Join-Path $evidenceDir "${pcName}_${driveLabel}.txt"
    $evidenceContent = New-BitLockerEvidence -Volume $blVolume -RecoveryKey $recoveryKey -PCName $pcName

    try {
        $evidenceContent | Out-File -FilePath $evidencePath -Encoding UTF8 -Force
        Show-Success "Evidence saved: $evidencePath"
    }
    catch {
        Show-Error "Failed to save evidence for ${driveLetter}: $_"
        $failCount++
        Write-Host ""
        continue
    }

    # --- Auto Unlock ---
    if ($drive.AutoUnlock -eq "TRUE") {
        Show-Info "Enabling Auto Unlock on $driveLetter..."
        try {
            Enable-BitLockerAutoUnlock -MountPoint $driveLetter -ErrorAction Stop
            Show-Success "Auto Unlock enabled on $driveLetter"
        }
        catch {
            Show-Warning "Failed to enable Auto Unlock on ${driveLetter}: $_"
            Show-Info "Auto Unlock may require the system drive to be fully encrypted first"
        }
    }

    $successCount++
    Write-Host ""
}

# ========================================
# Result Summary
# ========================================
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount -Title "BitLocker Configuration Results")
