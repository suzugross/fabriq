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
        if ($kp.KeyProtectorType -eq "TpmPin") {
            $line += "  (PIN not recorded for security)"
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

# ========================================
# Helper: Set FVE Group Policy Registry for TPM+PIN
# ========================================
function Set-FveRegistryPolicy {
    param(
        [switch]$EnableEnhancedPin
    )

    $fvePath = "HKLM:\SOFTWARE\Policies\Microsoft\FVE"

    if (-not (Test-Path $fvePath)) {
        try {
            New-Item -Path $fvePath -Force | Out-Null
        }
        catch {
            Show-Error "Failed to create FVE registry key: $_"
            return $false
        }
    }

    $fveValues = @(
        @{ Name = "UseAdvancedStartup"; Value = 1 }
        @{ Name = "EnableBDEWithNoTPM"; Value = 0 }
        @{ Name = "UseTPM";             Value = 0 }
        @{ Name = "UseTPMPIN";          Value = 1 }
        @{ Name = "UseTPMKey";          Value = 0 }
        @{ Name = "UseTPMKeyPIN";       Value = 0 }
    )

    if ($EnableEnhancedPin) {
        $fveValues += @{ Name = "UseEnhancedPin"; Value = 1 }
    }

    $failCount = 0
    foreach ($entry in $fveValues) {
        try {
            Set-ItemProperty -Path $fvePath -Name $entry.Name -Value $entry.Value -Type DWord -Force -ErrorAction Stop
            Show-Success "FVE: $($entry.Name) = $($entry.Value)"
        }
        catch {
            Show-Error "Failed to set FVE $($entry.Name): $_"
            $failCount++
        }
    }

    return ($failCount -eq 0)
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

# ========================================
# Resolve PIN for Each Drive (hostlist > CSV > none)
# ========================================
$csvColumns = $driveList[0].PSObject.Properties.Name
$hasPinColumn = 'Pin' -in $csvColumns
$hostPin = $env:SELECTED_PIN

$pinRequired = $false
$enhancedPinRequired = $false

foreach ($drive in $driveList) {
    # Priority 1: hostlist ($env:SELECTED_PIN)
    # Priority 2: module CSV (Pin column)
    # Priority 3: none
    $csvPin = $null
    if ($hasPinColumn -and -not [string]::IsNullOrWhiteSpace($drive.Pin)) {
        $csvPin = $drive.Pin
    }

    $resolvedPin = $null
    $pinSource   = $null

    if (-not [string]::IsNullOrWhiteSpace($hostPin)) {
        $resolvedPin = $hostPin
        $pinSource   = "hostlist"
    }
    elseif ($null -ne $csvPin) {
        $resolvedPin = $csvPin
        $pinSource   = "module CSV"
    }

    # Guard: reject unresolved ENC: values (decryption failed or passphrase unavailable)
    if ($null -ne $resolvedPin -and $resolvedPin.StartsWith('ENC:')) {
        Show-Error "PIN ($pinSource) is still encrypted (ENC:). Decryption may have failed or passphrase was not entered."
        return (New-ModuleResult -Status "Error" -Message "PIN decryption failed - encrypted value cannot be used as PIN")
    }

    # Store resolved PIN into the drive object for downstream use
    if ($null -ne $resolvedPin) {
        if ($hasPinColumn) {
            $drive.Pin = $resolvedPin
        }
        else {
            $drive | Add-Member -NotePropertyName 'Pin' -NotePropertyValue $resolvedPin -Force
        }
        $drive | Add-Member -NotePropertyName '_PinSource' -NotePropertyValue $pinSource -Force

        $pinRequired = $true
        if ($resolvedPin -match '[^0-9]') {
            $enhancedPinRequired = $true
        }
    }
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
    $drivePin = if ($drive.PSObject.Properties.Name -contains 'Pin') { $drive.Pin } else { $null }
    if (-not [string]::IsNullOrWhiteSpace($drivePin)) {
        $source = if ($drive.PSObject.Properties.Name -contains '_PinSource') { $drive._PinSource } else { "unknown" }
        Write-Host "    PIN Protector:       Yes (source: $source)" -ForegroundColor White
    }
    elseif ($hasPinColumn -or -not [string]::IsNullOrWhiteSpace($hostPin)) {
        Write-Host "    PIN Protector:       -" -ForegroundColor White
    }
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
# FVE Registry Policy (for TPM+PIN)
# ========================================
if ($pinRequired) {
    if ($enhancedPinRequired) {
        Show-Info "Setting FVE Group Policy registry for TPM+PIN support (Enhanced PIN)..."
    }
    else {
        Show-Info "Setting FVE Group Policy registry for TPM+PIN support..."
    }
    Write-Host ""

    $fveResult = Set-FveRegistryPolicy -EnableEnhancedPin:$enhancedPinRequired
    if (-not $fveResult) {
        Show-Error "Failed to set FVE registry policy. Cannot proceed with PIN configuration."
        return (New-ModuleResult -Status "Error" -Message "FVE registry policy setup failed")
    }

    Write-Host ""
}

# ========================================
# Evidence Directory Setup
# ========================================
$pcName = if (-not [string]::IsNullOrEmpty($env:SELECTED_NEW_PCNAME)) {
    $env:SELECTED_NEW_PCNAME
} else {
    $env:COMPUTERNAME
}

if (-not [string]::IsNullOrWhiteSpace($global:FabriqEvidenceBasePath)) {
    # Unified path mode: flat under bitlocker/
    $evidenceDir = Join-Path $global:FabriqEvidenceBasePath "bitlocker"
} else {
    # Fallback: legacy path
    $dateStr = Get-Date -Format "yyyy_MM_dd"
    $uid     = if ($global:FabriqUniqueId) { $global:FabriqUniqueId } else { Get-HardwareUniqueId }
    $evidenceDir = Join-Path $PSScriptRoot "..\..\..\evidence\bitlocker\${dateStr}_${uid}_${pcName}"
}

if (-not (Test-Path $evidenceDir)) {
    $null = New-Item -ItemType Directory -Path $evidenceDir -Force
}

# ========================================
# BitLocker Encryption Process
# ========================================
$successCount = 0
$skipCount = 0
$failCount = 0
$verifyFail = 0

foreach ($drive in $validDrives) {
    $driveLetter = $drive.TargetDrive
    $driveLabel = $driveLetter -replace ':', ''

    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "[$driveLetter] $($drive.Description)" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor White

    # --- Check if already applied ---
    # Match the module's own verification shapes (see Step 5.5 below):
    #   On                    - protection active
    #   EncryptionInProgress  - conversion underway (shape a)
    #   FullyDecrypted + RecoveryPassword protector
    #                         - deferred enable pending hardware test/reboot (shape b)
    # FullyEncrypted + Off (suspended protection) is deliberately NOT treated
    # as applied: that state needs operator attention and keeps surfacing as
    # an Error through the enable path, as before. Previously only "On" was
    # accepted, so a re-run inside the deferred window re-entered
    # Enable-BitLocker and produced a false Fail.
    $blVolume = Get-BitLockerVolume -MountPoint $driveLetter -ErrorAction SilentlyContinue
    $rpProtector = if ($blVolume) { $blVolume.KeyProtector | Where-Object { $_.KeyProtectorType -eq "RecoveryPassword" } } else { $null }
    $alreadyApplied = $blVolume -and (
        $blVolume.ProtectionStatus -eq "On" -or
        $blVolume.VolumeStatus -eq "EncryptionInProgress" -or
        ($blVolume.VolumeStatus -eq "FullyDecrypted" -and $rpProtector)
    )
    if ($alreadyApplied) {

        $hasPin = -not [string]::IsNullOrWhiteSpace($drive.Pin)
        $rowMutated = $false   # set when this branch changes protector state

        if ($hasPin) {
            # Check if TpmPin protector already exists
            $existingTpmPin = $blVolume.KeyProtector | Where-Object { $_.KeyProtectorType -eq "TpmPin" }
            if ($existingTpmPin) {
                # Self-heal: a bare Tpm protector left alongside TpmPin (e.g.
                # a previous upgrade whose Remove step failed) permanently
                # bypasses the PIN at boot. Detect and remove it instead of
                # skipping past it forever.
                $bareTpm = $blVolume.KeyProtector | Where-Object { $_.KeyProtectorType -eq "Tpm" }
                if ($bareTpm) {
                    Show-Warning "$driveLetter has a bare Tpm protector alongside TpmPin (PIN bypass); removing..."
                    try {
                        foreach ($tp in $bareTpm) {
                            $null = Remove-BitLockerKeyProtector -MountPoint $driveLetter -KeyProtectorId $tp.KeyProtectorId -ErrorAction Stop
                            Show-Success "Removed leftover Tpm-only protector: $($tp.KeyProtectorId)"
                        }
                        $rowMutated = $true
                    }
                    catch {
                        Show-Error "Failed to remove leftover Tpm protector on ${driveLetter}: $_ (PIN remains bypassable until fixed)"
                        $failCount++
                        $verifyFail++
                        Write-Host ""
                        continue
                    }
                }
                else {
                    Show-Skip "$driveLetter already has TpmPin protector"
                }
            }
            else {
                # Already applied with TPM-only -> Upgrade to TpmAndPin
                if ($blVolume.VolumeType -ne "OperatingSystem") {
                    Show-Warning "$driveLetter is not an OS drive. PIN protector only applies to OS drives. Skipping PIN."
                }
                else {
                    $pinSrc = if ($drive.PSObject.Properties.Name -contains '_PinSource') { $drive._PinSource } else { "unknown" }
                    Show-Info "$driveLetter is encrypted. Adding TpmAndPin protector (PIN source: $pinSrc)..."
                    try {
                        $securePin = ConvertTo-SecureString $drive.Pin -AsPlainText -Force
                        $null = Add-BitLockerKeyProtector -MountPoint $driveLetter -Pin $securePin -TpmAndPinProtector -ErrorAction Stop
                        Show-Success "Added TpmAndPin protector to $driveLetter"

                        # Remove old Tpm-only protector. Re-read the volume
                        # first: some Windows builds auto-drop the bare Tpm
                        # protector when TpmAndPin is added, which invalidated
                        # the stale KeyProtectorId and made the explicit
                        # Remove throw 0x80070490 (element not found) -> a
                        # false Error even though the goal state was reached.
                        # A missing protector is success (goal = no bare Tpm).
                        $reread = Get-BitLockerVolume -MountPoint $driveLetter -ErrorAction SilentlyContinue
                        $oldTpm = if ($reread) { $reread.KeyProtector | Where-Object { $_.KeyProtectorType -eq "Tpm" } } else { $null }
                        foreach ($tp in $oldTpm) {
                            try {
                                $null = Remove-BitLockerKeyProtector -MountPoint $driveLetter -KeyProtectorId $tp.KeyProtectorId -ErrorAction Stop
                                Show-Success "Removed old Tpm-only protector: $($tp.KeyProtectorId)"
                            }
                            catch {
                                # Re-check: tolerate "already gone", fail only
                                # if a bare Tpm genuinely persists (real PIN
                                # bypass) - the Step 5.5 read-back below is the
                                # authoritative gate on the final shape.
                                $still = (Get-BitLockerVolume -MountPoint $driveLetter -ErrorAction SilentlyContinue).KeyProtector |
                                    Where-Object { $_.KeyProtectorType -eq "Tpm" -and $_.KeyProtectorId -eq $tp.KeyProtectorId }
                                if ($still) {
                                    Show-Error "Failed to remove old Tpm-only protector on ${driveLetter}: $_ (PIN remains bypassable)"
                                }
                                else {
                                    Show-Info "Old Tpm-only protector already gone on $driveLetter (auto-removed with TpmAndPin add)"
                                }
                            }
                        }
                        $rowMutated = $true
                    }
                    catch {
                        # Add itself failed: no protector shape change.
                        Show-Error "Failed to upgrade $driveLetter to TpmAndPin: $_"
                        $failCount++
                        $verifyFail++
                        Write-Host ""
                        continue
                    }
                }
            }
        }
        else {
            if ($blVolume.ProtectionStatus -eq "On") {
                Show-Skip "$driveLetter is already encrypted (ProtectionStatus: On)"
            }
            else {
                Show-Skip "$driveLetter encryption already accepted (VolumeStatus: $($blVolume.VolumeStatus))"
            }
        }

        # Save recovery key for already-applied drives
        $blVolume = Get-BitLockerVolume -MountPoint $driveLetter -ErrorAction SilentlyContinue
        $existingKey = $blVolume.KeyProtector | Where-Object { $_.KeyProtectorType -eq "RecoveryPassword" } | Select-Object -First 1
        if ($null -ne $existingKey) {
            $evidencePath = Join-Path $evidenceDir "${pcName}_${driveLabel}.txt"
            $evidenceContent = New-BitLockerEvidence -Volume $blVolume -RecoveryKey $existingKey -PCName $pcName
            $evidenceContent | Out-File -FilePath $evidencePath -Encoding UTF8 -Force
            Show-Info "Recovery key saved: $evidencePath"
        }

        if ($rowMutated) {
            # Read back the protector shape this row was supposed to reach:
            # TpmPin present AND no bare Tpm protector (PIN not bypassable).
            $tpmPinOk    = [bool]($blVolume.KeyProtector | Where-Object { $_.KeyProtectorType -eq "TpmPin" })
            $bareTpmGone = -not ($blVolume.KeyProtector | Where-Object { $_.KeyProtectorType -eq "Tpm" })
            if ($tpmPinOk -and $bareTpmGone) {
                Show-Success "Verified: $driveLetter protector shape is TpmPin without bare Tpm"
            }
            else {
                Show-Warning "Verification: $driveLetter protector shape unexpected (TpmPin=$tpmPinOk, bareTpmAbsent=$bareTpmGone)"
                $verifyFail++
            }
            $successCount++
        }
        else {
            $skipCount++
        }
        Write-Host ""
        continue
    }

    # --- Enable BitLocker ---
    $hasPin = -not [string]::IsNullOrWhiteSpace($drive.Pin)

    # Validate: PIN only applies to OS drives
    if ($hasPin) {
        $volumeCheck = Get-BitLockerVolume -MountPoint $driveLetter -ErrorAction SilentlyContinue
        if ($volumeCheck -and $volumeCheck.VolumeType -ne "OperatingSystem") {
            Show-Warning "$driveLetter is not an OS drive. PIN protector only applies to OS drives. Using TPM-only."
            $hasPin = $false
        }
    }

    try {
        Show-Info "Enabling BitLocker on $driveLetter..."

        if ($hasPin) {
            # --- TPM+PIN path ---
            $pinSrc = if ($drive.PSObject.Properties.Name -contains '_PinSource') { $drive._PinSource } else { "unknown" }
            Show-Info "Using TpmAndPin protector (PIN source: $pinSrc)"
            $securePin = ConvertTo-SecureString $drive.Pin -AsPlainText -Force

            $blParams = @{
                MountPoint         = $driveLetter
                EncryptionMethod   = $drive.EncryptionMethod
                Pin                = $securePin
                TpmAndPinProtector = $true
                ErrorAction        = "Stop"
            }

            if ($drive.UsedSpaceOnly -eq "TRUE") {
                $blParams.UsedSpaceOnly = $true
            }
            if ($drive.SkipHardwareTest -eq "TRUE") {
                $blParams.SkipHardwareTest = $true
            }

            $null = Enable-BitLocker @blParams
            Show-Success "BitLocker enabled on $driveLetter with TpmAndPin protector"

            # Add RecoveryPassword protector separately (different parameter set)
            Show-Info "Adding RecoveryPassword protector..."
            $null = Add-BitLockerKeyProtector -MountPoint $driveLetter -RecoveryPasswordProtector -ErrorAction Stop
            Show-Success "RecoveryPassword protector added to $driveLetter"
        }
        else {
            # --- Original TPM-only path ---
            $blParams = @{
                MountPoint                = $driveLetter
                EncryptionMethod          = $drive.EncryptionMethod
                RecoveryPasswordProtector = $true
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

    # --- Post-apply verification: confirm the encryption was accepted ---
    # Do NOT use ProtectionStatus: it stays Off until the conversion completes (often only after
    # a reboot), which would yield a false FAIL. Confirm an enable-acceptance signature instead.
    # Two valid acceptance shapes:
    #   (a) immediate  - conversion is underway/done (data drives; OS drive with SkipHardwareTest):
    #                    VolumeStatus EncryptionInProgress/FullyEncrypted + a real method.
    #   (b) deferred   - OS drive without SkipHardwareTest schedules encryption for the next boot,
    #                    so VolumeStatus stays FullyDecrypted while a RecoveryPassword protector is
    #                    already in place. Per CLAUDE.md section 6, verify "acceptance" (protector
    #                    present) for such reboot-pending state rather than active conversion.
    $vbl = Get-BitLockerVolume -MountPoint $driveLetter -ErrorAction SilentlyContinue
    $rpPresent = ($null -ne $vbl) -and
                 [bool]($vbl.KeyProtector | Where-Object { $_.KeyProtectorType -eq "RecoveryPassword" })
    if ($null -eq $vbl) {
        Show-Warning "Verification: $driveLetter volume info unavailable"
        $verifyFail++
    }
    elseif (($vbl.VolumeStatus -in @("EncryptionInProgress", "FullyEncrypted")) -and
            ($vbl.EncryptionMethod -and $vbl.EncryptionMethod -ne "None")) {
        # (a) immediate: encryption underway or complete -> confirmed.
    }
    elseif ($rpPresent -and ($vbl.VolumeStatus -eq "FullyDecrypted")) {
        # (b) deferred: protectors in place, encryption scheduled for next boot -> accepted.
        Show-Info "$driveLetter encryption is scheduled (pending hardware test / reboot)"
    }
    else {
        Show-Warning "Verification: $driveLetter encryption not confirmed (VolumeStatus=$($vbl.VolumeStatus), recoveryProtector=$rpPresent)"
        $verifyFail++
    }

    # PIN rows: the protector shape must be TpmPin WITHOUT a bare Tpm
    # protector (a bare Tpm next to TpmPin makes the PIN bypassable at boot).
    if ($hasPin -and $null -ne $vbl) {
        $tpmPinOk    = [bool]($vbl.KeyProtector | Where-Object { $_.KeyProtectorType -eq "TpmPin" })
        $bareTpmGone = -not ($vbl.KeyProtector | Where-Object { $_.KeyProtectorType -eq "Tpm" })
        if (-not ($tpmPinOk -and $bareTpmGone)) {
            Show-Warning "Verification: $driveLetter protector shape unexpected (TpmPin=$tpmPinOk, bareTpmAbsent=$bareTpmGone)"
            $verifyFail++
        }
    }

    $successCount++
    Write-Host ""
}

# ========================================
# Result Summary
# ========================================
# Verified: True only when something ran and every drive that ran is confirmed (no readback
# mismatch and no apply failure). Skip-because-already-encrypted is a desired-state confirmation.
# $null when no drive ran (no valid drives -> early Skipped return above).
$verified = if (($successCount + $skipCount + $failCount) -gt 0) {
    (($verifyFail -eq 0) -and ($failCount -eq 0))
} else { $null }
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount -Verified $verified -Title "BitLocker Configuration Results")
