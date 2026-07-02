# ========================================
# Restore Point Configuration Script
# ========================================
# Configure Windows System Restore.
# Enable system protection, remove the 24-hour throttle, set the
# shadow-copy storage cap, and create restore points on demand.
#
# [NOTES]
# - Requires administrator privileges
# - Client OS only (Windows 10 / 11)
# ========================================

Write-Host ""
Show-Separator
Write-Host "Restore Point Configuration" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# ========================================
# Local Helper Functions
# ========================================
# No equivalent helper exists in common.ps1; this implementation
# follows the Test-RegistryValueMatch pattern from reg_hklm_config.
# ========================================

function Test-RestoreRegistryValue {
    param(
        [string]$Name,
        [int]$ExpectedValue
    )
    $regPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore"
    try {
        if (-not (Test-Path $regPath)) { return $false }
        $prop = Get-ItemProperty -Path $regPath -Name $Name -ErrorAction SilentlyContinue
        if ($null -eq $prop) { return $false }
        return ([int]$prop.$Name -eq $ExpectedValue)
    }
    catch { return $false }
}

function Test-SystemRestoreEnabled {
    # "SR enabled on the drive" indicator. Enable-ComputerRestore sets
    # RPSessionInterval to a non-zero value (1 on Win10/11); it is 0 when SR is
    # not enabled. The legacy DisableSR flag is ABSENT on modern Windows and
    # cannot be used as an indicator - verified empirically on the test VM
    # (clean=0 / disabled=0 / enabled=1; DisableSR absent in all three states).
    $regPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore"
    try {
        $v = (Get-ItemProperty -Path $regPath -Name "RPSessionInterval" -ErrorAction SilentlyContinue).RPSessionInterval
        return ($null -ne $v -and [int]$v -ge 1)
    }
    catch { return $false }
}

function Get-ShadowStorageInfo {
    param([string]$Drive)
    # Implementation modeled on ssid_config's netsh-output parsing.
    # Extracts the maximum size from `vssadmin list shadowstorage`.
    try {
        $output = vssadmin list shadowstorage /for=$Drive 2>&1 | Out-String
        if ($LASTEXITCODE -ne 0) { return $null }
        if ($output -match 'Maximum[^:]*:\s*(.+)') {
            return $Matches[1].Trim()
        }
        return $null
    }
    catch { return $null }
}


# ========================================
# Step 1: Load CSV
# ========================================
$csvPath = Join-Path $PSScriptRoot "restore_point_list.csv"

$enabledItems = Import-ModuleCsv -Path $csvPath -FilterEnabled `
    -RequiredColumns @("Enabled", "SettingName", "Description")

if ($null -eq $enabledItems) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load restore_point_list.csv")
}
if ($enabledItems.Count -eq 0) {
    return (New-ModuleResult -Status "Skipped" -Message "No enabled entries")
}


# ========================================
# Step 2: Prerequisite check (administrator privileges)
# ========================================
if (-not (Test-AdminPrivilege)) {
    Show-Error "Administrator privileges required"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Administrator privileges required")
}


# ========================================
# Step 3: Dry-run summary before execution
# ========================================
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "Restore Point Settings" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

$regPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore"

foreach ($item in $enabledItems) {
    $displayName = $item.Description
    $settingName = $item.SettingName

    $marker = "[APPLY]"
    $markerColor = "Yellow"

    switch ($settingName) {
        'enable_protection' {
            # Same indicator as the apply loop (RPSessionInterval >= 1).
            # The legacy DisableSR flag is absent on modern Windows, so the
            # old DisableSR-based preview always showed [APPLY] even when
            # the apply loop would skip.
            if (Test-SystemRestoreEnabled) {
                $marker = "[SKIP]"
                $markerColor = "Gray"
            }
            $drive = $item.Drive
            Write-Host "  $marker $displayName" -ForegroundColor $markerColor
            Write-Host "    Drive: $drive" -ForegroundColor DarkGray
        }
        'remove_24h_limit' {
            # 24h limit is already removed when SystemRestorePointCreationFrequency = 0
            if (Test-RestoreRegistryValue -Name "SystemRestorePointCreationFrequency" -ExpectedValue 0) {
                $marker = "[SKIP]"
                $markerColor = "Gray"
            }
            Write-Host "  $marker $displayName" -ForegroundColor $markerColor
            Write-Host "    Registry: SystemRestorePointCreationFrequency = 0" -ForegroundColor DarkGray
        }
        'set_storage_size' {
            $drive = $item.Drive
            $targetPercent = $item.Value
            $currentMax = Get-ShadowStorageInfo -Drive $drive
            if ($null -ne $currentMax) {
                Write-Host "  $marker $displayName" -ForegroundColor $markerColor
                Write-Host "    Drive: $drive  Current: $currentMax  -> Target: ${targetPercent}%" -ForegroundColor DarkGray
            }
            else {
                Write-Host "  $marker $displayName" -ForegroundColor $markerColor
                Write-Host "    Drive: $drive  Target: ${targetPercent}%" -ForegroundColor DarkGray
            }
        }
        'create_restore_point' {
            $rpType = if ($item.Value) { $item.Value } else { "MODIFY_SETTINGS" }
            Write-Host "  $marker $displayName" -ForegroundColor $markerColor
            Write-Host "    Type: $rpType" -ForegroundColor DarkGray
        }
        default {
            Write-Host "  [UNKNOWN] $displayName ($settingName)" -ForegroundColor Red
        }
    }
    Write-Host ""
}

Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""


# ========================================
# Step 4: User confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Apply the above restore point settings?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""


# ========================================
# Step 5: Apply-settings loop
# ========================================
$successCount = 0
$skipCount    = 0
$failCount    = 0
# Post-Apply Verification tally (read-back of applied/skipped settings).
# enable_protection / remove_24h_limit / restore-point creation are verified.
# Excluded (still counts toward success/fail): set_storage_size, whose exact
# maxsize% is only readable via locale-fragile vssadmin-list parsing.
$verifyPass = 0
$verifyFail = 0

foreach ($item in $enabledItems) {
    $displayName = $item.Description
    $settingName = $item.SettingName

    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "Processing: $displayName" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor White

    switch ($settingName) {

        'enable_protection' {
            # Idempotency: skip when SR is already enabled (RPSessionInterval >= 1).
            if (Test-SystemRestoreEnabled) {
                Show-Skip "System protection already enabled"
                $skipCount++
                $verifyPass++
                Write-Host ""
                continue
            }

            try {
                $drive = $item.Drive
                Enable-ComputerRestore -Drive $drive -ErrorAction Stop
                Show-Success "System protection enabled on $drive"
                $successCount++
                if (Test-SystemRestoreEnabled) {
                    $verifyPass++
                } else {
                    Show-Warning "Verify: System Restore not reported enabled (RPSessionInterval) after Enable-ComputerRestore"
                    $verifyFail++
                }
            }
            catch {
                Show-Error "Failed to enable system protection: $_"
                $failCount++
                $verifyFail++
            }
        }

        'remove_24h_limit' {
            # Idempotency: skip when the value is already 0
            if (Test-RestoreRegistryValue -Name "SystemRestorePointCreationFrequency" -ExpectedValue 0) {
                Show-Skip "24h limit already removed"
                $skipCount++
                $verifyPass++
                Write-Host ""
                continue
            }

            try {
                $path = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore"
                $prop = Get-ItemProperty -Path $path -Name "SystemRestorePointCreationFrequency" -ErrorAction SilentlyContinue

                if ($null -ne $prop) {
                    Set-ItemProperty -Path $path -Name "SystemRestorePointCreationFrequency" -Value 0 -Force -ErrorAction Stop
                }
                else {
                    New-ItemProperty -Path $path -Name "SystemRestorePointCreationFrequency" -Value 0 -PropertyType DWord -Force -ErrorAction Stop | Out-Null
                }
                Show-Success "24h creation limit removed"
                $successCount++
                if (Test-RestoreRegistryValue -Name "SystemRestorePointCreationFrequency" -ExpectedValue 0) {
                    $verifyPass++
                } else {
                    Show-Warning "Verify: SystemRestorePointCreationFrequency is not 0 after apply"
                    $verifyFail++
                }
            }
            catch {
                Show-Error "Failed to remove 24h limit: $_"
                $failCount++
                $verifyFail++
            }
        }

        'set_storage_size' {
            try {
                $drive = $item.Drive
                $targetPercent = $item.Value
                $output = vssadmin resize shadowstorage /for=$drive /on=$drive /maxsize=${targetPercent}% 2>&1 | Out-String

                if ($LASTEXITCODE -ne 0) {
                    Show-Error "vssadmin failed: $output"
                    $failCount++
                }
                else {
                    Show-Success "Shadow storage max size set to ${targetPercent}% on $drive"
                    $successCount++
                }
            }
            catch {
                Show-Error "Failed to set storage size: $_"
                $failCount++
            }
        }

        'create_restore_point' {
            try {
                $rpType = if ($item.Value) { $item.Value } else { "MODIFY_SETTINGS" }
                $rpDesc = $item.Description

                # Checkpoint-Computer does NOT throw when the 24h throttle
                # suppresses creation - it warns and returns normally,
                # which used to count as Success with no restore point on
                # disk. Record the latest SequenceNumber and read back.
                $beforeSeq = 0
                # Baseline = MAX SequenceNumber, not "last enumerated":
                # WMI enumeration order is not contractual, and an
                # understated baseline would let an OLD same-description
                # point pass the read-back below as freshly created.
                $allPoints = @(Get-ComputerRestorePoint -ErrorAction SilentlyContinue)
                if ($allPoints.Count -gt 0) {
                    $beforeSeq = [int64]($allPoints | Measure-Object -Property SequenceNumber -Maximum).Maximum
                }

                $cpWarnings = @()
                Checkpoint-Computer -Description $rpDesc -RestorePointType $rpType -ErrorAction Stop -WarningVariable cpWarnings

                # Fail-closed read-back: no new point (or an unreadable
                # list) is a Fail, whatever the cause (throttle, VSS).
                $created = Get-ComputerRestorePoint -ErrorAction SilentlyContinue |
                    Where-Object { [int64]$_.SequenceNumber -gt $beforeSeq -and $_.Description -eq $rpDesc } |
                    Select-Object -First 1

                if ($created) {
                    Show-Success "Restore point created: $rpDesc (SequenceNumber=$($created.SequenceNumber))"
                    $successCount++
                    $verifyPass++
                }
                else {
                    $warnText = if ($cpWarnings.Count -gt 0) { " ($($cpWarnings -join '; '))" } else { "" }
                    Show-Error "Restore point was NOT created: $rpDesc$warnText"
                    $failCount++
                    $verifyFail++
                }
            }
            catch {
                Show-Error "Failed to create restore point: $_"
                $failCount++
                $verifyFail++
            }
        }

        default {
            Show-Error "Unknown setting: $settingName"
            $failCount++
        }
    }

    Write-Host ""
}


# ========================================
# Step 6: Aggregate and return result
# ========================================
# -Verified covers the read-back-verifiable settings: SR enabled (RPSessionInterval),
# 24h-limit reads back as 0, restore-point creation confirmed by a strictly-newer
# SequenceNumber. $null when none were in scope, $true when all read back as
# expected, $false on any mismatch or apply failure among them.
$verified = if (($verifyPass + $verifyFail) -gt 0) { ($verifyFail -eq 0) } else { $null }

return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount `
    -Title "Restore Point Configuration Results" -Verified $verified)
