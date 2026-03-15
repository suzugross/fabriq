# ========================================
# Restore Point Configuration Script
# ========================================
# Windows の復元ポイントに関する設定を行う。
# システム保護の有効化、24時間制限の解除、
# シャドウコピー容量設定、復元ポイントの作成を制御する。
#
# [NOTES]
# - 管理者権限が必要
# - クライアントOS（Windows 10/11）でのみ動作
# ========================================

Write-Host ""
Show-Separator
Write-Host "Restore Point Configuration" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# ========================================
# Local Helper Functions
# ========================================
# common.ps1 に該当関数なし。reg_hklm_config の
# Test-RegistryValueMatch パターンを参考に実装。
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

function Get-ShadowStorageInfo {
    param([string]$Drive)
    # ssid_config の netsh 出力パースパターンを参考に実装。
    # vssadmin list shadowstorage の出力から最大容量を取得する。
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
# Step 1: CSV 読み込み
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
# Step 2: 前提条件チェック（管理者権限）
# ========================================
if (-not (Test-AdminPrivilege)) {
    Show-Error "Administrator privileges required"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Administrator privileges required")
}


# ========================================
# Step 3: 実行前の確認表示
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
            # DisableSR = 0 ならシステム保護有効
            if (Test-RestoreRegistryValue -Name "DisableSR" -ExpectedValue 0) {
                $marker = "[SKIP]"
                $markerColor = "Gray"
            }
            $drive = $item.Drive
            Write-Host "  $marker $displayName" -ForegroundColor $markerColor
            Write-Host "    Drive: $drive" -ForegroundColor DarkGray
        }
        'remove_24h_limit' {
            # SystemRestorePointCreationFrequency = 0 なら制限解除済み
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
# Step 4: 実行確認
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Apply the above restore point settings?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""


# ========================================
# Step 5: 設定適用ループ
# ========================================
$successCount = 0
$skipCount    = 0
$failCount    = 0

foreach ($item in $enabledItems) {
    $displayName = $item.Description
    $settingName = $item.SettingName

    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "Processing: $displayName" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor White

    switch ($settingName) {

        'enable_protection' {
            # 冪等性チェック: DisableSR が既に 0 ならスキップ
            if (Test-RestoreRegistryValue -Name "DisableSR" -ExpectedValue 0) {
                Show-Skip "System protection already enabled"
                $skipCount++
                Write-Host ""
                continue
            }

            try {
                $drive = $item.Drive
                Enable-ComputerRestore -Drive $drive -ErrorAction Stop
                Show-Success "System protection enabled on $drive"
                $successCount++
            }
            catch {
                Show-Error "Failed to enable system protection: $_"
                $failCount++
            }
        }

        'remove_24h_limit' {
            # 冪等性チェック: 既に 0 ならスキップ
            if (Test-RestoreRegistryValue -Name "SystemRestorePointCreationFrequency" -ExpectedValue 0) {
                Show-Skip "24h limit already removed"
                $skipCount++
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
            }
            catch {
                Show-Error "Failed to remove 24h limit: $_"
                $failCount++
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

                Checkpoint-Computer -Description $rpDesc -RestorePointType $rpType -ErrorAction Stop
                Show-Success "Restore point created: $rpDesc"
                $successCount++
            }
            catch {
                Show-Error "Failed to create restore point: $_"
                $failCount++
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
# Step 6: 結果集計・返却
# ========================================
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount `
    -Title "Restore Point Configuration Results")
