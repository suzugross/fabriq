# ========================================
# BitLocker Disable Script
# ========================================
# 対象ドライブの BitLocker を無効化（復号化）する。
#
# [NOTES]
# - 管理者権限が必要
# - bitlocker_list.csv の TargetDrive を参照（暗号化関連カラムは無視）
# - 冪等性: 既に復号済み・復号中のドライブはスキップする
# - 暗号化処理中のドライブは中断して復号化に転じる
# ========================================

Write-Host ""
Show-Separator
Write-Host "BitLocker Disable" -ForegroundColor Cyan
Show-Separator
Write-Host ""


# ========================================
# Step 1: CSV 読み込み
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
# Step 3: 実行前の確認表示（ドライラン）
# ========================================
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "BitLocker Disable Targets" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

$hasDisableTarget = $false

foreach ($item in $enabledItems) {
    $driveLetter = $item.TargetDrive
    $displayName = if ($item.Description) { "$driveLetter $($item.Description)" } else { $driveLetter }

    # ドライブ存在確認
    if (-not (Test-Path "${driveLetter}\")) {
        Write-Host "  [NOT FOUND] $displayName" -ForegroundColor DarkGray
        Write-Host "    Drive does not exist" -ForegroundColor DarkGray
        Write-Host ""
        continue
    }

    # BitLocker ステータス取得
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

    # 状態別マーカー表示
    if ($volumeStatus -eq "FullyDecrypted") {
        Write-Host "  [ALREADY OFF] $displayName" -ForegroundColor DarkGray
        Write-Host "    Status: $protectionStatus ($volumeStatus)" -ForegroundColor DarkGray
    }
    elseif ($volumeStatus -eq "DecryptionInProgress") {
        Write-Host "  [DECRYPTING] $displayName" -ForegroundColor DarkGray
        Write-Host "    Status: $protectionStatus ($volumeStatus) - ${encryptPercent}% remaining" -ForegroundColor DarkGray
    }
    elseif ($volumeStatus -eq "EncryptionInProgress") {
        Write-Host "  [DISABLE] $displayName" -ForegroundColor Yellow
        Write-Host "    Status: $protectionStatus ($volumeStatus) - Encryption will be interrupted" -ForegroundColor Yellow
        $hasDisableTarget = $true
    }
    else {
        # FullyEncrypted or other encrypted states
        Write-Host "  [DISABLE] $displayName" -ForegroundColor Yellow
        Write-Host "    Status: $protectionStatus ($volumeStatus)" -ForegroundColor DarkGray
        $hasDisableTarget = $true
    }
    Write-Host ""
}

Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

if (-not $hasDisableTarget) {
    Show-Skip "No drives require BitLocker disabling"
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "No drives require BitLocker disabling")
}


# ========================================
# Step 4: 実行確認
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Disable BitLocker on the above drives?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""


# ========================================
# Step 5: 設定適用ループ
# ========================================
$successCount = 0
$skipCount    = 0
$failCount    = 0

foreach ($item in $enabledItems) {
    $driveLetter = $item.TargetDrive
    $displayName = if ($item.Description) { "$driveLetter $($item.Description)" } else { $driveLetter }

    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "Processing: $displayName" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor White

    # Skip 判定: ドライブ未検出
    if (-not (Test-Path "${driveLetter}\")) {
        Show-Skip "Drive not found: $driveLetter"
        Write-Host ""
        $skipCount++
        continue
    }

    # BitLocker ステータス取得
    $blVolume = Get-BitLockerVolume -MountPoint $driveLetter -ErrorAction SilentlyContinue
    if ($null -eq $blVolume) {
        Show-Skip "Unable to get BitLocker status: $driveLetter"
        Write-Host ""
        $skipCount++
        continue
    }

    $volumeStatus = $blVolume.VolumeStatus

    # Skip 判定: 既に復号済み
    if ($volumeStatus -eq "FullyDecrypted") {
        Show-Skip "Already decrypted: $driveLetter"
        Write-Host ""
        $skipCount++
        continue
    }

    # Skip 判定: 復号化処理中
    if ($volumeStatus -eq "DecryptionInProgress") {
        Show-Skip "Decryption already in progress: $driveLetter ($($blVolume.EncryptionPercentage)% remaining)"
        Write-Host ""
        $skipCount++
        continue
    }

    # Warning: 暗号化処理中のドライブを中断して復号化
    if ($volumeStatus -eq "EncryptionInProgress") {
        Show-Warning "Encryption is in progress on $driveLetter. Interrupting encryption and starting decryption."
    }

    # メイン処理
    try {
        Show-Info "Disabling BitLocker on $driveLetter..."
        $null = Disable-BitLocker -MountPoint $driveLetter -ErrorAction Stop
        Show-Success "BitLocker disable initiated: $driveLetter (decryption started)"
        $successCount++
    }
    catch {
        Show-Error "Failed to disable BitLocker on ${driveLetter}: $_"
        $failCount++
    }

    Write-Host ""
}


# ========================================
# Step 6: 結果集計・返却
# ========================================
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount `
    -Title "BitLocker Disable Results")
