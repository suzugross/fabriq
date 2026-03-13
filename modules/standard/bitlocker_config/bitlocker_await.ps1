# ========================================
# BitLocker Await Script
# ========================================
# 暗号化処理中の BitLocker ボリュームが完了するまで待機する。
#
# [NOTES]
# - 管理者権限が必要
# - bitlocker_list.csv の TargetDrive を参照し、対象ドライブのみ監視する
# - 冪等性: 暗号化中のドライブがなければ Skipped を返す
# - ポーリング間隔: 30秒
# ========================================

Write-Host ""
Show-Separator
Write-Host "BitLocker Await" -ForegroundColor Cyan
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
Write-Host "BitLocker Encryption Status" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

$hasAwaitTarget = $false

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
    if ($volumeStatus -eq "FullyEncrypted") {
        Write-Host "  [COMPLETE] $displayName" -ForegroundColor DarkGray
        Write-Host "    Status: $protectionStatus ($volumeStatus)" -ForegroundColor DarkGray
    }
    elseif ($volumeStatus -eq "EncryptionInProgress") {
        Write-Host "  [AWAIT] $displayName" -ForegroundColor Yellow
        Write-Host "    Status: $protectionStatus ($volumeStatus) - ${encryptPercent}% encrypted" -ForegroundColor Yellow
        $hasAwaitTarget = $true
    }
    elseif ($volumeStatus -eq "FullyDecrypted") {
        Write-Host "  [NOT ENCRYPTED] $displayName" -ForegroundColor DarkGray
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
    Show-Skip "No drives are currently encrypting"
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "No drives are currently encrypting")
}


# ========================================
# Step 4: 実行確認
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Wait for encryption to complete?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""


# ========================================
# Step 5: 待機ループ
# ========================================
# タイムアウト: ドライブ単位で進捗を監視し、30分間
# 1% も進行しなかったドライブは停滞とみなして監視を打ち切る。
# 進行があればタイムアウトカウンタはリセットされる。
# ========================================
$pollIntervalSec  = 30
$staleTimeoutSec  = 1800   # 30分

# ドライブごとの進捗追跡テーブル
$driveTracker = @{}
foreach ($item in $enabledItems) {
    $dl = $item.TargetDrive
    $blVol = Get-BitLockerVolume -MountPoint $dl -ErrorAction SilentlyContinue
    if ($blVol -and $blVol.VolumeStatus -eq "EncryptionInProgress") {
        $driveTracker[$dl] = @{
            LastPercent  = $blVol.EncryptionPercentage
            StaleElapsed = 0
            TimedOut     = $false
        }
    }
}

Show-Info "Monitoring encryption progress (polling every ${pollIntervalSec}s, stale timeout ${staleTimeoutSec}s)..."
Write-Host ""

while ($true) {
    $pendingDrives = @()

    foreach ($item in $enabledItems) {
        $driveLetter = $item.TargetDrive

        # タイムアウト済みドライブはスキップ
        if ($driveTracker.ContainsKey($driveLetter) -and $driveTracker[$driveLetter].TimedOut) { continue }

        if (-not (Test-Path "${driveLetter}\")) { continue }

        $blVolume = Get-BitLockerVolume -MountPoint $driveLetter -ErrorAction SilentlyContinue
        if ($null -eq $blVolume) { continue }

        if ($blVolume.VolumeStatus -eq "EncryptionInProgress") {
            $currentPercent = $blVolume.EncryptionPercentage

            # 進捗判定
            if ($driveTracker.ContainsKey($driveLetter)) {
                $tracker = $driveTracker[$driveLetter]
                if ($currentPercent -gt $tracker.LastPercent) {
                    # 進行あり → リセット
                    $tracker.LastPercent  = $currentPercent
                    $tracker.StaleElapsed = 0
                }
                else {
                    # 進行なし → カウンタ加算
                    $tracker.StaleElapsed += $pollIntervalSec
                    if ($tracker.StaleElapsed -ge $staleTimeoutSec) {
                        $tracker.TimedOut = $true
                        $displayName = if ($item.Description) { "$driveLetter $($item.Description)" } else { $driveLetter }
                        Show-Error "Stale timeout: $displayName (no progress for $($staleTimeoutSec / 60) min at ${currentPercent}%)"
                        continue
                    }
                }
            }

            $pendingDrives += @{
                Drive   = $driveLetter
                Percent = $currentPercent
            }
        }
    }

    if ($pendingDrives.Count -eq 0) {
        break
    }

    # 進捗表示
    $timestamp = Get-Date -Format "HH:mm:ss"
    $progressParts = @()
    foreach ($pd in $pendingDrives) {
        $progressParts += "$($pd.Drive) $($pd.Percent)%"
    }
    $progressText = $progressParts -join " / "
    Write-Host "  [$timestamp] Encrypting: $progressText" -ForegroundColor DarkGray

    Start-Sleep -Seconds $pollIntervalSec
}

Write-Host ""


# ========================================
# Step 6: 結果集計・返却
# ========================================
# 最終確認: 各ドライブの完了状態をチェック
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
        $skipCount++
    }
    else {
        Show-Error "Unexpected status on ${displayName}: $($blVolume.VolumeStatus)"
        $failCount++
    }
}

Write-Host ""
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount `
    -Title "BitLocker Await Results")
