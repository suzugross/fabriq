# ========================================
# History Destroyer - CSV-Driven History Cleanup
# ========================================
# Deletes various Windows history, cache, and temporary data
# based on destroy_list.csv configuration.
#
# NOTES:
# - Requires administrator privileges for some operations
# - Explorer will be temporarily stopped during cleanup
# - Special handlers manage complex cleanup operations (browsers, Office, etc.)
# ========================================

Write-Host ""
Show-Separator
Write-Host "History Destroyer" -ForegroundColor Cyan
Show-Separator
Write-Host ""


# ========================================
# [P/Invoke] ごみ箱の完全消去に使用
# ========================================
Add-Type -MemberDefinition @'
    [DllImport("Shell32.dll")]
    public static extern int SHEmptyRecycleBin(IntPtr hwnd, string pszRootPath, int dwFlags);
'@ -Name Win32RecycleBin -Namespace HistoryDestroyer -ErrorAction SilentlyContinue


# ========================================
# プライベート関数群: Special ハンドラ
# ========================================

# ----------------------------------------
# ディスパッチャ: TargetPath → 対応ハンドラへ振り分け
# 戻り値: "Success" or "Skip"（失敗時は throw）
# ----------------------------------------
function Invoke-DestroyHandler {
    param([string]$HandlerName)
    switch ($HandlerName) {
        "clear-all-eventlogs" { return (Clear-AllEventLogs) }
        "recycle-bin"         { return (Clear-RecycleBinSafe) }
        "office-mru"          { return (Clear-OfficeMRU) }
        "edge-cleanup"        { return (Clear-BrowserData -Browser "Edge" -BasePath "$env:LOCALAPPDATA\Microsoft\Edge\User Data" -ProcessName "msedge") }
        "chrome-cleanup"      { return (Clear-BrowserData -Browser "Chrome" -BasePath "$env:LOCALAPPDATA\Google\Chrome\User Data" -ProcessName "chrome") }
        "search-index"        { return (Clear-SearchIndex) }
        "wifi-ssid"           { return (Clear-WiFiProfiles) }
        default               { throw "Unknown special handler: $HandlerName" }
    }
}

# ----------------------------------------
# (1) イベントログ全消去
# ----------------------------------------
function Clear-AllEventLogs {
    $logs = Get-WinEvent -ListLog * -Force -ErrorAction SilentlyContinue
    if ($null -eq $logs -or $logs.Count -eq 0) {
        Show-Skip "No event logs found"
        return "Skip"
    }

    $clearedCount = 0
    foreach ($log in $logs) {
        $null = & wevtutil.exe cl $log.LogName 2>&1
        if ($LASTEXITCODE -eq 0) { $clearedCount++ }
    }

    Show-Success "Cleared $clearedCount event logs"
    return "Success"
}

# ----------------------------------------
# (2) ごみ箱消去 (P/Invoke)
# ----------------------------------------
function Clear-RecycleBinSafe {
    # Flags: SHERB_NOCONFIRMATION(1) | SHERB_NOPROGRESSUI(2) | SHERB_NOSOUND(4) = 7
    $result = [HistoryDestroyer.Win32RecycleBin]::SHEmptyRecycleBin([IntPtr]::Zero, $null, 7)

    if ($result -eq 0) {
        Show-Success "Recycle Bin emptied"
    }
    else {
        # -2147418113 (0x8000FFFF) 等: ごみ箱が空の場合も正常
        Show-Info "Recycle Bin already empty or emptied (HRESULT: $result)"
    }
    return "Success"
}

# ----------------------------------------
# (3) Office MRU レジストリ動的列挙・削除
# ----------------------------------------
function Clear-OfficeMRU {
    $officeBase = "HKCU:\Software\Microsoft\Office"
    if (-not (Test-Path $officeBase)) {
        Show-Skip "Office registry not found"
        return "Skip"
    }

    $officeCleaned = 0
    $apps = @("Word", "Excel", "PowerPoint", "Access", "Publisher", "Visio")

    Get-ChildItem $officeBase -ErrorAction SilentlyContinue | ForEach-Object {
        $version = $_.PSChildName
        foreach ($app in $apps) {
            $placeMRU = "$officeBase\$version\$app\Place MRU"
            $fileMRU  = "$officeBase\$version\$app\File MRU"

            if (Test-Path $placeMRU) {
                $null = Remove-ItemProperty -Path $placeMRU -Name * -Force -ErrorAction SilentlyContinue
                $officeCleaned++
            }
            if (Test-Path $fileMRU) {
                $null = Remove-ItemProperty -Path $fileMRU -Name * -Force -ErrorAction SilentlyContinue
                $officeCleaned++
            }
        }
    }

    Show-Success "Office MRU cleaned ($officeCleaned entries)"
    return "Success"
}

# ----------------------------------------
# (4)(5) ブラウザデータ削除（Edge / Chrome 共通）
# ----------------------------------------
function Clear-BrowserData {
    param(
        [string]$Browser,
        [string]$BasePath,
        [string]$ProcessName
    )

    if (-not (Test-Path $BasePath)) {
        Show-Skip "$Browser not found"
        return "Skip"
    }

    # ブラウザプロセス停止
    $proc = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue
    if ($proc) {
        try {
            Stop-Process -Name $ProcessName -Force -ErrorAction Stop
            Start-Sleep -Seconds 2
        }
        catch {
            Show-Warning "Failed to stop $Browser process: $($_.Exception.Message)"
        }
    }

    # 削除対象（元コード 14種）
    $browserTargets = @(
        "Cache", "Code Cache", "GPUCache",
        "History", "Cookies", "Cookies-journal",
        "Top Sites", "Top Sites-journal",
        "Visited Links",
        "Web Data", "Web Data-journal",
        "Session Storage", "Local Storage"
    )

    # 全プロファイルを列挙（Default, Profile 1, Profile 2, ...）
    $profiles = Get-ChildItem $BasePath -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -eq "Default" -or $_.Name -match "^Profile " }

    $cleanedCount = 0
    foreach ($profile in $profiles) {
        foreach ($target in $browserTargets) {
            $targetPath = Join-Path $profile.FullName $target
            if (Test-Path $targetPath) {
                try {
                    Remove-Item $targetPath -Recurse -Force -ErrorAction Stop
                    $cleanedCount++
                }
                catch {
                    # ロック中ファイルは無視して継続
                }
            }
        }
    }

    Show-Success "$Browser data cleaned ($cleanedCount items)"
    return "Success"
}

# ----------------------------------------
# (6) Windows Search インデックス再構築
# ----------------------------------------
function Clear-SearchIndex {
    $wsearchService = Get-Service -Name "WSearch" -ErrorAction SilentlyContinue
    if (-not $wsearchService) {
        Show-Skip "Windows Search service not found"
        return "Skip"
    }

    $searchDbPath = "$env:ProgramData\Microsoft\Search\Data\Applications\Windows\Windows.edb"

    try {
        # サービス停止
        if ($wsearchService.Status -eq "Running") {
            Stop-Service -Name "WSearch" -Force -ErrorAction Stop
            Start-Sleep -Seconds 2
        }

        # インデックスDB 削除
        if (Test-Path $searchDbPath) {
            Remove-Item $searchDbPath -Force -ErrorAction Stop
            Show-Success "Search index deleted"
        }
        else {
            Show-Info "Search index file not found (already clean)"
        }
    }
    catch {
        # エラー発生時もサービス再起動を試行してから再 throw
        $null = Start-Service -Name "WSearch" -ErrorAction SilentlyContinue
        throw
    }
    finally {
        # 正常・異常問わずサービス再起動を保証
        $null = Start-Service -Name "WSearch" -ErrorAction SilentlyContinue
    }

    return "Success"
}

# ----------------------------------------
# (7) キッティング用 Wi-Fi プロファイル削除 (ssid_list.csv)
# ----------------------------------------
function Clear-WiFiProfiles {
    $ssidCsvPath = Join-Path $PSScriptRoot "ssid_list.csv"

    if (-not (Test-Path $ssidCsvPath)) {
        Show-Skip "ssid_list.csv not found"
        return "Skip"
    }

    $ssidItems = Import-ModuleCsv -Path $ssidCsvPath -FilterEnabled
    if ($null -eq $ssidItems -or $ssidItems.Count -eq 0) {
        # Import-ModuleCsv -FilterEnabled が既に Skip メッセージを出力済み
        return "Skip"
    }

    # Wi-Fi サービス確認
    $wlanSvc = Get-Service -Name "WlanSvc" -ErrorAction SilentlyContinue
    if (-not $wlanSvc -or $wlanSvc.Status -ne "Running") {
        Show-Skip "Wi-Fi service (WlanSvc) not available on this device"
        return "Skip"
    }

    $ssidDeleted = 0
    $ssidSkipped = 0
    $ssidErrors  = 0

    foreach ($ssidItem in $ssidItems) {
        $ssidName = $ssidItem.SSID
        $label = if ($ssidItem.Description) { "$ssidName ($($ssidItem.Description))" } else { $ssidName }

        # 冪等性: プロファイル存在確認
        $null = & netsh wlan show profile name="$ssidName" 2>&1
        if ($LASTEXITCODE -ne 0) {
            Show-Skip "Not found: $label"
            $ssidSkipped++
            continue
        }

        # 削除
        $null = & netsh wlan delete profile name="$ssidName" 2>&1
        if ($LASTEXITCODE -eq 0) {
            Show-Success "Deleted: $label"
            $ssidDeleted++
        }
        else {
            Show-Error "Failed to delete: $label"
            $ssidErrors++
        }
    }

    Show-Info "SSID cleanup: $ssidDeleted deleted, $ssidSkipped not found, $ssidErrors failed"

    if ($ssidErrors -gt 0) {
        throw "SSID cleanup had $ssidErrors failure(s)"
    }
    return "Success"
}


# ========================================
# Step 0: Explorer 停止（全操作の前提）
# ========================================
# Explorer を停止してファイルロックを解除する。
# メインループの外で手続き的に実行する。
# ========================================
Show-Warning "Explorer will be temporarily stopped during cleanup."
Write-Host "          The taskbar and desktop will disappear briefly." -ForegroundColor Red
Write-Host ""

try {
    Stop-Process -Name "explorer" -Force -ErrorAction Stop
    Show-Success "Explorer stopped"
}
catch {
    Show-Warning "Failed to stop Explorer: $($_.Exception.Message)"
    Write-Host "          Some locked files may not be deleted" -ForegroundColor Yellow
}
Write-Host ""


# ========================================
# Step 1: CSV 読み込み
# ========================================
$csvPath = Join-Path $PSScriptRoot "destroy_list.csv"

$enabledItems = Import-ModuleCsv -Path $csvPath -FilterEnabled `
    -RequiredColumns @("Enabled", "GroupName", "TargetName", "ActionType", "TargetPath")

if ($null -eq $enabledItems) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load destroy_list.csv")
}
if ($enabledItems.Count -eq 0) {
    return (New-ModuleResult -Status "Skipped" -Message "No enabled entries")
}

# DeletePath の環境変数を展開（file_delete パターン準拠）
foreach ($item in $enabledItems) {
    if ($item.ActionType -eq "DeletePath") {
        $item.TargetPath = [System.Environment]::ExpandEnvironmentVariables($item.TargetPath)
    }
}


# ========================================
# Step 2: 前提条件チェック（Early Return）
# ========================================
# fabriq は管理者権限で起動されることが前提のため、
# ここでは追加の前提条件チェックは省略する。


# ========================================
# Step 3: 実行前の確認表示（ドライラン）
# ========================================
Show-Info "Cleanup targets: $($enabledItems.Count) items"
Write-Host ""

Write-Host "========================================" -ForegroundColor Yellow
Write-Host "Destruction Targets" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

# GroupName でグルーピングして表示
$groups = $enabledItems | Group-Object -Property GroupName

foreach ($group in $groups) {
    Write-Host "  [$($group.Name)]" -ForegroundColor White
    foreach ($item in $group.Group) {
        $displayName = if ($item.Description) { $item.Description } else { $item.TargetName }
        Write-Host "    [DESTROY] $displayName" -ForegroundColor Yellow
        Write-Host "      $($item.ActionType): $($item.TargetPath)" -ForegroundColor DarkGray
    }
    Write-Host ""
}

Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""


# ========================================
# Step 4: 実行確認
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Proceed with history destruction?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""


# ========================================
# Step 5: メイン処理ループ
# ========================================
$successCount = 0
$skipCount    = 0
$failCount    = 0
$total        = $enabledItems.Count
$current      = 0

foreach ($item in $enabledItems) {
    $current++
    $displayName = if ($item.Description) { $item.Description } else { $item.TargetName }
    $ifNotFound  = if ($item.IfNotFound) { $item.IfNotFound } else { "Skip" }

    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "[$current/$total] $displayName" -ForegroundColor Cyan
    Write-Host "  $($item.ActionType): $($item.TargetPath)" -ForegroundColor DarkGray
    Write-Host "----------------------------------------" -ForegroundColor White

    try {
        switch ($item.ActionType) {

            "DeletePath" {
                # ベストエフォート削除: 存在確認 → 最大限削除 → 残存チェック
                if (-not (Test-Path $item.TargetPath)) {
                    if ($ifNotFound -eq "Error") {
                        Show-Error "Target not found: $($item.TargetPath)"
                        $failCount++
                    }
                    else {
                        Show-Skip "Not found - skipped"
                        $skipCount++
                    }
                    Write-Host ""
                    continue
                }

                # SilentlyContinue: ロック中ファイルをスキップしつつ削除可能なものはすべて削除
                Remove-Item -Path $item.TargetPath -Force -Recurse -ErrorAction SilentlyContinue

                # 事後チェック: 残存ファイルの有無で Success / Warning を判定
                if (Test-Path $item.TargetPath) {
                    Show-Warning "Partially deleted: $displayName (some files in use)"
                }
                else {
                    Show-Success "Deleted: $displayName"
                }
                $successCount++
            }

            "ClearRegistry" {
                # レジストリキーの値クリア（キー構造は保持、値のみ削除）
                if (-not (Test-Path $item.TargetPath)) {
                    if ($ifNotFound -eq "Error") {
                        Show-Error "Registry key not found: $($item.TargetPath)"
                        $failCount++
                    }
                    else {
                        Show-Skip "Registry key not found - skipped"
                        $skipCount++
                    }
                    Write-Host ""
                    continue
                }

                # 直下のプロパティ（値）をクリア
                $null = Remove-ItemProperty -Path $item.TargetPath -Name * -Force -ErrorAction SilentlyContinue

                # サブキーが存在する場合、各サブキーの値もクリア（キー構造は保持）
                $subKeys = Get-ChildItem -Path $item.TargetPath -ErrorAction SilentlyContinue
                foreach ($subKey in $subKeys) {
                    $null = Remove-ItemProperty -Path $subKey.PSPath -Name * -Force -ErrorAction SilentlyContinue
                }

                Show-Success "Registry cleared: $displayName"
                $successCount++
            }

            "Command" {
                # 単純ワンライナーコマンドの実行
                $null = Invoke-Expression $item.TargetPath
                Show-Success "Command executed: $displayName"
                $successCount++
            }

            "Special" {
                # ディスパッチャ経由でハンドラを呼び出し
                # 戻り値: "Success" → $successCount / "Skip" → $skipCount / throw → catch で $failCount
                $handlerResult = Invoke-DestroyHandler -HandlerName $item.TargetPath
                if ($handlerResult -eq "Skip") {
                    $skipCount++
                }
                else {
                    $successCount++
                }
            }

            default {
                Show-Error "Unknown ActionType: $($item.ActionType)"
                $failCount++
            }
        }
    }
    catch {
        Show-Error "Failed: $displayName : $_"
        $failCount++
    }

    Write-Host ""
}


# ========================================
# 最終Step: Explorer 再起動
# ========================================
Show-Info "Restarting Explorer..."

$maxWait = 15; $interval = 1; $elapsed = 0; $restarted = $false
while ($elapsed -lt $maxWait) {
    Start-Sleep -Seconds $interval
    $elapsed += $interval
    if (@(Get-Process -Name "explorer" -ErrorAction SilentlyContinue).Count -gt 0) {
        $restarted = $true; break
    }
}
if ($restarted) {
    Show-Success "Explorer restarted (${elapsed}s)"
}
else {
    # Windows の自動再起動が間に合わなかった場合のみ明示的に起動
    Start-Process "explorer.exe"
    Show-Warning "Explorer auto-restart timed out. Started manually."
}
Write-Host ""


# ========================================
# Step 6: 結果集計・返却
# ========================================
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount `
    -Title "History Destroyer Results")
