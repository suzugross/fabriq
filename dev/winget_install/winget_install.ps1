# ========================================
# Winget Batch Installer
# ========================================

Write-Host "Initializing Winget Installation..." -ForegroundColor Cyan
Write-Host ""

# ----------------------------------------
# 1. Internet Connection Check
# ----------------------------------------
Write-Host "Checking internet connection..." -ForegroundColor White
if (Test-Connection -ComputerName "8.8.8.8" -Count 1 -Quiet) {
    Write-Host "[SUCCESS] Internet Connection OK" -ForegroundColor Green
}
else {
    Write-Host "[ERROR] No Internet Connection (Ping 8.8.8.8 failed)" -ForegroundColor Red
    return (New-ModuleResult -Status "Error" -Message "No Internet Connection")
}

# ----------------------------------------
# 2. Check Winget Availability
# ----------------------------------------
if (-not (Get-Command "winget" -ErrorAction SilentlyContinue)) {
    Write-Host "[ERROR] 'winget' command not found. Please update App Installer." -ForegroundColor Red
    return (New-ModuleResult -Status "Error" -Message "winget command not found")
}

# ----------------------------------------
# 3. Load CSV
# ----------------------------------------
$csvPath = Join-Path $PSScriptRoot "app_list.csv"

if (-not (Test-Path $csvPath)) {
    Write-Host "[ERROR] app_list.csv not found: $csvPath" -ForegroundColor Red
    return (New-ModuleResult -Status "Error" -Message "app_list.csv not found")
}

try {
    $appList = @(Import-Csv -Path $csvPath -Encoding Default)
}
catch {
    Write-Host "[ERROR] Failed to load CSV: $_" -ForegroundColor Red
    return (New-ModuleResult -Status "Error" -Message "Failed to load CSV")
}

if ($appList.Count -eq 0) {
    Write-Host "[WARN] CSV contains no data" -ForegroundColor Yellow
    return (New-ModuleResult -Status "Skipped" -Message "No apps listed")
}

# ----------------------------------------
# 4. Installation Loop
# ----------------------------------------
$successCount = 0
$failCount = 0
$skipCount = 0

foreach ($app in $appList) {
    $appName = if ($app.Description) { $app.Description } else { $app.AppID }

    # --- Check Enabled Flag ---
    if ($app.Enabled -ne "1") {
        Write-Host "----------------------------------------" -ForegroundColor White
        Write-Host "Target: $appName" -ForegroundColor Cyan
        Write-Host "  [SKIP] Disabled in CSV" -ForegroundColor DarkGray
        $skipCount++
        continue
    }

    if ([string]::IsNullOrWhiteSpace($app.AppID)) { 
        continue 
    }

    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "Installing: $appName ($($app.AppID))" -ForegroundColor Cyan

    # Build Command
    # --id: ID指定
    # --exact: 完全一致
    # --silent: サイレントインストール
    # --accept-source-agreements --accept-package-agreements: 同意画面スキップ
    # --force: 念のため強制フラグ（任意）
    $wingetArgs = "install --id `"$($app.AppID)`" --exact --silent --accept-source-agreements --accept-package-agreements"
    
    if (-not [string]::IsNullOrWhiteSpace($app.Options)) {
        $wingetArgs += " $($app.Options)"
    }

    try {
        # Start-Processで実行し、終了を待機する
        $process = Start-Process -FilePath "winget" -ArgumentList $wingetArgs -Wait -NoNewWindow -PassThru
        
        if ($process.ExitCode -eq 0) {
            Write-Host "[SUCCESS] Installation completed." -ForegroundColor Green
            $successCount++
        }
        # エラーコードが出ても成功している場合があるため（再起動要求など）、
        # 必要に応じて特定のエラーコード（例: 3010=再起動保留）は成功扱いにする処理を入れても良い
        else {
            Write-Host "[ERROR] Installation failed. ExitCode: $($process.ExitCode)" -ForegroundColor Red
            $failCount++
        }
    }
    catch {
        Write-Host "[ERROR] Execution error: $_" -ForegroundColor Red
        $failCount++
    }
    
    Write-Host ""
}

# ----------------------------------------
# 5. Result Summary
# ----------------------------------------
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Execution Results" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Success: $successCount items" -ForegroundColor Green
Write-Host "  Skipped: $skipCount items" -ForegroundColor Yellow
Write-Host "  Failed:  $failCount items" -ForegroundColor $(if ($failCount -gt 0) { "Red" } else { "Green" })
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Determine Module Status
$overallStatus = if ($failCount -eq 0 -and $successCount -gt 0) { "Success" }
    elseif ($successCount -gt 0 -and $failCount -gt 0) { "Partial" }
    elseif ($failCount -eq 0 -and $skipCount -gt 0) { "Skipped" }
    elseif ($failCount -gt 0) { "Error" }
    else { "Skipped" }

return (New-ModuleResult -Status $overallStatus -Message "Success: $successCount, Skip: $skipCount, Fail: $failCount")