# ========================================
# Process Killer Script
# ========================================
# process_list.csv に定義されたプロセスを強制終了する。
# 対象プロセスが起動していない場合は何もしない（冪等性維持）。
#
# [NOTES]
# - ProcessName は Get-Process -Name に渡す名前（.exe 不要）
# - 管理者権限がない場合、他ユーザーのプロセスは終了できないことがある
# ========================================

Write-Host ""
Show-Separator
Write-Host "Process Killer" -ForegroundColor Cyan
Show-Separator
Write-Host ""


# ========================================
# Step 1: CSV 読み込み
# ========================================
$csvPath = Join-Path $PSScriptRoot "process_list.csv"

$enabledItems = Import-ModuleCsv -Path $csvPath -FilterEnabled `
    -RequiredColumns @("Enabled", "ProcessName", "Description")

if ($null -eq $enabledItems) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load process_list.csv")
}
if ($enabledItems.Count -eq 0) {
    return (New-ModuleResult -Status "Skipped" -Message "No enabled entries")
}


# ========================================
# Step 2: 前提条件チェック
# ========================================
# プロセス終了に外部リソースは不要のため省略


# ========================================
# Step 3: 実行前の確認表示（ドライラン）
# ========================================
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "Target Processes" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

foreach ($item in $enabledItems) {
    $processes = @(Get-Process -Name $item.ProcessName -ErrorAction SilentlyContinue)

    if ($processes.Count -gt 0) {
        Write-Host "  [Running] $($item.Description)" -ForegroundColor Yellow
        Write-Host "    Process: $($item.ProcessName)  ($($processes.Count) instance(s))" -ForegroundColor DarkGray
    }
    else {
        Write-Host "  [Not Running] $($item.Description)" -ForegroundColor DarkGray
        Write-Host "    Process: $($item.ProcessName)" -ForegroundColor DarkGray
    }
    Write-Host ""
}

Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""


# ========================================
# Step 4: 実行確認
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Terminate the above running processes?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""


# ========================================
# Step 5: 設定適用ループ
# ========================================
$successCount = 0
$skipCount    = 0
$failCount    = 0

foreach ($item in $enabledItems) {
    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "Processing: $($item.Description)" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor White

    # ----------------------------------------
    # 冪等性チェック：プロセスが存在しなければ Skip
    # ----------------------------------------
    $processes = @(Get-Process -Name $item.ProcessName -ErrorAction SilentlyContinue)
    if ($processes.Count -eq 0) {
        Show-Skip "Already not running: $($item.ProcessName)"
        Write-Host ""
        $skipCount++
        continue
    }

    # ----------------------------------------
    # メイン処理：強制終了
    # ----------------------------------------
    try {
        Stop-Process -Name $item.ProcessName -Force -ErrorAction Stop
        Show-Success "Terminated: $($item.ProcessName)  ($($processes.Count) instance(s))"
        $successCount++
    }
    catch {
        Show-Error "Failed to terminate: $($item.ProcessName) : $_"
        $failCount++
    }

    Write-Host ""
}


# ========================================
# Step 6: 結果集計・返却
# ========================================
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount `
    -Title "Process Killer Results")
