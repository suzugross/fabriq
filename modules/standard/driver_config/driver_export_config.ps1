# ========================================
# Driver Export Script
# ========================================
# PCのサードパーティドライバを driver/{モデル名}/ にエクスポートする。
#
# [NOTES]
# - 管理者権限が必要
# - Export-WindowsDriver はサードパーティドライバのみエクスポートする
# - 既存フォルダは中身をクリアして再エクスポートする
# ========================================

Write-Host ""
Show-Separator
Write-Host "Driver Export" -ForegroundColor Cyan
Show-Separator
Write-Host ""


# ========================================
# Step 1: CSV 読み込み
# ========================================
$csvPath = Join-Path $PSScriptRoot "driver.csv"

$enabledItems = Import-ModuleCsv -Path $csvPath -FilterEnabled `
    -RequiredColumns @("Enabled", "Id", "model")

if ($null -eq $enabledItems) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load driver.csv")
}
if ($enabledItems.Count -eq 0) {
    return (New-ModuleResult -Status "Skipped" -Message "No enabled entries")
}


# ========================================
# Step 2: 前提条件チェック（Early Return）
# ========================================
$driverDir = Join-Path $PSScriptRoot "driver"
if (-not (Test-Path $driverDir)) {
    try {
        New-Item -Path $driverDir -ItemType Directory -Force | Out-Null
        Show-Info "Created driver directory: $driverDir"
    }
    catch {
        Show-Error "Failed to create driver directory: $_"
        Write-Host ""
        return (New-ModuleResult -Status "Error" -Message "Failed to create driver directory")
    }
}


# ========================================
# Step 3: 実行前の確認表示（ドライラン）
# ========================================
# モデル名の動的取得
$systemModel = ""
try {
    $systemModel = Get-CimInstance -ClassName Win32_ComputerSystem |
        Select-Object -ExpandProperty Model
}
catch {
    Show-Warning "Failed to get system model: $_"
}

# モデル名サニタイズ関数
function Get-SafeModelName {
    param([string]$RawName)
    $safeName = $RawName -replace '\s', '_'
    $safeName = $safeName -replace '[\\/:*?"<>|]', ''
    $safeName = $safeName.Trim('_').Trim('.')
    if ($safeName.Length -gt 80) { $safeName = $safeName.Substring(0, 80) }
    if ([string]::IsNullOrWhiteSpace($safeName)) { $safeName = "Unknown_Model" }
    return $safeName
}

Write-Host "========================================" -ForegroundColor Yellow
Write-Host "Export Targets" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

foreach ($item in $enabledItems) {
    # モデル名の解決
    $modelName = if (-not [string]::IsNullOrWhiteSpace($item.model)) {
        $item.model
    }
    else {
        Get-SafeModelName -RawName $systemModel
    }
    $destPath = Join-Path $driverDir $modelName

    # 既存フォルダの有無で表示を切り替え
    if (Test-Path $destPath) {
        Write-Host "  [OVERWRITE] Id=$($item.Id) : $modelName" -ForegroundColor Red
    }
    else {
        Write-Host "  [EXPORT] Id=$($item.Id) : $modelName" -ForegroundColor Yellow
    }
    Write-Host "    Path: $destPath" -ForegroundColor DarkGray

    # モデル名の出所を表示
    if (-not [string]::IsNullOrWhiteSpace($item.model)) {
        Write-Host "    Source: CSV specified" -ForegroundColor DarkGray
    }
    else {
        Write-Host "    Source: Auto-detected ($systemModel)" -ForegroundColor DarkGray
    }
    Write-Host ""
}

Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""


# ========================================
# Step 4: 実行確認
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Export drivers to the above paths?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""


# ========================================
# Step 5: 設定適用ループ
# ========================================
$successCount = 0
$skipCount    = 0
$failCount    = 0

foreach ($item in $enabledItems) {
    # モデル名の解決
    $modelName = if (-not [string]::IsNullOrWhiteSpace($item.model)) {
        $item.model
    }
    else {
        Get-SafeModelName -RawName $systemModel
    }
    $destPath = Join-Path $driverDir $modelName

    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "Exporting: Id=$($item.Id) - $modelName" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor White

    try {
        # 既存フォルダのクリア
        if (Test-Path $destPath) {
            Show-Info "Clearing existing folder: $destPath"
            Remove-Item -Path $destPath -Recurse -Force
        }

        # フォルダ作成
        New-Item -Path $destPath -ItemType Directory -Force | Out-Null

        # ドライバエクスポート
        Show-Info "Running Export-WindowsDriver..."
        $null = Export-WindowsDriver -Online -Destination $destPath

        # エクスポート結果の確認
        $infCount = @(Get-ChildItem -Path $destPath -Filter "*.inf" -Recurse -File).Count
        Show-Success "Exported to: $destPath ($infCount .inf files)"
        $successCount++
    }
    catch {
        Show-Error "Failed: Id=$($item.Id) - $modelName : $_"
        $failCount++
    }

    Write-Host ""
}


# ========================================
# Step 6: 結果集計・返却
# ========================================
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount `
    -Title "Driver Export Results")
