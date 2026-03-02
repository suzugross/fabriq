# ========================================
# Driver Import Script
# ========================================
# driver/{モデル名}/ のドライバを pnputil でインポート（インストール）する。
#
# [NOTES]
# - 管理者権限が必要
# - pnputil.exe /add-driver を使用
# - 終了コード 3010 は再起動要求（成功扱い）
# ========================================

Write-Host ""
Show-Separator
Write-Host "Driver Import" -ForegroundColor Cyan
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
    Show-Error "Driver directory not found: $driverDir"
    Show-Error "Run Driver Export first to create driver backups."
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Driver directory not found")
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
Write-Host "Import Targets" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

$hasValidTarget = $false

foreach ($item in $enabledItems) {
    # モデル名の解決
    $modelName = if (-not [string]::IsNullOrWhiteSpace($item.model)) {
        $item.model
    }
    else {
        Get-SafeModelName -RawName $systemModel
    }
    $sourcePath = Join-Path $driverDir $modelName

    if (-not (Test-Path $sourcePath)) {
        Write-Host "  [NOT FOUND] Id=$($item.Id) : $modelName" -ForegroundColor DarkGray
        Write-Host "    Path: $sourcePath" -ForegroundColor DarkGray
    }
    else {
        $infCount = @(Get-ChildItem -Path $sourcePath -Filter "*.inf" -Recurse -File).Count
        if ($infCount -eq 0) {
            Write-Host "  [EMPTY] Id=$($item.Id) : $modelName" -ForegroundColor DarkGray
            Write-Host "    Path: $sourcePath (no .inf files)" -ForegroundColor DarkGray
        }
        else {
            Write-Host "  [IMPORT] Id=$($item.Id) : $modelName" -ForegroundColor Yellow
            Write-Host "    Path: $sourcePath ($infCount .inf files)" -ForegroundColor DarkGray
            $hasValidTarget = $true
        }
    }

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

if (-not $hasValidTarget) {
    Show-Skip "No valid driver folders found to import"
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "No valid driver folders found")
}


# ========================================
# Step 4: 実行確認
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Import drivers from the above paths?"
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
    $sourcePath = Join-Path $driverDir $modelName

    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "Importing: Id=$($item.Id) - $modelName" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor White

    # Skip 判定: フォルダが存在しない
    if (-not (Test-Path $sourcePath)) {
        Show-Skip "Folder not found: $sourcePath"
        Write-Host ""
        $skipCount++
        continue
    }

    # Skip 判定: .inf ファイルがない
    $infCount = @(Get-ChildItem -Path $sourcePath -Filter "*.inf" -Recurse -File).Count
    if ($infCount -eq 0) {
        Show-Skip "No .inf files in: $sourcePath"
        Write-Host ""
        $skipCount++
        continue
    }

    try {
        Show-Info "Running pnputil /add-driver ($infCount .inf files)..."
        $infPattern = Join-Path $sourcePath "*.inf"
        $pnputilResult = & pnputil.exe /add-driver $infPattern /subdirs /install 2>&1
        $exitCode = $LASTEXITCODE

        # 終了コード判定
        #   0    = 正常終了
        #   259  = 既に最新または追加対象なし (ERROR_NO_MORE_ITEMS)
        #   3010 = 要再起動 (成功扱い)
        #   その他 = エラー
        if ($exitCode -eq 0) {
            Show-Success "Imported: $modelName"
            $successCount++
        }
        elseif ($exitCode -eq 259) {
            Show-Skip "Already up to date: $modelName (exit code 259)"
            $skipCount++
        }
        elseif ($exitCode -eq 3010) {
            Show-Success "Imported: $modelName (restart required)"
            Show-Warning "A system restart is required to complete driver installation."
            $successCount++
        }
        else {
            Show-Error "pnputil exited with code $exitCode"
            $pnputilOutput = $pnputilResult | Out-String
            if (-not [string]::IsNullOrWhiteSpace($pnputilOutput)) {
                Write-Host $pnputilOutput -ForegroundColor DarkGray
            }
            $failCount++
        }
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
    -Title "Driver Import Results")
