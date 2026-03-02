# ========================================
# Default App Config Script
# ========================================
# エクスポート済みの XML を使用して既定のアプリ関連付けをインポートする。
#
# [NOTES]
# - 管理者権限が必要
# - DISM /Import-DefaultAppAssociations を使用
# - インポートした関連付けは新規ユーザープロファイル作成時に適用される
# ========================================

Write-Host ""
Show-Separator
Write-Host "Default App Config" -ForegroundColor Cyan
Show-Separator
Write-Host ""


# ========================================
# Step 1: CSV 読み込み
# ========================================
$csvPath = Join-Path $PSScriptRoot "default_app_list.csv"

$enabledItems = Import-ModuleCsv -Path $csvPath -FilterEnabled `
    -RequiredColumns @("Enabled", "XmlFile", "Description")

if ($null -eq $enabledItems) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load default_app_list.csv")
}
if ($enabledItems.Count -eq 0) {
    return (New-ModuleResult -Status "Skipped" -Message "No enabled entries")
}


# ========================================
# Step 2: 前提条件チェック（Early Return）
# ========================================
$xmlDir = Join-Path $PSScriptRoot "xml"
if (-not (Test-Path $xmlDir)) {
    Show-Error "XML directory not found: $xmlDir"
    Show-Error "Run Export App Associations first to create XML files."
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "XML directory not found")
}


# ========================================
# Step 3: 実行前の確認表示（ドライラン）
# ========================================
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "Import Targets" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

$hasValidTarget = $false

foreach ($item in $enabledItems) {
    $displayName = if ($item.Description) { $item.Description } else { $item.XmlFile }
    $xmlPath = Join-Path $xmlDir $item.XmlFile

    if (Test-Path $xmlPath) {
        $fileSize = (Get-Item $xmlPath).Length
        Write-Host "  [IMPORT] $displayName" -ForegroundColor Yellow
        Write-Host "    File: $xmlPath ($fileSize bytes)" -ForegroundColor DarkGray
        $hasValidTarget = $true
    }
    else {
        Write-Host "  [NOT FOUND] $displayName" -ForegroundColor DarkGray
        Write-Host "    File: $xmlPath" -ForegroundColor DarkGray
    }
    Write-Host ""
}

Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

if (-not $hasValidTarget) {
    Show-Skip "No valid XML files found to import"
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "No valid XML files found")
}


# ========================================
# Step 4: 実行確認
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Import the above app associations?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""


# ========================================
# Step 5: 設定適用ループ
# ========================================
$successCount = 0
$skipCount    = 0
$failCount    = 0

foreach ($item in $enabledItems) {
    $displayName = if ($item.Description) { $item.Description } else { $item.XmlFile }
    $xmlPath = Join-Path $xmlDir $item.XmlFile

    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "Importing: $displayName" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor White

    # Skip 判定: XML ファイルが存在しない
    if (-not (Test-Path $xmlPath)) {
        Show-Skip "File not found: $xmlPath"
        Write-Host ""
        $skipCount++
        continue
    }

    try {
        Show-Info "Running Dism /Import-DefaultAppAssociations..."
        $dismResult = & Dism.exe /Online /Import-DefaultAppAssociations:"$xmlPath" 2>&1
        $exitCode = $LASTEXITCODE

        if ($exitCode -eq 0) {
            Show-Success "Imported: $displayName"
            $successCount++
        }
        else {
            Show-Error "Dism exited with code $exitCode"
            $dismOutput = $dismResult | Out-String
            if (-not [string]::IsNullOrWhiteSpace($dismOutput)) {
                Write-Host $dismOutput -ForegroundColor DarkGray
            }
            $failCount++
        }
    }
    catch {
        Show-Error "Failed: $displayName : $_"
        $failCount++
    }

    Write-Host ""
}


# ========================================
# Step 6: 結果集計・返却
# ========================================
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount `
    -Title "Default App Config Results")
