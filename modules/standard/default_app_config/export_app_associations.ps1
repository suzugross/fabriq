# ========================================
# Export App Associations Script
# ========================================
# 現在のユーザーの既定のアプリ関連付けを XML にエクスポートする。
#
# [NOTES]
# - 管理者権限が必要
# - マスター PC での準備作業用（本番キッティングでは使用しない）
# - DISM /Export-DefaultAppAssociations を使用
# ========================================

Write-Host ""
Show-Separator
Write-Host "Export App Associations" -ForegroundColor Cyan
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
    try {
        New-Item -Path $xmlDir -ItemType Directory -Force | Out-Null
        Show-Info "Created xml directory: $xmlDir"
    }
    catch {
        Show-Error "Failed to create xml directory: $_"
        Write-Host ""
        return (New-ModuleResult -Status "Error" -Message "Failed to create xml directory")
    }
}


# ========================================
# Step 3: 実行前の確認表示（ドライラン）
# ========================================
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "Export Targets" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

foreach ($item in $enabledItems) {
    $displayName = if ($item.Description) { $item.Description } else { $item.XmlFile }
    $xmlPath = Join-Path $xmlDir $item.XmlFile

    if (Test-Path $xmlPath) {
        Write-Host "  [OVERWRITE] $displayName" -ForegroundColor Red
        Show-Warning "Existing file will be overwritten: $($item.XmlFile)"
    }
    else {
        Write-Host "  [EXPORT] $displayName" -ForegroundColor Yellow
    }
    Write-Host "    File: $xmlPath" -ForegroundColor DarkGray
    Write-Host ""
}

Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""


# ========================================
# Step 4: 実行確認
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Export app associations to the above files?"
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
    Write-Host "Exporting: $displayName" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor White

    try {
        Show-Info "Running Dism /Export-DefaultAppAssociations..."
        $dismResult = & Dism.exe /Online /Export-DefaultAppAssociations:"$xmlPath" 2>&1
        $exitCode = $LASTEXITCODE

        if ($exitCode -eq 0) {
            # ファイルが実際に生成されたか確認
            if (Test-Path $xmlPath) {
                $fileSize = (Get-Item $xmlPath).Length
                Show-Success "Exported: $xmlPath ($fileSize bytes)"
            }
            else {
                Show-Success "Exported: $xmlPath"
            }
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
    -Title "Export App Associations Results")
