# ========================================
# Taskbar Config - Taskbar Pin Layout Generator
# ========================================
# Generates LayoutModification.xml from taskbar_list.csv
# and deploys it to the Default User profile.
# Also copies the generated XML to sysprep_config/source/
# for use with Sysprep-based kitting workflows.
#
# [NOTES]
# - Requires administrator privileges (writes to Default User profile)
# - Does not affect existing user profiles
# ========================================

Write-Host ""
Show-Separator
Write-Host "Taskbar Config" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# ========================================
# Step 1: CSV 読み込み
# ========================================
$csvPath = Join-Path $PSScriptRoot "taskbar_list.csv"

$allItems = Import-ModuleCsv -Path $csvPath `
    -RequiredColumns @("Enabled", "Order", "LinkPath", "Description")

if ($null -eq $allItems) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load taskbar_list.csv")
}

# Enabled フィルタ + Order 昇順ソート（0件 = 全ピン解除として続行）
$items = @($allItems | Where-Object { $_.Enabled -eq "1" } | Sort-Object { [int]$_.Order })

# ========================================
# Step 2: 前提条件チェック（Deploy先ディレクトリ）
# ========================================
$deployDir = "C:\Users\Default\AppData\Local\Microsoft\Windows\Shell"
$deployPath = Join-Path $deployDir "LayoutModification.xml"

$sysprepSourceDir  = Join-Path $PSScriptRoot "..\sysprep_config\source"
$sysprepSourcePath = Join-Path $sysprepSourceDir "LayoutModification.xml"

if (-not (Test-Path $deployDir)) {
    try {
        $null = New-Item -ItemType Directory -Path $deployDir -Force -ErrorAction Stop
        Show-Info "Created deploy directory: $deployDir"
    }
    catch {
        Show-Error "Failed to create deploy directory: $deployDir - $_"
        Write-Host ""
        return (New-ModuleResult -Status "Error" -Message "Deploy directory creation failed")
    }
}

# ========================================
# Step 3: 実行前の確認表示
# ========================================
if ($items.Count -eq 0) {
    Show-Info "Taskbar pin targets: 0 apps (all default pins will be removed)"
}
else {
    Show-Info "Taskbar pin targets: $($items.Count) apps"
}
Write-Host ""

$index = 0
foreach ($item in $items) {
    $index++
    $expandedPath = [System.Environment]::ExpandEnvironmentVariables($item.LinkPath)

    if (Test-Path $expandedPath) {
        $marker = "[PIN]"
        $markerColor = "White"
    }
    else {
        $marker = "[NOT FOUND]"
        $markerColor = "Yellow"
    }

    Write-Host "  [$index] $($item.Description)  $marker" -ForegroundColor $markerColor
    Write-Host "      LinkPath: $($item.LinkPath)" -ForegroundColor DarkGray
    if ($marker -eq "[NOT FOUND]") {
        Show-Warning "Shortcut not found at: $expandedPath"
    }
    Write-Host ""
}

# 既存 XML の状態表示
if (Test-Path $deployPath) {
    Write-Host "  Deploy: $deployPath  [OVERWRITE]" -ForegroundColor Yellow
}
else {
    Write-Host "  Deploy: $deployPath  [NEW]" -ForegroundColor White
}

if (Test-Path $sysprepSourcePath) {
    Write-Host "  Copy:   $sysprepSourcePath  [OVERWRITE]" -ForegroundColor Yellow
}
else {
    Write-Host "  Copy:   $sysprepSourcePath  [NEW]" -ForegroundColor White
}
Write-Host ""

# ========================================
# Step 4: 実行確認
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Generate and deploy LayoutModification.xml?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""

# ========================================
# Step 5: XML 生成 & デプロイ
# ========================================

# 5-1: DesktopApp エントリを組み立て
$pinEntries = ""
foreach ($item in $items) {
    $pinEntries += "      <taskbar:DesktopApp DesktopApplicationLinkPath=`"$($item.LinkPath)`"/>`r`n"
}

# 5-2: XML 全体を組み立て
$xmlContent = @"
<?xml version="1.0" encoding="utf-8"?>
<LayoutModificationTemplate
    xmlns="http://schemas.microsoft.com/Start/2014/LayoutModification"
    xmlns:defaultlayout="http://schemas.microsoft.com/Start/2014/FullDefaultLayout"
    xmlns:start="http://schemas.microsoft.com/Start/2014/StartLayout"
    xmlns:taskbar="http://schemas.microsoft.com/Start/2014/TaskbarLayout"
    Version="1">
  <CustomTaskbarLayoutCollection PinListPlacement="Replace">
    <defaultlayout:TaskbarLayout>
      <taskbar:TaskbarPinList>
$pinEntries      </taskbar:TaskbarPinList>
    </defaultlayout:TaskbarLayout>
  </CustomTaskbarLayoutCollection>
</LayoutModificationTemplate>
"@

# 5-3: Default User プロファイルへ書き出し
try {
    $xmlContent | Out-File -FilePath $deployPath -Encoding UTF8 -Force -ErrorAction Stop

    # 書き出し後の検証
    if (-not (Test-Path $deployPath)) {
        Show-Error "XML file was not created: $deployPath"
        Write-Host ""
        return (New-ModuleResult -Status "Error" -Message "XML file not found after write")
    }

    $fileSize = (Get-Item $deployPath).Length
    if ($fileSize -eq 0) {
        Show-Error "XML file is empty: $deployPath"
        Write-Host ""
        return (New-ModuleResult -Status "Error" -Message "XML file is empty after write")
    }

    Show-Success "LayoutModification.xml deployed ($($items.Count) apps, $fileSize bytes)"
    Write-Host "  Path: $deployPath" -ForegroundColor DarkGray
    Write-Host ""
}
catch {
    Show-Error "Failed to write XML: $_"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Failed to write XML: $_")
}

# 5-4: sysprep_config/source/ へコピー
try {
    if (-not (Test-Path $sysprepSourceDir)) {
        $null = New-Item -ItemType Directory -Path $sysprepSourceDir -Force -ErrorAction Stop
        Show-Info "Created directory: $sysprepSourceDir"
    }

    $null = Copy-Item -Path $deployPath -Destination $sysprepSourcePath -Force -ErrorAction Stop

    Show-Success "Copied to sysprep_config/source/"
    Write-Host "  Path: $sysprepSourcePath" -ForegroundColor DarkGray
    Write-Host ""
}
catch {
    Show-Warning "Failed to copy to sysprep_config/source/: $_"
    Write-Host ""
}

# ========================================
# Step 6: 結果返却
# ========================================
return (New-ModuleResult -Status "Success" -Message "LayoutModification.xml deployed ($($items.Count) apps)")
