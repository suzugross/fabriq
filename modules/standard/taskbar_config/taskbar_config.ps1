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
# Step 1: Load CSV
# ========================================
$csvPath = Join-Path $PSScriptRoot "taskbar_list.csv"

$allItems = Import-ModuleCsv -Path $csvPath `
    -RequiredColumns @("Enabled", "Order", "LinkPath", "Description")

if ($null -eq $allItems) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load taskbar_list.csv")
}

# Filter Enabled + sort by Order ascending (zero rows means "unpin everything")
$items = @($allItems | Where-Object { $_.Enabled -eq "1" } | Sort-Object { [int]$_.Order })

# ========================================
# Step 2: Prerequisite check (deploy-target directory)
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
# Step 3: Dry-run summary before execution
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
    $expandedPath = Expand-UserEnvironmentVariables $item.LinkPath

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

# Show the state of any existing XML
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
# Step 4: User confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Generate and deploy LayoutModification.xml?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""

# ========================================
# Step 5: Generate XML and deploy
# ========================================

# 5-1: Build the DesktopApp entries
$pinEntries = ""
foreach ($item in $items) {
    $pinEntries += "      <taskbar:DesktopApp DesktopApplicationLinkPath=`"$($item.LinkPath)`"/>`r`n"
}

# 5-2: Build the full XML
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

# 5-3: Write into the Default User profile
$deployVerified = $null
try {
    $xmlContent | Out-File -FilePath $deployPath -Encoding UTF8 -Force -ErrorAction Stop

    # Post-write verification
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

    # Step 5.5: Post-apply verification (content-level)
    # Read the file back and confirm BOTH: (1) the entry count matches (no truncation / no extras),
    # and (2) every requested LinkPath value is actually present. Substring/.Contains is robust to
    # BOM / trailing newline that an exact-string compare would trip on, and ordinal .Contains avoids
    # wildcard/regex interpretation of path characters.
    $deployedRaw = Get-Content -Path $deployPath -Raw -ErrorAction Stop
    $pinCount = ([regex]::Matches($deployedRaw, 'DesktopApplicationLinkPath')).Count
    $allPathsPresent = $true
    foreach ($it in $items) {
        if (-not $deployedRaw.Contains([string]$it.LinkPath)) { $allPathsPresent = $false; break }
    }
    if (($pinCount -eq $items.Count) -and $allPathsPresent) {
        $deployVerified = $true
    }
    else {
        Show-Warning "Verification mismatch: deployed XML pins=$pinCount/$($items.Count), all paths present=$allPathsPresent"
        $deployVerified = $false
    }
}
catch {
    Show-Error "Failed to write XML: $_"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Failed to write XML: $_")
}

# 5-4: Copy into sysprep_config/source/
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
# Step 6: Return result
# ========================================
return (New-ModuleResult -Status "Success" -Message "LayoutModification.xml deployed ($($items.Count) apps)" -Verified $deployVerified)
