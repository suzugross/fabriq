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

# Resolve each row to exactly one pin type: LinkPath (DesktopApplicationLinkPath)
# or AppId (DesktopApplicationID). The AppId column is optional; when the column
# is absent every row is a LinkPath row (pre-1.2.0 CSVs keep working unchanged).
# A row with both or neither is ambiguous -> fail closed.
$invalidRows = @()
foreach ($item in $items) {
    $linkValue = ([string]$item.LinkPath).Trim()
    $appIdValue = ""
    if ($null -ne $item.PSObject.Properties['AppId']) {
        $appIdValue = ([string]$item.AppId).Trim()
    }

    if ($linkValue -ne "" -and $appIdValue -eq "") {
        Add-Member -InputObject $item -NotePropertyName PinType -NotePropertyValue "Link"
        Add-Member -InputObject $item -NotePropertyName PinValue -NotePropertyValue $linkValue
    }
    elseif ($linkValue -eq "" -and $appIdValue -ne "") {
        Add-Member -InputObject $item -NotePropertyName PinType -NotePropertyValue "AppId"
        Add-Member -InputObject $item -NotePropertyName PinValue -NotePropertyValue $appIdValue
    }
    else {
        $invalidRows += "Order=$($item.Order) ($($item.Description))"
    }
}
if ($invalidRows.Count -gt 0) {
    Show-Error "Each enabled row must have exactly one of LinkPath / AppId. Invalid rows: $($invalidRows -join ', ')"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Invalid taskbar_list.csv rows: $($invalidRows -join ', ')")
}

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

    if ($item.PinType -eq "AppId") {
        # An application ID cannot be existence-checked; Windows silently
        # skips pins whose app is not installed at first logon.
        Write-Host "  [$index] $($item.Description)  [PIN:APPID]" -ForegroundColor White
        Write-Host "      AppId: $($item.PinValue) (no existence check)" -ForegroundColor DarkGray
        Write-Host ""
        continue
    }

    $expandedPath = Expand-UserEnvironmentVariables $item.PinValue

    if (Test-Path $expandedPath) {
        $marker = "[PIN]"
        $markerColor = "White"
    }
    else {
        $marker = "[NOT FOUND]"
        $markerColor = "Yellow"
    }

    Write-Host "  [$index] $($item.Description)  $marker" -ForegroundColor $markerColor
    Write-Host "      LinkPath: $($item.PinValue)" -ForegroundColor DarkGray
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
# Attribute values are XML-escaped (same pattern as sysprep_config): a path
# containing & would otherwise produce invalid XML that Windows silently ignores.
$pinEntries = ""
foreach ($item in $items) {
    $escapedValue = [System.Security.SecurityElement]::Escape($item.PinValue)
    if ($item.PinType -eq "AppId") {
        $pinEntries += "      <taskbar:DesktopApp DesktopApplicationID=`"$escapedValue`"/>`r`n"
    }
    else {
        $pinEntries += "      <taskbar:DesktopApp DesktopApplicationLinkPath=`"$escapedValue`"/>`r`n"
    }
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
    # and (2) every requested pin value (LinkPath or AppId, XML-escaped form) is actually present.
    # Substring/.Contains is robust to BOM / trailing newline that an exact-string compare would
    # trip on, and ordinal .Contains avoids wildcard/regex interpretation of path characters.
    $deployedRaw = Get-Content -Path $deployPath -Raw -ErrorAction Stop
    $pinCount = ([regex]::Matches($deployedRaw, '<taskbar:DesktopApp ')).Count
    $allValuesPresent = $true
    foreach ($it in $items) {
        $escapedCheck = [System.Security.SecurityElement]::Escape([string]$it.PinValue)
        if (-not $deployedRaw.Contains($escapedCheck)) { $allValuesPresent = $false; break }
    }
    if (($pinCount -eq $items.Count) -and $allValuesPresent) {
        $deployVerified = $true
    }
    else {
        Show-Warning "Verification mismatch: deployed XML pins=$pinCount/$($items.Count), all values present=$allValuesPresent"
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
