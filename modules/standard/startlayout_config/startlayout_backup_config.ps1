# ========================================
# Start Layout Backup Script
# ========================================
# Exports the current start menu layout to a JSON file
# using the Export-StartLayout cmdlet (Windows 11).
# The exported JSON is stored under the json/ subdirectory
# and used as input for the build phase.
#
# [NOTES]
# - Requires Windows 11 (Export-StartLayout cmdlet)
# - Overwrites existing JSON files without prompt
# ========================================

Write-Host ""
Show-Separator
Write-Host "Start Layout Backup" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# ========================================
# Step 1: CSV reading
# ========================================
$csvPath = Join-Path $PSScriptRoot "startlayout_list.csv"

$enabledItems = Import-ModuleCsv -Path $csvPath -FilterEnabled `
    -RequiredColumns @("Enabled", "Id", "FileName")

if ($null -eq $enabledItems) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load startlayout_list.csv")
}
if ($enabledItems.Count -eq 0) {
    return (New-ModuleResult -Status "Skipped" -Message "No enabled entries")
}

# ========================================
# Step 2: Pre-flight check
# ========================================
if (-not (Get-Command "Export-StartLayout" -ErrorAction SilentlyContinue)) {
    Show-Error "Export-StartLayout cmdlet is not available on this system."
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Export-StartLayout cmdlet not found")
}

# ========================================
# Step 3: Pre-execution display
# ========================================
$jsonDir = Join-Path $PSScriptRoot "json"

Show-Info "Export targets: $($enabledItems.Count) item(s)"
Write-Host ""

foreach ($item in $enabledItems) {
    $outputPath = Join-Path $jsonDir "$($item.FileName).json"

    if (Test-Path $outputPath) {
        $marker = "[OVERWRITE]"
        $markerColor = "Yellow"
    }
    else {
        $marker = "[NEW]"
        $markerColor = "White"
    }

    Write-Host "  [Id:$($item.Id)] $($item.FileName).json  $marker" -ForegroundColor $markerColor
    Write-Host "    Output: $outputPath" -ForegroundColor DarkGray
    Write-Host ""
}

# ========================================
# Step 4: Confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Export current start layout?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""

# ========================================
# Step 5: Export execution
# ========================================
$successCount = 0
$skipCount    = 0
$failCount    = 0

# Create json/ directory if it does not exist
if (-not (Test-Path $jsonDir)) {
    try {
        $null = New-Item -ItemType Directory -Path $jsonDir -Force -ErrorAction Stop
        Show-Info "Created directory: $jsonDir"
    }
    catch {
        Show-Error "Failed to create json directory: $_"
        Write-Host ""
        return (New-ModuleResult -Status "Error" -Message "Failed to create json directory")
    }
}

foreach ($item in $enabledItems) {
    $outputPath = Join-Path $jsonDir "$($item.FileName).json"

    try {
        Export-StartLayout -Path $outputPath -ErrorAction Stop

        # Post-write validation
        if (-not (Test-Path $outputPath)) {
            Show-Error "JSON file was not created: $outputPath"
            $failCount++
            Write-Host ""
            continue
        }

        $fileSize = (Get-Item $outputPath).Length
        if ($fileSize -eq 0) {
            Show-Error "JSON file is empty: $outputPath"
            $failCount++
            Write-Host ""
            continue
        }

        Show-Success "Exported: $($item.FileName).json ($fileSize bytes)"
        Write-Host "  Path: $outputPath" -ForegroundColor DarkGray
        $successCount++
    }
    catch {
        Show-Error "Failed to export: $($item.FileName).json - $_"
        $failCount++
    }

    Write-Host ""
}

# ========================================
# Step 6: Result
# ========================================
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount `
    -Title "Start Layout Backup Results")
