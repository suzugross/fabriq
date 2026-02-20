# ========================================
# File Delete - CSV-Driven File/Folder Deletion
# ========================================
# Deletes files and folders listed in delete_list.csv.
# Supports environment variable expansion and per-entry
# missing-file behavior (Skip or Error).
# ========================================

Write-Host ""
Show-Separator
Write-Host "File Delete" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# ========================================
# Step 1: Load CSV
# ========================================
$csvPath = Join-Path $PSScriptRoot "delete_list.csv"
$items = Import-ModuleCsv -Path $csvPath -FilterEnabled
if ($null -eq $items) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load delete_list.csv")
}
if ($items.Count -eq 0) {
    return (New-ModuleResult -Status "Skipped" -Message "No enabled entries")
}

# Expand Windows-style environment variables (%USERPROFILE%, %TEMP%, etc.)
foreach ($item in $items) {
    $item.TargetPath = [System.Environment]::ExpandEnvironmentVariables($item.TargetPath)
}

# ========================================
# Step 2: Display deletion targets with status
# ========================================
Show-Info "Deletion targets: $($items.Count) items"
Write-Host ""

$index = 0
foreach ($item in $items) {
    $index++
    $targetPath = $item.TargetPath
    $ifNotFound = if ($item.IfNotFound) { $item.IfNotFound } else { "Skip" }

    if (Test-Path $targetPath) {
        $isDir = (Get-Item $targetPath -ErrorAction SilentlyContinue) -is [System.IO.DirectoryInfo]
        $typeLabel = if ($isDir) { "Dir" } else { "File" }
        $marker = "[$typeLabel][Exists]"
        $markerColor = "White"
    }
    else {
        $marker = "[Not Found]"
        $markerColor = if ($ifNotFound -eq "Error") { "Red" } else { "Gray" }
    }

    Write-Host "  [$index] $($item.Description)  $marker" -ForegroundColor $markerColor
    Write-Host "      Path: $targetPath" -ForegroundColor DarkGray
    if ($ifNotFound -eq "Error") {
        Write-Host "      IfNotFound: Error" -ForegroundColor DarkGray
    }
    Write-Host ""
}

# ========================================
# Step 3: Confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Delete the above targets?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""

# ========================================
# Step 4: Execute deletion
# ========================================
$successCount = 0
$skipCount    = 0
$failCount    = 0
$total        = $items.Count
$current      = 0

foreach ($item in $items) {
    $current++
    $targetPath = $item.TargetPath
    $ifNotFound = if ($item.IfNotFound) { $item.IfNotFound } else { "Skip" }

    Write-Host "[$current/$total] $($item.Description)" -ForegroundColor Cyan
    Write-Host "  Path: $targetPath" -ForegroundColor DarkGray

    if (-not (Test-Path $targetPath)) {
        if ($ifNotFound -eq "Error") {
            Show-Error "Target not found: $targetPath"
            $failCount++
        }
        else {
            Show-Skip "Not found — skipped"
            $skipCount++
        }
        Write-Host ""
        continue
    }

    try {
        Remove-Item -Path $targetPath -Force -Recurse -ErrorAction Stop
        Show-Success "Deleted"
        $successCount++
    }
    catch {
        Show-Error "Deletion failed: $_"
        $failCount++
    }

    Write-Host ""
}

# ========================================
# Result Summary
# ========================================
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount -Title "File Delete Results")
