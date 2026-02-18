# ========================================
# File Copy Config - File Distribution Tool
# ========================================
# Copies files/folders from the source/ directory
# to specified destinations based on copy_list.csv.
# Supports overwrite control per entry.
# ========================================

Write-Host ""
Show-Separator
Write-Host "File Copy Config" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# ========================================
# Step 1: Load CSV
# ========================================
$csvPath = Join-Path $PSScriptRoot "copy_list.csv"
$items = Import-ModuleCsv -Path $csvPath -FilterEnabled
if ($null -eq $items) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load copy_list.csv")
}
if ($items.Count -eq 0) {
    return (New-ModuleResult -Status "Skipped" -Message "No enabled entries")
}

# Expand Windows-style environment variables in DestPath (%USERPROFILE%, etc.)
foreach ($item in $items) {
    $item.DestPath = [System.Environment]::ExpandEnvironmentVariables($item.DestPath)
}

# ========================================
# Step 2: Validate source directory
# ========================================
$sourceDir = Join-Path $PSScriptRoot "source"
if (-not (Test-Path $sourceDir)) {
    Show-Error "source/ directory not found: $sourceDir"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "source/ directory not found")
}

# ========================================
# Step 4: Display copy targets with status
# ========================================
Show-Info "Copy targets: $($items.Count) items"
Write-Host ""

$index = 0
foreach ($item in $items) {
    $index++
    $srcPath = Join-Path $sourceDir $item.FileName
    $destPath = Join-Path $item.DestPath $item.FileName
    $isOverwrite = ($item.Overwrite -eq "1")

    if (-not (Test-Path $srcPath)) {
        $marker = "[Missing]"
        $markerColor = "Red"
    }
    elseif (Test-Path $destPath) {
        if ($isOverwrite) {
            $marker = "[Overwrite]"
            $markerColor = "Yellow"
        }
        else {
            $marker = "[Current]"
            $markerColor = "Gray"
        }
    }
    else {
        $marker = "[Copy]"
        $markerColor = "White"
    }

    Write-Host "  [$index] $($item.FileName)  $marker" -ForegroundColor $markerColor
    Write-Host "      Source: $srcPath" -ForegroundColor DarkGray
    Write-Host "      Dest:   $($item.DestPath)" -ForegroundColor DarkGray
    if ($item.Description) {
        Write-Host "      $($item.Description)" -ForegroundColor Gray
    }
    Write-Host ""
}

# ========================================
# Step 5: Confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Copy the above files?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""

# ========================================
# Step 6: Execute copy operations
# ========================================
$successCount = 0
$skipCount = 0
$failCount = 0
$total = $items.Count
$current = 0

foreach ($item in $items) {
    $current++
    $srcPath = Join-Path $sourceDir $item.FileName
    $destPath = Join-Path $item.DestPath $item.FileName
    $isOverwrite = ($item.Overwrite -eq "1")

    Write-Host "[$current/$total] $($item.FileName)" -ForegroundColor Cyan

    # Source existence check
    if (-not (Test-Path $srcPath)) {
        Show-Error "Source not found: $srcPath"
        $failCount++
        Write-Host ""
        continue
    }

    # Create destination directory if needed
    if (-not (Test-Path $item.DestPath)) {
        try {
            $null = New-Item -ItemType Directory -Path $item.DestPath -Force
            Show-Info "Created directory: $($item.DestPath)"
        }
        catch {
            Show-Error "Failed to create directory: $($item.DestPath) - $_"
            $failCount++
            Write-Host ""
            continue
        }
    }

    # Overwrite control
    if (Test-Path $destPath) {
        if (-not $isOverwrite) {
            Show-Skip "File already exists (overwrite disabled)"
            $skipCount++
            Write-Host ""
            continue
        }
    }

    # Execute copy
    try {
        Copy-Item -Path $srcPath -Destination $item.DestPath -Recurse -Force -ErrorAction Stop
        if (Test-Path $destPath) {
            Show-Success "Copied"
        }
        else {
            Show-Success "Copied"
        }
        $successCount++
    }
    catch {
        Show-Error "Copy failed: $_"
        $failCount++
    }

    Write-Host ""
}

# ========================================
# Result Summary
# ========================================
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount -Title "File Copy Results")
