# ========================================
# File Copy Config - File Distribution Tool
# ========================================
# Copies files/folders from the source/ directory
# to specified destinations based on copy_list.csv.
# Supports overwrite control per entry.
# ========================================

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "File Copy Config" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# ========================================
# Step 1: Load CSV
# ========================================
$csvPath = Join-Path $PSScriptRoot "copy_list.csv"
if (-not (Test-Path $csvPath)) {
    Write-Host "[ERROR] copy_list.csv not found: $csvPath" -ForegroundColor Red
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "copy_list.csv not found")
}

try {
    $allItems = @(Import-Csv -Path $csvPath -Encoding Default)
}
catch {
    Write-Host "[ERROR] Failed to read copy_list.csv: $_" -ForegroundColor Red
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Failed to read copy_list.csv")
}

# ========================================
# Step 2: Validate source directory
# ========================================
$sourceDir = Join-Path $PSScriptRoot "source"
if (-not (Test-Path $sourceDir)) {
    Write-Host "[ERROR] source/ directory not found: $sourceDir" -ForegroundColor Red
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "source/ directory not found")
}

# ========================================
# Step 3: Filter enabled entries
# ========================================
$items = @($allItems | Where-Object { $_.Enabled -eq "1" })
if ($items.Count -eq 0) {
    Write-Host "[INFO] No enabled entries in copy_list.csv" -ForegroundColor Gray
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "No enabled entries")
}

# ========================================
# Step 4: Display copy targets with status
# ========================================
Write-Host "[INFO] Copy targets: $($items.Count) items" -ForegroundColor Cyan
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
if (-not (Confirm-Execution -Message "Copy the above files?")) {
    Write-Host ""
    Write-Host "[INFO] Canceled" -ForegroundColor Yellow
    Write-Host ""
    return (New-ModuleResult -Status "Cancelled" -Message "User canceled")
}

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
        Write-Host "  [ERROR] Source not found: $srcPath" -ForegroundColor Red
        $failCount++
        Write-Host ""
        continue
    }

    # Create destination directory if needed
    if (-not (Test-Path $item.DestPath)) {
        try {
            $null = New-Item -ItemType Directory -Path $item.DestPath -Force
            Write-Host "  [INFO] Created directory: $($item.DestPath)" -ForegroundColor Gray
        }
        catch {
            Write-Host "  [ERROR] Failed to create directory: $($item.DestPath) - $_" -ForegroundColor Red
            $failCount++
            Write-Host ""
            continue
        }
    }

    # Overwrite control
    if (Test-Path $destPath) {
        if (-not $isOverwrite) {
            Write-Host "  [SKIP] File already exists (overwrite disabled)" -ForegroundColor Gray
            $skipCount++
            Write-Host ""
            continue
        }
    }

    # Execute copy
    try {
        Copy-Item -Path $srcPath -Destination $item.DestPath -Recurse -Force -ErrorAction Stop
        if (Test-Path $destPath) {
            Write-Host "  [SUCCESS] Copied" -ForegroundColor Green
        }
        else {
            Write-Host "  [SUCCESS] Copied" -ForegroundColor Green
        }
        $successCount++
    }
    catch {
        Write-Host "  [ERROR] Copy failed: $_" -ForegroundColor Red
        $failCount++
    }

    Write-Host ""
}

# ========================================
# Result Summary
# ========================================
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "File Copy Results" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
if ($successCount -gt 0) {
    Write-Host "  Success: $successCount items" -ForegroundColor Green
}
if ($skipCount -gt 0) {
    Write-Host "  Skipped: $skipCount items (Already exists)" -ForegroundColor Gray
}
if ($failCount -gt 0) {
    Write-Host "  Failed:  $failCount items" -ForegroundColor Red
}
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Return ModuleResult
$overallStatus = if ($failCount -eq 0 -and $successCount -gt 0) { "Success" }
    elseif ($failCount -eq 0 -and $successCount -eq 0 -and $skipCount -gt 0) { "Skipped" }
    elseif ($successCount -gt 0 -and $failCount -gt 0) { "Partial" }
    elseif ($successCount -gt 0 -and $skipCount -gt 0) { "Success" }
    else { "Error" }
return (New-ModuleResult -Status $overallStatus -Message "Success: $successCount, Skip: $skipCount, Fail: $failCount")
