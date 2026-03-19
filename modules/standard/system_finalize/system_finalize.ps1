# ========================================
# System Finalize Script
# ========================================
# Performs final kitting cleanup: re-registers shell32.dll,
# clears icon/thumbnail caches, and restarts Explorer.
#
# [NOTES]
# - Explorer will be temporarily stopped during cache cleanup
# ========================================

Write-Host ""
Show-Separator
Write-Host "System Finalize" -ForegroundColor Cyan
Show-Separator
Write-Host ""


# ========================================
# Step 1: Define finalization targets
# ========================================
$cachePaths = @(
    @{ Path = "$env:LOCALAPPDATA\Microsoft\Windows\Explorer\iconcache*.db"; Description = "Icon cache (lowercase)" }
    @{ Path = "$env:LOCALAPPDATA\Microsoft\Windows\Explorer\IconCache*.db"; Description = "Icon cache (mixed case)" }
    @{ Path = "$env:LOCALAPPDATA\Microsoft\Windows\Explorer\thumbcache*.db"; Description = "Thumbnail cache" }
    @{ Path = "$env:LOCALAPPDATA\IconCache.db"; Description = "Legacy icon cache" }
)


# ========================================
# Step 3: Dry-run display
# ========================================
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "Finalization Targets" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

# regsvr32
Write-Host "  [APPLY] Shell re-registration (regsvr32 /s /i:U shell32.dll)" -ForegroundColor Yellow
Write-Host ""

# Explorer stop/restart
Write-Host "  [APPLY] Stop Explorer -> Delete caches -> Restart Explorer" -ForegroundColor Yellow
Write-Host ""

# Cache paths
foreach ($cache in $cachePaths) {
    $exists = Test-Path $cache.Path
    if ($exists) {
        Write-Host "  [DELETE] $($cache.Description)" -ForegroundColor Yellow
        Write-Host "    Path: $($cache.Path)" -ForegroundColor DarkGray
    }
    else {
        Write-Host "  [NOT FOUND] $($cache.Description)" -ForegroundColor DarkGray
        Write-Host "    Path: $($cache.Path)" -ForegroundColor DarkGray
    }
    Write-Host ""
}

Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""


# ========================================
# Step 4: Confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Proceed with system finalization?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""


# ========================================
# Step 5: Execution
# ========================================
$successCount = 0
$skipCount    = 0
$failCount    = 0

# ----------------------------------------
# Phase A: regsvr32 shell32.dll
# ----------------------------------------
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host "Processing: Shell re-registration" -ForegroundColor Cyan
Write-Host "----------------------------------------" -ForegroundColor White

try {
    Start-Process "regsvr32" -ArgumentList "/s /i:U shell32.dll" -Wait -NoNewWindow
    Show-Success "Shell re-registration completed"
    $successCount++
}
catch {
    Show-Error "Shell re-registration failed: $_"
    $failCount++
}

Write-Host ""

# ----------------------------------------
# Phase B: Stop Explorer (not counted)
# ----------------------------------------
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host "Stopping Explorer" -ForegroundColor Cyan
Write-Host "----------------------------------------" -ForegroundColor White

Show-Warning "Explorer will be temporarily stopped during cleanup."
Write-Host "          The taskbar and desktop will disappear briefly." -ForegroundColor Red
Write-Host ""

try {
    Stop-Process -Name "explorer" -Force -ErrorAction Stop
    Show-Success "Explorer stopped"
}
catch {
    Show-Warning "Failed to stop Explorer: $($_.Exception.Message)"
    Write-Host "          Some locked files may not be deleted" -ForegroundColor Yellow
}

Write-Host ""

# ----------------------------------------
# Phase C: Delete cache files (4 items counted)
# ----------------------------------------
foreach ($cache in $cachePaths) {
    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "Processing: $($cache.Description)" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor White

    if (-not (Test-Path $cache.Path)) {
        Show-Skip "Not found: $($cache.Path)"
        Write-Host ""
        $skipCount++
        continue
    }

    try {
        Remove-Item -Path $cache.Path -Force -ErrorAction SilentlyContinue

        # Post-delete check
        if (Test-Path $cache.Path) {
            Show-Warning "Partially deleted (some files may be locked): $($cache.Description)"
            $successCount++
        }
        else {
            Show-Success "Deleted: $($cache.Description)"
            $successCount++
        }
    }
    catch {
        Show-Error "Failed to delete: $($cache.Description) : $_"
        $failCount++
    }

    Write-Host ""
}

# ----------------------------------------
# Phase D: Restart Explorer (not counted)
# ----------------------------------------
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host "Restarting Explorer" -ForegroundColor Cyan
Write-Host "----------------------------------------" -ForegroundColor White

Show-Info "Restarting Explorer..."

$maxWait = 15; $interval = 1; $elapsed = 0; $restarted = $false
while ($elapsed -lt $maxWait) {
    Start-Sleep -Seconds $interval
    $elapsed += $interval
    if (@(Get-Process -Name "explorer" -ErrorAction SilentlyContinue).Count -gt 0) {
        $restarted = $true; break
    }
}
if ($restarted) {
    Show-Success "Explorer restarted (${elapsed}s)"
}
else {
    Start-Process "explorer.exe"
    Show-Warning "Explorer auto-restart timed out. Started manually."
}
Write-Host ""


# ========================================
# Step 6: Result summary
# ========================================
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount `
    -Title "System Finalize Results")
