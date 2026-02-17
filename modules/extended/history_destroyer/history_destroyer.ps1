# ========================================
# History Destroyer - Windows Comprehensive History Cleaner
# ========================================

Write-Host ""
Show-Separator
Write-Host "History Destroyer" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# ========================================
# Target List
# ========================================
Write-Host "The following items will be deleted:" -ForegroundColor Yellow
Write-Host ""
Write-Host "  [Explorer]" -ForegroundColor White
Write-Host "    - Recent files / Jump Lists" -ForegroundColor Gray
Write-Host "    - Registry MRU (Run, TypedPaths, Search, OpenSave)" -ForegroundColor Gray
Write-Host ""
Write-Host "  [System]" -ForegroundColor White
Write-Host "    - Event Viewer logs (all)" -ForegroundColor Gray
Write-Host "    - IME prediction cache" -ForegroundColor Gray
Write-Host "    - Temporary files (User/System)" -ForegroundColor Gray
Write-Host "    - Clipboard / DNS cache" -ForegroundColor Gray
Write-Host "    - Recycle Bin" -ForegroundColor Gray
Write-Host ""
Write-Host "  [Applications]" -ForegroundColor White
Write-Host "    - Office MRU (Word, Excel, PowerPoint, etc.)" -ForegroundColor Gray
Write-Host "    - Edge browser data (Cache, History, Cookies)" -ForegroundColor Gray
Write-Host "    - Chrome browser data (Cache, History, Cookies)" -ForegroundColor Gray
Write-Host ""
Write-Host "  [Additional]" -ForegroundColor White
Write-Host "    - Windows Search index" -ForegroundColor Gray
Write-Host "    - Thumbnail cache" -ForegroundColor Gray
Write-Host "    - Prefetch data" -ForegroundColor Gray
Write-Host ""
Show-Separator
Write-Host ""
Show-Warning "Explorer will be temporarily stopped during cleanup."
Write-Host "          The taskbar and desktop will disappear briefly." -ForegroundColor Red
Write-Host ""

$cancelResult = Confirm-ModuleExecution -Message "Proceed with history destruction?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""

$successCount = 0
$skipCount = 0
$failCount = 0
$totalSteps = 12

# ========================================
# Helper: Execute cleanup action
# ========================================
function Invoke-CleanupAction {
    param(
        [int]$Step,
        [int]$Total,
        [string]$Description,
        [scriptblock]$Action
    )

    Write-Host "[$Step/$Total] $Description" -ForegroundColor Yellow

    try {
        & $Action
        return "Success"
    }
    catch {
        Show-Error "$($_.Exception.Message)"
        return "Error"
    }
}

# ========================================
# Step 1: Stop Explorer
# ========================================
Write-Host "[1/$totalSteps] Stopping Explorer process..." -ForegroundColor Yellow

try {
    Stop-Process -Name "explorer" -Force -ErrorAction Stop
    Show-Success "Explorer stopped"
}
catch {
    Show-Warning "Failed to stop Explorer: $($_.Exception.Message)"
    Write-Host "          Some locked files may not be deleted" -ForegroundColor Yellow
}

Write-Host ""

# ========================================
# Step 2: Explorer History (Recent / JumpList / Registry MRU)
# ========================================
Write-Host "[2/$totalSteps] Cleaning Explorer history..." -ForegroundColor Yellow

$step2Errors = 0

# Recent files
$recentPaths = @(
    "$env:APPDATA\Microsoft\Windows\Recent\*",
    "$env:APPDATA\Microsoft\Windows\Recent\AutomaticDestinations\*",
    "$env:APPDATA\Microsoft\Windows\Recent\CustomDestinations\*"
)

foreach ($rp in $recentPaths) {
    try {
        if (Test-Path (Split-Path $rp -Parent)) {
            Remove-Item $rp -Recurse -Force -ErrorAction Stop
        }
    }
    catch {
        $step2Errors++
    }
}

# Registry MRU
$registryPaths = @(
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\RunMRU",
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\TypedPaths",
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\WordWheelQuery",
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\ComDlg32\OpenSavePidlMRU",
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\ComDlg32\LastVisitedPidlMRU",
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\StreamMRU"
)

foreach ($regPath in $registryPaths) {
    try {
        if (Test-Path $regPath) {
            Remove-ItemProperty -Path $regPath -Name * -ErrorAction Stop
        }
    }
    catch {
        $step2Errors++
    }
}

if ($step2Errors -eq 0) {
    Show-Success "Explorer history cleaned"
    $successCount++
}
else {
    Show-Warning "Explorer history cleaned ($step2Errors items failed - likely locked)"
    $successCount++
}

Write-Host ""

# ========================================
# Step 3: Event Viewer Logs
# ========================================
Write-Host "[3/$totalSteps] Clearing Event Viewer logs..." -ForegroundColor Yellow

try {
    $logs = Get-WinEvent -ListLog * -Force -ErrorAction SilentlyContinue
    $clearedCount = 0
    foreach ($log in $logs) {
        $wevtResult = & wevtutil.exe cl $log.LogName 2>&1
        if ($LASTEXITCODE -eq 0) { $clearedCount++ }
    }
    Show-Success "Cleared $clearedCount event logs"
    $successCount++
}
catch {
    Show-Error "Failed to clear event logs: $($_.Exception.Message)"
    $failCount++
}

Write-Host ""

# ========================================
# Step 4: IME Prediction Cache
# ========================================
Write-Host "[4/$totalSteps] Cleaning IME prediction cache..." -ForegroundColor Yellow

$imePath = "$env:APPDATA\Microsoft\InputMethod"
if (Test-Path $imePath) {
    try {
        Remove-Item "$imePath\*" -Recurse -Force -ErrorAction Stop
        Show-Success "IME cache cleaned"
        $successCount++
    }
    catch {
        Show-Warning "Some IME cache files could not be deleted (in use)"
        $successCount++
    }
}
else {
    Show-Skip "IME cache folder not found"
    $skipCount++
}

Write-Host ""

# ========================================
# Step 5: Temporary Files
# ========================================
Write-Host "[5/$totalSteps] Cleaning temporary files..." -ForegroundColor Yellow

$tempErrors = 0

# User temp
try {
    Remove-Item "$env:TEMP\*" -Recurse -Force -ErrorAction Stop
}
catch { $tempErrors++ }

# System temp
try {
    Remove-Item "$env:windir\Temp\*" -Recurse -Force -ErrorAction Stop
}
catch { $tempErrors++ }

if ($tempErrors -eq 0) {
    Show-Success "Temporary files cleaned"
}
else {
    Show-Warning "Temporary files cleaned (some locked files skipped)"
}
$successCount++

Write-Host ""

# ========================================
# Step 6: Clipboard & DNS Cache
# ========================================
Write-Host "[6/$totalSteps] Clearing clipboard and DNS cache..." -ForegroundColor Yellow

try {
    Set-Clipboard $null
    Show-Success "Clipboard cleared"
}
catch {
    Show-Warning "Failed to clear clipboard"
}

try {
    Clear-DnsClientCache
    Show-Success "DNS cache cleared"
}
catch {
    Show-Warning "Failed to clear DNS cache"
}

$successCount++
Write-Host ""

# ========================================
# Step 7: Recycle Bin
# ========================================
Write-Host "[7/$totalSteps] Emptying Recycle Bin..." -ForegroundColor Yellow

try {
    $shellCode = @'
    [DllImport("Shell32.dll")]
    public static extern int SHEmptyRecycleBin(IntPtr hwnd, string pszRootPath, int dwFlags);
'@
    Add-Type -MemberDefinition $shellCode -Name Win32RecycleBin -Namespace HistoryDestroyer -ErrorAction SilentlyContinue
    # Flags: SHERB_NOCONFIRMATION(1) | SHERB_NOPROGRESSUI(2) | SHERB_NOSOUND(4) = 7
    $null = [HistoryDestroyer.Win32RecycleBin]::SHEmptyRecycleBin([IntPtr]::Zero, $null, 7)
    Show-Success "Recycle Bin emptied"
    $successCount++
}
catch {
    Show-Error "Failed to empty Recycle Bin: $($_.Exception.Message)"
    $failCount++
}

Write-Host ""

# ========================================
# Step 8: Office MRU
# ========================================
Write-Host "[8/$totalSteps] Cleaning Office recent file history..." -ForegroundColor Yellow

$officeBase = "HKCU:\Software\Microsoft\Office"
if (Test-Path $officeBase) {
    $officeCleaned = 0
    Get-ChildItem $officeBase -ErrorAction SilentlyContinue | ForEach-Object {
        $version = $_.PSChildName
        $apps = @("Word", "Excel", "PowerPoint", "Access", "Publisher", "Visio")
        foreach ($app in $apps) {
            $placeMRU = "$officeBase\$version\$app\Place MRU"
            $fileMRU  = "$officeBase\$version\$app\File MRU"

            if (Test-Path $placeMRU) {
                Remove-ItemProperty -Path $placeMRU -Name * -ErrorAction SilentlyContinue
                $officeCleaned++
            }
            if (Test-Path $fileMRU) {
                Remove-ItemProperty -Path $fileMRU -Name * -ErrorAction SilentlyContinue
                $officeCleaned++
            }
        }
    }
    Show-Success "Office MRU cleaned ($officeCleaned entries)"
    $successCount++
}
else {
    Show-Skip "Office registry not found"
    $skipCount++
}

Write-Host ""

# ========================================
# Step 9: Browser Data (Edge)
# ========================================
Write-Host "[9/$totalSteps] Cleaning Microsoft Edge data..." -ForegroundColor Yellow

$edgeBase = "$env:LOCALAPPDATA\Microsoft\Edge\User Data"
if (Test-Path $edgeBase) {
    # Stop Edge if running
    $edgeProc = Get-Process -Name "msedge" -ErrorAction SilentlyContinue
    if ($edgeProc) {
        Stop-Process -Name "msedge" -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
    }

    $edgeTargets = @("Cache", "Code Cache", "GPUCache", "History", "Cookies", "Cookies-journal",
                     "Top Sites", "Top Sites-journal", "Visited Links", "Web Data", "Web Data-journal",
                     "Session Storage", "Local Storage")
    $edgeCleaned = 0

    # Process all profiles (Default, Profile 1, etc.)
    $profiles = Get-ChildItem $edgeBase -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -eq "Default" -or $_.Name -match "^Profile " }

    foreach ($profile in $profiles) {
        foreach ($target in $edgeTargets) {
            $targetPath = Join-Path $profile.FullName $target
            if (Test-Path $targetPath) {
                try {
                    Remove-Item $targetPath -Recurse -Force -ErrorAction Stop
                    $edgeCleaned++
                }
                catch { }
            }
        }
    }

    Show-Success "Edge data cleaned ($edgeCleaned items)"
    $successCount++
}
else {
    Show-Skip "Edge not found"
    $skipCount++
}

Write-Host ""

# ========================================
# Step 10: Browser Data (Chrome)
# ========================================
Write-Host "[10/$totalSteps] Cleaning Google Chrome data..." -ForegroundColor Yellow

$chromeBase = "$env:LOCALAPPDATA\Google\Chrome\User Data"
if (Test-Path $chromeBase) {
    # Stop Chrome if running
    $chromeProc = Get-Process -Name "chrome" -ErrorAction SilentlyContinue
    if ($chromeProc) {
        Stop-Process -Name "chrome" -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
    }

    $chromeTargets = @("Cache", "Code Cache", "GPUCache", "History", "Cookies", "Cookies-journal",
                       "Top Sites", "Top Sites-journal", "Visited Links", "Web Data", "Web Data-journal",
                       "Session Storage", "Local Storage")
    $chromeCleaned = 0

    $profiles = Get-ChildItem $chromeBase -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -eq "Default" -or $_.Name -match "^Profile " }

    foreach ($profile in $profiles) {
        foreach ($target in $chromeTargets) {
            $targetPath = Join-Path $profile.FullName $target
            if (Test-Path $targetPath) {
                try {
                    Remove-Item $targetPath -Recurse -Force -ErrorAction Stop
                    $chromeCleaned++
                }
                catch { }
            }
        }
    }

    Show-Success "Chrome data cleaned ($chromeCleaned items)"
    $successCount++
}
else {
    Show-Skip "Chrome not found"
    $skipCount++
}

Write-Host ""

# ========================================
# Step 11: Windows Search Index
# ========================================
Write-Host "[11/$totalSteps] Resetting Windows Search index..." -ForegroundColor Yellow

try {
    $wsearchService = Get-Service -Name "WSearch" -ErrorAction SilentlyContinue
    if ($wsearchService) {
        if ($wsearchService.Status -eq "Running") {
            Stop-Service -Name "WSearch" -Force -ErrorAction Stop
            Start-Sleep -Seconds 2
        }

        $searchDbPath = "$env:ProgramData\Microsoft\Search\Data\Applications\Windows\Windows.edb"
        if (Test-Path $searchDbPath) {
            Remove-Item $searchDbPath -Force -ErrorAction Stop
            Show-Success "Search index deleted"
        }
        else {
            Show-Skip "Search index file not found"
        }

        Start-Service -Name "WSearch" -ErrorAction SilentlyContinue
        $successCount++
    }
    else {
        Show-Skip "Windows Search service not found"
        $skipCount++
    }
}
catch {
    Show-Error "Failed to reset search index: $($_.Exception.Message)"
    # Try to restart WSearch even on error
    Start-Service -Name "WSearch" -ErrorAction SilentlyContinue
    $failCount++
}

Write-Host ""

# ========================================
# Step 12: Thumbnail Cache + Prefetch
# ========================================
Write-Host "[12/$totalSteps] Cleaning thumbnail cache and prefetch..." -ForegroundColor Yellow

# Thumbnail cache
$thumbPath = "$env:LOCALAPPDATA\Microsoft\Windows\Explorer"
$thumbCleaned = 0
if (Test-Path $thumbPath) {
    $thumbFiles = Get-ChildItem $thumbPath -Filter "thumbcache_*.db" -ErrorAction SilentlyContinue
    foreach ($f in $thumbFiles) {
        try {
            Remove-Item $f.FullName -Force -ErrorAction Stop
            $thumbCleaned++
        }
        catch { }
    }
    # Also clean iconcache
    $iconFiles = Get-ChildItem $thumbPath -Filter "iconcache_*.db" -ErrorAction SilentlyContinue
    foreach ($f in $iconFiles) {
        try {
            Remove-Item $f.FullName -Force -ErrorAction Stop
            $thumbCleaned++
        }
        catch { }
    }
}
Show-Success "Thumbnail cache cleaned ($thumbCleaned files)"

# Prefetch
$prefetchPath = "$env:windir\Prefetch"
$prefetchCleaned = 0
if (Test-Path $prefetchPath) {
    $pfFiles = Get-ChildItem $prefetchPath -ErrorAction SilentlyContinue
    foreach ($f in $pfFiles) {
        try {
            Remove-Item $f.FullName -Force -ErrorAction Stop
            $prefetchCleaned++
        }
        catch { }
    }
}
Show-Success "Prefetch cleaned ($prefetchCleaned files)"

$successCount++
Write-Host ""

# ========================================
# Restart Explorer
# ========================================
Show-Info "Restarting Explorer..."

Start-Process "explorer.exe"
Start-Sleep -Seconds 2

$explorerRunning = Get-Process -Name "explorer" -ErrorAction SilentlyContinue
if ($explorerRunning) {
    Show-Success "Explorer restarted"
}
else {
    Show-Warning "Explorer may not have restarted - please check manually"
}

Write-Host ""

# ========================================
# Result Summary
# ========================================
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount -Title "History Destroyer - Results")
