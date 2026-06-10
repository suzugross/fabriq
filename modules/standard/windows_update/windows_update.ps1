# ========================================
# Windows Update - Single Pass Module
# ========================================
# Scans, downloads, and installs available Windows Updates
# using the Microsoft.Update.Session COM API.
#
# This module performs ONE iteration of scan/download/install.
# The reboot loop is managed by main.ps1 (Invoke-WindowsUpdateLoop)
# following the same pattern as Profile __RESTART__.
#
# Prerequisites: Administrator privileges, network connectivity
# ========================================

param(
    [string[]]$SkipKBs = @(),
    [switch]$AutoConfirm
)

# Start transcript log for WU operations
$wuLogDir = Join-Path $PSScriptRoot "..\..\..\logs"
if (-not (Test-Path $wuLogDir)) { New-Item -Path $wuLogDir -ItemType Directory -Force | Out-Null }
$wuTimestamp = Get-Date -Format "yyyy_MM_dd_HHmmss"
$wuUniqueId  = Get-HardwareUniqueId
$wuHostname  = $env:COMPUTERNAME
$wuLogFile   = Join-Path $wuLogDir "wu_${wuTimestamp}_${wuUniqueId}_${wuHostname}.log"
$wuTranscriptStarted = $false
try {
    Start-Transcript -Path $wuLogFile -Append -ErrorAction Stop | Out-Null
    $wuTranscriptStarted = $true
}
catch {
    # Transcript may already be running (e.g. launched from main.ps1)
}

Write-Host ""
Show-Separator
Write-Host "Windows Update" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# ========================================
# Step 1: Load Configuration
# ========================================
$csvPath = Join-Path $PSScriptRoot "windows_update_list.csv"

$configItems = Import-ModuleCsv -Path $csvPath -FilterEnabled `
    -RequiredColumns @("Enabled", "SettingName", "Value")

if ($null -eq $configItems) {
    return @{ Status = "Error"; RebootRequired = $false; InstalledCount = 0; FailedCount = 0; InstalledKBs = @(); FailedKBs = @(); UpdatesFound = 0 }
}

# Parse SettingName/Value pairs into hashtable
$config = @{}
foreach ($item in $configItems) {
    $config[$item.SettingName] = $item.Value
}

$suspendBitLocker = ($config["SuspendBitLocker"] -eq "1")
$includeOptional  = ($config["IncludeOptionalUpdates"] -eq "1")
$includeSeeker    = ($config["IncludeSeekerUpdates"] -eq "1")
$autoInstall      = ($config["AutoInstall"] -eq "1")

# ========================================
# Step 2: Prerequisite Checks
# ========================================

# Admin privilege check
if (-not (Test-AdminPrivilege)) {
    return @{ Status = "Error"; RebootRequired = $false; InstalledCount = 0; FailedCount = 0; InstalledKBs = @(); FailedKBs = @(); UpdatesFound = 0 }
}

# Wait for base system services
Wait-SystemReady -RequiredServices @("LanmanWorkstation", "Dnscache")
Write-Host ""

# Start Windows Update service
$wuSvc = Get-Service -Name "wuauserv" -ErrorAction SilentlyContinue
if ($null -eq $wuSvc) {
    Show-Error "Windows Update service (wuauserv) not found"
    return @{ Status = "Error"; RebootRequired = $false; InstalledCount = 0; FailedCount = 0; InstalledKBs = @(); FailedKBs = @(); UpdatesFound = 0 }
}
if ($wuSvc.Status -ne "Running") {
    Show-Info "Starting Windows Update service..."
    try {
        Start-Service -Name "wuauserv" -ErrorAction Stop
        Show-Success "Windows Update service started"
    }
    catch {
        Show-Error "Failed to start Windows Update service: $_"
        return @{ Status = "Error"; RebootRequired = $false; InstalledCount = 0; FailedCount = 0; InstalledKBs = @(); FailedKBs = @(); UpdatesFound = 0 }
    }
}
else {
    Show-Info "Windows Update service is already running"
}
Write-Host ""

# Wait for network connectivity
Wait-NetworkReady
Write-Host ""

# Prevent sleep during updates
# Sleep suppression is managed by main.ps1 (Enable on startup, Disable on Exit-Fabriq)

# Suspend BitLocker if configured
if ($suspendBitLocker) {
    try {
        $blVolume = Get-BitLockerVolume -MountPoint "C:" -ErrorAction SilentlyContinue
        if ($blVolume -and $blVolume.ProtectionStatus -eq "On") {
            $null = Suspend-BitLocker -MountPoint "C:" -RebootCount 1 -ErrorAction SilentlyContinue
            Show-Info "BitLocker suspended for 1 reboot cycle"
        }
    }
    catch {
        Show-Warning "BitLocker suspension failed (non-fatal): $_"
    }
}

# ========================================
# Step 3: COM API Scan
# ========================================
$updateSession = New-Object -ComObject Microsoft.Update.Session
$updateSearcher = $updateSession.CreateUpdateSearcher()

# Register Microsoft Update service for optional quality updates
if ($includeOptional) {
    try {
        $serviceManager = New-Object -ComObject Microsoft.Update.ServiceManager
        $serviceManager.ClientApplicationID = "fabriq"
        $null = $serviceManager.AddService2("7971f918-a847-4430-9279-4a52d1efe18d", 7, "")
        Show-Info "Microsoft Update service registered (optional updates enabled)"
    }
    catch {
        Show-Warning "Failed to register Microsoft Update service (non-fatal): $_"
    }

    $updateSearcher.ServerSelection = 3  # ssOthers
    $updateSearcher.ServiceID = "7971f918-a847-4430-9279-4a52d1efe18d"
}

# Inner loop for no-reboot re-scans only (cascading updates within same session)
$maxNoRebootScans = 3
$noRebootScanCount = 0
$totalSuccessCount = 0
$totalFailCount = 0
$allInstalledKBs = @()
$allFailedKBs = @()
$rebootRequired = $false

while ($true) {
    Show-Info "Scanning for available updates..."
    Write-Host ""

    $searchCriteria = "IsInstalled=0"
    if ($includeSeeker) {
        $searchCriteria = "IsInstalled=0 and DeploymentAction='Installation' or IsInstalled=0 and DeploymentAction='OptionalInstallation'"
    }
    try {
        $searchResult = $updateSearcher.Search($searchCriteria)
    }
    catch {
        Show-Error "Update scan failed: $_"

        return @{ Status = "Error"; RebootRequired = $false; InstalledCount = $totalSuccessCount; FailedCount = $totalFailCount; InstalledKBs = @($allInstalledKBs); FailedKBs = @($allFailedKBs); UpdatesFound = 0 }
    }

    $availableUpdates = $searchResult.Updates

    # Filter out skipped KBs (phantom re-appearing updates)
    if ($SkipKBs.Count -gt 0 -and $availableUpdates.Count -gt 0) {
        $filteredUpdates = New-Object -ComObject Microsoft.Update.UpdateColl
        $skippedCount = 0
        for ($i = 0; $i -lt $availableUpdates.Count; $i++) {
            $update = $availableUpdates.Item($i)
            $kb = ""
            if ($update.Title -match '(KB\d+)') { $kb = $Matches[1] }
            if ($kb -and $SkipKBs -contains $kb) {
                Show-Warning "Skipping $kb (re-appeared after previous install): $($update.Title)"
                $skippedCount++
            }
            else {
                $filteredUpdates.Add($update) | Out-Null
            }
        }
        if ($skippedCount -gt 0) {
            Write-Host ""
        }
        $availableUpdates = $filteredUpdates
    }

    # No updates available -> all done
    if ($availableUpdates.Count -eq 0) {
        Show-Success "No updates available. System is up to date."
        Write-Host ""
        break
    }

    # ========================================
    # Step 4: Display + Confirmation
    # ========================================
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host "Available Updates ($($availableUpdates.Count) found)" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host ""

    for ($i = 0; $i -lt $availableUpdates.Count; $i++) {
        $update = $availableUpdates.Item($i)
        $sizeMB = [math]::Round($update.MaxDownloadSize / 1MB, 1)
        $rebootFlag = if ($update.RebootRequired) { "Yes" } else { "No" }
        $optionalTag = if ($update.BrowseOnly) { " [Optional]" } else { "" }

        Write-Host "  [$($i + 1)] $($update.Title)$optionalTag" -ForegroundColor White
        Write-Host "      Size: ${sizeMB}MB | Reboot required: $rebootFlag" -ForegroundColor DarkGray
        Write-Host ""
    }

    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host ""

    if ($noRebootScanCount -eq 0 -and -not $AutoConfirm) {
        if ($autoInstall) {
            Show-Info "[AUTO-INSTALL] Skipping confirmation, proceeding automatically"
        }
        else {
            $cancelResult = Confirm-ModuleExecution -Message "Install all $($availableUpdates.Count) updates?"
            if ($null -ne $cancelResult) {
        
                return @{ Status = "Cancelled"; RebootRequired = $false; InstalledCount = 0; FailedCount = 0; InstalledKBs = @(); FailedKBs = @(); UpdatesFound = $availableUpdates.Count }
            }
        }
    }
    else {
        Show-Info "[AUTO] Proceeding automatically"
    }
    Write-Host ""

    # ========================================
    # Step 5: Download + Install
    # ========================================

    # --- Download Phase ---
    $downloader = $updateSession.CreateUpdateDownloader()
    $downloadNeeded = 0
    $downloadDone = 0
    $downloadFailed = 0

    for ($i = 0; $i -lt $availableUpdates.Count; $i++) {
        if (-not $availableUpdates.Item($i).IsDownloaded) { $downloadNeeded++ }
    }

    if ($downloadNeeded -gt 0) {
        Show-Info "Downloading $downloadNeeded updates..."
        Write-Host ""

        $dlIndex = 0
        for ($i = 0; $i -lt $availableUpdates.Count; $i++) {
            $update = $availableUpdates.Item($i)
            if ($update.IsDownloaded) { continue }

            $dlIndex++
            $null = $update.AcceptEula()
            Show-Info "[$dlIndex/$downloadNeeded] Downloading: $($update.Title)"

            # KB number for failure accounting (same extraction as install phase)
            $dlKb = ""
            if ($update.Title -match '(KB\d+)') { $dlKb = $Matches[1] }

            try {
                $singleDownload = New-Object -ComObject Microsoft.Update.UpdateColl
                $singleDownload.Add($update) | Out-Null
                $downloader.Updates = $singleDownload
                $dlResult = $downloader.Download()

                # OperationResultCode 4=Failed / 5=Aborted are RETURNED (not
                # thrown) by WUA under e.g. USO contention; only 2=Succeeded /
                # 3=SucceededWithErrors mean the payload is on disk. A failed
                # download must enter the failure accounting, otherwise the
                # update silently drops out of the install phase (IsDownloaded
                # stays false) and the module can finish as Success.
                if ($dlResult.ResultCode -eq 2 -or $dlResult.ResultCode -eq 3) {
                    Show-Success "Downloaded"
                    $downloadDone++
                }
                else {
                    $dlHresult = "N/A"
                    try { $dlHresult = "0x{0:X8}" -f $dlResult.HResult } catch { }
                    Show-Error "Download failed (ResultCode: $($dlResult.ResultCode) | HResult: $dlHresult)"
                    $downloadFailed++
                    $totalFailCount++
                    $allFailedKBs += @{ KB = $dlKb; Title = $update.Title; HResult = $dlHresult }
                }
            }
            catch {
                Show-Error "Download error: $_"
                $downloadFailed++
                $totalFailCount++
                $allFailedKBs += @{ KB = $dlKb; Title = $update.Title; HResult = "N/A" }
            }
        }

        Write-Host ""
        if ($downloadFailed -gt 0 -and $downloadDone -eq 0) {
            Show-Error "All downloads failed"

            # Aggregate totals, not this iteration's counters: on a no-reboot
            # re-scan iteration, earlier installs must survive into the report.
            return @{
                Status         = if ($totalSuccessCount -gt 0) { "Partial" } else { "Error" }
                RebootRequired = $rebootRequired
                InstalledCount = $totalSuccessCount
                FailedCount    = $totalFailCount
                InstalledKBs   = @($allInstalledKBs)
                FailedKBs      = @($allFailedKBs)
                UpdatesFound   = $availableUpdates.Count
            }
        }
        elseif ($downloadFailed -gt 0) {
            Show-Warning "Download completed: $downloadDone succeeded, $downloadFailed failed"
        }
        else {
            Show-Success "All $downloadDone downloads completed"
        }
    }
    else {
        Show-Info "All updates already downloaded"
    }

    Write-Host ""

    # --- Install Phase ---
    $installTargetCount = 0
    for ($i = 0; $i -lt $availableUpdates.Count; $i++) {
        if ($availableUpdates.Item($i).IsDownloaded) { $installTargetCount++ }
    }

    if ($installTargetCount -eq 0) {
        Show-Warning "No downloaded updates available to install"
        break
    }

    Show-Info "Installing $installTargetCount updates..."
    Write-Host ""

    $installer = $updateSession.CreateUpdateInstaller()
    $successCount = 0
    $failCount = 0
    $newKBs = @()
    $newFailedKBs = @()
    $instIndex = 0

    for ($i = 0; $i -lt $availableUpdates.Count; $i++) {
        $update = $availableUpdates.Item($i)
        if (-not $update.IsDownloaded) { continue }

        # Extract KB number from title
        $kb = ""
        if ($update.Title -match '(KB\d+)') { $kb = $Matches[1] }

        $instIndex++
        Show-Info "[$instIndex/$installTargetCount] Installing: $($update.Title)"

        try {
            $singleUpdate = New-Object -ComObject Microsoft.Update.UpdateColl
            $singleUpdate.Add($update) | Out-Null
            $installer.Updates = $singleUpdate
            $singleResult = $installer.Install()

            $itemResult = $singleResult.GetUpdateResult(0)

            if ($itemResult.ResultCode -eq 2) {
                Show-Success "$($update.Title)"
                $successCount++
                $newKBs += @{ KB = $kb; Title = $update.Title }

                # Explicit commit to finalize OptionalInstall updates
                # (COM API Install() may only stage the update without committing to CBS)
                try {
                    $installer.Commit(0)
                    Show-Info "  Commit called"
                }
                catch {
                    # Commit not supported or not needed on this version — non-fatal
                }
            }
            else {
                $hresult = "0x{0:X8}" -f $itemResult.HResult
                Show-Error "$($update.Title)"
                Show-Error "  ResultCode: $($itemResult.ResultCode) | HResult: $hresult"
                $failCount++
                $newFailedKBs += @{ KB = $kb; Title = $update.Title; HResult = $hresult }
            }

            # Only count reboot requirement from successfully installed updates
            if ($itemResult.ResultCode -eq 2) {
                if ($itemResult.RebootRequired) { $rebootRequired = $true }
                if ($singleResult.RebootRequired) { $rebootRequired = $true }
            }
        }
        catch {
            Show-Error "Install failed: $($update.Title) - $_"
            $failCount++
            $newFailedKBs += @{ KB = $kb; Title = $update.Title; HResult = "N/A" }
        }
    }

    $totalSuccessCount += $successCount
    $totalFailCount += $failCount
    $allInstalledKBs += $newKBs
    $allFailedKBs += $newFailedKBs

    Write-Host ""
    Show-Info "Result: $successCount succeeded, $failCount failed"

    # Check system-wide reboot pending (catches non-mandatory reboots missed by per-update check)
    if (-not $rebootRequired -and $successCount -gt 0) {
        $comType = [Type]::GetTypeFromProgID("Microsoft.Update.SystemInformation")
        if ($null -ne $comType) {
            try {
                $sysInfo = [Activator]::CreateInstance($comType)
                if ($sysInfo.RebootRequired) {
                    Show-Info "System-wide reboot pending detected"
                    $rebootRequired = $true
                }
            }
            catch { }
        }
    }

    Write-Host ""

    # Post-install decision (within same session, no reboot)
    if ($rebootRequired) {
        # Reboot needed — return to orchestrator (main.ps1) for reboot handling
        break
    }

    # No reboot needed: re-scan for cascading updates
    $noRebootScanCount++
    if ($noRebootScanCount -ge $maxNoRebootScans) {
        Show-Info "Max no-reboot re-scan limit reached ($maxNoRebootScans). Finishing."
        break
    }

    Show-Info "No reboot required. Re-scanning for additional updates..."
    Write-Host ""
    continue
}

# Sleep suppression cleanup is handled by Exit-Fabriq in main.ps1

# Stop WU transcript
if ($wuTranscriptStarted) {
    Show-Info "WU log saved: $wuLogFile"
    Stop-Transcript -ErrorAction SilentlyContinue | Out-Null
}

# Return result to orchestrator (main.ps1)
return @{
    Status         = if ($totalSuccessCount -gt 0 -and $totalFailCount -eq 0) { "Success" }
                     elseif ($totalSuccessCount -gt 0 -and $totalFailCount -gt 0) { "Partial" }
                     elseif ($totalSuccessCount -eq 0 -and $totalFailCount -gt 0) { "Error" }
                     else { "Success" }
    RebootRequired = $rebootRequired
    InstalledCount = $totalSuccessCount
    FailedCount    = $totalFailCount
    InstalledKBs   = @($allInstalledKBs)
    FailedKBs      = @($allFailedKBs)
    UpdatesFound   = $availableUpdates.Count
}
