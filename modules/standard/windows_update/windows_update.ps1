# ========================================
# Windows Update Automation Script
# ========================================
# Scans, downloads, and installs all available Windows Updates
# using the Microsoft.Update.Session COM API.
# Supports self-contained reboot loops with state persistence.
#
# Prerequisites: Administrator privileges, network connectivity
# ========================================

Write-Host ""
Show-Separator
Write-Host "Windows Update" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# ========================================
# Step 1: Load Configuration + State
# ========================================
$csvPath = Join-Path $PSScriptRoot "windows_update_list.csv"

$configItems = Import-ModuleCsv -Path $csvPath -FilterEnabled `
    -RequiredColumns @("Enabled", "SettingName", "Value")

if ($null -eq $configItems) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load windows_update_list.csv")
}

# Parse SettingName/Value pairs into hashtable
$config = @{}
foreach ($item in $configItems) {
    $config[$item.SettingName] = $item.Value
}

$maxLoops         = if ($config["MaxRebootLoops"])         { [int]$config["MaxRebootLoops"] }         else { 5 }
$scanTimeout      = if ($config["ScanTimeoutMinutes"])     { [int]$config["ScanTimeoutMinutes"] }     else { 30 }
$downloadTimeout  = if ($config["DownloadTimeoutMinutes"]) { [int]$config["DownloadTimeoutMinutes"] } else { 60 }
$installTimeout   = if ($config["InstallTimeoutMinutes"])  { [int]$config["InstallTimeoutMinutes"] }  else { 120 }
$suspendBitLocker = ($config["SuspendBitLocker"] -eq "1")
$rebootCountdown  = if ($config["RebootCountdownSeconds"]) { [int]$config["RebootCountdownSeconds"] } else { 15 }
$autoLaunchFabriq = ($config["AutoLaunchFabriq"] -eq "1")
$autoLogonEnabled = ($config["AutoLogonEnabled"] -eq "1")
$includeOptional  = ($config["IncludeOptionalUpdates"] -eq "1")
$includeSeeker    = ($config["IncludeSeekerUpdates"] -eq "1")

# Load state file (reboot loop persistence)
$statePath = Join-Path $PSScriptRoot "wu_state.json"
$isResumedLoop = $false

if (Test-Path $statePath) {
    try {
        $stateRaw = Get-Content $statePath -Raw -Encoding UTF8 | ConvertFrom-Json
        $isResumedLoop = $true

        # Convert to mutable structure
        $state = @{
            LoopCount        = [int]$stateRaw.LoopCount
            MaxLoops         = [int]$stateRaw.MaxLoops
            InstalledKBs     = @($stateRaw.InstalledKBs)
            StartTime        = $stateRaw.StartTime
            LastRebootReason = $stateRaw.LastRebootReason
            UserConfirmed    = [bool]$stateRaw.UserConfirmed
        }

        Show-Info "Resumed from reboot loop (Loop $($state.LoopCount) of $($state.MaxLoops))"
    }
    catch {
        Show-Warning "Failed to load wu_state.json, starting fresh: $_"
        $isResumedLoop = $false
    }
}

if (-not $isResumedLoop) {
    $state = @{
        LoopCount        = 0
        MaxLoops         = $maxLoops
        InstalledKBs     = @()
        StartTime        = (Get-Date).ToString("o")
        LastRebootReason = ""
        UserConfirmed    = $false
    }
}

# ========================================
# Step 2: Prerequisite Checks
# ========================================

# Admin privilege check
if (-not (Test-AdminPrivilege)) {
    return (New-ModuleResult -Status "Error" -Message "Administrator privileges required")
}

# Wait for base system services
Wait-SystemReady -RequiredServices @("LanmanWorkstation", "Dnscache")
Write-Host ""

# Start Windows Update service (Manual startup type, must be started explicitly)
$wuSvc = Get-Service -Name "wuauserv" -ErrorAction SilentlyContinue
if ($null -eq $wuSvc) {
    return (New-ModuleResult -Status "Error" -Message "Windows Update service (wuauserv) not found")
}
if ($wuSvc.Status -ne "Running") {
    Show-Info "Starting Windows Update service..."
    try {
        Start-Service -Name "wuauserv" -ErrorAction Stop
        Show-Success "Windows Update service started"
    }
    catch {
        Show-Error "Failed to start Windows Update service: $_"
        return (New-ModuleResult -Status "Error" -Message "Failed to start wuauserv: $_")
    }
}
else {
    Show-Info "Windows Update service is already running"
}
Write-Host ""

# Wait for network connectivity
Wait-NetworkReady
Write-Host ""

# Safety valve: max loop count
if ($state.LoopCount -ge $state.MaxLoops) {
    Show-Warning "Max reboot loop limit reached ($($state.MaxLoops)). Stopping."
    Write-Host ""

    # Clean up state file
    if (Test-Path $statePath) {
        Remove-Item $statePath -Force -ErrorAction SilentlyContinue
    }

    $totalInstalled = $state.InstalledKBs.Count
    $msg = "Max loop limit reached. $totalInstalled updates installed across $($state.LoopCount) loops"

    if ($isResumedLoop) {
        Write-ExecutionHistory -ModuleName "Windows Update" -Category "Maintenance" -Status "Partial" -Message $msg
    }

    return (New-ModuleResult -Status "Partial" -Message $msg)
}

# Prevent sleep during updates
Enable-SleepSuppression

# Suspend BitLocker if configured
if ($suspendBitLocker) {
    try {
        $blVolume = Get-BitLockerVolume -MountPoint "C:" -ErrorAction SilentlyContinue
        if ($blVolume -and $blVolume.ProtectionStatus -eq "On") {
            Suspend-BitLocker -MountPoint "C:" -RebootCount 1 -ErrorAction SilentlyContinue
            Show-Info "BitLocker suspended for 1 reboot cycle"
        }
    }
    catch {
        Show-Warning "BitLocker suspension failed (non-fatal): $_"
    }
}

# ========================================
# Step 3-6: Scan, Confirm, Install Loop
# ========================================
# Inner loop handles no-reboot re-scans
$maxNoRebootScans = 3
$noRebootScanCount = 0
$totalSuccessCount = 0
$totalFailCount = 0

# Create COM session once
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

while ($true) {
    # ========================================
    # Step 3: COM API Scan + Dry-Run Display
    # ========================================
    Show-Info "Scanning for available updates..."
    Write-Host ""

    # Primary scan (uses Microsoft Update service if configured)
    # When IncludeSeekerUpdates is enabled, expand criteria to include
    # OptionalInstallation updates (seeker/gradual rollout patches that
    # the default "IsInstalled=0" query does not return).
    $searchCriteria = "IsInstalled=0"
    if ($includeSeeker) {
        $searchCriteria = "IsInstalled=0 and DeploymentAction='Installation' or IsInstalled=0 and DeploymentAction='OptionalInstallation'"
    }
    try {
        $searchResult = $updateSearcher.Search($searchCriteria)
    }
    catch {
        Show-Error "Update scan failed: $_"
        Disable-SleepSuppression

        if ($isResumedLoop) {
            Write-ExecutionHistory -ModuleName "Windows Update" -Category "Maintenance" -Status "Error" -Message "Scan failed: $_"
        }

        return (New-ModuleResult -Status "Error" -Message "Update scan failed: $_")
    }

    $availableUpdates = $searchResult.Updates

    # No updates available -> all done
    if ($availableUpdates.Count -eq 0) {
        Show-Success "No updates available. System is up to date."
        Write-Host ""
        break
    }

    # Display available updates (dry-run)
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

    # ========================================
    # Step 4: User Confirmation
    # ========================================
    if (-not $state.UserConfirmed) {
        $cancelResult = Confirm-ModuleExecution -Message "Install all $($availableUpdates.Count) updates?"
        if ($null -ne $cancelResult) {
            Disable-SleepSuppression
            # Clean up state file if exists
            if (Test-Path $statePath) {
                Remove-Item $statePath -Force -ErrorAction SilentlyContinue
            }
            return $cancelResult
        }
        $state.UserConfirmed = $true
    }
    else {
        Show-Info "[AUTO-RESUME] Loop $($state.LoopCount + 1) of $($state.MaxLoops) - proceeding automatically"
    }

    Write-Host ""

    # ========================================
    # Step 5: Download + Install
    # ========================================

    # --- Download Phase (per-update for progress visibility) ---
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
            $update.AcceptEula()
            Show-Info "[$dlIndex/$downloadNeeded] Downloading: $($update.Title)"

            try {
                $singleDownload = New-Object -ComObject Microsoft.Update.UpdateColl
                $singleDownload.Add($update) | Out-Null
                $downloader.Updates = $singleDownload
                $dlResult = $downloader.Download()

                if ($dlResult.ResultCode -ge 2) {
                    Show-Success "Downloaded"
                    $downloadDone++
                }
                else {
                    Show-Error "Download failed (ResultCode: $($dlResult.ResultCode))"
                    $downloadFailed++
                }
            }
            catch {
                Show-Error "Download error: $_"
                $downloadFailed++
            }
        }

        Write-Host ""
        if ($downloadFailed -gt 0 -and $downloadDone -eq 0) {
            Show-Error "All downloads failed"
            Disable-SleepSuppression

            if ($isResumedLoop) {
                Write-ExecutionHistory -ModuleName "Windows Update" -Category "Maintenance" -Status "Error" -Message "All downloads failed"
            }

            return (New-ModuleResult -Status "Error" -Message "All downloads failed")
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

    # --- Install Phase (per-update for progress visibility) ---
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
    $rebootRequired = $false
    $newKBs = @()
    $instIndex = 0

    for ($i = 0; $i -lt $availableUpdates.Count; $i++) {
        $update = $availableUpdates.Item($i)
        if (-not $update.IsDownloaded) { continue }

        $instIndex++
        Show-Info "[$instIndex/$installTargetCount] Installing: $($update.Title)"

        # Extract KB number from title
        $kb = ""
        if ($update.Title -match '(KB\d+)') { $kb = $Matches[1] }

        try {
            $singleUpdate = New-Object -ComObject Microsoft.Update.UpdateColl
            $singleUpdate.Add($update) | Out-Null
            $installer.Updates = $singleUpdate
            $singleResult = $installer.Install()

            $itemResult = $singleResult.GetUpdateResult(0)

            if ($itemResult.ResultCode -eq 2) {
                Show-Success "$($update.Title)"
                $successCount++
                $newKBs += @{ KB = $kb; Title = $update.Title; Loop = ($state.LoopCount + 1) }
            }
            else {
                Show-Error "$($update.Title) (ResultCode: $($itemResult.ResultCode))"
                $failCount++
            }

            if ($itemResult.RebootRequired) { $rebootRequired = $true }
            if ($singleResult.RebootRequired) { $rebootRequired = $true }
        }
        catch {
            Show-Error "Install failed: $($update.Title) - $_"
            $failCount++
        }
    }

    $totalSuccessCount += $successCount
    $totalFailCount += $failCount

    Write-Host ""
    Show-Info "Iteration result: $successCount succeeded, $failCount failed"
    Write-Host ""

    # Update state with newly installed KBs
    $state.InstalledKBs = @($state.InstalledKBs) + $newKBs
    $state.LoopCount = $state.LoopCount + 1

    # ========================================
    # Step 6: Post-Install Decision
    # ========================================

    if ($rebootRequired -and $state.LoopCount -lt $state.MaxLoops) {
        # --- Reboot required: save state and restart ---
        $state.LastRebootReason = "Updates require restart"

        # Save state file
        $stateJson = [PSCustomObject]@{
            LoopCount        = $state.LoopCount
            MaxLoops         = $state.MaxLoops
            InstalledKBs     = $state.InstalledKBs
            StartTime        = $state.StartTime
            LastRebootReason = $state.LastRebootReason
            UserConfirmed    = $state.UserConfirmed
        } | ConvertTo-Json -Depth 5

        $stateJson | Out-File -FilePath $statePath -Encoding UTF8 -Force
        Show-Info "State saved: Loop $($state.LoopCount), Total KBs: $($state.InstalledKBs.Count)"

        # Register RunOnce for auto-resume after reboot
        $launcherPath = Join-Path $PSScriptRoot "wu_launcher.bat"
        $runOncePath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"
        $runOnceName = "FabriqWindowsUpdate"
        $runOnceValue = "cmd /c `"$launcherPath`""

        try {
            if (-not (Test-Path $runOncePath)) {
                New-Item -Path $runOncePath -Force | Out-Null
            }

            $existing = Get-ItemProperty -Path $runOncePath -Name $runOnceName -ErrorAction SilentlyContinue
            if ($existing) {
                Set-ItemProperty -Path $runOncePath -Name $runOnceName -Value $runOnceValue -Type String -Force -ErrorAction Stop
            }
            else {
                New-ItemProperty -Path $runOncePath -Name $runOnceName -Value $runOnceValue -PropertyType String -Force -ErrorAction Stop | Out-Null
            }

            Show-Success "RunOnce registered: $runOnceName"
        }
        catch {
            Show-Error "Failed to register RunOnce: $_"
            Disable-SleepSuppression
            return (New-ModuleResult -Status "Error" -Message "Failed to register RunOnce: $_")
        }

        # Set one-time AutoLogon before reboot (matching autologon_config pattern)
        if ($autoLogonEnabled) {
            $alCsvPath = Join-Path $PSScriptRoot "..\..\standard\autologon_config\autologon_list.csv"
            if (Test-Path $alCsvPath) {
                try {
                    $alEntries = Import-ModuleCsv -Path $alCsvPath -RequiredColumns @("Enabled", "No", "User", "Password")
                    $alEnabled = @($alEntries | Where-Object { $_.Enabled -eq "1" })
                    $currentUser = $env:USERNAME
                    $alTarget = $alEnabled | Where-Object { $_.User -eq $currentUser } | Select-Object -First 1

                    if ($alTarget) {
                        $winlogonPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
                        Set-ItemProperty -Path $winlogonPath -Name "AutoAdminLogon" -Value "1" -Type String -Force -ErrorAction Stop
                        Set-ItemProperty -Path $winlogonPath -Name "DefaultUserName" -Value $alTarget.User -Type String -Force -ErrorAction Stop
                        Set-ItemProperty -Path $winlogonPath -Name "DefaultPassword" -Value $alTarget.Password -Type String -Force -ErrorAction Stop
                        Set-ItemProperty -Path $winlogonPath -Name "AutoLogonCount" -Value 1 -Type DWord -Force -ErrorAction Stop
                        if (-not [string]::IsNullOrWhiteSpace($alTarget.Domain)) {
                            Set-ItemProperty -Path $winlogonPath -Name "DefaultDomainName" -Value $alTarget.Domain -Type String -Force -ErrorAction Stop
                        }
                        Show-Success "AutoLogon configured for '$currentUser' (one-time)"
                    }
                    else {
                        Show-Warning "AutoLogon: no matching entry for '$currentUser' in autologon_list.csv"
                    }
                }
                catch {
                    Show-Warning "AutoLogon configuration failed (non-fatal): $_"
                }
            }
            else {
                Show-Warning "AutoLogon: autologon_list.csv not found"
            }
        }

        Write-Host ""

        # Write execution history (wu_launcher.bat does not go through main.ps1)
        if ($isResumedLoop) {
            Write-ExecutionHistory -ModuleName "Windows Update" -Category "Maintenance" -Status "Success" -Message "Loop $($state.LoopCount): $successCount installed, rebooting"
        }

        # Return result BEFORE reboot (matches restart_config pattern)
        $result = New-ModuleResult -Status "Success" -Message "Loop $($state.LoopCount): $successCount installed, $failCount failed. Rebooting..."

        Invoke-CountdownRestart -Seconds $rebootCountdown

        return $result
    }
    elseif (-not $rebootRequired) {
        # --- No reboot needed: re-scan for cascading updates ---
        $noRebootScanCount++

        if ($noRebootScanCount -ge $maxNoRebootScans) {
            Show-Info "Max no-reboot re-scan limit reached ($maxNoRebootScans). Finishing."
            break
        }

        Show-Info "No reboot required. Re-scanning for additional updates..."
        Write-Host ""
        continue
    }
    else {
        # Max loops reached with reboot still required
        Show-Warning "Max reboot loop limit reached ($($state.MaxLoops)). Additional updates may remain."
        break
    }
}

# ========================================
# All Done: Final Summary
# ========================================

# Clean up state file
if (Test-Path $statePath) {
    Remove-Item $statePath -Force -ErrorAction SilentlyContinue
}

Disable-SleepSuppression

$allKBs = @($state.InstalledKBs)
$totalInstalled = $allKBs.Count
$elapsed = (Get-Date) - [datetime]$state.StartTime
$elapsedMinutes = [math]::Round($elapsed.TotalMinutes, 1)

# Display final summary
Show-Separator
Write-Host "Windows Update Complete" -ForegroundColor Cyan
Show-Separator
Write-Host ""
Write-Host "  Total loops:     $($state.LoopCount)" -ForegroundColor White
Write-Host "  Total installed: $totalInstalled updates" -ForegroundColor Green
Write-Host "  Total failed:    $totalFailCount updates" -ForegroundColor $(if ($totalFailCount -gt 0) { "Red" } else { "White" })
Write-Host "  Elapsed time:    ${elapsedMinutes} minutes" -ForegroundColor White
Write-Host ""

if ($state.LoopCount -ge $state.MaxLoops) {
    Show-Warning "Max loop limit reached. Additional updates may remain."
    Write-Host ""
}

if ($allKBs.Count -gt 0) {
    Write-Host "  Installed updates:" -ForegroundColor DarkGray
    foreach ($kb in $allKBs) {
        $kbId = if ($kb.KB) { $kb.KB } else { "N/A" }
        $kbTitle = if ($kb.Title) { $kb.Title } else { "Unknown" }
        Write-Host "    [Loop $($kb.Loop)] $kbId - $kbTitle" -ForegroundColor DarkGray
    }
    Write-Host ""
}

Show-Separator
Write-Host ""

# Write execution history for resumed loops
if ($isResumedLoop) {
    $historyMsg = "$totalInstalled updates installed ($($state.LoopCount) loops, ${elapsedMinutes}min)"
    $historyStatus = if ($totalFailCount -eq 0 -and $totalInstalled -gt 0) { "Success" }
        elseif ($totalInstalled -gt 0 -and $totalFailCount -gt 0) { "Partial" }
        elseif ($totalInstalled -eq 0 -and $totalFailCount -gt 0) { "Error" }
        else { "Success" }
    Write-ExecutionHistory -ModuleName "Windows Update" -Category "Maintenance" -Status $historyStatus -Message $historyMsg
}

# Save completion result for next fabriq session import
$completedPath = Join-Path $PSScriptRoot "wu_completed.json"
$completedData = [PSCustomObject]@{
    TotalInstalled = $totalInstalled
    TotalFailed    = $totalFailCount
    TotalLoops     = $state.LoopCount
    ElapsedMinutes = $elapsedMinutes
    InstalledKBs   = @($allKBs | ForEach-Object { $_.KB })
    CompletedAt    = (Get-Date).ToString("o")
} | ConvertTo-Json -Depth 3

$completedData | Out-File -FilePath $completedPath -Encoding UTF8 -Force
Show-Info "Completion results saved: wu_completed.json"
Write-Host ""

# Auto-launch Fabriq.bat if configured
if ($autoLaunchFabriq) {
    $fabriqRoot = (Resolve-Path (Join-Path $PSScriptRoot "..\..\..")).Path
    $fabriqBat = Join-Path $fabriqRoot "Fabriq.bat"

    if (Test-Path $fabriqBat) {
        Show-Info "Launching Fabriq.bat..."
        Start-Process cmd -ArgumentList "/c `"$fabriqBat`"" -Verb RunAs
    }
    else {
        Show-Warning "Fabriq.bat not found: $fabriqBat"
    }
}

return (New-BatchResult -Success $totalInstalled -Skip 0 -Fail $totalFailCount `
    -Title "Windows Update Results" `
    -MessageSuffix "(Loops: $($state.LoopCount))")
