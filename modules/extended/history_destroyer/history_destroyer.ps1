# ========================================
# History Destroyer - CSV-Driven History Cleanup
# ========================================
# Deletes various Windows history, cache, and temporary data
# based on destroy_list.csv configuration.
#
# NOTES:
# - Requires administrator privileges for some operations
# - Explorer will be temporarily stopped during cleanup
# - Special handlers manage complex cleanup operations (browsers, Office, etc.)
# ========================================

Write-Host ""
Show-Separator
Write-Host "History Destroyer" -ForegroundColor Cyan
Show-Separator
Write-Host ""


# ========================================
# [P/Invoke] used for fully emptying the Recycle Bin
# ========================================
Add-Type -MemberDefinition @'
    [DllImport("Shell32.dll")]
    public static extern int SHEmptyRecycleBin(IntPtr hwnd, string pszRootPath, int dwFlags);
'@ -Name Win32RecycleBin -Namespace HistoryDestroyer -ErrorAction SilentlyContinue


# ========================================
# Private function group: Special handlers
# ========================================

# ----------------------------------------
# Dispatcher: route TargetPath -> matching handler
# Returns "Success" or "Skip" (handlers throw on failure)
# ----------------------------------------
function Invoke-DestroyHandler {
    param([string]$HandlerName)
    switch ($HandlerName) {
        "clear-all-eventlogs" { return (Clear-AllEventLogs) }
        "recycle-bin"         { return (Clear-RecycleBinSafe) }
        "office-mru"          { return (Clear-OfficeMRU) }
        "edge-cleanup"        { return (Clear-BrowserData -Browser "Edge" -BasePath "$env:LOCALAPPDATA\Microsoft\Edge\User Data" -ProcessName "msedge") }
        "chrome-cleanup"      { return (Clear-BrowserData -Browser "Chrome" -BasePath "$env:LOCALAPPDATA\Google\Chrome\User Data" -ProcessName "chrome") }
        "search-index"        { return (Clear-SearchIndex) }
        "wifi-ssid"           { return (Clear-WiFiProfiles) }
        default               { throw "Unknown special handler: $HandlerName" }
    }
}

# ----------------------------------------
# (1) Clear every event log
# ----------------------------------------
function Clear-AllEventLogs {
    $logs = Get-WinEvent -ListLog * -Force -ErrorAction SilentlyContinue
    if ($null -eq $logs -or $logs.Count -eq 0) {
        Show-Skip "No event logs found"
        return "Skip"
    }

    $clearedCount = 0
    foreach ($log in $logs) {
        $null = & wevtutil.exe cl $log.LogName 2>&1
        if ($LASTEXITCODE -eq 0) { $clearedCount++ }
    }

    Show-Success "Cleared $clearedCount event logs"
    return "Success"
}

# ----------------------------------------
# (2) Empty Recycle Bin (P/Invoke)
# ----------------------------------------
function Clear-RecycleBinSafe {
    # Flags: SHERB_NOCONFIRMATION(1) | SHERB_NOPROGRESSUI(2) | SHERB_NOSOUND(4) = 7
    $result = [HistoryDestroyer.Win32RecycleBin]::SHEmptyRecycleBin([IntPtr]::Zero, $null, 7)

    if ($result -eq 0) {
        Show-Success "Recycle Bin emptied"
    }
    else {
        # HRESULT values like -2147418113 (0x8000FFFF) just mean the bin was already empty
        Show-Info "Recycle Bin already empty or emptied (HRESULT: $result)"
    }
    return "Success"
}

# ----------------------------------------
# (3) Enumerate Office MRU registry keys dynamically and delete them
# ----------------------------------------
function Clear-OfficeMRU {
    $officeBase = "HKCU:\Software\Microsoft\Office"
    if (-not (Test-Path $officeBase)) {
        Show-Skip "Office registry not found"
        return "Skip"
    }

    $officeCleaned = 0
    $apps = @("Word", "Excel", "PowerPoint", "Access", "Publisher", "Visio")

    Get-ChildItem $officeBase -ErrorAction SilentlyContinue | ForEach-Object {
        $version = $_.PSChildName
        foreach ($app in $apps) {
            $placeMRU = "$officeBase\$version\$app\Place MRU"
            $fileMRU  = "$officeBase\$version\$app\File MRU"

            if (Test-Path $placeMRU) {
                $null = Remove-ItemProperty -Path $placeMRU -Name * -Force -ErrorAction SilentlyContinue
                $officeCleaned++
            }
            if (Test-Path $fileMRU) {
                $null = Remove-ItemProperty -Path $fileMRU -Name * -Force -ErrorAction SilentlyContinue
                $officeCleaned++
            }
        }
    }

    Show-Success "Office MRU cleaned ($officeCleaned entries)"
    return "Success"
}

# ----------------------------------------
# (4)(5) Browser data cleanup (shared Edge / Chrome implementation)
# ----------------------------------------
function Clear-BrowserData {
    param(
        [string]$Browser,
        [string]$BasePath,
        [string]$ProcessName
    )

    if (-not (Test-Path $BasePath)) {
        Show-Skip "$Browser not found"
        return "Skip"
    }

    # Stop the browser process
    $proc = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue
    if ($proc) {
        try {
            Stop-Process -Name $ProcessName -Force -ErrorAction Stop
            Start-Sleep -Seconds 2
        }
        catch {
            Show-Warning "Failed to stop $Browser process: $($_.Exception.Message)"
        }
    }

    # Targets to delete (14 entries from the original implementation)
    $browserTargets = @(
        "Cache", "Code Cache", "GPUCache",
        "History", "Cookies", "Cookies-journal",
        "Top Sites", "Top Sites-journal",
        "Visited Links",
        "Web Data", "Web Data-journal",
        "Session Storage", "Local Storage"
    )

    # Enumerate every profile (Default, Profile 1, Profile 2, ...)
    $profiles = Get-ChildItem $BasePath -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -eq "Default" -or $_.Name -match "^Profile " }

    $cleanedCount = 0
    foreach ($profile in $profiles) {
        foreach ($target in $browserTargets) {
            $targetPath = Join-Path $profile.FullName $target
            if (Test-Path $targetPath) {
                try {
                    Remove-Item $targetPath -Recurse -Force -ErrorAction Stop
                    $cleanedCount++
                }
                catch {
                    # Ignore locked files and continue
                }
            }
        }
    }

    Show-Success "$Browser data cleaned ($cleanedCount items)"
    return "Success"
}

# ----------------------------------------
# (6) Rebuild the Windows Search index
# ----------------------------------------
function Clear-SearchIndex {
    $wsearchService = Get-Service -Name "WSearch" -ErrorAction SilentlyContinue
    if (-not $wsearchService) {
        Show-Skip "Windows Search service not found"
        return "Skip"
    }

    $searchDbPath = "$env:ProgramData\Microsoft\Search\Data\Applications\Windows\Windows.edb"

    try {
        # Stop the service
        if ($wsearchService.Status -eq "Running") {
            Stop-Service -Name "WSearch" -Force -ErrorAction Stop
            Start-Sleep -Seconds 2
        }

        # Delete the index database
        if (Test-Path $searchDbPath) {
            Remove-Item $searchDbPath -Force -ErrorAction Stop
            Show-Success "Search index deleted"
        }
        else {
            Show-Info "Search index file not found (already clean)"
        }
    }
    catch {
        # On error, still try to restart the service before re-throwing
        $null = Start-Service -Name "WSearch" -ErrorAction SilentlyContinue
        throw
    }
    finally {
        # Guarantee service restart on both success and failure paths
        $null = Start-Service -Name "WSearch" -ErrorAction SilentlyContinue
    }

    return "Success"
}

# ----------------------------------------
# (7) Delete kitting-time Wi-Fi profiles (ssid_list.csv)
# ----------------------------------------
function Clear-WiFiProfiles {
    $ssidCsvPath = Join-Path $PSScriptRoot "ssid_list.csv"

    if (-not (Test-Path $ssidCsvPath)) {
        Show-Skip "ssid_list.csv not found"
        return "Skip"
    }

    $ssidItems = Import-ModuleCsv -Path $ssidCsvPath -FilterEnabled
    if ($null -eq $ssidItems -or $ssidItems.Count -eq 0) {
        # Import-ModuleCsv -FilterEnabled already printed the Skip message
        return "Skip"
    }

    # Confirm the Wi-Fi service is available
    $wlanSvc = Get-Service -Name "WlanSvc" -ErrorAction SilentlyContinue
    if (-not $wlanSvc -or $wlanSvc.Status -ne "Running") {
        Show-Skip "Wi-Fi service (WlanSvc) not available on this device"
        return "Skip"
    }

    $ssidDeleted = 0
    $ssidSkipped = 0
    $ssidErrors  = 0

    foreach ($ssidItem in $ssidItems) {
        $ssidName = $ssidItem.SSID
        $label = if ($ssidItem.Description) { "$ssidName ($($ssidItem.Description))" } else { $ssidName }

        # Idempotency: confirm the profile exists before deleting
        $null = & netsh wlan show profile name="$ssidName" 2>&1
        if ($LASTEXITCODE -ne 0) {
            Show-Skip "Not found: $label"
            $ssidSkipped++
            continue
        }

        # Delete
        $null = & netsh wlan delete profile name="$ssidName" 2>&1
        if ($LASTEXITCODE -eq 0) {
            Show-Success "Deleted: $label"
            $ssidDeleted++
        }
        else {
            Show-Error "Failed to delete: $label"
            $ssidErrors++
        }
    }

    Show-Info "SSID cleanup: $ssidDeleted deleted, $ssidSkipped not found, $ssidErrors failed"

    if ($ssidErrors -gt 0) {
        throw "SSID cleanup had $ssidErrors failure(s)"
    }
    return "Success"
}


# ========================================
# Step 0: Stop Explorer (prerequisite for every other operation)
# ========================================
# Stopping Explorer releases the file locks it holds.
# Run this once procedurally, outside the main loop.
# ========================================
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


# ========================================
# Step 1: Load CSV
# ========================================
$csvPath = Join-Path $PSScriptRoot "destroy_list.csv"

$enabledItems = Import-ModuleCsv -Path $csvPath -FilterEnabled `
    -RequiredColumns @("Enabled", "GroupName", "TargetName", "ActionType", "TargetPath")

if ($null -eq $enabledItems) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load destroy_list.csv")
}
if ($enabledItems.Count -eq 0) {
    return (New-ModuleResult -Status "Skipped" -Message "No enabled entries")
}

# Expand environment variables in DeletePath targets (file_delete-style)
foreach ($item in $enabledItems) {
    if ($item.ActionType -eq "DeletePath") {
        $item.TargetPath = Expand-UserEnvironmentVariables $item.TargetPath
    }
}

# Destructive path guard (CLAUDE.md section 8): evaluate once per
# DeletePath row here; consumed by both the preview display and the
# execution loop. Blocked rows are recorded as Fail (config error),
# never deleted — the guard sits OUTSIDE the confirmation gate so it
# also holds under AutoPilot auto-confirm.
foreach ($item in $enabledItems) {
    if ($item.ActionType -eq "DeletePath") {
        $guard = Test-FabriqProtectedPath -Path $item.TargetPath
        $item | Add-Member -NotePropertyName "_GuardBlocked" -NotePropertyValue (-not $guard.IsSafe)
        $item | Add-Member -NotePropertyName "_GuardReason"  -NotePropertyValue $guard.Reason
    }
}


# ========================================
# Step 2: Prerequisite check (early return)
# ========================================
# fabriq is always launched with admin privileges, so no extra
# prerequisite check is needed here.


# ========================================
# Step 3: Dry-run summary before execution
# ========================================
Show-Info "Cleanup targets: $($enabledItems.Count) items"
Write-Host ""

Write-Host "========================================" -ForegroundColor Yellow
Write-Host "Destruction Targets" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

# Group by GroupName for display
$groups = $enabledItems | Group-Object -Property GroupName

foreach ($group in $groups) {
    Write-Host "  [$($group.Name)]" -ForegroundColor White
    foreach ($item in $group.Group) {
        $displayName = if ($item.Description) { $item.Description } else { $item.TargetName }
        if ($item._GuardBlocked) {
            Write-Host "    [BLOCKED] $displayName" -ForegroundColor Red
            Write-Host "      $($item.ActionType): $($item.TargetPath)" -ForegroundColor DarkGray
            Write-Host "      Reason: $($item._GuardReason) - will be recorded as Fail" -ForegroundColor Red
        }
        else {
            Write-Host "    [DESTROY] $displayName" -ForegroundColor Yellow
            Write-Host "      $($item.ActionType): $($item.TargetPath)" -ForegroundColor DarkGray
        }
    }
    Write-Host ""
}

Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""


# ========================================
# Step 4: User confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Proceed with history destruction?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""


# ========================================
# Step 5: Main processing loop
# ========================================
$successCount = 0
$skipCount    = 0
$failCount    = 0
$total        = $enabledItems.Count
$current      = 0

foreach ($item in $enabledItems) {
    $current++
    $displayName = if ($item.Description) { $item.Description } else { $item.TargetName }
    $ifNotFound  = if ($item.IfNotFound) { $item.IfNotFound } else { "Skip" }

    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "[$current/$total] $displayName" -ForegroundColor Cyan
    Write-Host "  $($item.ActionType): $($item.TargetPath)" -ForegroundColor DarkGray
    Write-Host "----------------------------------------" -ForegroundColor White

    try {
        switch ($item.ActionType) {

            "DeletePath" {
                # Destructive path guard (CLAUDE.md section 8): blocked
                # targets are config errors - record Fail, never delete.
                if ($item._GuardBlocked) {
                    Show-Error "Blocked protected path: $($item.TargetPath) ($($item._GuardReason))"
                    $failCount++
                    Write-Host ""
                    continue
                }

                # Best-effort delete: probe -> delete what we can -> verify residue
                if (-not (Test-Path $item.TargetPath)) {
                    if ($ifNotFound -eq "Error") {
                        Show-Error "Target not found: $($item.TargetPath)"
                        $failCount++
                    }
                    else {
                        Show-Skip "Not found - skipped"
                        $skipCount++
                    }
                    Write-Host ""
                    continue
                }

                # SilentlyContinue: skip locked files, delete everything else
                Remove-Item -Path $item.TargetPath -Force -Recurse -ErrorAction SilentlyContinue

                # Post-check: decide Success / Warning based on residue
                if (Test-Path $item.TargetPath) {
                    Show-Warning "Partially deleted: $displayName (some files in use)"
                }
                else {
                    Show-Success "Deleted: $displayName"
                }
                $successCount++
            }

            "ClearRegistry" {
                # Clear registry key values (preserve the key tree, drop only the values)
                if (-not (Test-Path $item.TargetPath)) {
                    if ($ifNotFound -eq "Error") {
                        Show-Error "Registry key not found: $($item.TargetPath)"
                        $failCount++
                    }
                    else {
                        Show-Skip "Registry key not found - skipped"
                        $skipCount++
                    }
                    Write-Host ""
                    continue
                }

                # Clear the immediate properties (values)
                $null = Remove-ItemProperty -Path $item.TargetPath -Name * -Force -ErrorAction SilentlyContinue

                # If subkeys exist, also clear their values (still preserving the key tree)
                $subKeys = Get-ChildItem -Path $item.TargetPath -ErrorAction SilentlyContinue
                foreach ($subKey in $subKeys) {
                    $null = Remove-ItemProperty -Path $subKey.PSPath -Name * -Force -ErrorAction SilentlyContinue
                }

                Show-Success "Registry cleared: $displayName"
                $successCount++
            }

            "Command" {
                # Execute a simple one-liner command
                $null = Invoke-Expression $item.TargetPath
                Show-Success "Command executed: $displayName"
                $successCount++
            }

            "Special" {
                # Invoke the handler via the dispatcher.
                # Return values: "Success" -> $successCount, "Skip" -> $skipCount,
                # throw -> caught below as $failCount.
                $handlerResult = Invoke-DestroyHandler -HandlerName $item.TargetPath
                if ($handlerResult -eq "Skip") {
                    $skipCount++
                }
                else {
                    $successCount++
                }
            }

            default {
                Show-Error "Unknown ActionType: $($item.ActionType)"
                $failCount++
            }
        }
    }
    catch {
        Show-Error "Failed: $displayName : $_"
        $failCount++
    }

    Write-Host ""
}


# ========================================
# Final step: restart Explorer
# ========================================
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
    # Only force-start when Windows did not auto-revive Explorer in time
    Start-Process "explorer.exe"
    Show-Warning "Explorer auto-restart timed out. Started manually."
}
Write-Host ""


# ========================================
# Step 6: Aggregate and return result
# ========================================
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount `
    -Title "History Destroyer Results")
