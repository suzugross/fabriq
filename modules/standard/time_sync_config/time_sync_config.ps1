# ========================================
# Time Sync Config Script
# ========================================
# Configures the Windows Time service (W32Time) with NTP
# servers defined in time_sync_list.csv and triggers a
# time synchronization.
#
# [NOTES]
# - Requires administrator privileges
# - If no NTP servers are enabled in CSV, syncs with current config only
# - Uses w32tm for NTP configuration and resync
# - Appends ,0x9 flag to each server (SpecialPollInterval + Client)
# ========================================

Write-Host ""
Show-Separator
Write-Host "Time Sync" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# ========================================
# Step 1: CSV reading
# ========================================
$csvPath = Join-Path $PSScriptRoot "time_sync_list.csv"

# Load without -FilterEnabled to support 0-entry mode (sync-only)
$allItems = Import-ModuleCsv -Path $csvPath `
    -RequiredColumns @("Enabled", "NtpServer", "Description")

if ($null -eq $allItems) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load time_sync_list.csv")
}

$enabledItems = @($allItems | Where-Object { $_.Enabled -eq "1" })

# ========================================
# Step 2: Pre-flight check
# ========================================
$w32timeSvc = Get-Service -Name "W32Time" -ErrorAction SilentlyContinue
if (-not $w32timeSvc) {
    Show-Error "Windows Time service (W32Time) not found on this system."
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "W32Time service not found")
}

# ========================================
# Step 3: Pre-execution display
# ========================================

# Current service state
Show-Info "W32Time service: $($w32timeSvc.Status) (StartType: $($w32timeSvc.StartType))"
Write-Host ""

if ($enabledItems.Count -eq 0) {
    Show-Info "No NTP servers specified. Will sync with current configuration only."
    Write-Host ""
}
else {
    Show-Info "NTP servers: $($enabledItems.Count) server(s)"
    Write-Host ""

    foreach ($item in $enabledItems) {
        # Connectivity check (ICMP)
        $reachable = Test-Connection -ComputerName $item.NtpServer -Count 1 -Quiet -ErrorAction SilentlyContinue

        if ($reachable) {
            $marker = "[REACHABLE]"
            $markerColor = "Green"
        }
        else {
            $marker = "[UNREACHABLE]"
            $markerColor = "Yellow"
        }

        Write-Host "  $($item.Description)  $marker" -ForegroundColor $markerColor
        Write-Host "    Server: $($item.NtpServer)" -ForegroundColor DarkGray
        Write-Host ""
    }
}

# ========================================
# Step 4: Confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Configure NTP and synchronize time?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""

# ========================================
# Step 5: Execution
# ========================================

# ----------------------------------------
# 5-1: Ensure W32Time service is running
# ----------------------------------------
try {
    if ($w32timeSvc.Status -ne "Running") {
        Start-Service -Name "W32Time" -ErrorAction Stop
        Show-Success "W32Time service started"
    }
    else {
        Show-Info "W32Time service is already running"
    }

    $null = Set-Service -Name "W32Time" -StartupType Automatic -ErrorAction Stop
    Show-Info "W32Time startup type set to Automatic"
}
catch {
    Show-Error "Failed to start W32Time service: $_"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Failed to start W32Time service: $_")
}

Write-Host ""

# ----------------------------------------
# 5-2: Connectivity check (informational)
# ----------------------------------------
if ($enabledItems.Count -gt 0) {
    $reachableCount = 0
    foreach ($item in $enabledItems) {
        $reachable = Test-Connection -ComputerName $item.NtpServer -Count 1 -Quiet -ErrorAction SilentlyContinue
        if ($reachable) { $reachableCount++ }
    }

    if ($reachableCount -eq 0) {
        Show-Warning "All NTP servers are unreachable via ICMP. Proceeding anyway (ICMP may be blocked)."
    }
    else {
        Show-Info "Reachable servers: $reachableCount / $($enabledItems.Count)"
    }
    Write-Host ""
}

# ----------------------------------------
# 5-3: Configure NTP servers (skip if 0 entries)
# ----------------------------------------
if ($enabledItems.Count -gt 0) {
    # Build manualpeerlist with ,0x9 flag for each server
    $peerList = ($enabledItems | ForEach-Object { "$($_.NtpServer),0x9" }) -join " "

    Show-Info "Configuring NTP: $peerList"

    $configOutput = cmd /c "w32tm /config /manualpeerlist:`"$peerList`" /syncfromflags:manual /reliable:yes /update 2>&1"
    $configExitCode = $LASTEXITCODE

    if ($configOutput) {
        foreach ($line in $configOutput) {
            Write-Host "  $line" -ForegroundColor DarkGray
        }
    }

    if ($configExitCode -ne 0) {
        Show-Error "w32tm /config failed with exit code $configExitCode"
        Write-Host ""
        return (New-ModuleResult -Status "Error" -Message "w32tm /config failed (exit code: $configExitCode)")
    }

    Show-Success "NTP server configured"

    # Wait for configuration to propagate before resync
    Show-Info "Waiting for NTP configuration to take effect..."
    Start-Sleep -Seconds 3

    Write-Host ""
}
else {
    Show-Info "Skipping NTP server configuration (no servers specified)"
    Write-Host ""
}

# ----------------------------------------
# 5-4: Trigger time synchronization (with retry)
# ----------------------------------------
$maxRetries  = 5
$retryWait   = 3
$syncSuccess = $false

for ($i = 1; $i -le $maxRetries; $i++) {
    if ($i -eq 1) {
        Show-Info "Triggering time synchronization..."
    }
    else {
        Show-Info "Retrying time synchronization... ($i/$maxRetries)"
    }

    # Execute resync (output is discarded — judgment is based on /query /status)
    $null = w32tm /resync /force 2>&1

    # Query current sync status
    $status = w32tm /query /status 2>&1 | Out-String

    # Display status
    foreach ($line in ($status.Trim() -split "\r?\n")) {
        Write-Host "  $line" -ForegroundColor DarkGray
    }

    # Check if source is still local clock (not yet synced to NTP)
    $isLocalClock = ($status -like "*Local CMOS Clock*") -or
                    ($status -like "*LOCL*") -or
                    ($status -like "*Free-running System Clock*")

    if (-not $isLocalClock) {
        # Source has switched to an NTP server → success
        Show-Success "Time synchronization completed"
        $syncSuccess = $true
        break
    }

    # Still local clock — retry if attempts remain
    if ($i -lt $maxRetries) {
        Show-Warning "Source is still local clock. Waiting $retryWait seconds before retry..."
        Start-Sleep -Seconds $retryWait
    }
}

if (-not $syncSuccess) {
    Show-Warning "Sync source did not change after $maxRetries attempts. Proceeding anyway."
}

Write-Host ""

# ----------------------------------------
# 5-5: Display sync status (informational)
# ----------------------------------------
Show-Info "Current sync status:"

$finalStatus = w32tm /query /status 2>&1 | Out-String
foreach ($line in ($finalStatus.Trim() -split "\r?\n")) {
    Write-Host "  $line" -ForegroundColor DarkGray
}

Write-Host ""

# ========================================
# Step 6: Result
# ========================================
if ($enabledItems.Count -gt 0) {
    $serverNames = ($enabledItems | ForEach-Object { $_.NtpServer }) -join ", "
    $msg = "NTP configured ($serverNames) and time synced"
}
else {
    $msg = "Time synced with current configuration"
}

if (-not $syncSuccess) {
    return (New-ModuleResult -Status "Partial" -Message "$msg (sync not confirmed after $maxRetries attempts)")
}

return (New-ModuleResult -Status "Success" -Message $msg)
