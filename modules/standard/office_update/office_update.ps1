# ========================================
# Office Update Script
# ========================================
# Triggers Click-to-Run Office update and waits for completion
# using hybrid detection (version change + scenario registry + process).
#
# Prerequisites: Click-to-Run Office installed, network connectivity
# ========================================

Write-Host ""
Show-Separator
Write-Host "Office Update" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# ========================================
# Step 1: Load Configuration
# ========================================
$csvPath = Join-Path $PSScriptRoot "office_update_list.csv"

$configItems = Import-ModuleCsv -Path $csvPath -FilterEnabled `
    -RequiredColumns @("Enabled", "SettingName", "Value")

if ($null -eq $configItems) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load office_update_list.csv")
}

# Parse SettingName/Value pairs into hashtable
$config = @{}
foreach ($item in $configItems) {
    $config[$item.SettingName] = $item.Value
}

$timeoutMinutes    = if ($config["TimeoutMinutes"])      { [int]$config["TimeoutMinutes"] }      else { 60 }
$pollIntervalSec   = if ($config["PollIntervalSeconds"]) { [int]$config["PollIntervalSeconds"] } else { 10 }
$forceAppShutdown  = ($config["ForceAppShutdown"] -eq "1")
$displayLevel      = ($config["DisplayLevel"] -eq "1")


# ========================================
# Step 2: Prerequisite Checks
# ========================================

# Check Click-to-Run installation
$c2rConfigPath = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
$c2rConfig = Get-ItemProperty -Path $c2rConfigPath -ErrorAction SilentlyContinue

if ($null -eq $c2rConfig) {
    Show-Skip "Click-to-Run Office not installed"
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "Click-to-Run Office not installed")
}

# Check OfficeC2RClient.exe
$c2rClientPath = "C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe"
if (-not (Test-Path $c2rClientPath)) {
    Show-Error "OfficeC2RClient.exe not found: $c2rClientPath"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "OfficeC2RClient.exe not found")
}

# Record current version
$beforeVersion = $c2rConfig.VersionToReport
$productIds    = $c2rConfig.ProductReleaseIds
$updateChannel = $c2rConfig.UpdateChannel

if ([string]::IsNullOrWhiteSpace($beforeVersion)) {
    Show-Error "Unable to read current Office version from registry"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Cannot read VersionToReport")
}

# Wait for network
Wait-NetworkReady
Write-Host ""


# ========================================
# Step 3: Dry-Run Display
# ========================================
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "Office Update Settings" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""
Write-Host "  Product:     $productIds" -ForegroundColor White
Write-Host "  Version:     $beforeVersion" -ForegroundColor White
Write-Host "  Channel:     $updateChannel" -ForegroundColor White
Write-Host "  Executable:  $c2rClientPath" -ForegroundColor DarkGray
Write-Host "  Timeout:     $timeoutMinutes minutes" -ForegroundColor DarkGray
Write-Host "  Display UI:  $(if ($displayLevel) { 'Yes' } else { 'No' })" -ForegroundColor DarkGray
Write-Host "  Force close: $(if ($forceAppShutdown) { 'Yes' } else { 'No' })" -ForegroundColor DarkGray
Write-Host ""
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""


# ========================================
# Step 4: User Confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Trigger Office update?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""


# ========================================
# Step 5: Execute Update + Wait
# ========================================

# 5a. Force-close Office applications
if ($forceAppShutdown) {
    $officeProcesses = @("WINWORD", "EXCEL", "POWERPNT", "OUTLOOK", "ONENOTE", "MSACCESS", "MSPUB", "VISIO")
    $closedApps = @()

    foreach ($procName in $officeProcesses) {
        $procs = Get-Process -Name $procName -ErrorAction SilentlyContinue
        if ($procs) {
            $procs | Stop-Process -Force -ErrorAction SilentlyContinue
            $closedApps += $procName
        }
    }

    if ($closedApps.Count -gt 0) {
        Show-Info "Closed Office apps: $($closedApps -join ', ')"
    }
    else {
        Show-Info "No Office apps running"
    }
    Write-Host ""
}

# 5b. Trigger update
$displayArg = if ($displayLevel) { "displaylevel=True" } else { "displaylevel=False" }
$shutdownArg = if ($forceAppShutdown) { "forceappshutdown=True" } else { "forceappshutdown=False" }
$c2rArgs = "/update user $displayArg $shutdownArg updatepromptuser=False"

Show-Info "Triggering Office update..."
Show-Info "Command: OfficeC2RClient.exe $c2rArgs"

try {
    Start-Process -FilePath $c2rClientPath -ArgumentList $c2rArgs -ErrorAction Stop
}
catch {
    Show-Error "Failed to start OfficeC2RClient.exe: $_"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Failed to start update: $_")
}

Write-Host ""

# 5c. Wait for completion (hybrid detection)
$timeoutSec = $timeoutMinutes * 60
$elapsed = 0
$minWaitSec = 30  # minimum wait before checking idle state (avoid false negative on startup)
$scenarioPath = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Scenario"

Show-Info "Waiting for update to complete (timeout: ${timeoutMinutes}min, polling: ${pollIntervalSec}s)..."
Write-Host ""

while ($elapsed -lt $timeoutSec) {
    Start-Sleep -Seconds $pollIntervalSec
    $elapsed += $pollIntervalSec

    # Signal 1: Version change (definitive completion)
    $currentVersion = (Get-ItemProperty -Path $c2rConfigPath -Name "VersionToReport" -ErrorAction SilentlyContinue).VersionToReport
    if ($currentVersion -ne $beforeVersion) {
        Show-Success "Version changed: $beforeVersion -> $currentVersion"
        break
    }

    # Signal 2: OfficeC2RClient process check
    $c2rProc = Get-Process -Name "OfficeC2RClient" -ErrorAction SilentlyContinue
    $c2rActive = ($null -ne $c2rProc)

    # Signal 3: Scenario registry (active during update operations)
    $scenarioKeys = Get-ChildItem -Path $scenarioPath -ErrorAction SilentlyContinue
    $hasActiveScenario = ($null -ne $scenarioKeys -and $scenarioKeys.Count -gt 0)

    # Progress display (every 30 seconds)
    if ($elapsed % 30 -eq 0) {
        $timestamp = Get-Date -Format "HH:mm:ss"
        $statusParts = @()
        if ($c2rActive) { $statusParts += "C2R process active" }
        if ($hasActiveScenario) { $statusParts += "Scenario active" }
        if ($statusParts.Count -eq 0) { $statusParts += "idle" }
        $statusText = $statusParts -join ", "
        Write-Host "  [$timestamp] ${elapsed}s elapsed - $statusText" -ForegroundColor DarkGray
    }

    # Completion check: no C2R process AND no active scenario (after minimum wait)
    if (-not $c2rActive -and -not $hasActiveScenario -and $elapsed -ge $minWaitSec) {
        Show-Info "Update process completed (no active C2R process or scenario)"
        break
    }
}

Write-Host ""


# ========================================
# Step 6: Result
# ========================================
$afterVersion = (Get-ItemProperty -Path $c2rConfigPath -Name "VersionToReport" -ErrorAction SilentlyContinue).VersionToReport

if ($afterVersion -ne $beforeVersion) {
    Show-Success "Office updated: $beforeVersion -> $afterVersion"
    Write-Host ""
    return (New-ModuleResult -Status "Success" -Message "Updated: $beforeVersion -> $afterVersion")
}
elseif ($elapsed -ge $timeoutSec) {
    Show-Error "Update timed out after $timeoutMinutes minutes"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Update timeout (${timeoutMinutes}min)")
}
else {
    Show-Info "No update available (already up to date: $beforeVersion)"
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "Already up to date: $beforeVersion")
}
