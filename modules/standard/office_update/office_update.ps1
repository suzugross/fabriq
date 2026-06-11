# ========================================
# Office Update Script
# ========================================
# Triggers Click-to-Run Office update and waits for completion
# using hybrid detection (version change + scenario registry + process).
#
# Prerequisites: Click-to-Run Office installed, network connectivity
# ========================================

function Test-C2RScenarioActive {
    # Scenario subkeys persist after completion (field-confirmed: the
    # INSTALL scenario from the original ODT install stays forever with
    # all tasks TASKSTATE_COMPLETED), so key existence does NOT mean
    # "running" - that assumption made the idle exit unreachable and
    # every no-update run a 60min false timeout. A scenario is active
    # only if a TasksState entry is neither COMPLETED, FAILED, nor
    # CANCELLED; unknown states count as active (timeout is the backstop).
    param([string]$ScenarioPath)

    $scenarioKeys = Get-ChildItem -Path $ScenarioPath -ErrorAction SilentlyContinue
    foreach ($sk in $scenarioKeys) {
        $ts = Get-ItemProperty -Path (Join-Path $sk.PSPath "TasksState") -ErrorAction SilentlyContinue
        if ($null -eq $ts) { continue }
        foreach ($prop in $ts.PSObject.Properties) {
            if ($prop.Name -like 'PS*') { continue }
            $state = "$($prop.Value)"
            if ($state -and $state -notmatch 'COMPLETED|FAILED|CANCELLED') { return $true }
        }
    }
    return $false
}

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

# Network check: single bounded probe instead of Wait-NetworkReady
# (which loops forever) - an unattended run must fail fast and let the
# FlexProfile dashboard / AutoPilot ErrorMode decide what to do next.
$netReachable = Test-Connection -ComputerName "8.8.8.8" -Count 2 -Quiet -ErrorAction SilentlyContinue
if (-not $netReachable) {
    Show-Error "Network unreachable (8.8.8.8)"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Network unreachable")
}
Show-Success "Network connectivity OK (8.8.8.8)"
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
# Baseline for the post-mortem "did detection even run" check in Step 6.
# The value is a FILETIME-derived decimal string (field-confirmed REG_SZ,
# e.g. 13425636454200); only before/after advancement matters, and a
# missing value (population-dependent) disables that check.
$updatesRegPath = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Updates"
$detectBefore = $null
try {
    $detectBefore = [int64](Get-ItemProperty -Path $updatesRegPath -Name "UpdateDetectionLastRunTime" -ErrorAction Stop).UpdateDetectionLastRunTime
}
catch { }

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

    # Signal 3: Scenario registry, state-aware (see Test-C2RScenarioActive)
    $hasActiveScenario = Test-C2RScenarioActive -ScenarioPath $scenarioPath

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
    # Idle exit without a version change: "no update available" and
    # "update failed silently" both land here. Post-mortem on the
    # ClickToRun\Updates key separates the detectable failure classes;
    # values are population-dependent (field-confirmed on 2 machines),
    # so a missing value falls through to Skip rather than inventing a
    # verdict.
    $upd = Get-ItemProperty -Path $updatesRegPath -ErrorAction SilentlyContinue

    # (1) The UPDATE scenario itself recorded an error - this also
    # catches a post-detection download abort (field-confirmed value;
    # '0' = no error)
    $lastUpdErr = "$((Get-ItemProperty -Path (Join-Path $scenarioPath 'UPDATE') -ErrorAction SilentlyContinue).LastUpdateError)".Trim()
    if ($lastUpdErr -ne '' -and $lastUpdErr -ne '0') {
        Show-Error "Update failed (LastUpdateError=$lastUpdErr)"
        Write-Host ""
        return (New-ModuleResult -Status "Error" -Message "Update failed (LastUpdateError=$lastUpdErr)")
    }

    # (2) Update downloaded/staged but never applied
    $readyToApply = "$($upd.UpdatesReadyToApply)".Trim()
    if ($readyToApply -ne '' -and $readyToApply -ne '0') {
        Show-Error "An update is staged but was not applied (UpdatesReadyToApply=$readyToApply)"
        Write-Host ""
        return (New-ModuleResult -Status "Error" -Message "Update staged but not applied (UpdatesReadyToApply=$readyToApply)")
    }

    # (3) Detection never ran - the client exited before even checking
    $detectAfter = $null
    try { $detectAfter = [int64]"$($upd.UpdateDetectionLastRunTime)" } catch { }
    if ($null -ne $detectBefore -and $null -ne $detectAfter -and $detectAfter -le $detectBefore) {
        Show-Error "Update client exited without running detection (UpdateDetectionLastRunTime unchanged)"
        Write-Host ""
        return (New-ModuleResult -Status "Error" -Message "Update detection did not run (client exited early)")
    }

    # (4) Detection ran, no error recorded, nothing staged, version
    # unchanged -> treat as up to date. The Skip is not sticky - a
    # re-run retries the update.
    Show-Info "No version change detected (treated as up to date: $beforeVersion)"
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "No version change (treated as up to date: $beforeVersion)")
}
