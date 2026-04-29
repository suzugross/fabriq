# ========================================
# Process Killer Script
# ========================================
# Forcibly stop the processes listed in process_list.csv.
# Does nothing when the target process is not running (idempotent).
#
# [NOTES]
# - ProcessName is passed to Get-Process -Name (no .exe suffix)
# - Without admin privileges, processes owned by other users may not be killable
# ========================================

Write-Host ""
Show-Separator
Write-Host "Process Killer" -ForegroundColor Cyan
Show-Separator
Write-Host ""


# ========================================
# Step 1: Load CSV
# ========================================
$csvPath = Join-Path $PSScriptRoot "process_list.csv"

$enabledItems = Import-ModuleCsv -Path $csvPath -FilterEnabled `
    -RequiredColumns @("Enabled", "ProcessName", "Description")

if ($null -eq $enabledItems) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load process_list.csv")
}
if ($enabledItems.Count -eq 0) {
    return (New-ModuleResult -Status "Skipped" -Message "No enabled entries")
}


# ========================================
# Step 2: Prerequisite check
# ========================================
# Stopping a process does not depend on external resources; skip.


# ========================================
# Step 3: Dry-run summary before execution
# ========================================
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "Target Processes" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

foreach ($item in $enabledItems) {
    $processes = @(Get-Process -Name $item.ProcessName -ErrorAction SilentlyContinue)

    if ($processes.Count -gt 0) {
        Write-Host "  [Running] $($item.Description)" -ForegroundColor Yellow
        Write-Host "    Process: $($item.ProcessName)  ($($processes.Count) instance(s))" -ForegroundColor DarkGray
    }
    else {
        Write-Host "  [Not Running] $($item.Description)" -ForegroundColor DarkGray
        Write-Host "    Process: $($item.ProcessName)" -ForegroundColor DarkGray
    }
    Write-Host ""
}

Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""


# ========================================
# Step 4: User confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Terminate the above running processes?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""


# ========================================
# Step 5: Apply-settings loop
# ========================================
$successCount = 0
$skipCount    = 0
$failCount    = 0

foreach ($item in $enabledItems) {
    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "Processing: $($item.Description)" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor White

    # ----------------------------------------
    # Idempotency: skip when the process is not running
    # ----------------------------------------
    $processes = @(Get-Process -Name $item.ProcessName -ErrorAction SilentlyContinue)
    if ($processes.Count -eq 0) {
        Show-Skip "Already not running: $($item.ProcessName)"
        Write-Host ""
        $skipCount++
        continue
    }

    # ----------------------------------------
    # Main work: force-stop
    # ----------------------------------------
    try {
        Stop-Process -Name $item.ProcessName -Force -ErrorAction Stop
        Show-Success "Terminated: $($item.ProcessName)  ($($processes.Count) instance(s))"
        $successCount++
    }
    catch {
        Show-Error "Failed to terminate: $($item.ProcessName) : $_"
        $failCount++
    }

    Write-Host ""
}


# ========================================
# Step 6: Aggregate and return result
# ========================================
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount `
    -Title "Process Killer Results")
