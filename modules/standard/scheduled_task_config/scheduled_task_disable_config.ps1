# ========================================
# Scheduled Task Disable Script
# ========================================
# Disables scheduled tasks defined in task_list.csv.
# Tasks that are already disabled are skipped to
# maintain idempotency.
#
# [NOTES]
# - Requires administrator privileges
# - Tasks not found on the system are reported as failures
# ========================================

Write-Host ""
Show-Separator
Write-Host "Scheduled Task Disable" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# ========================================
# Step 1: CSV reading
# ========================================
$csvPath = Join-Path $PSScriptRoot "task_list.csv"

$enabledItems = Import-ModuleCsv -Path $csvPath -FilterEnabled `
    -RequiredColumns @("Enabled", "TaskPath", "TaskName", "Description")

if ($null -eq $enabledItems) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load task_list.csv")
}
if ($enabledItems.Count -eq 0) {
    return (New-ModuleResult -Status "Skipped" -Message "No enabled entries")
}

# ========================================
# Step 2: Pre-flight check
# ========================================
# Scheduled task cmdlets are built into Windows; no external dependency.

# ========================================
# Step 3: Pre-execution display
# ========================================
Show-Info "Disable targets: $($enabledItems.Count) task(s)"
Write-Host ""

foreach ($item in $enabledItems) {
    $task = Get-ScheduledTask -TaskName $item.TaskName -TaskPath $item.TaskPath -ErrorAction SilentlyContinue

    if (-not $task) {
        $marker = "[NOT FOUND]"
        $markerColor = "Red"
    }
    elseif ($task.State -eq "Disabled") {
        $marker = "[Disabled]"
        $markerColor = "DarkGray"
    }
    else {
        $marker = "[$($task.State)]"
        $markerColor = "Yellow"
    }

    Write-Host "  $($item.Description)  $marker" -ForegroundColor $markerColor
    Write-Host "    Task: $($item.TaskPath)$($item.TaskName)" -ForegroundColor DarkGray
    Write-Host ""
}

# ========================================
# Step 4: Confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Disable the above scheduled tasks?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""

# ========================================
# Step 5: Disable execution
# ========================================
$successCount = 0
$skipCount    = 0
$failCount    = 0

foreach ($item in $enabledItems) {
    $displayName = if ($item.Description) { $item.Description } else { $item.TaskName }

    # ----------------------------------------
    # Pre-check: task existence
    # ----------------------------------------
    $task = Get-ScheduledTask -TaskName $item.TaskName -TaskPath $item.TaskPath -ErrorAction SilentlyContinue

    if (-not $task) {
        Show-Warning "Task not found: $($item.TaskPath)$($item.TaskName)"
        $failCount++
        Write-Host ""
        continue
    }

    # ----------------------------------------
    # Idempotency: skip if already disabled
    # ----------------------------------------
    if ($task.State -eq "Disabled") {
        Show-Skip "Already disabled: $displayName"
        $skipCount++
        Write-Host ""
        continue
    }

    # ----------------------------------------
    # Main: disable task
    # ----------------------------------------
    try {
        $null = Disable-ScheduledTask -TaskName $item.TaskName -TaskPath $item.TaskPath -ErrorAction Stop
        Show-Success "Disabled: $displayName"
        $successCount++
    }
    catch {
        Show-Error "Failed to disable: $displayName - $_"
        $failCount++
    }

    Write-Host ""
}

# ========================================
# Step 6: Result
# ========================================
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount `
    -Title "Scheduled Task Disable Results")
