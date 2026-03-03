# ========================================
# Scheduled Task Enable Script
# ========================================
# Enables scheduled tasks defined in task_list.csv.
# Tasks that are already enabled (Ready/Running) are
# skipped to maintain idempotency.
#
# [NOTES]
# - Requires administrator privileges
# - Tasks not found on the system are reported as failures
# ========================================

Write-Host ""
Show-Separator
Write-Host "Scheduled Task Enable" -ForegroundColor Cyan
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
Show-Info "Enable targets: $($enabledItems.Count) task(s)"
Write-Host ""

foreach ($item in $enabledItems) {
    $task = Get-ScheduledTask -TaskName $item.TaskName -TaskPath $item.TaskPath -ErrorAction SilentlyContinue

    if (-not $task) {
        $marker = "[NOT FOUND]"
        $markerColor = "Red"
    }
    elseif ($task.State -eq "Disabled") {
        $marker = "[Disabled]"
        $markerColor = "Yellow"
    }
    else {
        $marker = "[$($task.State)]"
        $markerColor = "DarkGray"
    }

    Write-Host "  $($item.Description)  $marker" -ForegroundColor $markerColor
    Write-Host "    Task: $($item.TaskPath)$($item.TaskName)" -ForegroundColor DarkGray
    Write-Host ""
}

# ========================================
# Step 4: Confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Enable the above scheduled tasks?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""

# ========================================
# Step 5: Enable execution
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
    # Idempotency: skip if already enabled
    # ----------------------------------------
    if ($task.State -ne "Disabled") {
        Show-Skip "Already enabled: $displayName ($($task.State))"
        $skipCount++
        Write-Host ""
        continue
    }

    # ----------------------------------------
    # Main: enable task
    # ----------------------------------------
    try {
        $null = Enable-ScheduledTask -TaskName $item.TaskName -TaskPath $item.TaskPath -ErrorAction Stop
        Show-Success "Enabled: $displayName"
        $successCount++
    }
    catch {
        Show-Error "Failed to enable: $displayName - $_"
        $failCount++
    }

    Write-Host ""
}

# ========================================
# Step 6: Result
# ========================================
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount `
    -Title "Scheduled Task Enable Results")
