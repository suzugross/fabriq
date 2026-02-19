# ========================================
# Fabriq App Launcher Script
# ========================================
# Reads target_apps.csv and launches the listed
# fabriq GUI apps from the apps/ directory.
# Each app is launched as a separate PowerShell
# process. Wait=1 blocks until the app closes.
#
# NOTES:
# - Copy and rename this module to create
#   multiple launcher configurations in profiles.
# - Apps must follow the apps/{Name}/{Name}.ps1
#   naming convention.
# ========================================

Write-Host ""
Show-Separator
Write-Host "Fabriq App Launcher" -ForegroundColor Cyan
Show-Separator
Write-Host ""


# ========================================
# Step 1: CSV load
# ========================================
$csvPath = Join-Path $PSScriptRoot "target_apps.csv"

$appList = Import-ModuleCsv -Path $csvPath -FilterEnabled `
    -RequiredColumns @("Enabled", "AppName")

if ($null -eq $appList) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load target_apps.csv")
}
if ($appList.Count -eq 0) {
    return (New-ModuleResult -Status "Skipped" -Message "No enabled entries in target_apps.csv")
}


# ========================================
# Step 2: Pre-flight check
# ========================================
$appsDir = Join-Path $PSScriptRoot "..\..\..\apps"
$appsDir = [System.IO.Path]::GetFullPath($appsDir)

if (-not (Test-Path $appsDir)) {
    Show-Error "apps/ directory not found: $appsDir"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "apps/ directory not found")
}


# ========================================
# Step 3: Pre-execution display
# ========================================
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host "Apps to Launch" -ForegroundColor Cyan
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""

foreach ($app in $appList) {
    $displayName = if ($app.Description) { $app.Description } else { $app.AppName }
    $scriptPath  = Join-Path $appsDir "$($app.AppName)\$($app.AppName).ps1"
    $waitLabel   = if ($app.Wait -eq "0") { "No" } else { "Yes" }

    if (Test-Path $scriptPath) {
        Write-Host "  [READY]   $displayName" -ForegroundColor Yellow
        Write-Host "    Script: $($app.AppName)\$($app.AppName).ps1" -ForegroundColor DarkGray
        Write-Host "    Wait:   $waitLabel" -ForegroundColor DarkGray
    }
    else {
        Write-Host "  [MISSING] $displayName" -ForegroundColor DarkGray
        Write-Host "    Script: $($app.AppName)\$($app.AppName).ps1  (will be skipped)" -ForegroundColor DarkGray
    }
    Write-Host ""
}

Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""


# ========================================
# Step 4: Confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Launch the listed apps?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""


# ========================================
# Step 5: Launch loop
# ========================================
$successCount = 0
$skipCount    = 0
$failCount    = 0

foreach ($app in $appList) {
    $displayName = if ($app.Description) { $app.Description } else { $app.AppName }
    $scriptPath  = Join-Path $appsDir "$($app.AppName)\$($app.AppName).ps1"
    $waitMode    = $app.Wait -ne "0"

    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "Launching: $displayName" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor White

    # --- Script existence check (outside try) ---
    if (-not (Test-Path $scriptPath)) {
        Show-Skip "Script not found: $($app.AppName)\$($app.AppName).ps1"
        Write-Host ""
        $skipCount++
        continue
    }

    Show-Info "Script: $($app.AppName)\$($app.AppName).ps1"
    Show-Info "Wait:   $(if ($waitMode) { 'Yes (blocking)' } else { 'No (background)' })"

    try {
        $psArgs = "-NoProfile -ExecutionPolicy Unrestricted -File `"$scriptPath`""

        if ($waitMode) {
            $null = Start-Process powershell -ArgumentList $psArgs -Wait -PassThru -ErrorAction Stop
            Show-Success "Completed: $displayName"
        }
        else {
            Start-Process powershell -ArgumentList $psArgs -ErrorAction Stop
            Show-Success "Launched:  $displayName (running in background)"
        }

        $successCount++
    }
    catch {
        Show-Error "Failed to launch $displayName : $_"
        $failCount++
    }

    Write-Host ""
}


# ========================================
# Step 6: Result
# ========================================
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount `
    -Title "Fabriq App Launcher Results")
