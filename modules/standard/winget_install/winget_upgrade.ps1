# ========================================
# Winget Batch Upgrader
# ========================================
# Upgrades applications via winget based on
# app_list.csv configuration.
# Features:
#   - Pre-check (only upgrades installed apps)
#   - ExitCode 3010 treated as success (reboot pending)
#   - ExitCode -1978335212 treated as Skipped (already up to date)
#   - Not-installed apps are skipped (handled by winget_install.ps1)
# ========================================

# Known ExitCode for "No applicable update found"
$WINGET_NO_UPGRADE = -1978335212

Write-Host ""
Show-Separator
Write-Host "Winget Batch Upgrader" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# ----------------------------------------
# 1. Internet Connection Check
# ----------------------------------------
Wait-NetworkReady

# ----------------------------------------
# 2. Check Winget Availability
# ----------------------------------------
Write-Host "Checking winget availability..." -ForegroundColor White
if (-not (Get-Command "winget" -ErrorAction SilentlyContinue)) {
    Show-Error "'winget' command not found. Please update App Installer."
    return (New-ModuleResult -Status "Error" -Message "winget command not found")
}
Show-Success "winget is available"
Write-Host ""

# ----------------------------------------
# 2.5. Winget Source Reset
# ----------------------------------------
Show-Info "Resetting winget sources..."
$resetProcess = Start-Process -FilePath "winget" `
    -ArgumentList "source reset --force" `
    -Wait -NoNewWindow -PassThru

if ($resetProcess.ExitCode -eq 0) {
    Show-Success "winget source reset completed"
}
else {
    Show-Warning "winget source reset exited with code $($resetProcess.ExitCode) - continuing anyway"
}
Write-Host ""

# ----------------------------------------
# 3. Load CSV
# ----------------------------------------
$csvPath = Join-Path $PSScriptRoot "app_list.csv"

# Load all entries (without -FilterEnabled, needed for disabled app display)
$appList = Import-ModuleCsv -Path $csvPath -RequiredColumns @("Enabled", "AppID")
if ($null -eq $appList) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load app_list.csv")
}

$enabledApps = @($appList | Where-Object { $_.Enabled -eq "1" -and -not [string]::IsNullOrWhiteSpace($_.AppID) })

if ($enabledApps.Count -eq 0) {
    Show-Info "No enabled apps in app_list.csv"
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "No enabled apps")
}

# ----------------------------------------
# 4. Pre-upgrade Check & Target Display
# ----------------------------------------
Write-Host "Checking installed status..." -ForegroundColor Cyan
Write-Host ""

$toUpgrade = @()
$notInstalled = @()

foreach ($app in $enabledApps) {
    $appName = if ($app.Description) { $app.Description } else { $app.AppID }

    # Check if installed via winget list
    $listOutput = & winget list --id $app.AppID --exact --accept-source-agreements 2>&1 | Out-String
    if ($listOutput -match [regex]::Escape($app.AppID)) {
        Write-Host "  [UPGRADE] $appName ($($app.AppID))" -ForegroundColor White
        $toUpgrade += $app
    }
    else {
        Show-Skip "$appName ($($app.AppID)) - not installed (skipped)"
        $notInstalled += $app
    }
}

Write-Host ""

# Show disabled apps
$disabledApps = @($appList | Where-Object { $_.Enabled -ne "1" -and -not [string]::IsNullOrWhiteSpace($_.AppID) })
foreach ($app in $disabledApps) {
    $appName = if ($app.Description) { $app.Description } else { $app.AppID }
    Write-Host "  [DISABLED] $appName ($($app.AppID))" -ForegroundColor DarkGray
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "  To Upgrade:     $($toUpgrade.Count) apps" -ForegroundColor White
Write-Host "  Not Installed:  $($notInstalled.Count) apps" -ForegroundColor Gray
Write-Host "  Disabled:       $($disabledApps.Count) apps" -ForegroundColor DarkGray
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

if ($toUpgrade.Count -eq 0) {
    Show-Skip "No installed apps to upgrade"
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "No installed apps to upgrade (Not installed: $($notInstalled.Count))")
}

# ----------------------------------------
# 5. Confirmation
# ----------------------------------------
$cancelResult = Confirm-ModuleExecution -Message "Upgrade the above $($toUpgrade.Count) app(s)?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""

# ----------------------------------------
# 6. Upgrade Loop
# ----------------------------------------
$successCount = 0
$failCount = 0
$skipCount = $notInstalled.Count

foreach ($app in $toUpgrade) {
    $appName = if ($app.Description) { $app.Description } else { $app.AppID }

    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "Upgrading: $appName ($($app.AppID))" -ForegroundColor Cyan

    # Build winget arguments
    $wingetArgs = "upgrade --id `"$($app.AppID)`" --exact --silent --accept-source-agreements --accept-package-agreements"

    if (-not [string]::IsNullOrWhiteSpace($app.Options)) {
        $wingetArgs += " $($app.Options)"
    }

    try {
        $process = Start-Process -FilePath "winget" -ArgumentList $wingetArgs -Wait -NoNewWindow -PassThru

        switch ($process.ExitCode) {
            0 {
                Show-Success "Upgrade completed"
                $successCount++
            }
            3010 {
                # 3010 = reboot pending, upgrade itself succeeded
                Show-Success "Upgrade completed (reboot pending)"
                $successCount++
            }
            $WINGET_NO_UPGRADE {
                Show-Skip "Already up to date"
                $skipCount++
            }
            default {
                Show-Error "Upgrade failed. ExitCode: $($process.ExitCode)"
                $failCount++
            }
        }
    }
    catch {
        Show-Error "Execution error: $_"
        $failCount++
    }

    Write-Host ""
}

# ----------------------------------------
# 7. Result Summary
# ----------------------------------------
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount -Title "Upgrade Results")
