# ========================================
# Winget Batch Installer
# ========================================
# Installs applications via winget based on
# app_list.csv configuration.
# Features:
#   - Pre-install check (skip already installed)
#   - ExitCode 3010 treated as success (reboot pending)
#   - Idempotency via winget list --id check
# ========================================

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Winget Batch Installer" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# ----------------------------------------
# 1. Internet Connection Check
# ----------------------------------------
Write-Host "Checking internet connection..." -ForegroundColor White
if (Test-Connection -ComputerName "8.8.8.8" -Count 1 -Quiet) {
    Write-Host "[SUCCESS] Internet connection OK" -ForegroundColor Green
}
else {
    Write-Host "[ERROR] No internet connection (Ping 8.8.8.8 failed)" -ForegroundColor Red
    return (New-ModuleResult -Status "Error" -Message "No internet connection")
}

# ----------------------------------------
# 2. Check Winget Availability
# ----------------------------------------
Write-Host "Checking winget availability..." -ForegroundColor White
if (-not (Get-Command "winget" -ErrorAction SilentlyContinue)) {
    Write-Host "[ERROR] 'winget' command not found. Please update App Installer." -ForegroundColor Red
    return (New-ModuleResult -Status "Error" -Message "winget command not found")
}
Write-Host "[SUCCESS] winget is available" -ForegroundColor Green
Write-Host ""

# ----------------------------------------
# 3. Load CSV
# ----------------------------------------
$csvPath = Join-Path $PSScriptRoot "app_list.csv"

$appList = Import-CsvSafe -Path $csvPath -Description "app_list.csv"
if ($null -eq $appList) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load app_list.csv")
}

if (-not (Test-CsvColumns -CsvData $appList -RequiredColumns @("Enabled", "AppID") -CsvName "app_list.csv")) {
    return (New-ModuleResult -Status "Error" -Message "app_list.csv missing required columns")
}

$enabledApps = @($appList | Where-Object { $_.Enabled -eq "1" -and -not [string]::IsNullOrWhiteSpace($_.AppID) })

if ($enabledApps.Count -eq 0) {
    Write-Host "[INFO] No enabled apps in app_list.csv" -ForegroundColor Yellow
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "No enabled apps")
}

# ----------------------------------------
# 4. Pre-install Check & Target Display
# ----------------------------------------
Write-Host "Checking installed status..." -ForegroundColor Cyan
Write-Host ""

$toInstall = @()
$alreadyInstalled = @()

foreach ($app in $enabledApps) {
    $appName = if ($app.Description) { $app.Description } else { $app.AppID }

    # Check if already installed via winget list
    $listOutput = & winget list --id $app.AppID --exact --accept-source-agreements 2>&1 | Out-String
    if ($listOutput -match [regex]::Escape($app.AppID)) {
        Write-Host "  [SKIP] $appName ($($app.AppID)) - already installed" -ForegroundColor Gray
        $alreadyInstalled += $app
    }
    else {
        Write-Host "  [INSTALL] $appName ($($app.AppID))" -ForegroundColor White
        $toInstall += $app
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
Write-Host "  To Install:        $($toInstall.Count) apps" -ForegroundColor White
Write-Host "  Already Installed: $($alreadyInstalled.Count) apps" -ForegroundColor Gray
Write-Host "  Disabled:          $($disabledApps.Count) apps" -ForegroundColor DarkGray
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

if ($toInstall.Count -eq 0) {
    Write-Host "[INFO] All enabled apps are already installed" -ForegroundColor Green
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "All apps already installed (Skipped: $($alreadyInstalled.Count))")
}

# ----------------------------------------
# 5. Confirmation
# ----------------------------------------
if (-not (Confirm-Execution -Message "Install the above $($toInstall.Count) app(s)?")) {
    Write-Host ""
    Write-Host "[INFO] Cancelled" -ForegroundColor Cyan
    Write-Host ""
    return (New-ModuleResult -Status "Cancelled" -Message "User cancelled")
}

Write-Host ""

# ----------------------------------------
# 6. Installation Loop
# ----------------------------------------
$successCount = 0
$failCount = 0
$skipCount = $alreadyInstalled.Count

foreach ($app in $toInstall) {
    $appName = if ($app.Description) { $app.Description } else { $app.AppID }

    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "Installing: $appName ($($app.AppID))" -ForegroundColor Cyan

    # Build winget arguments
    $wingetArgs = "install --id `"$($app.AppID)`" --exact --silent --accept-source-agreements --accept-package-agreements"

    if (-not [string]::IsNullOrWhiteSpace($app.Options)) {
        $wingetArgs += " $($app.Options)"
    }

    try {
        $process = Start-Process -FilePath "winget" -ArgumentList $wingetArgs -Wait -NoNewWindow -PassThru

        if ($process.ExitCode -eq 0) {
            Write-Host "[SUCCESS] Installation completed" -ForegroundColor Green
            $successCount++
        }
        elseif ($process.ExitCode -eq 3010) {
            # 3010 = reboot pending, installation itself succeeded
            Write-Host "[SUCCESS] Installation completed (reboot pending)" -ForegroundColor Green
            $successCount++
        }
        else {
            Write-Host "[ERROR] Installation failed. ExitCode: $($process.ExitCode)" -ForegroundColor Red
            $failCount++
        }
    }
    catch {
        Write-Host "[ERROR] Execution error: $_" -ForegroundColor Red
        $failCount++
    }

    Write-Host ""
}

# ----------------------------------------
# 7. Result Summary
# ----------------------------------------
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Installation Results" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
if ($successCount -gt 0) {
    Write-Host "  Success: $successCount apps" -ForegroundColor Green
}
if ($skipCount -gt 0) {
    Write-Host "  Skipped: $skipCount apps (already installed)" -ForegroundColor Gray
}
if ($failCount -gt 0) {
    Write-Host "  Failed:  $failCount apps" -ForegroundColor Red
}
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Determine Module Status
$overallStatus = if ($failCount -eq 0 -and $successCount -gt 0) { "Success" }
    elseif ($successCount -gt 0 -and $failCount -gt 0) { "Partial" }
    elseif ($failCount -eq 0 -and $skipCount -gt 0 -and $successCount -eq 0) { "Skipped" }
    elseif ($failCount -gt 0) { "Error" }
    else { "Skipped" }

return (New-ModuleResult -Status $overallStatus -Message "Success: $successCount, Skip: $skipCount, Fail: $failCount")
