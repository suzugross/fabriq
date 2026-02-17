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
Show-Separator
Write-Host "Winget Batch Installer" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# ----------------------------------------
# 1. Internet Connection Check
# ----------------------------------------
Write-Host "Checking internet connection..." -ForegroundColor White
if (Test-Connection -ComputerName "8.8.8.8" -Count 1 -Quiet) {
    Show-Success "Internet connection OK"
}
else {
    Show-Error "No internet connection (Ping 8.8.8.8 failed)"
    return (New-ModuleResult -Status "Error" -Message "No internet connection")
}

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
    Show-Info "No enabled apps in app_list.csv"
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
        Show-Skip "$appName ($($app.AppID)) - already installed"
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
    Show-Skip "All enabled apps are already installed"
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "All apps already installed (Skipped: $($alreadyInstalled.Count))")
}

# ----------------------------------------
# 5. Confirmation
# ----------------------------------------
$cancelResult = Confirm-ModuleExecution -Message "Install the above $($toInstall.Count) app(s)?"
if ($null -ne $cancelResult) { return $cancelResult }

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
            Show-Success "Installation completed"
            $successCount++
        }
        elseif ($process.ExitCode -eq 3010) {
            # 3010 = reboot pending, installation itself succeeded
            Show-Success "Installation completed (reboot pending)"
            $successCount++
        }
        else {
            Show-Error "Installation failed. ExitCode: $($process.ExitCode)"
            $failCount++
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
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount -Title "Installation Results")
