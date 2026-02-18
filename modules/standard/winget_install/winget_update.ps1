# ========================================
# Winget App Installer Update
# ========================================
# Upgrades winget itself (Microsoft.AppInstaller)
# via the winget source.
# Intended to run before winget_install.ps1,
# typically separated by __RESTART__ in Profile.
# ========================================

# Known ExitCode for "No applicable update found"
$WINGET_ALREADY_UPTODATE = -1978335212

Write-Host ""
Show-Separator
Write-Host "Winget App Installer Update" -ForegroundColor Cyan
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
    Show-Error "'winget' command not found. Please update App Installer manually."
    return (New-ModuleResult -Status "Error" -Message "winget command not found")
}
Show-Success "winget is available"
Write-Host ""

# ----------------------------------------
# 3. Display Target
# ----------------------------------------
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "Update Target" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""
Write-Host "  Package : Microsoft.AppInstaller" -ForegroundColor White
Write-Host "  Source  : winget" -ForegroundColor White
Write-Host ""
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

# ----------------------------------------
# 4. Confirmation
# ----------------------------------------
$cancelResult = Confirm-ModuleExecution -Message "Update winget (Microsoft.AppInstaller)?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""

# ----------------------------------------
# 5. Execute Upgrade
# ----------------------------------------
Show-Info "Running winget upgrade..."
Write-Host ""

$wingetArgs = "upgrade Microsoft.AppInstaller --source winget --silent --accept-source-agreements --accept-package-agreements"

try {
    $process = Start-Process -FilePath "winget" `
        -ArgumentList $wingetArgs `
        -Wait -NoNewWindow -PassThru

    switch ($process.ExitCode) {
        0 {
            Show-Success "winget updated successfully"
            Write-Host ""
            return (New-ModuleResult -Status "Success" -Message "Microsoft.AppInstaller updated successfully")
        }
        3010 {
            Show-Success "winget updated successfully (reboot pending)"
            Write-Host ""
            return (New-ModuleResult -Status "Success" -Message "Microsoft.AppInstaller updated (reboot pending)")
        }
        $WINGET_ALREADY_UPTODATE {
            Show-Skip "winget is already up to date"
            Write-Host ""
            return (New-ModuleResult -Status "Skipped" -Message "Microsoft.AppInstaller is already up to date")
        }
        default {
            Show-Error "winget upgrade exited with code: $($process.ExitCode)"
            Write-Host ""
            return (New-ModuleResult -Status "Error" -Message "winget upgrade failed (ExitCode: $($process.ExitCode))")
        }
    }
}
catch {
    Show-Error "Execution error: $_"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "winget upgrade failed: $_")
}
