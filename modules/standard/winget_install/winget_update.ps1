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
# Single bounded probe instead of Wait-NetworkReady (which loops
# forever) - fail fast and let the AutoPilot ErrorMode / dialog decide
# (retry/skip) instead of hanging an unattended run.
$netReachable = Test-Connection -ComputerName "8.8.8.8" -Count 2 -Quiet -ErrorAction SilentlyContinue
if (-not $netReachable) {
    Show-Error "Network unreachable (8.8.8.8)"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Network unreachable")
}
Show-Success "Network connectivity OK (8.8.8.8)"
Write-Host ""

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

# Record the current winget version: updating Microsoft.AppInstaller
# replaces the running winget.exe itself, so the upgrade can exit with
# a non-standard code even when it succeeded. The version read-back in
# step 5 is the real verdict for unknown exit codes.
$versionBefore = ""
try { $versionBefore = "$(& winget --version 2>$null)".Trim() } catch { }

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
            # The self-update can kill the running winget.exe and surface
            # a non-standard exit code even on success, so the exit code
            # is not trustworthy here. Read the version back (a fresh
            # process runs the NEW binary) and let that decide. Retries
            # absorb MSIX registration latency.
            # Probe budget widened to ~18s (6 x 3s): MSIX re-registration of
            # the new App Installer can exceed the old 6s on slow disks and
            # was producing a false Error. The break still fires as soon as a
            # changed version is observed, so a fast machine is unaffected.
            $versionAfter = ""
            for ($attempt = 1; $attempt -le 6; $attempt++) {
                Start-Sleep -Seconds 3
                try { $versionAfter = "$(& winget --version 2>$null)".Trim() } catch { $versionAfter = "" }
                if ($versionAfter -and $versionAfter -ne $versionBefore) { break }
            }

            # The success verdict requires a KNOWN baseline. If $versionBefore
            # could not be read before the upgrade (empty), a non-empty
            # $versionAfter only proves winget still runs, not that it was
            # upgraded - claiming Success there would be a false positive. In
            # that case fall through to the exit-code-based Error (fail-closed).
            if ($versionBefore -and $versionAfter -and $versionAfter -ne $versionBefore) {
                Show-Success "winget updated successfully ($versionBefore -> $versionAfter, ExitCode: $($process.ExitCode))"
                Write-Host ""
                return (New-ModuleResult -Status "Success" -Message "Microsoft.AppInstaller updated ($versionBefore -> $versionAfter)" -Verified $true)
            }

            $baselineNote = if ($versionBefore) { "version unchanged: $versionBefore" } else { "pre-upgrade version was unreadable - cannot confirm update" }
            Show-Error "winget upgrade exited with code $($process.ExitCode); $baselineNote"
            Write-Host ""
            return (New-ModuleResult -Status "Error" -Message "winget upgrade failed (ExitCode: $($process.ExitCode), $baselineNote)")
        }
    }
}
catch {
    Show-Error "Execution error: $_"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "winget upgrade failed: $_")
}
