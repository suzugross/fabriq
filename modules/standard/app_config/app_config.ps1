# ========================================
# Application Installation Script
# ========================================

Show-Info "Executing application installation..."
Write-Host ""

# ========================================
# Load CSV
# ========================================
$csvPath = Join-Path $PSScriptRoot "app_list.csv"

$appList = Import-ModuleCsv -Path $csvPath -RequiredColumns @("Enabled", "AppName", "FileName", "Type")
if ($null -eq $appList) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load app_list.csv")
}

$enabledApps = @($appList | Where-Object { $_.Enabled -eq "1" })
$disabledApps = @($appList | Where-Object { $_.Enabled -ne "1" })

if ($enabledApps.Count -eq 0) {
    Show-Info "No enabled apps in app_list.csv"
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "No enabled apps")
}
Write-Host ""

# ========================================
# Installer Directory
# ========================================
$fileDir = Join-Path $PSScriptRoot "file"

if (-not (Test-Path $fileDir)) {
    Show-Error "'file' directory not found: $fileDir"
    return (New-ModuleResult -Status "Error" -Message "'file' directory not found")
}

# ========================================
# List Applications + File Check
# ========================================
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host "Installation List" -ForegroundColor Cyan
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""

$missingCount = 0

foreach ($app in $enabledApps) {
    $appName = if ($app.Description) { $app.Description } else { $app.AppName }
    $installerPath = Join-Path $fileDir $app.FileName
    $exists = Test-Path $installerPath

    if ($exists) {
        Write-Host "  $appName" -ForegroundColor Yellow
        Write-Host "    File: $($app.FileName) / Type: $($app.Type) / Args: $($app.SilentArgs)"
    }
    else {
        Write-Host "  $appName [NOT FOUND]" -ForegroundColor Red
        Write-Host "    File: $($app.FileName) is missing"
        $missingCount++
    }
    Write-Host ""
}

foreach ($app in $disabledApps) {
    $appName = if ($app.Description) { $app.Description } else { $app.AppName }
    Write-Host "  [DISABLED] $appName ($($app.FileName))" -ForegroundColor DarkGray
}

Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""

if ($missingCount -gt 0) {
    Show-Warning "$missingCount installers are missing"
    Show-Info "Missing applications will be skipped"
    Write-Host ""
}

# ========================================
# Confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Proceed with installation?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""

# ========================================
# Installation Process
# ========================================
$successCount = 0
$skipCount = 0
$failCount = 0

foreach ($app in $enabledApps) {
    $appName = if ($app.Description) { $app.Description } else { $app.AppName }
    $installerPath = Join-Path $fileDir $app.FileName

    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "Installing: $appName" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor White

    # File existence check
    if (-not (Test-Path $installerPath)) {
        Show-Skip "Installer not found: $($app.FileName)"
        Write-Host ""
        $skipCount++
        continue
    }

    try {
        $process = $null

        switch ($app.Type.ToLower()) {
            "msi" {
                Show-Info "Executing MSI installation..."
                $installerDir = Split-Path $installerPath -Parent
                $msiArgs = "/i `"$installerPath`" $($app.SilentArgs)"
                $process = Start-Process msiexec -ArgumentList $msiArgs -WorkingDirectory $installerDir -Wait -PassThru
            }
            "exe" {
                Show-Info "Executing EXE installation..."
                $installerDir = Split-Path $installerPath -Parent
                if ([string]::IsNullOrWhiteSpace($app.SilentArgs)) {
                    $process = Start-Process $installerPath -WorkingDirectory $installerDir -Wait -PassThru
                }
                else {
                    $process = Start-Process $installerPath -ArgumentList $app.SilentArgs -WorkingDirectory $installerDir -Wait -PassThru
                }
            }
            "bat" {
                Show-Info "Executing BAT installation..."
                $installerDir = Split-Path $installerPath -Parent
                if ([string]::IsNullOrWhiteSpace($app.SilentArgs)) {
                    $batArgs = "/c `"$installerPath`""
                }
                else {
                    $batArgs = "/c `"$installerPath`" $($app.SilentArgs)"
                }
                $process = Start-Process cmd.exe -ArgumentList $batArgs -WorkingDirectory $installerDir -Wait -PassThru
            }
            default {
                Show-Error "Unsupported type: $($app.Type)"
                $failCount++
                Write-Host ""
                continue
            }
        }

        if ($process.ExitCode -eq 0) {
            Show-Success "Installed $appName (ExitCode: 0)"
            $successCount++
        }
        else {
            Show-Error "Failed to install $appName (ExitCode: $($process.ExitCode))"
            $failCount++
        }
    }
    catch {
        Show-Error "Error during installation of $appName : $_"
        $failCount++
    }

    Write-Host ""
}

# ========================================
# Result Summary
# ========================================
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount -Title "Execution Results")