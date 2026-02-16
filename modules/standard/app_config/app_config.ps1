# ========================================
# Application Installation Script
# ========================================

Show-Info "Executing application installation..."
Write-Host ""

# ========================================
# Load CSV
# ========================================
$csvPath = Join-Path $PSScriptRoot "app_list.csv"

if (-not (Test-Path $csvPath)) {
    Show-Error "app_list.csv not found: $csvPath"
    return (New-ModuleResult -Status "Error" -Message "app_list.csv not found")
}

try {
    $appList = @(Import-Csv -Path $csvPath -Encoding Default)
}
catch {
    Show-Error "Failed to load app_list.csv: $_"
    return (New-ModuleResult -Status "Error" -Message "Failed to load app_list.csv: $_")
}

if ($appList.Count -eq 0) {
    Show-Error "app_list.csv contains no data"
    return (New-ModuleResult -Status "Error" -Message "app_list.csv contains no data")
}

Show-Info "Loaded $($appList.Count) application definitions"
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

foreach ($app in $appList) {
    $installerPath = Join-Path $fileDir $app.FileName
    $exists = Test-Path $installerPath

    if ($exists) {
        Write-Host "  $($app.AppName)" -ForegroundColor Yellow
        Write-Host "    File: $($app.FileName) / Type: $($app.Type) / Args: $($app.SilentArgs)"
    }
    else {
        Write-Host "  $($app.AppName) [NOT FOUND]" -ForegroundColor Red
        Write-Host "    File: $($app.FileName) is missing"
        $missingCount++
    }
    Write-Host ""
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
if (-not (Confirm-Execution -Message "Proceed with installation?")) {
    Write-Host ""
    Show-Info "Canceled"
    Write-Host ""
    return (New-ModuleResult -Status "Cancelled" -Message "User canceled")
}

Write-Host ""

# ========================================
# Installation Process
# ========================================
$successCount = 0
$skipCount = 0
$failCount = 0

foreach ($app in $appList) {
    $appName = $app.AppName
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
                $msiArgs = "/i `"$installerPath`" $($app.SilentArgs)"
                $process = Start-Process msiexec -ArgumentList $msiArgs -Wait -PassThru
            }
            "exe" {
                Show-Info "Executing EXE installation..."
                $process = Start-Process $installerPath -ArgumentList $app.SilentArgs -Wait -PassThru
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
Show-Separator
Write-Host "Execution Results" -ForegroundColor Cyan
Show-Separator
Write-Host "  Success: $successCount items" -ForegroundColor Green
Write-Host "  Skipped: $skipCount items" -ForegroundColor Yellow
Write-Host "  Failed: $failCount items" -ForegroundColor $(if ($failCount -gt 0) { "Red" } else { "Green" })
Show-Separator
Write-Host ""

# Return ModuleResult
$overallStatus = if ($failCount -eq 0 -and $successCount -gt 0) { "Success" }
    elseif ($successCount -gt 0 -and $failCount -gt 0) { "Partial" }
    elseif ($failCount -eq 0 -and $skipCount -gt 0) { "Skipped" }
    else { "Error" }
return (New-ModuleResult -Status $overallStatus -Message "Success: $successCount, Skip: $skipCount, Fail: $failCount")