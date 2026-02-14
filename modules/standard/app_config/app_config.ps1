# ========================================
# Application Installation Script
# ========================================

Write-Host "Executing application installation..." -ForegroundColor Cyan
Write-Host ""

# ========================================
# Load CSV
# ========================================
$csvPath = Join-Path $PSScriptRoot "app_list.csv"

if (-not (Test-Path $csvPath)) {
    Write-Host "[ERROR] app_list.csv not found: $csvPath" -ForegroundColor Red
    return (New-ModuleResult -Status "Error" -Message "app_list.csv not found")
}

try {
    $appList = @(Import-Csv -Path $csvPath -Encoding Default)
}
catch {
    Write-Host "[ERROR] Failed to load app_list.csv: $_" -ForegroundColor Red
    return (New-ModuleResult -Status "Error" -Message "Failed to load app_list.csv: $_")
}

if ($appList.Count -eq 0) {
    Write-Host "[ERROR] app_list.csv contains no data" -ForegroundColor Red
    return (New-ModuleResult -Status "Error" -Message "app_list.csv contains no data")
}

Write-Host "[INFO] Loaded $($appList.Count) application definitions" -ForegroundColor Cyan
Write-Host ""

# ========================================
# Installer Directory
# ========================================
$fileDir = Join-Path $PSScriptRoot "file"

if (-not (Test-Path $fileDir)) {
    Write-Host "[ERROR] 'file' directory not found: $fileDir" -ForegroundColor Red
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
    Write-Host "[WARNING] $missingCount installers are missing" -ForegroundColor Yellow
    Write-Host "[INFO] Missing applications will be skipped" -ForegroundColor Yellow
    Write-Host ""
}

# ========================================
# Confirmation
# ========================================
if (-not (Confirm-Execution -Message "Proceed with installation?")) {
    Write-Host ""
    Write-Host "[INFO] Canceled" -ForegroundColor Yellow
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
        Write-Host "[SKIP] Installer not found: $($app.FileName)" -ForegroundColor Yellow
        Write-Host ""
        $skipCount++
        continue
    }

    try {
        $process = $null

        switch ($app.Type.ToLower()) {
            "msi" {
                Write-Host "[INFO] Executing MSI installation..." -ForegroundColor Gray
                $msiArgs = "/i `"$installerPath`" $($app.SilentArgs)"
                $process = Start-Process msiexec -ArgumentList $msiArgs -Wait -PassThru
            }
            "exe" {
                Write-Host "[INFO] Executing EXE installation..." -ForegroundColor Gray
                $process = Start-Process $installerPath -ArgumentList $app.SilentArgs -Wait -PassThru
            }
            default {
                Write-Host "[ERROR] Unsupported type: $($app.Type)" -ForegroundColor Red
                $failCount++
                Write-Host ""
                continue
            }
        }

        if ($process.ExitCode -eq 0) {
            Write-Host "[SUCCESS] Installed $appName (ExitCode: 0)" -ForegroundColor Green
            $successCount++
        }
        else {
            Write-Host "[ERROR] Failed to install $appName (ExitCode: $($process.ExitCode))" -ForegroundColor Red
            $failCount++
        }
    }
    catch {
        Write-Host "[ERROR] Error during installation of $appName : $_" -ForegroundColor Red
        $failCount++
    }

    Write-Host ""
}

# ========================================
# Result Summary
# ========================================
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Execution Results" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Success: $successCount items" -ForegroundColor Green
Write-Host "  Skipped: $skipCount items" -ForegroundColor Yellow
Write-Host "  Failed: $failCount items" -ForegroundColor $(if ($failCount -gt 0) { "Red" } else { "Green" })
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Return ModuleResult
$overallStatus = if ($failCount -eq 0 -and $successCount -gt 0) { "Success" }
    elseif ($successCount -gt 0 -and $failCount -gt 0) { "Partial" }
    elseif ($failCount -eq 0 -and $skipCount -gt 0) { "Skipped" }
    else { "Error" }
return (New-ModuleResult -Status $overallStatus -Message "Success: $successCount, Skip: $skipCount, Fail: $failCount")