# ========================================
# Generic Batch Runner Script
# ========================================

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Generic Batch Runner" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# ========================================
# Load CSV
# ========================================
$csvPath = Join-Path $PSScriptRoot "batch_list.csv"

$csvData = Import-CsvSafe $csvPath
if (-not $csvData) {
    return (New-ModuleResult -Status "Error" -Message "batch_list.csv not found or empty")
}

if (-not (Test-CsvColumns -CsvData $csvData -RequiredColumns @("Enabled","Description","BatchPath","Arguments","TimeoutSec","SuccessCodes","Encoding") -CsvName "batch_list.csv")) {
    return (New-ModuleResult -Status "Error" -Message "batch_list.csv has invalid columns")
}

# Filter enabled entries
$batchList = @($csvData | Where-Object { $_.Enabled -eq '1' })

if ($batchList.Count -eq 0) {
    Write-Host "[INFO] No enabled batch entries found" -ForegroundColor Yellow
    return (New-ModuleResult -Status "Skipped" -Message "No enabled batch entries")
}

Write-Host "[INFO] Loaded $($batchList.Count) batch definition(s)" -ForegroundColor Cyan
Write-Host ""

# ========================================
# Batch List Display + File Check
# ========================================
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host "Batch Execution List" -ForegroundColor Cyan
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""

$missingCount = 0

foreach ($batch in $batchList) {
    # Resolve batch path
    $batchPath = $batch.BatchPath
    if ([System.IO.Path]::IsPathRooted($batchPath)) {
        $fullPath = $batchPath
    } else {
        $fullPath = Join-Path $PSScriptRoot $batchPath
    }

    $exists = Test-Path $fullPath
    $timeout = if ([string]::IsNullOrWhiteSpace($batch.TimeoutSec) -or $batch.TimeoutSec -eq '0') { "None" } else { "$($batch.TimeoutSec)s" }
    $codes = if ([string]::IsNullOrWhiteSpace($batch.SuccessCodes)) { "0" } else { $batch.SuccessCodes }

    if ($exists) {
        Write-Host "  $($batch.Description)" -ForegroundColor Yellow
        Write-Host "    Path: $batchPath / Timeout: $timeout / SuccessCodes: $codes"
        if (-not [string]::IsNullOrWhiteSpace($batch.Arguments)) {
            Write-Host "    Args: $($batch.Arguments)"
        }
    }
    else {
        Write-Host "  $($batch.Description) [NOT FOUND]" -ForegroundColor Red
        Write-Host "    Path: $batchPath"
        $missingCount++
    }
    Write-Host ""
}

Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""

if ($missingCount -gt 0) {
    Write-Host "[WARNING] $missingCount batch file(s) not found" -ForegroundColor Yellow
    Write-Host "[INFO] Missing batches will be skipped" -ForegroundColor Yellow
    Write-Host ""
}

# ========================================
# Confirmation
# ========================================
if (-not (Confirm-Execution -Message "Proceed with batch execution?")) {
    Write-Host ""
    Write-Host "[INFO] Canceled" -ForegroundColor Yellow
    Write-Host ""
    return (New-ModuleResult -Status "Cancelled" -Message "User canceled")
}

Write-Host ""

# ========================================
# Batch Execution
# ========================================
$successCount = 0
$skipCount = 0
$failCount = 0

foreach ($batch in $batchList) {
    $desc = $batch.Description

    # Resolve batch path
    $batchPath = $batch.BatchPath
    if ([System.IO.Path]::IsPathRooted($batchPath)) {
        $fullPath = $batchPath
    } else {
        $fullPath = Join-Path $PSScriptRoot $batchPath
    }

    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "Executing: $desc" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor White

    # File existence check
    if (-not (Test-Path $fullPath)) {
        Write-Host "[SKIP] Batch file not found: $batchPath" -ForegroundColor Yellow
        Write-Host ""
        $skipCount++
        continue
    }

    # Parse settings with defaults
    $successCodes = if ([string]::IsNullOrWhiteSpace($batch.SuccessCodes)) { "0" } else { $batch.SuccessCodes }
    $timeoutSec = if ([string]::IsNullOrWhiteSpace($batch.TimeoutSec) -or $batch.TimeoutSec -eq '0') { 0 } else { [int]$batch.TimeoutSec }

    # Parse success codes list
    $successList = @($successCodes -split ',' | ForEach-Object { [int]$_.Trim() })

    # Build argument list for Start-Process
    $spArgs = "/c `"$fullPath`""
    if (-not [string]::IsNullOrWhiteSpace($batch.Arguments)) {
        $spArgs = "/c `"$fullPath`" $($batch.Arguments)"
    }

    try {
        # Simple execution: no output redirection, output goes straight to console
        $proc = Start-Process -FilePath "cmd.exe" -ArgumentList $spArgs `
            -WorkingDirectory $PSScriptRoot -PassThru -NoNewWindow

        Write-Host "[INFO] Process started (PID: $($proc.Id))" -ForegroundColor Gray

        # Wait with timeout
        $timedOut = $false
        if ($timeoutSec -gt 0) {
            $null = $proc | Wait-Process -Timeout $timeoutSec -ErrorAction SilentlyContinue
            if (-not $proc.HasExited) {
                $timedOut = $true
                Write-Host "[ERROR] Timeout ($($timeoutSec)s exceeded). Killing process..." -ForegroundColor Red
                $proc | Stop-Process -Force -ErrorAction SilentlyContinue
            }
        } else {
            $proc | Wait-Process
        }

        $exitCode = $proc.ExitCode

        # Result judgment
        if ($timedOut) {
            Write-Host "[ERROR] $desc : Timed out after $($timeoutSec)s" -ForegroundColor Red
            $failCount++
        }
        elseif ($exitCode -in $successList) {
            Write-Host "[SUCCESS] $desc (ExitCode: $exitCode)" -ForegroundColor Green
            $successCount++
        }
        else {
            Write-Host "[ERROR] $desc (ExitCode: $exitCode, Expected: $successCodes)" -ForegroundColor Red
            $failCount++
        }
    }
    catch {
        Write-Host "[ERROR] $desc : $($_.Exception.Message)" -ForegroundColor Red
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
