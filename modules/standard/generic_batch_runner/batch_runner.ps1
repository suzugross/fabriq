# ========================================
# Generic Batch Runner Script
# ========================================

Write-Host ""
Show-Separator
Write-Host "Generic Batch Runner" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# ========================================
# Load CSV
# ========================================
$csvPath = Join-Path $PSScriptRoot "batch_list.csv"

$batchList = Import-ModuleCsv -Path $csvPath -FilterEnabled -RequiredColumns @("Enabled","Description","BatchPath","Arguments","TimeoutSec","SuccessCodes","Encoding")
if ($null -eq $batchList) { return (New-ModuleResult -Status "Error" -Message "Failed to load batch_list.csv") }
if ($batchList.Count -eq 0) { return (New-ModuleResult -Status "Skipped" -Message "No enabled batch entries") }
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
    Show-Warning "$missingCount batch file(s) not found"
    Show-Info "Missing batches will be skipped"
    Write-Host ""
}

# ========================================
# Confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Proceed with batch execution?"
if ($null -ne $cancelResult) { return $cancelResult }

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
        Show-Skip "Batch file not found: $batchPath"
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

        Show-Info "Process started (PID: $($proc.Id))"

        # Wait with timeout
        $timedOut = $false
        if ($timeoutSec -gt 0) {
            $null = $proc | Wait-Process -Timeout $timeoutSec -ErrorAction SilentlyContinue
            if (-not $proc.HasExited) {
                $timedOut = $true
                Show-Error "Timeout ($($timeoutSec)s exceeded). Killing process..."
                $proc | Stop-Process -Force -ErrorAction SilentlyContinue
            }
        } else {
            $proc | Wait-Process
        }

        $exitCode = $proc.ExitCode

        # Result judgment
        if ($timedOut) {
            Show-Error "$desc : Timed out after $($timeoutSec)s"
            $failCount++
        }
        elseif ($exitCode -in $successList) {
            Show-Success "$desc (ExitCode: $exitCode)"
            $successCount++
        }
        else {
            Show-Error "$desc (ExitCode: $exitCode, Expected: $successCodes)"
            $failCount++
        }
    }
    catch {
        Show-Error "$desc : $($_.Exception.Message)"
        $failCount++
    }

    Write-Host ""
}

# ========================================
# Result Summary
# ========================================
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount -Title "Execution Results")
