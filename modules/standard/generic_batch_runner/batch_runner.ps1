# ========================================
# Generic Batch Runner Script
# ========================================
# Runs .bat files listed in batch_list.csv.
#
# [NOTES]
# - BatchPath accepts an absolute path, %VAR% environment variables, or a
#   path relative to the MODULE DATA ROOT.
# - The relative form is resolved through Resolve-ModuleDataPath, so a profile
#   data folder (PDF) can ship batch_list.csv AND the .bat files it names
#   (dev/PROFILE_DATA_OVERLAY_PLAN.md). Without an overlay context the
#   resolution is an identity mapping = pre-overlay behavior.
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

$batchList = Import-ModuleCsv -Path $csvPath -FilterEnabled -RequiredColumns @("Enabled","Description","BatchPath","Arguments","TimeoutSec","SuccessCodes")
if ($null -eq $batchList) { return (New-ModuleResult -Status "Error" -Message "Failed to load batch_list.csv") }
if ($batchList.Count -eq 0) { return (New-ModuleResult -Status "Skipped" -Message "No enabled batch entries") }
Write-Host ""

# ========================================
# Path Resolution
# ========================================
# Resolved once up front so the preview list and the execution loop can never
# disagree about which file a row points at (script_looper uses the same
# _ResolvedPath pattern).
function Resolve-BatchPath {
    param([string]$RawPath)

    # Blank must short-circuit: Expand-UserEnvironmentVariables rejects an
    # empty mandatory string, and an empty relative path would resolve to the
    # module data root itself (a directory) and look like a valid target.
    if ([string]::IsNullOrWhiteSpace($RawPath)) { return "" }

    $expanded = Expand-UserEnvironmentVariables $RawPath
    if ([System.IO.Path]::IsPathRooted($expanded)) { return $expanded }
    return (Resolve-ModuleDataPath -Path (Join-Path $PSScriptRoot $expanded))
}

foreach ($batch in $batchList) {
    $null = $batch | Add-Member -NotePropertyName "_ResolvedPath" `
        -NotePropertyValue (Resolve-BatchPath $batch.BatchPath) -Force
}

# Working directory for every batch: the module data root - the profile-side
# module folder while an overlay is active, otherwise $PSScriptRoot. Keeping
# "CWD = module root" preserves the semantics existing batches were written
# against while pointing CWD-relative references at the active data set.
$batchWorkDir = Resolve-ModuleDataPath -Path $PSScriptRoot

# ========================================
# Batch List Display + File Check
# ========================================
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host "Batch Execution List" -ForegroundColor Cyan
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""
Write-Host "  WorkDir: $batchWorkDir" -ForegroundColor DarkGray
Write-Host ""

$missingCount = 0

foreach ($batch in $batchList) {
    $fullPath = $batch._ResolvedPath
    $exists = (-not [string]::IsNullOrWhiteSpace($fullPath)) -and (Test-Path -LiteralPath $fullPath -PathType Leaf)
    $timeout = if ([string]::IsNullOrWhiteSpace($batch.TimeoutSec) -or $batch.TimeoutSec -eq '0') { "None" } else { "$($batch.TimeoutSec)s" }
    $codes = if ([string]::IsNullOrWhiteSpace($batch.SuccessCodes)) { "0" } else { $batch.SuccessCodes }

    if ($exists) {
        Write-Host "  $($batch.Description)" -ForegroundColor Yellow
        Write-Host "    Path: $fullPath / Timeout: $timeout / SuccessCodes: $codes"
        if (-not [string]::IsNullOrWhiteSpace($batch.Arguments)) {
            Write-Host "    Args: $($batch.Arguments)"
        }
    }
    else {
        Write-Host "  $($batch.Description) [NOT FOUND]" -ForegroundColor Red
        if ([string]::IsNullOrWhiteSpace($fullPath)) {
            Write-Host "    Path: (BatchPath is empty)"
        }
        else {
            Write-Host "    Path: $fullPath"
        }
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
    $fullPath = $batch._ResolvedPath

    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "Executing: $desc" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor White

    # File existence check
    if ([string]::IsNullOrWhiteSpace($fullPath)) {
        Show-Skip "BatchPath is empty"
        Write-Host ""
        $skipCount++
        continue
    }
    if (-not (Test-Path -LiteralPath $fullPath -PathType Leaf)) {
        Show-Skip "Batch file not found: $fullPath"
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
            -WorkingDirectory $batchWorkDir -PassThru -NoNewWindow

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
