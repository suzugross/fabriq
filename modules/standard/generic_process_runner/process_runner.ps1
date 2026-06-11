# ========================================
# Generic Process Runner Script
# ========================================
# Generic module that runs arbitrary EXE files defined in a CSV.
#
# [NOTES]
# - ExecutablePath accepts absolute paths, environment variables,
#   or paths relative to WorkingDirectory
# - When WorkingDirectory is empty, paths are resolved relative to $PSScriptRoot
# ========================================

Write-Host ""
Show-Separator
Write-Host "Generic Process Runner" -ForegroundColor Cyan
Show-Separator
Write-Host ""


# ========================================
# Step 1: Load CSV
# ========================================
$csvPath = Join-Path $PSScriptRoot "process_list.csv"

$processList = Import-ModuleCsv -Path $csvPath -FilterEnabled `
    -RequiredColumns @("Enabled", "Description", "ExecutablePath", "Arguments", "WorkingDirectory", "TimeoutSec", "SuccessCodes", "NoNewWindow")

if ($null -eq $processList) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load process_list.csv")
}
if ($processList.Count -eq 0) {
    return (New-ModuleResult -Status "Skipped" -Message "No enabled process entries")
}
Write-Host ""


# ========================================
# Step 3: Dry-run summary before execution
# ========================================
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host "Process Execution List" -ForegroundColor Cyan
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""

$missingCount = 0

foreach ($proc in $processList) {
    # --- Resolve path ---
    # Expand environment variables
    $exePath = Expand-UserEnvironmentVariables ($proc.ExecutablePath)
    $workDir = if (-not [string]::IsNullOrWhiteSpace($proc.WorkingDirectory)) {
        Expand-UserEnvironmentVariables ($proc.WorkingDirectory)
    } else {
        ""
    }

    # Resolve absolute vs relative path
    if ([System.IO.Path]::IsPathRooted($exePath)) {
        $fullPath = $exePath
    } elseif (-not [string]::IsNullOrWhiteSpace($workDir)) {
        $fullPath = Join-Path $workDir $exePath
    } else {
        $fullPath = Join-Path $PSScriptRoot $exePath
    }

    # Format parameters for display
    $exists  = Test-Path $fullPath
    $timeout = if ([string]::IsNullOrWhiteSpace($proc.TimeoutSec) -or $proc.TimeoutSec -eq '0') { "None" } else { "$($proc.TimeoutSec)s" }
    $codes   = if ([string]::IsNullOrWhiteSpace($proc.SuccessCodes)) { "0" } else { $proc.SuccessCodes }
    $window  = if ($proc.NoNewWindow -eq '1') { "NoNewWindow" } else { "NewWindow" }
    $dispWorkDir = if (-not [string]::IsNullOrWhiteSpace($workDir)) { $workDir } else { "(default: $PSScriptRoot)" }

    if ($exists) {
        Write-Host "  $($proc.Description)" -ForegroundColor Yellow
        Write-Host "    EXE:      $fullPath"
        Write-Host "    WorkDir:  $dispWorkDir"
        Write-Host "    Timeout:  $timeout / SuccessCodes: $codes / Window: $window"
        if (-not [string]::IsNullOrWhiteSpace($proc.Arguments)) {
            Write-Host "    Args:     $($proc.Arguments)"
        }
        if (-not [string]::IsNullOrWhiteSpace($proc.WaitProcessName)) {
            Write-Host "    WaitFor:  $($proc.WaitProcessName) (poll until process exits)"
        }
    }
    else {
        Write-Host "  $($proc.Description) [NOT FOUND]" -ForegroundColor Red
        Write-Host "    EXE:      $fullPath"
        $missingCount++
    }
    Write-Host ""
}

Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""

if ($missingCount -gt 0) {
    Show-Warning "$missingCount executable(s) not found"
    Show-Info "Missing executables will be skipped during execution"
    Write-Host ""
}


# ========================================
# Step 4: User confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Proceed with process execution?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""


# ========================================
# Step 5: Process-execution loop
# ========================================
$successCount = 0
$skipCount    = 0
$failCount    = 0

foreach ($proc in $processList) {
    $desc = $proc.Description

    # --- Resolve path (same logic as Step 3) ---
    $exePath = Expand-UserEnvironmentVariables ($proc.ExecutablePath)
    $workDir = if (-not [string]::IsNullOrWhiteSpace($proc.WorkingDirectory)) {
        Expand-UserEnvironmentVariables ($proc.WorkingDirectory)
    } else {
        ""
    }

    if ([System.IO.Path]::IsPathRooted($exePath)) {
        $fullPath = $exePath
    } elseif (-not [string]::IsNullOrWhiteSpace($workDir)) {
        $fullPath = Join-Path $workDir $exePath
    } else {
        $fullPath = Join-Path $PSScriptRoot $exePath
    }

    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "Executing: $desc" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor White

    # --- Existence check ---
    if (-not (Test-Path $fullPath)) {
        Show-Skip "Executable not found: $fullPath"
        Write-Host ""
        $skipCount++
        continue
    }

    # --- Format parameters ---
    $successCodes = if ([string]::IsNullOrWhiteSpace($proc.SuccessCodes)) { "0" } else { $proc.SuccessCodes }
    $timeoutSec   = if ([string]::IsNullOrWhiteSpace($proc.TimeoutSec) -or $proc.TimeoutSec -eq '0') { 0 } else { [int]$proc.TimeoutSec }
    $successList  = @($successCodes -split ',' | ForEach-Object { [int]$_.Trim() })

    # --- Decide the working directory ---
    $execWorkDir = if (-not [string]::IsNullOrWhiteSpace($workDir)) {
        $workDir
    } else {
        $PSScriptRoot
    }

    # --- Build the Start-Process arguments ---
    $spParams = @{
        FilePath         = $fullPath
        WorkingDirectory = $execWorkDir
        PassThru         = $true
    }

    if (-not [string]::IsNullOrWhiteSpace($proc.Arguments)) {
        $spParams.ArgumentList = $proc.Arguments
    }

    if ($proc.NoNewWindow -eq '1') {
        $spParams.NoNewWindow = $true
    }

    try {
        $process = Start-Process @spParams

        Show-Info "Process started (PID: $($process.Id))"

        # --- Timeout watch ---
        $timedOut = $false
        $stillRunning = $false
        # Default ceiling for the WaitProcessName polling phase when
        # TimeoutSec=0 (the Guide promises that phase is hang-proof).
        # The MAIN Wait-Process below stays unlimited by documented
        # design (operator-interactive GUI installers).
        $defaultPollCeilingSec = 3600
        if ($timeoutSec -gt 0) {
            $null = $process | Wait-Process -Timeout $timeoutSec -ErrorAction SilentlyContinue
            if (-not $process.HasExited) {
                $timedOut = $true
                Show-Error "Timeout ($($timeoutSec)s exceeded). Killing process..."
                $process | Stop-Process -Force -ErrorAction SilentlyContinue
            }
        } else {
            $process | Wait-Process
        }

        # --- Post-execution process polling (WaitProcessName) ---
        $waitProcessName = if (
            ($proc.PSObject.Properties.Name -contains 'WaitProcessName') -and
            (-not [string]::IsNullOrWhiteSpace($proc.WaitProcessName))
        ) { $proc.WaitProcessName } else { "" }

        if (-not [string]::IsNullOrWhiteSpace($waitProcessName) -and -not $timedOut) {
            Show-Info "Waiting for process '$waitProcessName' to complete..."
            $pollInterval = 5
            $elapsed = 0
            # Explicit TimeoutSec keeps its documented kill-on-timeout
            # behavior. TimeoutSec=0 used to poll forever here (while a
            # zombie child has no operator-visible UI to act on); it now
            # gets the default ceiling WITHOUT killing - a legitimately
            # slow installer must not be shot at an arbitrary cap, so the
            # item is counted as Error and the process is left running
            # for the operator / ErrorMode to decide.
            $pollCeilingSec = if ($timeoutSec -gt 0) { $timeoutSec } else { $defaultPollCeilingSec }
            while ($true) {
                $running = Get-Process -Name $waitProcessName -ErrorAction SilentlyContinue
                if (-not $running) { break }
                $elapsed += $pollInterval
                if ($elapsed -ge $pollCeilingSec) {
                    if ($timeoutSec -gt 0) {
                        $timedOut = $true
                        Show-Error "Timeout ($($timeoutSec)s exceeded) while waiting for '$waitProcessName'"
                        $running | Stop-Process -Force -ErrorAction SilentlyContinue
                    }
                    else {
                        $stillRunning = $true
                        Show-Error "'$waitProcessName' still running after default ceiling (${defaultPollCeilingSec}s) - NOT killed"
                    }
                    break
                }
                Start-Sleep -Seconds $pollInterval
            }
            if (-not $timedOut -and -not $stillRunning) {
                Show-Success "Process '$waitProcessName' has completed"
            }
        }

        $exitCode = $process.ExitCode

        # --- Decide the result ---
        if ($timedOut) {
            Show-Error "$desc : Timed out after $($timeoutSec)s"
            $failCount++
        }
        elseif ($stillRunning) {
            Show-Error "$desc : '$waitProcessName' still running after ${defaultPollCeilingSec}s ceiling (left running)"
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
# Step 6: Aggregate and return result
# ========================================
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount `
    -Title "Process Runner Results")
