# ========================================
# Script Looper
# ========================================
# Executes specified scripts with retry/loop logic based on
# CSV-defined conditions (OnError / Always).
#
# [NOTES]
# - Target scripts should return New-ModuleResult for proper status detection.
# - Legacy scripts (no ModuleResult) are also supported:
#   exception = Error, exit code != 0 = Error, otherwise Success.
#   A .ps1 child that runs native commands should end with an explicit
#   `exit 0`, or a failing native command mid-script will be reported
#   as the child's exit code (fail-closed by design).
# - Uses the same dual-detection pattern as Invoke-KittingScript
#   (pipeline output + $global:_LastModuleResult fallback).
# ========================================

Write-Host ""
Show-Separator
Write-Host "Script Looper" -ForegroundColor Cyan
Show-Separator
Write-Host ""


# ========================================
# Step 1: Load CSV
# ========================================
$csvPath = Join-Path $PSScriptRoot "looper_list.csv"

$enabledItems = Import-ModuleCsv -Path $csvPath -FilterEnabled `
    -RequiredColumns @("Enabled", "ScriptPath", "MaxRetry", "IntervalSec", "Condition")

if ($null -eq $enabledItems) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load looper_list.csv")
}
if ($enabledItems.Count -eq 0) {
    return (New-ModuleResult -Status "Skipped" -Message "No enabled entries")
}


# ========================================
# Step 2: Prerequisite check (resolve paths and validate parameters)
# ========================================
$baseDir = (Get-Location).Path

foreach ($item in $enabledItems) {
    # --- Resolve ScriptPath (absolute, or relative to the fabriq root) ---
    $resolvedPath = if ([System.IO.Path]::IsPathRooted($item.ScriptPath)) {
        $item.ScriptPath
    } else {
        Join-Path $baseDir $item.ScriptPath
    }
    $null = $item | Add-Member -NotePropertyName "_ResolvedPath" -NotePropertyValue $resolvedPath -Force
    $null = $item | Add-Member -NotePropertyName "_PathValid" -NotePropertyValue (Test-Path $resolvedPath) -Force

    # --- Validate MaxRetry / IntervalSec / Condition ---
    $maxRetry = 0
    $intervalSec = 0
    $validParams = $true

    if (-not [int]::TryParse($item.MaxRetry, [ref]$maxRetry) -or $maxRetry -lt 1) {
        $validParams = $false
    }
    if (-not [int]::TryParse($item.IntervalSec, [ref]$intervalSec) -or $intervalSec -lt 0) {
        $validParams = $false
    }
    if ($item.Condition -notin @("OnError", "Always")) {
        $validParams = $false
    }

    $null = $item | Add-Member -NotePropertyName "_ValidParams" -NotePropertyValue $validParams -Force
}


# ========================================
# Step 3: Dry-run summary before execution
# ========================================
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "Loop Targets" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

foreach ($item in $enabledItems) {
    $displayName = if ($item.Description) { $item.Description } else { $item.ScriptPath }

    if (-not $item._PathValid) {
        Write-Host "  [NOT FOUND] $displayName" -ForegroundColor Red
        Write-Host "    Path: $($item._ResolvedPath)" -ForegroundColor DarkGray
    }
    elseif (-not $item._ValidParams) {
        Write-Host "  [INVALID] $displayName" -ForegroundColor Red
        Write-Host "    Path: $($item._ResolvedPath)" -ForegroundColor DarkGray
        Write-Host "    Check: MaxRetry >= 1, IntervalSec >= 0, Condition = OnError|Always" -ForegroundColor DarkGray
    }
    else {
        Write-Host "  [READY] $displayName" -ForegroundColor Yellow
        Write-Host "    Path: $($item.ScriptPath)  |  Condition: $($item.Condition)  |  Max: $($item.MaxRetry)  |  Interval: $($item.IntervalSec)s" -ForegroundColor DarkGray
    }
    Write-Host ""
}

Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""


# ========================================
# Step 4: User confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Execute the above loop targets?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""


# ========================================
# Step 5: Loop execution
# ========================================
$successCount = 0
$skipCount    = 0
$failCount    = 0

foreach ($item in $enabledItems) {
    $displayName = if ($item.Description) { $item.Description } else { $item.ScriptPath }

    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "Looping: $displayName" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor White

    # ----------------------------------------
    # Per-item prerequisite check (Skip path)
    # ----------------------------------------
    if (-not $item._PathValid) {
        Show-Skip "Script not found: $($item._ResolvedPath)"
        Write-Host ""
        $skipCount++
        continue
    }

    if (-not $item._ValidParams) {
        Show-Skip "Invalid parameters: $displayName"
        Write-Host ""
        $skipCount++
        continue
    }

    $maxRetry    = [int]$item.MaxRetry
    $intervalSec = [int]$item.IntervalSec
    $condition   = $item.Condition
    $scriptPath  = $item._ResolvedPath

    # ----------------------------------------
    # Main work (retry loop)
    # ----------------------------------------
    $lastStatus  = "Error"
    $lastMessage = ""

    for ($attempt = 1; $attempt -le $maxRetry; $attempt++) {
        Show-Info "Attempt $attempt/$maxRetry : $displayName"

        # --- Run the target script (dual ModuleResult detection) ---
        $moduleResult = $null
        try {
            $global:_LastModuleResult = $null
            # Reset so a code left by an EARLIER command is never mistaken
            # for this child's verdict.
            $global:LASTEXITCODE = 0

            $output = & $scriptPath

            # Find a ModuleResult in the pipeline output
            if ($null -ne $output) {
                foreach ($outItem in @($output)) {
                    if ($outItem -is [PSCustomObject] -and $outItem._IsModuleResult -eq $true) {
                        $moduleResult = $outItem
                    }
                }
            }

            # Fallback: $global:_LastModuleResult
            if (-not $moduleResult -and $null -ne $global:_LastModuleResult) {
                $moduleResult = $global:_LastModuleResult
            }
            $global:_LastModuleResult = $null

            if ($moduleResult) {
                # Explicit contract wins - the exit code is not consulted.
                $lastStatus  = $moduleResult.Status
                $lastMessage = $moduleResult.Message
            }
            elseif ($LASTEXITCODE -ne 0) {
                # Legacy script signalling failure the classic way
                # (exit 1 from a .bat / .exe / .ps1). Without this check
                # a failing child counted as Success and the OnError
                # retry - this module's whole purpose - never fired.
                $lastStatus  = "Error"
                $lastMessage = "(legacy - exit code $LASTEXITCODE)"
            }
            else {
                # Legacy script: no ModuleResult, no exception, exit 0 → Success
                $lastStatus  = "Success"
                $lastMessage = "(legacy - no ModuleResult)"
            }
        }
        catch {
            $lastStatus  = "Error"
            $lastMessage = $_.Exception.Message
        }

        # --- Log this attempt's result ---
        switch ($lastStatus) {
            "Success"   { Show-Success "Attempt ${attempt}: Success - $lastMessage" }
            "Error"     { Show-Error   "Attempt ${attempt}: Error - $lastMessage" }
            "Skipped"   { Show-Skip    "Attempt ${attempt}: Skipped - $lastMessage" }
            "Cancelled" { Show-Info    "Attempt ${attempt}: Cancelled - $lastMessage" }
            "Partial"   { Show-Warning "Attempt ${attempt}: Partial - $lastMessage" }
            default     { Show-Info    "Attempt ${attempt}: $lastStatus - $lastMessage" }
        }

        # --- Condition check: should we retry? ---
        $shouldRetry = $false
        switch ($condition) {
            "OnError" {
                # Retry only on Error; any other status ends the loop
                if ($lastStatus -eq "Error") {
                    $shouldRetry = $true
                }
            }
            "Always" {
                # Always retry (the final attempt is gated below)
                $shouldRetry = $true
            }
        }

        # Do not retry past the final attempt
        if ($attempt -ge $maxRetry) {
            $shouldRetry = $false
        }

        # --- Log the reason the loop ends ---
        if (-not $shouldRetry) {
            if ($lastStatus -eq "Error" -and $attempt -ge $maxRetry) {
                Show-Warning "Max retry reached ($maxRetry). Giving up: $displayName"
            }
            elseif ($condition -eq "OnError" -and $lastStatus -ne "Error") {
                Show-Info "No error detected. Loop complete: $displayName"
            }
            elseif ($condition -eq "Always" -and $attempt -ge $maxRetry) {
                Show-Info "All $maxRetry iterations complete: $displayName"
            }
            break
        }

        # --- Wait before retrying ---
        if ($intervalSec -gt 0) {
            Show-Warning "Retrying in ${intervalSec}s... (attempt $attempt/$maxRetry)"
            Start-Sleep -Seconds $intervalSec
        }
        else {
            Show-Warning "Retrying immediately... (attempt $attempt/$maxRetry)"
        }
    }

    # --- Final verdict for this entry ---
    if ($lastStatus -eq "Error") {
        $failCount++
    }
    else {
        $successCount++
    }

    Write-Host ""
}


# ========================================
# Step 6: Aggregate and return result
# ========================================
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount `
    -Title "[Script Looper] Results")
