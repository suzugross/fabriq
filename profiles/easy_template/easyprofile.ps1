# ========================================
# EasyProfile Runner
# ========================================
# Lightweight AutoPilot runner.
# Reads easyprofile.csv and executes the
# listed module scripts in order.
#
# NO logging / NO evidence / NO checklist.
# NO session initialization / NO history.
#
# CSV columns: Enabled, Script, Description, Segment
#   - Segment (optional): when set, exported as
#     $env:FABRIQ_SEGMENT before invocation so
#     Segment-aware modules (Import-ModuleCsv path)
#     filter their _list.csv accordingly. Cleared
#     after each module to avoid pollution.
#
# NOTES:
# - Place this directory at profiles/<name>/
# - fabriq root is resolved as 2 levels up
#   from this script's location.
# - Copy this directory to create profiles:
#     profiles/easy_profile_A/
#     profiles/easy_profile_B/
# ========================================

# ========================================
# Resolve fabriq root and load common.ps1
# ========================================
# profiles/<name>/ -> ../.. = fabriq root
$fabriqRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot "..\.."))
$commonPath  = Join-Path $fabriqRoot "kernel\common.ps1"

if (-not (Test-Path $commonPath)) {
    Write-Host "[ERROR] common.ps1 not found: $commonPath" -ForegroundColor Red
    Write-Host "Ensure this script is located at: profiles/<name>/easyprofile.ps1" -ForegroundColor Gray
    Read-Host "Press Enter to exit"
    exit 1
}

. $commonPath

# ========================================
# AutoPilot: auto-confirm all module prompts
# ========================================
$global:AutoPilotMode    = $true
$global:AutoPilotWaitSec = 0

# ========================================
# Console setup
# ========================================
Set-ConsoleSize -Columns 75 -Lines 35
Enable-SleepSuppression

Write-Host ""
Show-Separator
Write-Host "EasyProfile Runner" -ForegroundColor Green
Write-Host "  Profile : $(Split-Path $PSScriptRoot -Leaf)" -ForegroundColor White
Show-Separator
Write-Host ""


# ========================================
# Load easyprofile.csv
# ========================================
# Plain Import-CsvSafe is used instead of Import-ModuleCsv because
# Import-ModuleCsv applies a strict-match Segment filter to the CSV
# itself when a Segment column is present (intended for module
# _list.csv files). In EasyProfile, the Segment column carries per-row
# parameters for module dispatch, so that filter would silently drop
# every row whose Segment differs from $env:FABRIQ_SEGMENT (typically
# unset on entry). This mirrors what Resolve-ProfileModules does for
# the Linear profile CSV in kernel/common.ps1.
$csvPath = Join-Path $PSScriptRoot "easyprofile.csv"

$allEntries = Import-CsvSafe -Path $csvPath -Description "easyprofile.csv"

if ($null -eq $allEntries) {
    Read-Host "Press Enter to exit"
    Disable-SleepSuppression
    exit 1
}

# Empty CSV (header only). Import-CsvSafe already emitted Show-Warning;
# treat as "no entries" and exit cleanly so Test-CsvColumns doesn't
# silently fail on an empty array.
if ($allEntries.Count -eq 0) {
    Read-Host "Press Enter to exit"
    Disable-SleepSuppression
    exit 0
}

if (-not (Test-CsvColumns -CsvData $allEntries -RequiredColumns @("Enabled", "Script") -CsvName "easyprofile.csv")) {
    Read-Host "Press Enter to exit"
    Disable-SleepSuppression
    exit 1
}

$scriptList = @($allEntries | Where-Object { $_.Enabled -eq "1" })

if ($scriptList.Count -eq 0) {
    Show-Info "No enabled entries in easyprofile.csv"
    Read-Host "Press Enter to exit"
    Disable-SleepSuppression
    exit 0
}


# ========================================
# Pre-execution display
# ========================================
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host "Scripts to Execute" -ForegroundColor Cyan
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""

$idx = 0
foreach ($entry in $scriptList) {
    $idx++
    $segment    = if ($entry.PSObject.Properties.Name -contains 'Segment') { "$($entry.Segment)".Trim() } else { "" }
    $segLabel   = if ($segment) { " [seg:$segment]" } else { "" }
    $baseName   = if ($entry.Description) { $entry.Description } else { [System.IO.Path]::GetFileName($entry.Script) }
    $displayName = "$baseName$segLabel"
    $scriptPath  = Join-Path $fabriqRoot $entry.Script

    if (Test-Path $scriptPath) {
        Write-Host "  [$idx] $displayName" -ForegroundColor Yellow
        Write-Host "      $($entry.Script)" -ForegroundColor DarkGray
    }
    else {
        Write-Host "  [$idx] $displayName  [NOT FOUND]" -ForegroundColor DarkGray
        Write-Host "      $($entry.Script)  (will be skipped)" -ForegroundColor DarkGray
    }
    Write-Host ""
}

Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""


# ========================================
# Execution loop
# ========================================
$successCount = 0
$skipCount    = 0
$failCount    = 0

$idx = 0
foreach ($entry in $scriptList) {
    $idx++
    $segment    = if ($entry.PSObject.Properties.Name -contains 'Segment') { "$($entry.Segment)".Trim() } else { "" }
    $segLabel   = if ($segment) { " [seg:$segment]" } else { "" }
    $baseName   = if ($entry.Description) { $entry.Description } else { [System.IO.Path]::GetFileName($entry.Script) }
    $displayName = "$baseName$segLabel"
    $scriptPath  = [string](Join-Path $fabriqRoot $entry.Script)

    Write-Host ""
    Show-Separator
    Write-Host "[$idx/$($scriptList.Count)] $displayName" -ForegroundColor Green
    Show-Separator
    Write-Host ""

    # --- Script existence check (outside try) ---
    if (-not (Test-Path $scriptPath)) {
        Show-Skip "Script not found: $($entry.Script)"
        Write-Host ""
        $skipCount++
        continue
    }

    # --- Segment parameter passing via $env:FABRIQ_SEGMENT ---
    # Same env-var contract as kernel/main.ps1 (Linear path). Modules
    # consume it implicitly through Import-ModuleCsv's default
    # -Segment parameter. Cleared in `finally` so exceptions or early
    # returns inside the module do not leak the value to the next row.
    if ($segment) {
        $env:FABRIQ_SEGMENT = $segment
    }

    # --- Execute with ModuleResult detection ---
    # Uses the same detection logic as Invoke-KittingScript in main.ps1.
    # Write-ExecutionHistory / Capture-ScreenEvidence are intentionally
    # NOT called — this runner produces no logs or evidence.
    try {
        $global:_LastModuleResult = $null

        $output = & $scriptPath

        # Detect ModuleResult from pipeline output
        $moduleResult = $null
        if ($null -ne $output) {
            foreach ($item in @($output)) {
                if ($item -is [PSCustomObject] -and $item._IsModuleResult -eq $true) {
                    $moduleResult = $item
                }
            }
        }

        # Fallback: global variable (pipeline capture failure)
        if (-not $moduleResult -and $null -ne $global:_LastModuleResult) {
            $moduleResult = $global:_LastModuleResult
        }
        $global:_LastModuleResult = $null

        if ($moduleResult) {
            $status  = $moduleResult.Status
            $message = $moduleResult.Message

            switch ($status) {
                "Success"   { Write-Host ""; Show-Success "Completed: $displayName" }
                "Error"     { Write-Host ""; Show-Error   "Error: $displayName — $message" }
                "Cancelled" { Write-Host ""; Show-Info    "Cancelled: $displayName" }
                "Skipped"   { Write-Host ""; Show-Skip    "Skipped: $displayName — $message" }
                "Partial"   { Write-Host ""; Show-Warning "Partial: $displayName — $message" }
            }

            if ($status -eq "Success" -or $status -eq "Partial") { $successCount++ }
            elseif ($status -eq "Skipped" -or $status -eq "Cancelled") { $skipCount++ }
            else { $failCount++ }
        }
        else {
            Write-Host ""
            Show-Warning "Completed (no ModuleResult returned)"
            $successCount++
        }
    }
    catch {
        Write-Host ""
        Show-Error "Exception in '$displayName': $_"
        $failCount++
    }
    finally {
        if ($segment) {
            $env:FABRIQ_SEGMENT = $null
        }
    }
}


# ========================================
# Final summary
# ========================================
Write-Host ""
Show-Separator
Write-Host "EasyProfile Results" -ForegroundColor Green
Show-Separator
if ($successCount -gt 0) { Write-Host "  Success: $successCount scripts" -ForegroundColor Green }
if ($skipCount    -gt 0) { Write-Host "  Skipped: $skipCount scripts"    -ForegroundColor Gray  }
if ($failCount    -gt 0) { Write-Host "  Failed:  $failCount scripts"    -ForegroundColor Red   }
Show-Separator
Write-Host ""

Disable-SleepSuppression
Read-Host "Press Enter to exit"
