# ========================================
# Test Harness Module
# ========================================
# Data-driven test simulator for verifying fabriq framework features
# without polluting the system. Reproduces every ModuleResult Status,
# Verified flag, AutoPilot ErrorMode (Skip / Retry / Ask) interaction,
# and provides a per-attempt counter for retry-loop testing.
#
# All scenarios are defined in test_harness_list.csv. Profile callers
# select a scenario bundle via the Segment column (auto-filtered by
# Import-ModuleCsv).
#
# [NOTES]
# - No registry / file system / process side effects (only Start-Sleep).
# - Cross-invocation state lives in $global:FabriqTestHarnessState
#   and is discarded when the PowerShell process exits (e.g. on
#   __RESTART__ or session end).
# ========================================

Write-Host ""
Show-Separator
Write-Host "Test Harness" -ForegroundColor Cyan
Show-Separator
Write-Host ""


# ========================================
# Step 1: Load scenario CSV
# ========================================
$csvPath = Join-Path $PSScriptRoot "test_harness_list.csv"

$enabledItems = Import-ModuleCsv -Path $csvPath -FilterEnabled `
    -RequiredColumns @("Enabled", "TestName", "Behavior")

if ($null -eq $enabledItems) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load test_harness_list.csv")
}
if ($enabledItems.Count -eq 0) {
    return (New-ModuleResult -Status "Skipped" -Message "No enabled scenarios for current segment")
}


# ========================================
# Step 3: Dry-run preview
# ========================================
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "Test Scenarios" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

foreach ($item in $enabledItems) {
    $verifiedDisp   = if ([string]::IsNullOrWhiteSpace($item.Verified))   { "(null)" } else { $item.Verified }
    $failFirstDisp  = if ([string]::IsNullOrWhiteSpace($item.FailFirstN)) { "0"      } else { $item.FailFirstN }
    $delayDisp      = if ([string]::IsNullOrWhiteSpace($item.DelaySec))   { "0"      } else { $item.DelaySec }

    Write-Host "  [TEST] $($item.TestName)" -ForegroundColor Yellow
    Write-Host "    Behavior=$($item.Behavior)  Verified=$verifiedDisp  FailFirstN=$failFirstDisp  DelaySec=$delayDisp" -ForegroundColor DarkGray
    if ($item.Description) {
        Write-Host "    $($item.Description)" -ForegroundColor DarkGray
    }
    Write-Host ""
}

Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""


# ========================================
# Step 4: Confirmation (auto-Y in AutoPilot)
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Run test harness scenarios?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""


# ========================================
# Step 5: Process each scenario item
# ========================================
if ($null -eq $global:FabriqTestHarnessState) {
    $global:FabriqTestHarnessState = @{}
}

# Resolve current segment from environment for state keying
$currentSegment = if ([string]::IsNullOrWhiteSpace($env:FABRIQ_SEGMENT)) { "" } else { $env:FABRIQ_SEGMENT.Trim() }

$successCount = 0
$skipCount    = 0
$failCount    = 0

# Verified aggregation: collect each item's Verified value
$verifiedValues = @()

foreach ($item in $enabledItems) {
    $displayName = if ($item.Description) { $item.Description } else { $item.TestName }

    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "Processing: $displayName" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor White

    # Optional delay
    $delaySec = 0
    if (-not [string]::IsNullOrWhiteSpace($item.DelaySec)) {
        $delaySec = [int]$item.DelaySec
    }
    if ($delaySec -gt 0) {
        Show-Info "Sleeping ${delaySec}s..."
        Start-Sleep -Seconds $delaySec
    }

    # Cancel behavior - immediately abort the whole module
    if ($item.Behavior -eq "cancel") {
        Write-Host ""
        Show-Warning "[TEST] Simulating user cancellation"
        Write-Host ""
        return (New-ModuleResult -Status "Cancelled" -Message "Simulated cancellation by test harness")
    }

    # FailFirstN handling - bump per-(segment+testname) attempt counter
    $stateKey = "${currentSegment}::$($item.TestName)"
    if (-not $global:FabriqTestHarnessState.ContainsKey($stateKey)) {
        $global:FabriqTestHarnessState[$stateKey] = 0
    }
    $global:FabriqTestHarnessState[$stateKey]++
    $attempt = $global:FabriqTestHarnessState[$stateKey]

    $failFirstN = 0
    if (-not [string]::IsNullOrWhiteSpace($item.FailFirstN)) {
        $failFirstN = [int]$item.FailFirstN
    }

    $effectiveBehavior = $item.Behavior
    if ($failFirstN -gt 0 -and $attempt -le $failFirstN) {
        Show-Warning "[TEST] Forcing fail (attempt $attempt of FailFirstN=$failFirstN)"
        $effectiveBehavior = "fail"
    }

    # Apply behavior
    switch ($effectiveBehavior) {
        "success" {
            Show-Success "Completed: $displayName"
            $successCount++
            $verifiedValues += ,$item.Verified
        }
        "fail" {
            Show-Error "Failed: $displayName"
            $failCount++
            $verifiedValues += ,$item.Verified
        }
        "skip" {
            Show-Skip "Skipped: $displayName"
            $skipCount++
            $verifiedValues += ,$item.Verified
        }
        default {
            Show-Error "Unknown Behavior '$($item.Behavior)' in $($item.TestName)"
            $failCount++
            $verifiedValues += ,$item.Verified
        }
    }

    Write-Host ""
}


# ========================================
# Step 5.5: Aggregate Verified flag
# ========================================
# Rules:
#   - any "false" -> $false (FAIL)
#   - all "true"  -> $true (PASS)
#   - any blank with no false -> $null (no verification reported)
$hasFalse = $false
$hasBlank = $false
$hasTrue  = $false

foreach ($v in $verifiedValues) {
    if ([string]::IsNullOrWhiteSpace($v)) {
        $hasBlank = $true
    }
    elseif ($v -match '^(?i)false$') {
        $hasFalse = $true
    }
    elseif ($v -match '^(?i)true$') {
        $hasTrue = $true
    }
    else {
        $hasBlank = $true
    }
}

$verified = $null
if ($hasFalse) {
    $verified = $false
}
elseif ($hasTrue -and -not $hasBlank) {
    $verified = $true
}
else {
    $verified = $null
}


# ========================================
# Step 6: Aggregate result
# ========================================
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount `
    -Title "Test Harness Results" -Verified $verified)
