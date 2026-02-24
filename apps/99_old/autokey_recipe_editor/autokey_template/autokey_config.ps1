# ========================================
# Auto Keyboard - Automated Input Tool
# ========================================
# Template module for keyboard automation.
# Copy this directory and edit recipe.csv to create
# custom automation modules.
# ========================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic

# ========================================
# Helper Functions
# ========================================

function Invoke-ActionOpen {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }

    try {
        if ($Value -match " ") {
            $parts = $Value -split " ", 2
            try {
                $null = Start-Process -FilePath $parts[0] -ArgumentList $parts[1] -ErrorAction Stop -PassThru
            }
            catch {
                # Fallback: shell execution (for URLs etc)
                $null = Start-Process -FilePath "cmd" -ArgumentList "/c start $($parts[0]) $($parts[1])" -WindowStyle Hidden
            }
        }
        else {
            $null = Start-Process $Value -ErrorAction Stop -PassThru
        }
        return $true
    }
    catch {
        Write-Host "  [ERROR] Failed to open: $Value - $_" -ForegroundColor Red
        return $false
    }
}

function Invoke-ActionWaitWin {
    param([string]$Title, [int]$TimeoutMs)
    if ($TimeoutMs -le 0) { $TimeoutMs = 10000 }
    $elapsed = 0

    Write-Host "  Waiting for window: '$Title' (max $($TimeoutMs / 1000)s)..." -NoNewline -ForegroundColor Gray

    while ($elapsed -lt $TimeoutMs) {
        $proc = Get-Process | Where-Object { $_.MainWindowTitle -like "*$Title*" } | Select-Object -First 1
        if ($proc) {
            Write-Host " Found!" -ForegroundColor Green
            try {
                $script:WsShell.AppActivate($proc.Id) | Out-Null
                Start-Sleep -Milliseconds 500
            } catch {}
            return $true
        }
        Start-Sleep -Milliseconds 500
        $elapsed += 500
        Write-Host "." -NoNewline
    }

    Write-Host " Timeout!" -ForegroundColor Yellow
    return $false
}

function Invoke-ActionAppFocus {
    param([string]$Title)
    try {
        $result = $script:WsShell.AppActivate($Title)
        return [bool]$result
    }
    catch { return $false }
}

function Invoke-ActionType {
    param([string]$Text)
    try {
        [System.Windows.Forms.SendKeys]::SendWait($Text)
        return $true
    }
    catch { return $false }
}

function Invoke-ActionKey {
    param([string]$KeySequence)
    try {
        [System.Windows.Forms.SendKeys]::SendWait($KeySequence)
        return $true
    }
    catch { return $false }
}

# ========================================
# Main Process
# ========================================

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Auto Keyboard - Automated Input" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# ========================================
# Step 1: Load Recipe CSV
# ========================================
$csvPath = Join-Path $PSScriptRoot "recipe.csv"
if (-not (Test-Path $csvPath)) {
    Write-Host "[ERROR] recipe.csv not found: $csvPath" -ForegroundColor Red
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "recipe.csv not found")
}

try {
    $allSteps = @(Import-Csv -Path $csvPath -Encoding Default)
}
catch {
    Write-Host "[ERROR] Failed to read recipe.csv: $_" -ForegroundColor Red
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Failed to read recipe.csv")
}

# Filter empty rows
$steps = @($allSteps | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Action) })
if ($steps.Count -eq 0) {
    Write-Host "[INFO] No steps defined in recipe.csv" -ForegroundColor Gray
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "No steps defined")
}

# ========================================
# Step 2: Display Recipe Steps
# ========================================
Write-Host "[INFO] Recipe: $($steps.Count) steps" -ForegroundColor Cyan
Write-Host ""

foreach ($step in $steps) {
    $stepNum = if ($step.Step) { $step.Step } else { "-" }
    $actionDisplay = switch ($step.Action) {
        "Open"     { "Open:     $($step.Value)" }
        "WaitWin"  { "WaitWin:  '$($step.Value)' (timeout: $($step.Wait)ms)" }
        "AppFocus" { "Focus:    '$($step.Value)'" }
        "Type"     { "Type:     '$($step.Value)'" }
        "Key"      { "Key:      $($step.Value)" }
        "Wait"     { "Wait:     $($step.Value)ms" }
        default    { "$($step.Action): $($step.Value)" }
    }
    Write-Host "  [$stepNum] $actionDisplay" -ForegroundColor White
    if ($step.Note) {
        Write-Host "       $($step.Note)" -ForegroundColor Gray
    }
}

Write-Host ""

# ========================================
# Step 3: Confirmation
# ========================================
if (-not (Confirm-Execution -Message "Execute the above automation steps?")) {
    Write-Host ""
    Write-Host "[INFO] Canceled" -ForegroundColor Yellow
    Write-Host ""
    return (New-ModuleResult -Status "Cancelled" -Message "User canceled")
}

Write-Host ""

# ========================================
# Step 4: Initialize COM
# ========================================
$script:WsShell = New-Object -ComObject WScript.Shell

# ========================================
# Step 5: Execute Recipe
# ========================================
$successCount = 0
$failCount = 0
$total = $steps.Count
$current = 0

foreach ($step in $steps) {
    $current++
    $stepNum = if ($step.Step) { $step.Step } else { $current }
    $actionResult = $false

    Write-Host "[$current/$total] Step $stepNum - $($step.Action): $($step.Note)" -ForegroundColor Cyan

    switch ($step.Action) {
        "Open" {
            $actionResult = Invoke-ActionOpen -Value $step.Value
            if ($actionResult) {
                Write-Host "  [SUCCESS] Opened: $($step.Value)" -ForegroundColor Green
            }
            else {
                Write-Host "  [ERROR] Failed to open: $($step.Value)" -ForegroundColor Red
            }
        }

        "WaitWin" {
            $timeoutMs = 10000
            if ($step.Wait) {
                $parsed = 0
                if ([int]::TryParse($step.Wait, [ref]$parsed) -and $parsed -gt 0) {
                    $timeoutMs = $parsed
                }
            }
            $actionResult = Invoke-ActionWaitWin -Title $step.Value -TimeoutMs $timeoutMs
            if (-not $actionResult) {
                Write-Host "  [WARNING] Window not found within timeout" -ForegroundColor Yellow
            }
        }

        "AppFocus" {
            $actionResult = Invoke-ActionAppFocus -Title $step.Value
            if ($actionResult) {
                Write-Host "  [SUCCESS] Focused: $($step.Value)" -ForegroundColor Green
            }
            else {
                Write-Host "  [WARNING] Focus failed: $($step.Value)" -ForegroundColor Yellow
            }
        }

        "Type" {
            $actionResult = Invoke-ActionType -Text $step.Value
            if ($actionResult) {
                Write-Host "  [SUCCESS] Typed text" -ForegroundColor Green
            }
            else {
                Write-Host "  [ERROR] Type failed" -ForegroundColor Red
            }
        }

        "Key" {
            $actionResult = Invoke-ActionKey -KeySequence $step.Value
            if ($actionResult) {
                Write-Host "  [SUCCESS] Key sent: $($step.Value)" -ForegroundColor Green
            }
            else {
                Write-Host "  [ERROR] Key send failed: $($step.Value)" -ForegroundColor Red
            }
        }

        "Wait" {
            $ms = 0
            if ($step.Value) {
                $null = [int]::TryParse($step.Value, [ref]$ms)
            }
            if ($ms -gt 0) {
                Start-Sleep -Milliseconds $ms
            }
            $actionResult = $true
            Write-Host "  [OK] Waited ${ms}ms" -ForegroundColor Gray
        }

        default {
            Write-Host "  [WARNING] Unknown action: $($step.Action)" -ForegroundColor Yellow
            $actionResult = $false
        }
    }

    if ($actionResult) { $successCount++ } else { $failCount++ }

    # Post-step wait (WaitWin and Wait handle their own timing)
    if ($step.Action -ne "WaitWin" -and $step.Action -ne "Wait") {
        $waitMs = 0
        if ($step.Wait) {
            $null = [int]::TryParse($step.Wait, [ref]$waitMs)
        }
        if ($waitMs -gt 0) {
            Start-Sleep -Milliseconds $waitMs
        }
    }

    Write-Host ""
}

# ========================================
# Step 6: Cleanup COM
# ========================================
try {
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:WsShell) | Out-Null
} catch {}

# ========================================
# Result Summary
# ========================================
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Automation Results" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
if ($successCount -gt 0) {
    Write-Host "  Success: $successCount steps" -ForegroundColor Green
}
if ($failCount -gt 0) {
    Write-Host "  Failed:  $failCount steps" -ForegroundColor Red
}
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Return ModuleResult
$overallStatus = if ($failCount -eq 0 -and $successCount -gt 0) { "Success" }
    elseif ($successCount -gt 0 -and $failCount -gt 0) { "Partial" }
    elseif ($failCount -gt 0) { "Error" }
    else { "Success" }
return (New-ModuleResult -Status $overallStatus -Message "Success: $successCount, Fail: $failCount")
