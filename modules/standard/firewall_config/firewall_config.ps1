# ========================================
# Firewall Configuration Script
# ========================================

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Firewall Configuration" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# ========================================
# Step 0: Load Configuration from CSV
# ========================================
$csvPath = Join-Path $PSScriptRoot "firewall_list.csv"
$autoConfig = $null

if (Test-Path $csvPath) {
    $csvData = Import-CsvSafe $csvPath
    if ($csvData -and (Test-CsvColumns -CsvData $csvData -RequiredColumns @("Enabled","status","id","description") -CsvName "firewall_list.csv")) {
        $activeConfig = $csvData | Where-Object { $_.Enabled -eq '1' } | Select-Object -First 1
        if ($activeConfig) {
            Write-Host "[INFO] Loaded configuration from firewall_list.csv" -ForegroundColor Cyan
            $autoConfig = $activeConfig.status
        } else {
            Write-Host "[INFO] firewall_list.csv found but no enabled entries. Switching to manual mode." -ForegroundColor Gray
        }
    }
} else {
    Write-Host "[INFO] firewall_list.csv not found. Switching to manual mode." -ForegroundColor Gray
}

# ========================================
# Step 1: Get Current Status
# ========================================
Write-Host ""
Write-Host "[INFO] Getting current firewall status..." -ForegroundColor Cyan
Write-Host ""

try {
    $profiles = Get-NetFirewallProfile -ErrorAction Stop
}
catch {
    Write-Host "[ERROR] Failed to get firewall info: $_" -ForegroundColor Red
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Failed to get firewall info: $_")
}

# Display status for each profile
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "Current Firewall Status" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

foreach ($p in $profiles) {
    $statusText = if ($p.Enabled) { "ON (Enabled)" } else { "OFF (Disabled)" }
    $statusColor = if ($p.Enabled) { "Green" } else { "Red" }
    Write-Host "  $($p.Name): " -NoNewline -ForegroundColor White
    Write-Host "$statusText" -ForegroundColor $statusColor
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

# Determine status of all profiles
$allEnabled = ($profiles | Where-Object { $_.Enabled -eq $true }).Count -eq $profiles.Count
$allDisabled = ($profiles | Where-Object { $_.Enabled -eq $false }).Count -eq $profiles.Count

# ========================================
# Step 2: Select Operation (Auto/Manual)
# ========================================
$choice = $null

if ($autoConfig) {
    Write-Host "Auto-configuration selected from CSV: " -NoNewline -ForegroundColor Cyan
    Write-Host $autoConfig.ToUpper() -ForegroundColor Yellow
    Write-Host ""

    if ($autoConfig -eq 'off') {
        $choice = '1'
    } elseif ($autoConfig -eq 'on') {
        $choice = '2'
    } else {
        Write-Host "[WARNING] Unknown status in CSV: $autoConfig. Falling back to manual." -ForegroundColor Yellow
    }
}

# Manual Selection (Fallback)
if ([string]::IsNullOrEmpty($choice)) {
    Write-Host "Please select an operation:" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  [1] Disable all profiles (OFF)" -ForegroundColor White
    Write-Host "  [2] Enable all profiles (ON)" -ForegroundColor White
    Write-Host "  [0] Cancel" -ForegroundColor White
    Write-Host ""

    if ($allDisabled) {
        Write-Host "  * Currently all disabled" -ForegroundColor Gray
    }
    elseif ($allEnabled) {
        Write-Host "  * Currently all enabled" -ForegroundColor Gray
    }
    else {
        Write-Host "  * Status varies by profile" -ForegroundColor Yellow
    }

    Write-Host ""
    Write-Host -NoNewline "Enter number: "
    $choice = Read-Host
}

if ($choice -eq '0') {
    Write-Host ""
    Write-Host "[INFO] Canceled" -ForegroundColor Cyan
    Write-Host ""
    return (New-ModuleResult -Status "Cancelled" -Message "User canceled")
}

$targetEnabled = $null
$actionText = ""

switch ($choice) {
    '1' {
        $targetEnabled = "False"
        $actionText = "Disabled (OFF)"
    }
    '2' {
        $targetEnabled = "True"
        $actionText = "Enabled (ON)"
    }
    default {
        Write-Host ""
        Write-Host "[ERROR] Invalid number" -ForegroundColor Red
        Write-Host ""
        return (New-ModuleResult -Status "Error" -Message "Invalid number")
    }
}

# ========================================
# Idempotency Check
# ========================================
$targetBool = ($targetEnabled -eq "True")
$alreadyMatch = ($profiles | Where-Object { $_.Enabled -eq $targetBool }).Count -eq $profiles.Count

if ($alreadyMatch) {
    Write-Host ""
    Write-Host "[INFO] All profiles are already $actionText. Skipping." -ForegroundColor Green
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "Already $actionText")
}

# ========================================
# Step 3: Confirmation
# ========================================
Write-Host ""
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "Set firewall for all profiles to $actionText" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

if (-not (Confirm-Execution -Message "Are you sure you want to execute?")) {
    Write-Host ""
    Write-Host "[INFO] Canceled" -ForegroundColor Cyan
    Write-Host ""
    return (New-ModuleResult -Status "Cancelled" -Message "User canceled")
}

Write-Host ""

# ========================================
# Step 4: Apply Settings
# ========================================
$successCount = 0
$errorCount = 0

foreach ($p in $profiles) {
    $profileName = $p.Name
    Write-Host "[INFO] Changing $profileName profile to $actionText..." -ForegroundColor Cyan

    try {
        Set-NetFirewallProfile -Name $profileName -Enabled $targetEnabled -ErrorAction Stop
        Write-Host "[SUCCESS] ${profileName}: $actionText" -ForegroundColor Green
        $successCount++
    }
    catch {
        Write-Host "[ERROR] ${profileName}: $_" -ForegroundColor Red
        $errorCount++
    }
}

Write-Host ""

# ========================================
# Step 5: Verify Status After Change
# ========================================
Write-Host "[INFO] Verifying status after change..." -ForegroundColor Cyan
Write-Host ""

try {
    $afterProfiles = Get-NetFirewallProfile -ErrorAction Stop

    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Firewall Status After Change" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""

    foreach ($p in $afterProfiles) {
        $statusText = if ($p.Enabled) { "ON (Enabled)" } else { "OFF (Disabled)" }
        $statusColor = if ($p.Enabled) { "Green" } else { "Red" }
        Write-Host "  $($p.Name): " -NoNewline -ForegroundColor White
        Write-Host "$statusText" -ForegroundColor $statusColor
    }

    Write-Host ""
}
catch {
    Write-Host "[WARNING] Failed to verify status after change" -ForegroundColor Yellow
}

# Result Summary
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Configuration Results" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Success: $successCount items" -ForegroundColor Green
if ($errorCount -gt 0) {
    Write-Host "Failed: $errorCount items" -ForegroundColor Red
}
Write-Host ""

# Return ModuleResult
$overallStatus = if ($errorCount -eq 0 -and $successCount -gt 0) { "Success" }
    elseif ($successCount -gt 0 -and $errorCount -gt 0) { "Partial" }
    else { "Error" }
return (New-ModuleResult -Status $overallStatus -Message "Success: $successCount, Fail: $errorCount")
