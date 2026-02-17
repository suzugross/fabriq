# ========================================
# Firewall Configuration Script
# ========================================

Write-Host ""
Show-Separator
Write-Host "Firewall Configuration" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# ========================================
# Step 0: Load Configuration from CSV
# ========================================
$csvPath = Join-Path $PSScriptRoot "firewall_list.csv"
$autoConfig = $null

if (Test-Path $csvPath) {
    $csvData = Import-ModuleCsv -Path $csvPath -RequiredColumns @("Enabled", "status")
    if ($null -ne $csvData) {
        $activeConfig = $csvData | Where-Object { $_.Enabled -eq '1' } | Select-Object -First 1
        if ($activeConfig) {
            $autoConfig = $activeConfig.status
        } else {
            Show-Info "firewall_list.csv found but no enabled entries. Switching to manual mode."
        }
    }
} else {
    Show-Info "firewall_list.csv not found. Switching to manual mode."
}

# ========================================
# Step 1: Get Current Status
# ========================================
Write-Host ""
Show-Info "Getting current firewall status..."
Write-Host ""

try {
    $profiles = Get-NetFirewallProfile -ErrorAction Stop
}
catch {
    Show-Error "Failed to get firewall info: $_"
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
        Show-Warning "Unknown status in CSV: $autoConfig. Falling back to manual."
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
    Show-Info "Canceled"
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
        Show-Error "Invalid number"
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
    Show-Skip "All profiles are already $actionText. Skipping."
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

$cancelResult = Confirm-ModuleExecution -Message "Are you sure you want to execute?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""

# ========================================
# Step 4: Apply Settings
# ========================================
$successCount = 0
$failCount = 0

foreach ($p in $profiles) {
    $profileName = $p.Name
    Show-Info "Changing $profileName profile to $actionText..."

    try {
        Set-NetFirewallProfile -Name $profileName -Enabled $targetEnabled -ErrorAction Stop
        Show-Success "${profileName}: $actionText"
        $successCount++
    }
    catch {
        Show-Error "${profileName}: $_"
        $failCount++
    }
}

Write-Host ""

# ========================================
# Step 5: Verify Status After Change
# ========================================
Show-Info "Verifying status after change..."
Write-Host ""

try {
    $afterProfiles = Get-NetFirewallProfile -ErrorAction Stop

    Show-Separator
    Write-Host "Firewall Status After Change" -ForegroundColor Cyan
    Show-Separator
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
    Show-Warning "Failed to verify status after change"
}

return (New-BatchResult -Success $successCount -Fail $failCount -Title "Configuration Results")
