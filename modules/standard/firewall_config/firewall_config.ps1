# ========================================
# Firewall Configuration Script
# ========================================
# [PURPOSE]
# Configure Windows Firewall profiles (Domain/Private/Public) individually.
# Supports per-profile on/off via CSV, legacy all-at-once CSV, and manual menu.
# ========================================

Write-Host ""
Show-Separator
Write-Host "Firewall Configuration" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# ========================================
# Step 1: Get Current Status
# ========================================
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

# ========================================
# Step 2: Load CSV and Build Target Configuration
# ========================================
# $targets = array of { Profile, TargetEnabled(bool), ActionText }
$targets = @()
$csvMode = $false

$csvPath = Join-Path $PSScriptRoot "firewall_list.csv"

if (Test-Path $csvPath) {
    $csvData = Import-ModuleCsv -Path $csvPath -FilterEnabled `
        -RequiredColumns @("Enabled", "Status")

    if ($null -ne $csvData) {
        $csvArray = @($csvData)
        $hasProfileColumn = $csvArray[0].PSObject.Properties.Name -contains 'Profile'

        if ($hasProfileColumn) {
            # New format: per-profile rows
            foreach ($row in $csvArray) {
                $profileName = $row.Profile.Trim()
                $status = $row.Status.Trim().ToLower()

                if ($profileName -notin @('Domain', 'Private', 'Public')) {
                    Show-Warning "Unknown profile name: $profileName (skipping)"
                    continue
                }
                if ($status -notin @('on', 'off')) {
                    Show-Warning "Unknown status: $status for $profileName (skipping)"
                    continue
                }

                $targets += [PSCustomObject]@{
                    Profile       = $profileName
                    TargetEnabled = ($status -eq 'on')
                    ActionText    = if ($status -eq 'on') { "Enabled (ON)" } else { "Disabled (OFF)" }
                }
            }

            if ($targets.Count -gt 0) {
                $csvMode = $true
                Show-Info "CSV mode: per-profile configuration ($($targets.Count) profile(s))"
            }
        }
        else {
            # Legacy format: single status for all profiles
            $activeConfig = $csvArray | Select-Object -First 1
            $status = $activeConfig.Status.Trim().ToLower()

            if ($status -in @('on', 'off')) {
                $targetBool = ($status -eq 'on')
                $actionText = if ($targetBool) { "Enabled (ON)" } else { "Disabled (OFF)" }

                foreach ($p in $profiles) {
                    $targets += [PSCustomObject]@{
                        Profile       = $p.Name
                        TargetEnabled = $targetBool
                        ActionText    = $actionText
                    }
                }
                $csvMode = $true
                Show-Info "CSV mode: legacy format - all profiles $actionText"
            }
            else {
                Show-Warning "Unknown status in CSV: $status. Falling back to manual."
            }
        }
    }
}
else {
    Show-Info "firewall_list.csv not found. Manual mode."
}

# ========================================
# Step 2.5: Manual Mode (Fallback)
# ========================================
if (-not $csvMode) {
    Write-Host "Please select an operation:" -ForegroundColor Cyan
    Write-Host ""

    $idx = 1
    foreach ($p in $profiles) {
        $current = if ($p.Enabled) { "ON" } else { "OFF" }
        $toggle  = if ($p.Enabled) { "OFF" } else { "ON" }
        Write-Host "  [$idx] $($p.Name) : $current -> $toggle" -ForegroundColor White
        $idx++
    }
    Write-Host ""
    Write-Host "  [4] All OFF" -ForegroundColor White
    Write-Host "  [5] All ON" -ForegroundColor White
    Write-Host "  [0] Cancel" -ForegroundColor White
    Write-Host ""

    Write-Host -NoNewline "Enter number (comma-separated for multiple, e.g. 1,3): "
    $choice = Read-Host

    if ($choice -eq '0') {
        Write-Host ""
        Show-Info "Canceled"
        Write-Host ""
        return (New-ModuleResult -Status "Cancelled" -Message "User canceled")
    }

    if ($choice -eq '4') {
        foreach ($p in $profiles) {
            $targets += [PSCustomObject]@{
                Profile       = $p.Name
                TargetEnabled = $false
                ActionText    = "Disabled (OFF)"
            }
        }
    }
    elseif ($choice -eq '5') {
        foreach ($p in $profiles) {
            $targets += [PSCustomObject]@{
                Profile       = $p.Name
                TargetEnabled = $true
                ActionText    = "Enabled (ON)"
            }
        }
    }
    else {
        $selections = $choice -split ',' | ForEach-Object { $_.Trim() }
        foreach ($sel in $selections) {
            $num = 0
            if (-not [int]::TryParse($sel, [ref]$num) -or $num -lt 1 -or $num -gt $profiles.Count) {
                Show-Error "Invalid selection: $sel"
                Write-Host ""
                return (New-ModuleResult -Status "Error" -Message "Invalid selection: $sel")
            }

            $p = $profiles[$num - 1]
            $newState = -not $p.Enabled
            $targets += [PSCustomObject]@{
                Profile       = $p.Name
                TargetEnabled = $newState
                ActionText    = if ($newState) { "Enabled (ON)" } else { "Disabled (OFF)" }
            }
        }
    }
}

if ($targets.Count -eq 0) {
    Show-Info "No profiles to configure"
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "No profiles to configure")
}

# ========================================
# Step 3: Idempotency Check & Confirmation
# ========================================
Write-Host ""
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "Planned Changes" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

$changesToApply = @()
$skipCount = 0

foreach ($t in @($targets)) {
    $current = $profiles | Where-Object { $_.Name -eq $t.Profile }
    $currentText = if ($current.Enabled) { "ON" } else { "OFF" }
    $targetText  = if ($t.TargetEnabled) { "ON" } else { "OFF" }

    if ($current.Enabled -eq $t.TargetEnabled) {
        Write-Host "  [SKIP] $($t.Profile): already $($t.ActionText)" -ForegroundColor Gray
        $skipCount++
    }
    else {
        Write-Host "  [APPLY] $($t.Profile): $currentText -> $targetText" -ForegroundColor Yellow
        $changesToApply += $t
    }
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

if ($changesToApply.Count -eq 0) {
    Show-Skip "All profiles are already at the expected state. Skipping."
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "Already at expected state" -Verified $true)
}

$cancelResult = Confirm-ModuleExecution -Message "Are you sure you want to execute?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""

# ========================================
# Step 4: Apply Settings
# ========================================
$successCount = 0
$failCount = 0

foreach ($t in @($changesToApply)) {
    Show-Info "Changing $($t.Profile) to $($t.ActionText)..."
    $enabledValue = if ($t.TargetEnabled) { "True" } else { "False" }

    try {
        Set-NetFirewallProfile -Name $t.Profile -Enabled $enabledValue -ErrorAction Stop
        Show-Success "$($t.Profile): $($t.ActionText)"
        $successCount++
    }
    catch {
        Show-Error "$($t.Profile): $_"
        $failCount++
    }
}

Write-Host ""

# ========================================
# Step 5.5: Post-Apply Verification
# ========================================
Show-Info "Verifying applied settings..."
Write-Host ""

$verified = $null
try {
    $afterProfiles = Get-NetFirewallProfile -ErrorAction Stop
    $verifyPass = 0
    $verifyFail = 0

    foreach ($t in @($targets)) {
        $actual = $afterProfiles | Where-Object { $_.Name -eq $t.Profile }
        $actualText = if ($actual.Enabled) { "ON (Enabled)" } else { "OFF (Disabled)" }

        if ($actual.Enabled -eq $t.TargetEnabled) {
            Write-Host "  [VERIFIED] $($t.Profile): $actualText" -ForegroundColor Green
            $verifyPass++
        }
        else {
            Write-Host "  [VERIFY FAILED] $($t.Profile): expected $($t.ActionText), got $actualText" -ForegroundColor Red
            $verifyFail++
        }
    }

    Write-Host ""
    $verified = ($verifyFail -eq 0)
}
catch {
    Show-Warning "Failed to verify status after change: $_"
}

return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount `
    -Title "Configuration Results" -Verified $verified)
