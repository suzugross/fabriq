# ========================================
# IPv6 Configuration Script
# ========================================

# Check Administrator Privileges
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "Administrator privileges are required."
    Write-Warning "Please run PowerShell as Administrator and try again."
    return (New-ModuleResult -Status "Error" -Message "Administrator privileges required")
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "IPv6 Configuration" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# ========================================
# Load CSV
# ========================================
$csvPath = Join-Path $PSScriptRoot "ipv6_list.csv"

if (-not (Test-Path $csvPath)) {
    Write-Host "[ERROR] ipv6_list.csv not found: $csvPath" -ForegroundColor Red
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "ipv6_list.csv not found")
}

try {
    $ipv6List = @(Import-Csv -Path $csvPath -Encoding Default)
}
catch {
    Write-Host "[ERROR] Failed to load ipv6_list.csv: $_" -ForegroundColor Red
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Failed to load ipv6_list.csv: $_")
}

if ($ipv6List.Count -eq 0) {
    Write-Host "[ERROR] ipv6_list.csv contains no data" -ForegroundColor Red
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "ipv6_list.csv contains no data")
}

Write-Host "[INFO] Loaded $($ipv6List.Count) settings" -ForegroundColor Cyan
Write-Host ""

# ========================================
# List Settings
# ========================================
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host "Target Adapters List" -ForegroundColor Cyan
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""

foreach ($item in $ipv6List) {
    $stateStr = if ($item.Enabled -eq "1") { "Enable" } else { "Disable" }
    $stateColor = if ($item.Enabled -eq "1") { "Green" } else { "Yellow" }
    
    Write-Host "  Pattern: $($item.AdapterPattern)" -ForegroundColor White
    Write-Host "  Action:  $stateStr" -ForegroundColor $stateColor
    Write-Host ""
}

Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""

# ========================================
# Confirmation
# ========================================
if (-not (Confirm-Execution -Message "Apply the above IPv6 settings?")) {
    Write-Host ""
    Write-Host "[INFO] Canceled" -ForegroundColor Yellow
    Write-Host ""
    return (New-ModuleResult -Status "Cancelled" -Message "User canceled")
}

Write-Host ""

# ========================================
# Apply Settings
# ========================================
$successCount = 0
$failCount = 0
$skipCount = 0

foreach ($item in $ipv6List) {
    $pattern = $item.AdapterPattern
    $targetState = ($item.Enabled -eq "1")
    $actionName = if ($targetState) { "Enable" } else { "Disable" }

    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "Processing Pattern: $pattern ($actionName)" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor White

    # Find adapters matching the pattern
    $adapters = Get-NetAdapter | Where-Object { $_.Name -like $pattern }

    if (-not $adapters) {
        Write-Host "[SKIP] No adapters found matching '$pattern'" -ForegroundColor Yellow
        $skipCount++
        Write-Host ""
        continue
    }

    foreach ($adapter in $adapters) {
        try {
            Write-Host "[INFO] Configuring: $($adapter.Name)" -ForegroundColor Gray
            
            Set-NetAdapterBinding -Name $adapter.Name -ComponentID ms_tcpip6 -Enabled $targetState -ErrorAction Stop
            
            Write-Host "[SUCCESS] $($adapter.Name): IPv6 $actionName" -ForegroundColor Green
            $successCount++
        }
        catch {
            Write-Host "[ERROR] Failed to configure $($adapter.Name): $_" -ForegroundColor Red
            $failCount++
        }
    }
    Write-Host ""
}

# ========================================
# Result Summary
# ========================================
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Execution Results" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Success: $successCount items" -ForegroundColor Green
Write-Host "  Skipped: $skipCount patterns" -ForegroundColor Yellow
Write-Host "  Failed:  $failCount items" -ForegroundColor $(if ($failCount -gt 0) { "Red" } else { "Green" })
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Return ModuleResult
$overallStatus = if ($failCount -eq 0 -and $successCount -gt 0) { "Success" }
    elseif ($successCount -gt 0 -and $failCount -gt 0) { "Partial" }
    elseif ($failCount -eq 0 -and $skipCount -gt 0) { "Skipped" }
    else { "Error" }
return (New-ModuleResult -Status $overallStatus -Message "Success: $successCount, Skip: $skipCount, Fail: $failCount")