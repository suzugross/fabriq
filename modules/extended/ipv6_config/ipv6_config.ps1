# ========================================
# IPv6 Configuration Script
# ========================================

# Check Administrator Privileges
if (-not (Test-AdminPrivilege)) {
    Show-Error "This script requires administrator privileges."
    return (New-ModuleResult -Status "Error" -Message "Administrator privileges required")
}

Write-Host ""
Show-Separator
Write-Host "IPv6 Configuration" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# ========================================
# Load CSV
# ========================================
$csvPath = Join-Path $PSScriptRoot "ipv6_list.csv"

$ipv6List = Import-CsvSafe -Path $csvPath -Description "ipv6_list.csv"
if ($null -eq $ipv6List -or $ipv6List.Count -eq 0) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load ipv6_list.csv")
}

Show-Info "Loaded $($ipv6List.Count) settings"
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
$cancelResult = Confirm-ModuleExecution -Message "Apply the above IPv6 settings?"
if ($null -ne $cancelResult) { return $cancelResult }

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
        Show-Skip "No adapters found matching '$pattern'"
        $skipCount++
        Write-Host ""
        continue
    }

    foreach ($adapter in $adapters) {
        try {
            Show-Info "Configuring: $($adapter.Name)"
            
            Set-NetAdapterBinding -Name $adapter.Name -ComponentID ms_tcpip6 -Enabled $targetState -ErrorAction Stop
            
            Show-Success "$($adapter.Name): IPv6 $actionName"
            $successCount++
        }
        catch {
            Show-Error "Failed to configure $($adapter.Name): $_"
            $failCount++
        }
    }
    Write-Host ""
}

# ========================================
# Result Summary
# ========================================
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount -Title "Execution Results")