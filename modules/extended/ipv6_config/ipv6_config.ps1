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

$ipv6List = Import-ModuleCsv -Path $csvPath -RequiredColumns @("Enabled", "AdapterPattern", "IPv6State")
if ($null -eq $ipv6List) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load ipv6_list.csv")
}

$enabledItems = @($ipv6List | Where-Object { $_.Enabled -eq "1" })
$disabledItems = @($ipv6List | Where-Object { $_.Enabled -ne "1" })

if ($enabledItems.Count -eq 0) {
    Show-Info "No enabled entries in ipv6_list.csv"
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "No enabled entries")
}
Write-Host ""

# ========================================
# List Settings
# ========================================
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host "Target Adapters List" -ForegroundColor Cyan
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""

foreach ($item in $enabledItems) {
    $stateStr = if ($item.IPv6State -eq "1") { "Enable" } else { "Disable" }
    $stateColor = if ($item.IPv6State -eq "1") { "Green" } else { "Yellow" }
    $descStr = if ($item.Description) { " ($($item.Description))" } else { "" }

    Write-Host "  Pattern: $($item.AdapterPattern)$descStr" -ForegroundColor White
    Write-Host "  Action:  $stateStr" -ForegroundColor $stateColor
    Write-Host ""
}

foreach ($item in $disabledItems) {
    $descStr = if ($item.Description) { " ($($item.Description))" } else { "" }
    Write-Host "  [DISABLED] $($item.AdapterPattern)$descStr" -ForegroundColor DarkGray
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

foreach ($item in $enabledItems) {
    $pattern = $item.AdapterPattern
    $targetState = ($item.IPv6State -eq "1")
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