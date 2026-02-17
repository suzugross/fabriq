# ========================================
# Hostname Change Script
# ========================================

$HOSTLIST_CSV = Join-Path $PSScriptRoot "..\..\..\kernel\csv\hostlist.csv"

Write-Host ""
Show-Separator
Write-Host "Hostname Change" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# Load CSV
$hostItems = Import-CsvSafe -Path $HOSTLIST_CSV -Description "hostlist.csv"
if ($null -eq $hostItems) {
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Failed to load hostlist.csv")
}

# Filter items with NewPCName
$validItems = @()
foreach ($item in $hostItems) {
    if (-not [string]::IsNullOrEmpty($item.'NewPCName')) {
        $validItems += $item
    }
}

if ($validItems.Count -eq 0) {
    Show-Error "No NewPCName registered in hostlist.csv"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "No NewPCName registered in hostlist.csv")
}

# Current Hostname
$currentHostname = $env:COMPUTERNAME
Write-Host "  Current Hostname: $currentHostname" -ForegroundColor White
Write-Host ""

# List New PC Names
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "New Hostname List" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

for ($i = 0; $i -lt $validItems.Count; $i++) {
    $no = $validItems[$i].'AdminID'
    $oldName = $validItems[$i].'OldPCName'
    $newName = $validItems[$i].'NewPCName'
    Write-Host "  [$($i + 1)] $newName  (AdminID: $no / OldName: $oldName)" -ForegroundColor White
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

# Select Number
Write-Host -NoNewline "Select the number for the new PC name (0 to cancel): "
$selection = Read-Host

# Cancel
if ($selection -eq '0') {
    Write-Host ""
    Show-Info "Canceled"
    Write-Host ""
    return (New-ModuleResult -Status "Cancelled" -Message "User canceled")
}

# Validate Number
$selNum = 0
if (-not [int]::TryParse($selection, [ref]$selNum) -or $selNum -lt 1 -or $selNum -gt $validItems.Count) {
    Write-Host ""
    Show-Error "Invalid number"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Invalid number")
}

# Get Selected PC Name
$newHostname = $validItems[$selNum - 1].'NewPCName'

Write-Host ""

# Check Same Name
if ($currentHostname -eq $newHostname) {
    Show-Skip "Current hostname is the same. No change needed."
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "Current hostname is the same")
}

Write-Host "  $currentHostname -> $newHostname" -ForegroundColor White
Write-Host ""

# Confirm Execution
$cancelResult = Confirm-ModuleExecution -Message "Do you want to change the hostname?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""

# Change Hostname
try {
    Rename-Computer -NewName $newHostname -Force -ErrorAction Stop
    Show-Success "Hostname changed successfully: $currentHostname -> $newHostname"
}
catch {
    Show-Error "Failed to change hostname: $_"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Failed to change hostname: $_")
}

Write-Host ""

Show-Warning "Restart is required to apply the hostname change."
Write-Host ""

return (New-ModuleResult -Status "Success" -Message "Hostname changed: $currentHostname -> $newHostname (restart required)")