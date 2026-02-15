# ========================================
# Hostname Change Script
# ========================================

$HOSTLIST_CSV = Join-Path $PSScriptRoot "..\..\..\kernel\csv\hostlist.csv"

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Hostname Change" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Check if CSV file exists
if (-not (Test-Path $HOSTLIST_CSV)) {
    Write-Host "[ERROR] hostlist.csv not found: $HOSTLIST_CSV" -ForegroundColor Red
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "hostlist.csv not found")
}

# Load CSV
try {
    $hostItems = Import-Csv -Path $HOSTLIST_CSV -Encoding Default
}
catch {
    Write-Host "[ERROR] Failed to load CSV: $_" -ForegroundColor Red
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Failed to load CSV: $_")
}

# Filter items with NewPCName
$validItems = @()
foreach ($item in $hostItems) {
    if (-not [string]::IsNullOrEmpty($item.'NewPCName')) {
        $validItems += $item
    }
}

if ($validItems.Count -eq 0) {
    Write-Host "[ERROR] No NewPCName registered in hostlist.csv" -ForegroundColor Red
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
    Write-Host "[INFO] Canceled" -ForegroundColor Cyan
    Write-Host ""
    return (New-ModuleResult -Status "Cancelled" -Message "User canceled")
}

# Validate Number
$selNum = 0
if (-not [int]::TryParse($selection, [ref]$selNum) -or $selNum -lt 1 -or $selNum -gt $validItems.Count) {
    Write-Host ""
    Write-Host "[ERROR] Invalid number" -ForegroundColor Red
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Invalid number")
}

# Get Selected PC Name
$newHostname = $validItems[$selNum - 1].'NewPCName'

Write-Host ""

# Check Same Name
if ($currentHostname -eq $newHostname) {
    Write-Host "[INFO] Current hostname is the same. No change needed." -ForegroundColor Yellow
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "Current hostname is the same")
}

Write-Host "  $currentHostname -> $newHostname" -ForegroundColor White
Write-Host ""

# Confirm Execution
if (-not (Confirm-Execution -Message "Do you want to change the hostname?")) {
    Write-Host ""
    Write-Host "[INFO] Canceled" -ForegroundColor Cyan
    Write-Host ""
    return (New-ModuleResult -Status "Cancelled" -Message "User canceled")
}

Write-Host ""

# Change Hostname
try {
    Rename-Computer -NewName $newHostname -Force -ErrorAction Stop
    Write-Host "[SUCCESS] Hostname changed successfully: $currentHostname -> $newHostname" -ForegroundColor Green
}
catch {
    Write-Host "[ERROR] Failed to change hostname: $_" -ForegroundColor Red
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Failed to change hostname: $_")
}

Write-Host ""

Write-Host "[INFO] Restart is required to apply the hostname change." -ForegroundColor Yellow
Write-Host ""

return (New-ModuleResult -Status "Success" -Message "Hostname changed: $currentHostname -> $newHostname (restart required)")