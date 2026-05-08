# ========================================
# Domain Join Script
# ========================================

Show-Info "Executing domain join process..."
Write-Host ""

# ========================================
# Load domain.csv
# ========================================
$csvPath = Join-Path $PSScriptRoot "domain.csv"

$domainList = Import-ModuleCsv -Path $csvPath -RequiredColumns @("Enabled", "domain", "user", "pass", "dns")
if ($null -eq $domainList) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load domain.csv")
}

$domainEntry = $domainList | Where-Object { $_.Enabled -eq '1' } | Select-Object -First 1
if ($null -eq $domainEntry) {
    Show-Info "No enabled entries in domain.csv"
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "No enabled entries")
}
$DOMAIN = $domainEntry.'domain'
$USER = $domainEntry.'user'
$PASS = $domainEntry.'pass'
$DNS = $domainEntry.'dns'

# ========================================
# DNS Connectivity Pre-Check (bounded, fail-fast)
# ========================================
# Single bounded probe instead of Wait-NetworkReady (which blocks
# indefinitely). On failure we return Error so the FlexProfile dashboard
# / AutoPilot ErrorMode dispatcher can decide what to do next (retry,
# skip, or operator dialog).
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host "DNS Connection Check" -ForegroundColor Cyan
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""

$dnsReachable = Test-Connection -ComputerName $DNS -Count 2 -Quiet -ErrorAction SilentlyContinue
if (-not $dnsReachable) {
    Show-Error "DNS unreachable: $DNS"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "DNS unreachable: $DNS")
}
Show-Success "DNS reachable: $DNS"
Write-Host ""

# ========================================
# Domain Join
# ========================================
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host "Domain Join Process" -ForegroundColor Cyan
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""

Write-Host "Executing domain join: $DOMAIN / $USER" -ForegroundColor Yellow
Write-Host ""

# Local Stop preference so Add-Computer's non-terminating errors
# (DNS resolution failure, auth failure, DC unreachable, etc.) are
# routed into the catch block.
$ErrorActionPreference = 'Stop'

try {
    $securePassword = ConvertTo-SecureString $PASS -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($USER, $securePassword)

    Add-Computer -DomainName $DOMAIN -Credential $credential -Force

    Write-Host ""
    Show-Success "Domain join completed"
    Write-Host ""
    return (New-ModuleResult -Status "Success" -Message "Domain join completed")
}
catch {
    $errorMsg = $_.Exception.Message
    Write-Host ""
    Show-Error "Domain join failed: $errorMsg"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Domain join failed: $errorMsg")
}
