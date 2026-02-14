# ========================================
# Domain Join Script
# ========================================

Write-Host "Executing domain join process..." -ForegroundColor Cyan
Write-Host ""

# ========================================
# Load domain.csv
# ========================================
$csvPath = Join-Path $PSScriptRoot "domain.csv"

if (-not (Test-Path $csvPath)) {
    Write-Host "[ERROR] domain.csv not found: $csvPath" -ForegroundColor Red
    return (New-ModuleResult -Status "Error" -Message "domain.csv not found")
}

try {
    $domainList = @(Import-Csv -Path $csvPath -Encoding Default)
}
catch {
    Write-Host "[ERROR] Failed to load domain.csv: $_" -ForegroundColor Red
    return (New-ModuleResult -Status "Error" -Message "Failed to load domain.csv: $_")
}

if ($domainList.Count -eq 0) {
    Write-Host "[ERROR] domain.csv contains no data" -ForegroundColor Red
    return (New-ModuleResult -Status "Error" -Message "domain.csv contains no data")
}

# Note: CSV headers must match these keys exactly
$domainEntry = $domainList[0]
$DOMAIN = $domainEntry.'domain'
$USER = $domainEntry.'user'
$PASS = $domainEntry.'pass'
$DNS = $domainEntry.'dns'

# ========================================
# Check DNS Connection
# ========================================
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host "DNS Connection Check" -ForegroundColor Cyan
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""

Write-Host "[INFO] Checking connection to DNS server ($DNS)..." -ForegroundColor Cyan

try {
    $pingResult = Test-Connection -ComputerName $DNS -Count 2 -Quiet

    if ($pingResult) {
        Write-Host "[SUCCESS] Ping to DNS server succeeded" -ForegroundColor Green
    }
    else {
        Write-Host "[ERROR] Ping to DNS server failed" -ForegroundColor Red
        Write-Host "[ERROR] Aborting domain join" -ForegroundColor Red
        return (New-ModuleResult -Status "Error" -Message "Ping to DNS server failed")
    }
}
catch {
    Write-Host "[ERROR] Error checking DNS connection: $_" -ForegroundColor Red
    Write-Host "[ERROR] Aborting domain join" -ForegroundColor Red
    return (New-ModuleResult -Status "Error" -Message "DNS connection check failed: $_")
}

Write-Host ""

# ========================================
# Domain Join Process
# ========================================
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host "Domain Join Process" -ForegroundColor Cyan
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""

Write-Host "Executing domain join: $DOMAIN / $USER" -ForegroundColor Yellow
Write-Host ""

$ErrorActionPreference = 'Stop'

try {
    # Create credentials
    $securePassword = ConvertTo-SecureString $PASS -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($USER, $securePassword)

    # Join domain
    Add-Computer -DomainName $DOMAIN -Credential $credential -Force

    Write-Host ""
    Write-Host "[SUCCESS] Domain join completed" -ForegroundColor Green
    Write-Host ""
    return (New-ModuleResult -Status "Success" -Message "Domain join completed")
}
catch {
    Write-Host ""
    Write-Host "[ERROR] Error occurred during domain join: $_" -ForegroundColor Red
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Domain join failed: $_")
}