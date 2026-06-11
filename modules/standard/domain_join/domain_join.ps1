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
# Idempotency Check (before the DNS probe - an already-joined machine
# must Skip even when the kitting network is currently unreachable)
# ========================================
$cs = Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue
if ($cs -and $cs.PartOfDomain) {
    if ("$($cs.Domain)" -ieq "$DOMAIN") {
        Show-Skip "Already a member of domain '$($cs.Domain)'"
        Write-Host ""
        return (New-ModuleResult -Status "Skipped" -Message "Already joined to $($cs.Domain)" -Verified $true)
    }
    # Joined to a DIFFERENT domain: a CSV/reality contradiction. Fail
    # closed with both names instead of skipping or re-joining silently.
    Show-Error "Machine is joined to a DIFFERENT domain (current: $($cs.Domain), target: $DOMAIN)"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Joined to different domain (current: $($cs.Domain), target: $DOMAIN)")
}

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

    # Step 5.5: Post-Apply Verification - the join is reflected in
    # Win32_ComputerSystem immediately (the reboot only completes it).
    $verified = $null
    $csAfter = Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue
    if ($csAfter) {
        $verified = ($csAfter.PartOfDomain -and ("$($csAfter.Domain)" -ieq "$DOMAIN"))
        if ($verified) {
            Write-Host "  [VERIFIED] Member of $($csAfter.Domain) (reboot pending)" -ForegroundColor Green
        } else {
            Write-Host "  [VERIFY FAILED] PartOfDomain=$($csAfter.PartOfDomain), Domain=$($csAfter.Domain)" -ForegroundColor Red
        }
    }
    Write-Host ""
    return (New-ModuleResult -Status "Success" -Message "Domain join completed" -Verified $verified)
}
catch {
    $errorMsg = $_.Exception.Message
    Write-Host ""
    Show-Error "Domain join failed: $errorMsg"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Domain join failed: $errorMsg")
}
