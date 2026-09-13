# ========================================
# Domain Join Script
# ========================================

# ========================================
# Local Helpers
# ========================================
function Get-PendingComputerName {
    # hostname_config renames with Rename-Computer, which only STAGES the new
    # name: ComputerName holds the pending name, ActiveComputerName the one the
    # machine is still running under. Joining while a rename is staged creates
    # the account under the OLD name; after the reboot the machine calls itself
    # by the NEW one and can no longer authenticate, so the join has to know.
    # Returns the pending name, or $null when no rename is staged.
    try {
        $base = "HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName"
        $pending = (Get-ItemProperty -Path "$base\ComputerName" -Name "ComputerName" -ErrorAction Stop).ComputerName
        $active = (Get-ItemProperty -Path "$base\ActiveComputerName" -Name "ComputerName" -ErrorAction Stop).ComputerName
        if (-not [string]::IsNullOrWhiteSpace($pending) -and ($pending -ne $active)) {
            return "$pending"
        }
    }
    catch { }
    return $null
}

function Write-JoinDiagnostics {
    # Operator-facing diagnostics for a failed join. This deliberately does NOT
    # match on the exception text: the Win32 tail of an Add-Computer error
    # message is localized (Japanese on a Japanese Windows), so substring
    # matching would silently stop firing on the machines that matter. Instead
    # we print a fixed checklist and lift the NON-localized status code out of
    # netsetup.log.
    param(
        [string]$OuPath
    )

    Write-Host ""
    Write-Host "  Check the following:" -ForegroundColor Yellow
    if (-not [string]::IsNullOrWhiteSpace($OuPath)) {
        Write-Host "   - The OU must already exist (the join never creates it)." -ForegroundColor White
        Write-Host "   - Creating an account in an OU works either with 'Create Computer" -ForegroundColor White
        Write-Host "     Objects' delegated on that OU (unlimited), or with the plain 'Add" -ForegroundColor White
        Write-Host "     workstations to domain' right - but that one is capped by" -ForegroundColor White
        Write-Host "     ms-DS-MachineAccountQuota (10 by default) PER ACCOUNT, so a" -ForegroundColor White
        Write-Host "     long-serving join account can simply run out." -ForegroundColor White
    }
    Write-Host "   - A computer account with this name may already exist in AD." -ForegroundColor White
    if (-not [string]::IsNullOrWhiteSpace($OuPath)) {
        Write-Host "     With 'ou' set, that object must ALSO already be in exactly that" -ForegroundColor White
        Write-Host "     OU - the join refuses to re-use one that lives elsewhere and never" -ForegroundColor White
        Write-Host "     moves it (status 0x8b0)." -ForegroundColor White
    }
    Write-Host "     Re-using one has to clear TWO more gates: the KB5020276 policy (the" -ForegroundColor White
    Write-Host "     joining account must be the one that CREATED it, or the object" -ForegroundColor White
    Write-Host "     must have been created by an admin - being a Domain Admin at join" -ForegroundColor White
    Write-Host "     time is NOT an exemption) and write access on the object itself" -ForegroundColor White
    Write-Host "     (reset password / validated dNSHostName+SPN / account" -ForegroundColor White
    Write-Host "     restrictions - 'Create Computer Objects' alone is not enough)." -ForegroundColor White
    Write-Host "     Deleting the stale object first always works." -ForegroundColor White
    Write-Host "   - Verify the credential and that a DC is reachable." -ForegroundColor White

    # netsetup.log is a debug log, not a localized UI surface: the status line
    # reads the same on every language of Windows. Best-effort only - a missing,
    # locked or truncated log must never change the result the module returns.
    try {
        $logPath = Join-Path $env:SystemRoot "debug\netsetup.log"
        if (Test-Path -LiteralPath $logPath) {
            $tail = @(Get-Content -LiteralPath $logPath -Tail 200 -ErrorAction Stop)
            $statusLine = $tail | Where-Object { $_ -match 'NetpDoDomainJoin:\s*status:\s*0x[0-9a-fA-F]+' } | Select-Object -Last 1
            if ($statusLine -and $statusLine -match '(0x[0-9a-fA-F]+)') {
                $code = $Matches[1].ToLower()
                # Measured codes take priority over guessed ones: 0x2 is what a
                # missing OU path and a non-OU path both return in practice
                # (Windows Server 2025 DC, 2026-09-13), not the ERROR_DS_*
                # range one would expect.
                $hint = switch ($code) {
                    '0x2'    { "ERROR_FILE_NOT_FOUND - the OU path does not exist, or is not an OU" }
                    '0x8b0'  { "NERR_UserExists - an account of this name already exists, AND (when 'ou' is set) it is in a DIFFERENT OU" }
                    '0xaac'  { "NERR_AccountReuseBlockedByPolicy - re-use blocked by policy (KB5020276)" }
                    '0x5'    { "ERROR_ACCESS_DENIED - insufficient rights on the domain or the target OU" }
                    '0x534'  { "ERROR_NONE_MAPPED - the account or OU could not be resolved" }
                    '0x54b'  { "ERROR_NO_SUCH_DOMAIN - the domain was not found" }
                    '0x6ba'  { "RPC_S_SERVER_UNAVAILABLE - the DC could not be contacted" }
                    '0x525'  { "ERROR_NO_SUCH_USER - no DC holds this computer account yet (normal on a first join)" }
                    '0x2030' { "ERROR_DS_NO_SUCH_OBJECT - the OU path does not exist" }
                    default  { "see the log for details" }
                }
                Write-Host ""
                Write-Host "  netsetup.log status: $code ($hint)" -ForegroundColor Yellow

                # 0x8b0 with an OU requested is the one failure whose real
                # cause the Win32 message actively hides: it says only "the
                # account already exists", while the join actually refused
                # because the existing object sits somewhere else. The three
                # ways out are not obvious, so spell them out.
                if ($code -eq '0x8b0' -and -not [string]::IsNullOrWhiteSpace($OuPath)) {
                    Write-Host "    The join compares the requested OU with where the existing" -ForegroundColor White
                    Write-Host "    object actually is, and refuses when they differ. Look for" -ForegroundColor White
                    Write-Host "    'NetpGetComputerObjectDn' in the log - it prints both DNs." -ForegroundColor White
                    Write-Host "    Ways out: delete the stale object in AD, or set 'ou' to the" -ForegroundColor White
                    Write-Host "    OU the object is really in, or leave 'ou' empty (with no OU" -ForegroundColor White
                    Write-Host "    requested the object is re-used wherever it already lives)." -ForegroundColor White
                }
            }
            Write-Host "  Full log: $logPath" -ForegroundColor White
        }
    }
    catch { }

    Write-Host ""
}

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

# 'ou' is an OPTIONAL column. A domain.csv written before the column existed -
# or a profile data overlay that still ships the old header - must keep working,
# so the value is read through a property guard instead of a direct access.
$OU = ""
if ($null -ne $domainEntry.PSObject.Properties['ou'] -and -not [string]::IsNullOrWhiteSpace($domainEntry.'ou')) {
    $OU = "$($domainEntry.'ou')".Trim()
}

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
# OU Path Validation (only when an OU is requested)
# ========================================
# The OU is honoured only while the computer account is CREATED, so this runs
# AFTER the idempotency check on purpose: an already-joined machine must still
# Skip cleanly even when the CSV carries a typo in 'ou'.
if ($OU -ne "") {
    # RFC 1779 distinguished name: comma-separated <attr>=<value> components,
    # where a comma INSIDE a value is backslash-escaped (OU=Sales\,EMEA).
    $dnComponent = '[A-Za-z][A-Za-z0-9-]*=(?:[^,\\]|\\.)+'
    $dnPattern = "^$dnComponent(?:,\s*$dnComponent)*$"
    if (($OU -notmatch $dnPattern) -or ($OU -notmatch '(?i)(^|,)\s*DC=')) {
        Show-Error "Invalid OU path (a full distinguished name is required): $OU"
        Write-Host "  Expected form: OU=Kitting,OU=PC,DC=example,DC=local" -ForegroundColor Yellow
        Write-Host "  A DN contains commas - quote the whole field in domain.csv." -ForegroundColor Yellow
        Write-Host "  Leave 'ou' empty to use the domain default computer container." -ForegroundColor Yellow
        Write-Host ""
        return (New-ModuleResult -Status "Error" -Message "Invalid OU path: $OU")
    }
    if ($OU -notmatch '(?i)^\s*OU=') {
        # Measured against Windows Server 2025 / Windows 11 (2026-09-13): the
        # join API resolves the path and rejects anything that is not an
        # organizational unit, logging "Specified path '<dn>' is not an OU" and
        # failing with status 0x2. A container (CN=Computers and friends) can
        # therefore never work, so stop here instead of paying a network round
        # trip for an opaque error.
        Show-Error "OU path must point at an organizational unit (OU=...): $OU"
        Write-Host "  A container such as CN=Computers is rejected by the join API." -ForegroundColor Yellow
        Write-Host "  Leave 'ou' empty to use the domain default computer container." -ForegroundColor Yellow
        Write-Host ""
        return (New-ModuleResult -Status "Error" -Message "OU path is not an organizational unit: $OU")
    }
}

# ========================================
# Staged Rename Detection
# ========================================
# A profile that runs hostname_config and domain_join without a __RESTART__
# between them reaches this point with the rename still staged. Joining under
# the old name would leave the machine with an account name that stops matching
# after the reboot, which breaks the secure channel before the PC is even
# delivered. Joining under the NEW name is unambiguously what the hostlist
# asked for, so the option is applied automatically rather than handed to the
# operator as a decision.
$pendingName = Get-PendingComputerName
$joinName = $env:COMPUTERNAME
$joinWithNewName = $false
if ($null -ne $pendingName) {
    $joinName = $pendingName
    $joinWithNewName = $true
}

# ========================================
# Settings Preview
# ========================================
$ouLabel = if ($OU -eq "") { "(domain default container)" } else { $OU }
$nameLabel = if ($joinWithNewName) { "$env:COMPUTERNAME -> $joinName (staged rename)" } else { $joinName }
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host "Domain Join Settings" -ForegroundColor Cyan
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host "  Computer: " -NoNewline -ForegroundColor White
Write-Host $nameLabel -ForegroundColor Yellow
Write-Host "  Domain  : " -NoNewline -ForegroundColor White
Write-Host $DOMAIN -ForegroundColor Yellow
Write-Host "  Account : " -NoNewline -ForegroundColor White
Write-Host $USER -ForegroundColor Yellow
Write-Host "  DNS     : " -NoNewline -ForegroundColor White
Write-Host $DNS -ForegroundColor Yellow
Write-Host "  OU      : " -NoNewline -ForegroundColor White
Write-Host $ouLabel -ForegroundColor Yellow
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""

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

    # -OUPath is passed straight through to the AccountOU parameter of the
    # underlying JoinDomainOrWorkgroup / NetJoinDomain call. It is omitted
    # entirely when no OU is configured, which is what places the account in
    # the domain default container (an empty string is rejected as a bad DN).
    $joinParams = @{
        DomainName = $DOMAIN
        Credential = $credential
        Force      = $true
    }
    if ($OU -ne "") {
        $joinParams['OUPath'] = $OU
    }
    if ($joinWithNewName) {
        # -Options REPLACES the cmdlet default instead of adding to it, and
        # that default is AccountCreate. Passing JoinWithNewName on its own
        # would silently drop account creation and make every first-time join
        # fail against a domain where the account is not pre-staged, so the
        # two flags have to travel together.
        $joinParams['Options'] = @('AccountCreate', 'JoinWithNewName')
        Show-Warning "A rename to '$joinName' is staged - joining under the new name"
        Write-Host ""
    }

    Add-Computer @joinParams

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

    # OU placement is deliberately NOT read back. The join API already enforces
    # it, and it fails closed in every direction (measured on Windows Server
    # 2025 + Windows 11, 2026-09-13):
    #   - OU missing, or not an organizational unit -> status 0x2
    #   - an account of this name already exists in a DIFFERENT OU -> status
    #     0x8b0, rejected by NetpGetComputerObjectDn, which compares the
    #     requested DN against the existing object's DN (by length first, then
    #     by string - both paths were exercised)
    # So a successful join with 'ou' set can only mean the object is in that
    # OU. A directory query added here could return "matches" and nothing else,
    # and a check that cannot fail is not a check - it would only buy an LDAP
    # dependency. The requested OU is therefore reported as information.
    if ($OU -ne "") {
        Write-Host "  [INFO] Requested OU: $OU (enforced by the join API, not re-read here)" -ForegroundColor White
    }
    Write-Host ""
    $successMessage = if ($OU -ne "") { "Domain join completed (OU: $OU)" } else { "Domain join completed" }
    return (New-ModuleResult -Status "Success" -Message $successMessage -Verified $verified)
}
catch {
    $errorMsg = $_.Exception.Message
    Write-Host ""
    Show-Error "Domain join failed: $errorMsg"
    Write-JoinDiagnostics -OuPath $OU
    return (New-ModuleResult -Status "Error" -Message "Domain join failed: $errorMsg")
}
