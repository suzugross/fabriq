# ========================================
# Fabriq AD lab builder -- run ON the machine that becomes the lab domain controller
# ========================================
# Turns a freshly installed Windows Server into the AD lab the module test rig
# needs (domain, OU tree, service accounts, OU delegation, pre-created computer
# accounts), driven entirely by ad_lab.json.
#
# LAB ONLY. Passwords live in the JSON in clear text on purpose - this builds a
# throw-away verification domain, never anything that touches production AD.
#
# Two phases, because promoting a domain controller forces a reboot:
#   phase 1 (this run)  : network -> AD DS feature -> forest promotion -> reboot
#   phase 2 (after boot): OUs -> users -> delegation -> computers -> settings
# Phase 2 runs by itself: phase 1 registers a startup task that re-invokes this
# script. There is no state file - the phase is derived from the machine itself
# (is it a DC yet?), so every step is idempotent and re-running is always safe.
#
# Usage (from an elevated prompt, or just double-click ad_lab.bat):
#   powershell.exe -NoProfile -ExecutionPolicy Bypass -File ad_lab.ps1
#   ... -Status      report what already exists, change nothing
#   ... -Config <path>
#   ... -NoReboot    do everything except the reboot at the end of phase 1
# ========================================
param(
    [string]$Config = '',
    [switch]$Status,
    [switch]$Resume,
    [switch]$NoReboot,
    [switch]$LoadOnly
)

$ErrorActionPreference = 'Stop'

$script:WorkDir   = 'C:\ad_lab'
$script:TaskName  = 'AdLabResume'
$script:FailCount = 0

# Well-known AD schema / extended-right GUIDs. These are used instead of dsacls
# rights names on purpose: dsacls takes LOCALIZED names ("Reset Password" is
# Japanese on a Japanese Server), while the GUIDs are the same everywhere.
$script:GuidComputerClass       = [guid]'bf967a86-0de6-11d0-a285-00aa003049e2'
$script:GuidResetPassword       = [guid]'00299570-246d-11d0-a768-00aa006e0529'
$script:GuidValidatedDnsName    = [guid]'72e39547-7b18-11d1-adef-00c04fd8d5cd'
$script:GuidValidatedSpn        = [guid]'f3a64788-5306-11d1-a9c5-0000f80367c1'
$script:GuidAccountRestrictions = [guid]'4c164200-20c0-11d0-a768-00aa006e0529'

# ========================================
# Output helpers (this runs on a bare server - no fabriq kernel available)
# ========================================
function Write-Step { param([string]$Text) Write-Host ""; Write-Host "== $Text" -ForegroundColor Cyan }
function Write-Ok   { param([string]$Text) Write-Host "   [OK]   $Text" -ForegroundColor Green }
function Write-Skip { param([string]$Text) Write-Host "   [SKIP] $Text" -ForegroundColor DarkGray }
function Write-Warn { param([string]$Text) Write-Host "   [WARN] $Text" -ForegroundColor Yellow }
function Write-Fail {
    # Phase 2 keeps going after a failed item, so the count is what decides the
    # exit code at the end - a half-built lab must never look like a clean run.
    param([string]$Text)
    $script:FailCount++
    Write-Host "   [FAIL] $Text" -ForegroundColor Red
}
function Write-Note { param([string]$Text) Write-Host "          $Text" -ForegroundColor DarkGray }

function Stop-WithError {
    param([string]$Message)
    Write-Host ""
    Write-Fail $Message
    Write-Host ""
    exit 1
}

# ========================================
# Config
# ========================================
function Read-LabConfig {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) { Stop-WithError "Config not found: $Path" }
    try {
        $cfg = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json
    }
    catch {
        Stop-WithError "Config is not valid JSON: $Path ($($_.Exception.Message))"
    }

    if (-not $cfg.domain) { Stop-WithError "Config has no 'domain' section" }
    foreach ($f in @('fqdn', 'netbios', 'dsrmPassword')) {
        if ([string]::IsNullOrWhiteSpace($cfg.domain.$f)) { Stop-WithError "Config: domain.$f is required" }
    }
    if ($cfg.domain.fqdn -notmatch '^[A-Za-z0-9][A-Za-z0-9-]*(\.[A-Za-z0-9][A-Za-z0-9-]*)+$') {
        Stop-WithError "Config: domain.fqdn must be a dotted DNS name (got '$($cfg.domain.fqdn)')"
    }
    if ($cfg.domain.netbios.Length -gt 15) {
        Stop-WithError "Config: domain.netbios must be 15 characters or fewer (got '$($cfg.domain.netbios)')"
    }
    foreach ($u in @($cfg.users)) {
        if ([string]::IsNullOrWhiteSpace($u.name))     { Stop-WithError "Config: a users[] entry has no 'name'" }
        if ([string]::IsNullOrWhiteSpace($u.password)) { Stop-WithError "Config: user '$($u.name)' has no 'password'" }
        if ($u.name.Length -gt 20) { Stop-WithError "Config: user '$($u.name)' exceeds the 20 character sAMAccountName limit" }
    }
    foreach ($d in @($cfg.delegations)) {
        if ($d.right -notin @('CreateComputer', 'FullJoin')) {
            Stop-WithError "Config: delegation right must be 'CreateComputer' or 'FullJoin' (got '$($d.right)')"
        }
    }
    foreach ($n in @($cfg.settings.domainAdmins)) {
        if ([string]::IsNullOrWhiteSpace($n)) { Stop-WithError "Config: settings.domainAdmins contains an empty entry" }
    }
    return $cfg
}

function Get-DomainDn {
    param([string]$Fqdn)
    return (($Fqdn -split '\.' | ForEach-Object { "DC=$_" }) -join ',')
}

function Resolve-LabDn {
    # OU paths in the JSON may be written relative to the domain root
    # ("OU=Kitting,OU=PC") so the same file survives a domain rename. A value
    # that already carries its own DC= components is taken as-is.
    param([string]$Path, [string]$DomainDn)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $DomainDn }
    if ($Path -match '(?i)(^|,)\s*DC=') { return $Path }
    return "$Path,$DomainDn"
}

# ========================================
# Environment probes
# ========================================
function Test-IsAdmin {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    return (New-Object Security.Principal.WindowsPrincipal($id)).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Test-IsServerSku {
    # ProductType: 1 = workstation, 2 = domain controller, 3 = server
    return ((Get-CimInstance Win32_OperatingSystem).ProductType -ne 1)
}

function Test-IsDomainController {
    # DomainRole: 4 = backup DC, 5 = primary DC
    return ((Get-CimInstance Win32_ComputerSystem).DomainRole -in @(4, 5))
}

function Split-LabOuDn {
    # Splits "OU=Kitting,OU=PC,DC=lab,DC=local" into the leaf name New-ADOrganizationalUnit
    # wants ("Kitting") and the parent DN it goes under. The leaf RDN may carry a
    # backslash-escaped comma (OU=Sales\,EMEA -> name "Sales,EMEA"), which is why
    # this is not a plain -split ','.
    param([string]$Dn)
    if ($Dn -notmatch '^(?<rdn>(?:[^,\\]|\\.)+),(?<parent>.+)$') { return $null }
    # Capture both groups BEFORE any further -match: the next match overwrites $Matches.
    $rdn    = $Matches['rdn']
    $parent = $Matches['parent']
    if ($rdn -notmatch '(?i)^OU=') { return $null }
    return [pscustomobject]@{
        Name   = ($rdn -replace '(?i)^OU=', '') -replace '\\(.)', '$1'
        Parent = $parent
    }
}

function Test-LabObject {
    # -Identity is exact and needs no filter parsing. Note that the AD module's
    # -Filter scriptblock form cannot resolve member access ($u.name), so every
    # lookup here either uses -Identity or a string filter built from a local.
    param([string]$Dn)
    # SilentlyContinue as well as the catch: a "not found" here is an expected
    # answer, and letting it write an error record just fills the transcript
    # with noise that reads like a failure.
    try { return [bool](Get-ADObject -Identity $Dn -ErrorAction SilentlyContinue) }
    catch { return $false }
}

function Get-PendingRebootReason {
    # Install-ADDSForest runs its own prerequisite check, and that check REFUSES
    # to promote while a reboot is pending ("role change is in progress or this
    # computer needs to be restarted"). Installing the AD DS role sets exactly
    # that flag, and so does a Windows Update run, so the promotion has to be
    # done on the far side of a reboot.
    $reasons = @()
    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') {
        $reasons += 'Component Based Servicing (a role or feature was installed)'
    }
    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\PackagesPending') {
        $reasons += 'Servicing packages pending'
    }
    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired') {
        $reasons += 'Windows Update'
    }
    $sm = Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name 'PendingFileRenameOperations' -ErrorAction SilentlyContinue
    if ($sm -and $sm.PendingFileRenameOperations) { $reasons += 'Pending file rename operations' }

    $active = (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ActiveComputerName' -Name 'ComputerName' -ErrorAction SilentlyContinue).ComputerName
    $next   = (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName' -Name 'ComputerName' -ErrorAction SilentlyContinue).ComputerName
    if ($active -and $next -and ($active -ne $next)) { $reasons += "Computer rename pending ($active -> $next)" }

    return $reasons
}

function Get-RebootLoopCount {
    # Phase detection stays stateless; this counter exists only so a reboot flag
    # that never clears cannot bounce the machine forever.
    $p = Join-Path $script:WorkDir 'reboot_count.txt'
    if (Test-Path -LiteralPath $p) { return [int](Get-Content -LiteralPath $p -Raw).Trim() }
    return 0
}

function Set-RebootLoopCount {
    param([int]$Value)
    Set-Content -LiteralPath (Join-Path $script:WorkDir 'reboot_count.txt') -Value $Value -Encoding Ascii
}

function Wait-LabDirectory {
    param([int]$TimeoutSeconds = 900)

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)

    # Probe with plain LDAP FIRST. Importing the ActiveDirectory module while
    # ADWS is still starting leaves the module loaded but WITHOUT its AD: drive
    # ("could not find a default server ... Active Directory Web Services"), and
    # because a later Import-Module is a no-op for an already-loaded module the
    # drive never appears - which is exactly what broke the delegation step.
    Write-Step "Waiting for the directory to answer (LDAP)"
    $up = $false
    while ((Get-Date) -lt $deadline) {
        try {
            $rootDse = New-Object DirectoryServices.DirectoryEntry("LDAP://RootDSE")
            if ($rootDse.Properties['defaultNamingContext'].Value) { $up = $true; break }
        }
        catch { }
        Start-Sleep -Seconds 10
    }
    if (-not $up) { Stop-WithError "LDAP did not answer within $TimeoutSeconds seconds" }
    Write-Ok "LDAP is up"

    Write-Step "Waiting for Active Directory Web Services"
    while ((Get-Date) -lt $deadline) {
        try {
            Import-Module ActiveDirectory -Force -ErrorAction Stop
            $d = Get-ADDomain -ErrorAction Stop
            Write-Ok "Directory is up: $($d.DNSRoot)"
            return $d
        }
        catch {
            Start-Sleep -Seconds 10
        }
    }
    Stop-WithError "Active Directory did not answer within $TimeoutSeconds seconds"
}

# ========================================
# Phase 1 steps
# ========================================
function Set-LabNetwork {
    param($Network)

    if (-not $Network -or [string]::IsNullOrWhiteSpace($Network.ip)) {
        Write-Skip "No network.ip configured - leaving networking untouched"
        return
    }

    $adapters = @(Get-NetAdapter -Physical | Where-Object { $_.Status -eq 'Up' })
    if ($adapters.Count -ne 1) {
        Write-Warn "Expected exactly one connected adapter, found $($adapters.Count) - skipping network setup"
        Write-Note "Configure the address by hand, then re-run this script."
        return
    }
    $adapter = $adapters[0]

    $current = @(Get-NetIPAddress -InterfaceIndex $adapter.ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue)

    # "The address is already there" is NOT enough to skip: a DC has to own its
    # address statically and with the configured mask. A machine that is sitting
    # on the target address via DHCP, or with a different prefix length, gets
    # converted rather than left alone.
    $good = @($current | Where-Object {
        $_.IPAddress -eq $Network.ip -and
        $_.PrefixLength -eq [int]$Network.prefixLength -and
        $_.PrefixOrigin -eq 'Manual'
    })
    if ($good.Count -gt 0) {
        Write-Skip "Address $($Network.ip)/$($Network.prefixLength) is already set statically on '$($adapter.Name)'"
    }
    else {
        $held = @($current | Where-Object { $_.IPAddress -eq $Network.ip })
        if ($held.Count -gt 0) {
            Write-Warn "Address $($Network.ip) is present as $($held[0].PrefixOrigin) /$($held[0].PrefixLength) - converting it to a static /$($Network.prefixLength)"
        }
        Write-Note "Reconfiguring '$($adapter.Name)' now - a remote session over this NIC would drop here."
        Get-NetIPAddress -InterfaceIndex $adapter.ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue |
            Remove-NetIPAddress -Confirm:$false -ErrorAction SilentlyContinue
        Remove-NetRoute -InterfaceIndex $adapter.ifIndex -DestinationPrefix '0.0.0.0/0' -Confirm:$false -ErrorAction SilentlyContinue
        # Without this a DHCP lease can simply reclaim the interface later.
        Set-NetIPInterface -InterfaceIndex $adapter.ifIndex -AddressFamily IPv4 -Dhcp Disabled -ErrorAction SilentlyContinue
        $params = @{
            InterfaceIndex = $adapter.ifIndex
            IPAddress      = $Network.ip
            PrefixLength   = [int]$Network.prefixLength
            AddressFamily  = 'IPv4'
        }
        if (-not [string]::IsNullOrWhiteSpace($Network.gateway)) { $params['DefaultGateway'] = $Network.gateway }
        New-NetIPAddress @params | Out-Null
        Write-Ok "Address set: $($Network.ip)/$($Network.prefixLength) on '$($adapter.Name)'"
    }

    # A DC must resolve itself: point the client resolver at its own address.
    Set-DnsClientServerAddress -InterfaceIndex $adapter.ifIndex -ServerAddresses $Network.ip
    Write-Ok "DNS client points at itself ($($Network.ip))"
}

function Test-LabBuiltinAdmin {
    # Promoting the first DC of a forest turns the standalone server's SAM into
    # the directory: the built-in Administrator (SID *-500) becomes the DOMAIN
    # Administrator, and the local SAM stops being usable for logon (DSRM only).
    # Two ways that bites:
    #   - the promotion prerequisite check refuses an Administrator whose
    #     password is blank / not required
    #   - other local accounts come across as ORDINARY domain users, and an
    #     ordinary user cannot log on to a domain controller - so an operator
    #     who only knows a non-builtin local account can lock themselves out
    param([string]$Netbios, $DomainAdmins)

    Write-Step "Built-in Administrator check"
    $admin = $null
    try { $admin = Get-LocalUser | Where-Object { $_.SID.Value -match '-500$' } | Select-Object -First 1 } catch { }
    if (-not $admin) {
        Write-Warn "Could not read the local Administrator account - skipping this check"
        return
    }

    if (-not $admin.PasswordRequired) {
        Stop-WithError ("The built-in Administrator '{0}' is set to 'password not required', which the forest promotion refuses. Fix it first:  net user {0} <password> /passwordreq:yes" -f $admin.Name)
    }
    if (-not $admin.Enabled) {
        # Only fatal when nothing else is lined up to get back in: a configured
        # settings.domainAdmins entry is a deliberate alternative.
        if (@($DomainAdmins).Count -gt 0) {
            Write-Warn "The built-in Administrator '$($admin.Name)' is disabled - logon after the promotion depends on settings.domainAdmins ($((@($DomainAdmins)) -join ', '))"
            Write-Note "If those accounts do not come across, DSRM is the only way back in."
        }
        else {
            Stop-WithError ("The built-in Administrator '{0}' is disabled and settings.domainAdmins is empty, so nothing would be able to log on to the DC. Enable it (and know its password):  net user {0} <password> /active:yes" -f $admin.Name)
        }
    }
    else {
        Write-Ok "Built-in Administrator '$($admin.Name)' is enabled and requires a password"
        Write-Note "After the promotion, log on as $Netbios\$($admin.Name) using its CURRENT password."
    }

    $me = $env:USERNAME
    if ($me -and ($me -ne $admin.Name)) {
        Write-Warn "You are running as the local account '$me'."
        Write-Note "Local accounts migrate into the directory as ORDINARY users, and an ordinary"
        Write-Note "user cannot log on to a domain controller. To keep using '$me' after the"
        Write-Note "promotion, put it in settings.domainAdmins in the JSON (phase 2 adds it to"
        Write-Note "Domain Admins). Otherwise use $Netbios\$($admin.Name)."
    }
}

function Install-LabAdFeature {
    Write-Step "Installing the AD DS role"
    $feature = Get-WindowsFeature -Name AD-Domain-Services
    if ($feature.Installed) {
        Write-Skip "AD-Domain-Services is already installed"
        return
    }
    $r = Install-WindowsFeature -Name AD-Domain-Services -IncludeManagementTools
    if (-not $r.Success) { Stop-WithError "Install-WindowsFeature failed: $($r.ExitCode)" }
    Write-Ok "AD-Domain-Services installed"
    if ($r.RestartNeeded -ne 'No') { Write-Note "The role install wants a restart before the promotion." }
}

function Invoke-LabPromotion {
    param($Cfg)

    Write-Step "Promoting to a new forest: $($Cfg.domain.fqdn)"
    Write-Note "This takes several minutes and ends with a reboot."
    Import-Module ADDSDeployment -ErrorAction Stop
    $dsrm = ConvertTo-SecureString $Cfg.domain.dsrmPassword -AsPlainText -Force
    $promo = @{
        DomainName                    = $Cfg.domain.fqdn
        DomainNetbiosName             = $Cfg.domain.netbios
        SafeModeAdministratorPassword = $dsrm
        InstallDns                    = $true
        CreateDnsDelegation           = $false
        NoRebootOnCompletion          = $true
        Force                         = $true
    }
    if (-not [string]::IsNullOrWhiteSpace($Cfg.domain.forestMode)) { $promo['ForestMode'] = $Cfg.domain.forestMode }
    if (-not [string]::IsNullOrWhiteSpace($Cfg.domain.domainMode)) { $promo['DomainMode'] = $Cfg.domain.domainMode }

    $result = Install-ADDSForest @promo
    if ($result.Status -eq 'Error') { Stop-WithError "Install-ADDSForest reported an error - see the log above" }
    Write-Ok "Promotion staged (status: $($result.Status))"
}

function Register-ResumeTask {
    param([string]$ScriptPath, [string]$ConfigPath)

    $arg = '-NoProfile -ExecutionPolicy Bypass -File "{0}" -Config "{1}" -Resume' -f $ScriptPath, $ConfigPath
    $action    = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $arg
    $trigger   = New-ScheduledTaskTrigger -AtStartup
    $principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
    $settings  = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit ([TimeSpan]::FromHours(2))

    Unregister-ScheduledTask -TaskName $script:TaskName -Confirm:$false -ErrorAction SilentlyContinue
    Register-ScheduledTask -TaskName $script:TaskName -Action $action -Trigger $trigger `
        -Principal $principal -Settings $settings -Description 'Fabriq AD lab - phase 2 after the promotion reboot' | Out-Null
    Write-Ok "Startup task '$($script:TaskName)' registered (phase 2 runs by itself after the reboot)"
}

function Unregister-ResumeTask {
    $t = Get-ScheduledTask -TaskName $script:TaskName -ErrorAction SilentlyContinue
    if ($t) {
        Unregister-ScheduledTask -TaskName $script:TaskName -Confirm:$false
        Write-Ok "Startup task '$($script:TaskName)' removed"
    }
}

# ========================================
# Phase 2 steps
# ========================================
function New-LabOu {
    param([string[]]$Paths, [string]$DomainDn)

    Write-Step "Organizational units"
    if (@($Paths).Count -eq 0) { Write-Skip "None configured"; return }

    # Parents first: a shallower DN has fewer components.
    $ordered = @($Paths | Sort-Object { ($_ -split ',').Count }
    )
    foreach ($p in $ordered) {
        $dn = Resolve-LabDn -Path $p -DomainDn $DomainDn
        try {
            if (Test-LabObject -Dn $dn) {
                Write-Skip "OU exists: $dn"
                continue
            }
            $split = Split-LabOuDn -Dn $dn
            if (-not $split) {
                Write-Fail "Cannot parse OU DN: $dn"
                continue
            }
            New-ADOrganizationalUnit -Name $split.Name -Path $split.Parent -ProtectedFromAccidentalDeletion $false
            Write-Ok "OU created: $dn"
        }
        catch {
            Write-Fail "Could not create $dn : $($_.Exception.Message)"
        }
    }
}

function New-LabUser {
    param($Users, [string]$Fqdn)

    Write-Step "Service accounts"
    if (@($Users).Count -eq 0) { Write-Skip "None configured"; return }

    foreach ($u in @($Users)) {
        $sam = $u.name
        try {
            $existing = Get-ADUser -Filter "SamAccountName -eq '$sam'" -ErrorAction SilentlyContinue
            if ($existing) {
                Write-Skip "User exists: $($u.name)"
            }
            else {
                New-ADUser -Name $u.name -SamAccountName $u.name `
                    -UserPrincipalName ("{0}@{1}" -f $u.name, $Fqdn) `
                    -AccountPassword (ConvertTo-SecureString $u.password -AsPlainText -Force) `
                    -Description ("$($u.description)") `
                    -Enabled $true -PasswordNeverExpires $true -CannotChangePassword $true
                Write-Ok "User created: $($u.name)"
            }

            foreach ($g in @($u.groups)) {
                if ([string]::IsNullOrWhiteSpace($g)) { continue }
                $members = @(Get-ADGroupMember -Identity $g -ErrorAction SilentlyContinue | Select-Object -ExpandProperty SamAccountName)
                if ($members -contains $u.name) {
                    Write-Skip "$($u.name) is already in '$g'"
                }
                else {
                    Add-ADGroupMember -Identity $g -Members $u.name
                    Write-Ok "$($u.name) added to '$g'"
                }
            }
        }
        catch {
            Write-Fail "User '$sam' failed: $($_.Exception.Message)"
        }
    }
}

function Add-LabDomainAdmin {
    # Accounts that must stay usable at the DC console after the promotion.
    # Domain Admins is addressed by its well-known RID (-512) rather than by
    # name, for the same reason the delegation uses schema GUIDs: names are
    # localized, RIDs are not.
    param($Names, $Domain)

    Write-Step "Domain Admins membership"
    if (@($Names).Count -eq 0) { Write-Skip "None configured"; return }

    $group = Get-ADGroup -Identity ("{0}-512" -f $Domain.DomainSID.Value)
    $members = @(Get-ADGroupMember -Identity $group -ErrorAction SilentlyContinue | Select-Object -ExpandProperty SamAccountName)

    foreach ($n in @($Names)) {
        if ([string]::IsNullOrWhiteSpace($n)) { continue }
        try {
            if ($members -contains $n) {
                Write-Skip "$n is already in $($group.Name)"
                continue
            }
            $u = Get-ADUser -Filter "SamAccountName -eq '$n'" -ErrorAction SilentlyContinue
            if (-not $u) {
                Write-Warn "$n was not found in the directory - skipped (was it a local account that did not migrate?)"
                continue
            }
            Add-ADGroupMember -Identity $group -Members $u
            Write-Ok "$n added to $($group.Name)"
        }
        catch {
            Write-Fail "Could not add '$n' to Domain Admins: $($_.Exception.Message)"
        }
    }
}

function Get-LabDelegationAce {
    # Returns the ACEs that make up one delegation level.
    #   CreateComputer : just enough to CREATE computer accounts in the OU
    #   FullJoin       : create + the writes a join needs to RE-USE an existing
    #                    account (reset password, validated dNSHostName / SPN,
    #                    account restrictions) - the classic "join delegation".
    param([Security.Principal.SecurityIdentifier]$Sid, [string]$Right)

    $aces = @(
        New-Object DirectoryServices.ActiveDirectoryAccessRule(
            $Sid, 'CreateChild', 'Allow', $script:GuidComputerClass, 'All')
    )
    if ($Right -eq 'FullJoin') {
        $aces += New-Object DirectoryServices.ActiveDirectoryAccessRule(
            $Sid, 'DeleteChild', 'Allow', $script:GuidComputerClass, 'All')
        $aces += New-Object DirectoryServices.ActiveDirectoryAccessRule(
            $Sid, 'ExtendedRight', 'Allow', $script:GuidResetPassword, 'Descendents', $script:GuidComputerClass)
        $aces += New-Object DirectoryServices.ActiveDirectoryAccessRule(
            $Sid, 'Self', 'Allow', $script:GuidValidatedDnsName, 'Descendents', $script:GuidComputerClass)
        $aces += New-Object DirectoryServices.ActiveDirectoryAccessRule(
            $Sid, 'Self', 'Allow', $script:GuidValidatedSpn, 'Descendents', $script:GuidComputerClass)
        $aces += New-Object DirectoryServices.ActiveDirectoryAccessRule(
            $Sid, 'WriteProperty', 'Allow', $script:GuidAccountRestrictions, 'Descendents', $script:GuidComputerClass)
    }
    return $aces
}

function Set-LabDelegation {
    param($Delegations, [string]$DomainDn, [string]$Netbios)

    Write-Step "OU delegation"
    if (@($Delegations).Count -eq 0) { Write-Skip "None configured"; return }

    foreach ($d in @($Delegations)) {
        $dn = Resolve-LabDn -Path $d.ou -DomainDn $DomainDn
        # Per-item try/catch: one bad delegation must not abandon the rest of
        # the lab (it used to abort the whole phase and leave it half-built).
        try {
            if (-not (Test-LabObject -Dn $dn)) {
                Write-Fail "Delegation target does not exist: $dn"
                continue
            }
            $sid = $null
            try { $sid = (New-Object Security.Principal.NTAccount($Netbios, $d.user)).Translate([Security.Principal.SecurityIdentifier]) }
            catch {
                Write-Fail "Cannot resolve '$Netbios\$($d.user)' - is the user configured above it?"
                continue
            }

            # Plain ADSI rather than Get-Acl/Set-Acl on the AD: drive: that drive
            # is absent whenever the ActiveDirectory module was first imported
            # before ADWS was ready, and its path syntax also mangles a DN with
            # an escaped comma (OU=Sales\,EMEA).
            $entry = New-Object DirectoryServices.DirectoryEntry("LDAP://$dn")
            $sec = $entry.ObjectSecurity
            # Compare identities as SIDs on BOTH sides: existing rules come back
            # as NTAccount by default, so a SID-vs-name comparison never matched
            # and every re-run appended duplicate ACEs.
            $existing = @($sec.GetAccessRules($true, $false, [Security.Principal.SecurityIdentifier]))

            $added = 0
            foreach ($ace in (Get-LabDelegationAce -Sid $sid -Right $d.right)) {
                # Rights are compared as FLAGS, not for equality: AD merges ACEs
                # that differ only in rights (CreateChild + DeleteChild on the
                # same object type come back as one "CreateChild, DeleteChild"
                # rule), so an equality test never matches on a re-run and every
                # run would append the ACE again.
                $dup = @($existing | Where-Object {
                    $_.IdentityReference.Value -eq $sid.Value -and
                    (($_.ActiveDirectoryRights -band $ace.ActiveDirectoryRights) -eq $ace.ActiveDirectoryRights) -and
                    $_.ObjectType -eq $ace.ObjectType -and
                    $_.InheritedObjectType -eq $ace.InheritedObjectType
                })
                if ($dup.Count -gt 0) { continue }
                $sec.AddAccessRule($ace)
                $added++
            }
            if ($added -eq 0) {
                Write-Skip "$($d.right) for $($d.user) already delegated on $dn"
            }
            else {
                $entry.ObjectSecurity = $sec
                $entry.CommitChanges()
                Write-Ok "$($d.right) delegated to $($d.user) on $dn ($added ACE)"
            }
        }
        catch {
            Write-Fail "Delegation for $($d.user) on $dn failed: $($_.Exception.Message)"
        }
    }
}

function New-LabComputer {
    param($Computers, [string]$DomainDn, [string]$Netbios, $Users)

    Write-Step "Pre-created computer accounts"
    if (@($Computers).Count -eq 0) { Write-Skip "None configured"; return }

    foreach ($c in @($Computers)) {
        $dn = Resolve-LabDn -Path $c.ou -DomainDn $DomainDn
        $cn = $c.name
        try {
            $existing = Get-ADComputer -Filter "Name -eq '$cn'" -ErrorAction SilentlyContinue
            if ($existing) {
                Write-Skip "Computer exists: $($existing.DistinguishedName)"
                continue
            }

            $params = @{ Name = $c.name; Path = $dn }
            if (-not [string]::IsNullOrWhiteSpace($c.createdBy)) {
                # The OWNER of the object is what the KB5020276 re-use check looks
                # at, so who creates it is the whole point of this entry.
                $owner = @($Users | Where-Object { $_.name -eq $c.createdBy })[0]
                if (-not $owner) {
                    Write-Fail "createdBy '$($c.createdBy)' is not one of the configured users"
                    continue
                }
                $params['Credential'] = New-Object pscredential(
                    ("{0}\{1}" -f $Netbios, $owner.name),
                    (ConvertTo-SecureString $owner.password -AsPlainText -Force))
            }
            New-ADComputer @params
            $note = if ($c.createdBy) { " (created by $($c.createdBy))" } else { "" }
            Write-Ok "Computer created: $($c.name) in $dn$note"
        }
        catch {
            Write-Fail "Could not create $cn in $dn : $($_.Exception.Message)"
            if ($c.createdBy) { Write-Note "createdBy '$($c.createdBy)' needs create rights there - check the delegation above." }
        }
    }
}

function Set-LabSettings {
    param($Settings, $Network, [string]$DomainDn)

    Write-Step "Domain settings"
    $touched = $false

    if ($Settings -and $null -ne $Settings.machineAccountQuota) {
        $current = (Get-ADObject -Identity $DomainDn -Properties 'ms-DS-MachineAccountQuota').'ms-DS-MachineAccountQuota'
        if ("$current" -eq "$($Settings.machineAccountQuota)") {
            Write-Skip "ms-DS-MachineAccountQuota is already $current"
        }
        else {
            Set-ADObject -Identity $DomainDn -Replace @{ 'ms-DS-MachineAccountQuota' = [int]$Settings.machineAccountQuota }
            Write-Ok "ms-DS-MachineAccountQuota: $current -> $($Settings.machineAccountQuota)"
        }
        $touched = $true
    }

    if ($Network -and -not [string]::IsNullOrWhiteSpace($Network.dnsForwarder)) {
        $fwd = @(Get-DnsServerForwarder -ErrorAction SilentlyContinue | Select-Object -ExpandProperty IPAddress | ForEach-Object { $_.IPAddressToString })
        if ($fwd -contains $Network.dnsForwarder) {
            Write-Skip "DNS forwarder $($Network.dnsForwarder) is already set"
        }
        else {
            Add-DnsServerForwarder -IPAddress $Network.dnsForwarder -PassThru | Out-Null
            Write-Ok "DNS forwarder added: $($Network.dnsForwarder)"
        }
        $touched = $true
    }

    if (-not $touched) { Write-Skip "None configured" }
}

# ========================================
# Status report
# ========================================
function Show-LabStatus {
    param($Cfg)

    Write-Host ""
    Write-Host "========================================" -ForegroundColor White
    Write-Host " AD lab status" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor White

    if (-not (Test-IsDomainController)) {
        Write-Host "  This machine is NOT a domain controller yet." -ForegroundColor Yellow
        Write-Host "  Server SKU : $(Test-IsServerSku)"
        Write-Host "  Run this script without -Status to build the lab."
        Write-Host ""
        return
    }

    Import-Module ActiveDirectory -ErrorAction Stop
    $domain = Get-ADDomain
    $domainDn = $domain.DistinguishedName
    Write-Host "  Domain     : $($domain.DNSRoot) ($($domain.NetBIOSName))"
    Write-Host "  DC         : $env:COMPUTERNAME"
    Write-Host "  Domain DN  : $domainDn"
    $quota = (Get-ADObject -Identity $domainDn -Properties 'ms-DS-MachineAccountQuota').'ms-DS-MachineAccountQuota'
    Write-Host "  Quota      : ms-DS-MachineAccountQuota = $quota"
    $da = @(Get-ADGroupMember -Identity ("{0}-512" -f $domain.DomainSID.Value) -ErrorAction SilentlyContinue |
            Select-Object -ExpandProperty SamAccountName)
    Write-Host "  Domain Admins : $($da -join ', ')"
    Write-Host "  Log on as     : $($domain.NetBIOSName)\<any Domain Admins member> - a plain domain user cannot log on to a DC"
    Write-Host ""

    Write-Host "  OUs" -ForegroundColor Cyan
    foreach ($p in @($Cfg.ous)) {
        $dn = Resolve-LabDn -Path $p -DomainDn $domainDn
        $ok = Test-LabObject -Dn $dn
        Write-Host ("    [{0}] {1}" -f $(if ($ok) { 'x' } else { ' ' }), $dn) -ForegroundColor $(if ($ok) { 'Green' } else { 'Red' })
    }

    Write-Host "  Users" -ForegroundColor Cyan
    foreach ($u in @($Cfg.users)) {
        $sam = $u.name
        $ok = [bool](Get-ADUser -Filter "SamAccountName -eq '$sam'" -ErrorAction SilentlyContinue)
        Write-Host ("    [{0}] {1}" -f $(if ($ok) { 'x' } else { ' ' }), $u.name) -ForegroundColor $(if ($ok) { 'Green' } else { 'Red' })
    }

    Write-Host "  Delegations" -ForegroundColor Cyan
    foreach ($d in @($Cfg.delegations)) {
        $dn = Resolve-LabDn -Path $d.ou -DomainDn $domainDn
        $ok = $false
        try {
            $sid = (New-Object Security.Principal.NTAccount($domain.NetBIOSName, $d.user)).Translate([Security.Principal.SecurityIdentifier])
            $entry = New-Object DirectoryServices.DirectoryEntry("LDAP://$dn")
            $ok = @($entry.ObjectSecurity.GetAccessRules($true, $false, [Security.Principal.SecurityIdentifier]) | Where-Object {
                $_.IdentityReference.Value -eq $sid.Value -and
                $_.ActiveDirectoryRights -match 'CreateChild' -and
                $_.ObjectType -eq $script:GuidComputerClass
            }).Count -gt 0
        }
        catch { }
        Write-Host ("    [{0}] {1} -> {2} on {3}" -f $(if ($ok) { 'x' } else { ' ' }), $d.user, $d.right, $dn) -ForegroundColor $(if ($ok) { 'Green' } else { 'Red' })
    }

    Write-Host "  Computers" -ForegroundColor Cyan
    foreach ($c in @($Cfg.computers)) {
        $cn = $c.name
        $obj = Get-ADComputer -Filter "Name -eq '$cn'" -ErrorAction SilentlyContinue
        $ok = [bool]$obj
        $where = if ($obj) { $obj.DistinguishedName } else { "$($c.name) (missing)" }
        Write-Host ("    [{0}] {1}" -f $(if ($ok) { 'x' } else { ' ' }), $where) -ForegroundColor $(if ($ok) { 'Green' } else { 'Red' })
    }
    Write-Host ""
}

# ========================================
# Main
# ========================================
# -LoadOnly stops here so the pure-logic helpers above can be dot-sourced and
# tested off a domain controller (see ad_lab_logic.tests.ps1).
if ($LoadOnly) { return }

if ([string]::IsNullOrWhiteSpace($Config)) { $Config = Join-Path $PSScriptRoot 'ad_lab.json' }
$Config = (Resolve-Path -LiteralPath $Config -ErrorAction SilentlyContinue).Path
if (-not $Config) { Stop-WithError "Config not found. Pass -Config <path> or put ad_lab.json next to this script." }
$cfg = Read-LabConfig -Path $Config

# -Status only reads, so it runs unelevated too (handy for a quick check from a
# normal shell). Everything that changes the machine needs administrator.
if ($Status) {
    Show-LabStatus -Cfg $cfg
    exit 0
}

if (-not (Test-IsAdmin)) { Stop-WithError "Administrator privileges are required (right-click ad_lab.bat -> Run as administrator)" }

if (-not (Test-Path -LiteralPath $script:WorkDir)) { New-Item -ItemType Directory -Path $script:WorkDir | Out-Null }
try { Start-Transcript -Path (Join-Path $script:WorkDir 'ad_lab.log') -Append | Out-Null } catch { }

Write-Host ""
Write-Host "========================================" -ForegroundColor White
Write-Host " Fabriq AD lab builder" -ForegroundColor Cyan
Write-Host " LAB ONLY - never point this at production AD" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor White
Write-Host "  Config : $Config"
Write-Host "  Domain : $($cfg.domain.fqdn) ($($cfg.domain.netbios))"
Write-Host "  Mode   : $(if (Test-IsDomainController) { 'phase 2 (configure)' } else { 'phase 1 (promote)' })$(if ($Resume) { ' [auto-resume]' } else { '' })"

try {
    if (-not (Test-IsDomainController)) {
        # ---------- Phase 1 ----------
        if (-not (Test-IsServerSku)) {
            Stop-WithError "This is a client SKU. Run the lab builder on Windows Server."
        }

        # Run the resume from a fixed local path so it survives the script being
        # launched from a share, a mounted ISO or a folder that later moves.
        $stagedScript = Join-Path $script:WorkDir 'ad_lab.ps1'
        $stagedConfig = Join-Path $script:WorkDir 'ad_lab.json'
        Copy-Item -LiteralPath $PSCommandPath -Destination $stagedScript -Force
        Copy-Item -LiteralPath $Config -Destination $stagedConfig -Force

        Test-LabBuiltinAdmin -Netbios $cfg.domain.netbios -DomainAdmins $cfg.settings.domainAdmins
        Set-LabNetwork -Network $cfg.network
        Install-LabAdFeature

        # The promotion prerequisite check refuses to run while a reboot is
        # pending, so clear that first and come back. Re-entering phase 1 after
        # the reboot is free: the network and the role both report SKIP.
        $pending = @(Get-PendingRebootReason)
        if ($pending.Count -gt 0) {
            $loops = Get-RebootLoopCount
            Write-Step "A restart is required before the promotion"
            foreach ($p in $pending) { Write-Note "- $p" }
            if ($loops -ge 3) {
                Stop-WithError "Still pending after $loops restarts - clear it by hand (check Windows Update), then run this again."
            }
            Set-RebootLoopCount -Value ($loops + 1)
            Register-ResumeTask -ScriptPath $stagedScript -ConfigPath $stagedConfig

            Write-Host ""
            Write-Host "Restarting now. The promotion continues by itself after the reboot." -ForegroundColor Cyan
            Write-Host ""
            try { Stop-Transcript | Out-Null } catch { }
            if ($NoReboot) {
                Write-Warn "-NoReboot given: restart by hand, then run this again."
                exit 0
            }
            Restart-Computer -Force
            exit 0
        }

        Invoke-LabPromotion -Cfg $cfg
        Set-RebootLoopCount -Value 0
        Register-ResumeTask -ScriptPath $stagedScript -ConfigPath $stagedConfig

        Write-Host ""
        Write-Host "Phase 1 done. Rebooting to finish the promotion." -ForegroundColor Cyan
        Write-Host "After the reboot, phase 2 runs by itself (log: $(Join-Path $script:WorkDir 'ad_lab.log'))." -ForegroundColor Cyan
        Write-Host ""
        try { Stop-Transcript | Out-Null } catch { }

        if ($NoReboot) {
            Write-Warn "-NoReboot given: reboot by hand to continue."
            exit 0
        }
        Restart-Computer -Force
        exit 0
    }

    # ---------- Phase 2 ----------
    $domain = Wait-LabDirectory
    if ($domain.DNSRoot -ine $cfg.domain.fqdn) {
        Stop-WithError "This DC serves '$($domain.DNSRoot)' but the config says '$($cfg.domain.fqdn)'"
    }
    $domainDn = $domain.DistinguishedName

    New-LabOu        -Paths $cfg.ous -DomainDn $domainDn
    New-LabUser      -Users $cfg.users -Fqdn $cfg.domain.fqdn
    Add-LabDomainAdmin -Names $cfg.settings.domainAdmins -Domain $domain
    Set-LabDelegation -Delegations $cfg.delegations -DomainDn $domainDn -Netbios $domain.NetBIOSName
    New-LabComputer  -Computers $cfg.computers -DomainDn $domainDn -Netbios $domain.NetBIOSName -Users $cfg.users
    Set-LabSettings  -Settings $cfg.settings -Network $cfg.network -DomainDn $domainDn

    if ($script:FailCount -eq 0) { Unregister-ResumeTask }
    Show-LabStatus -Cfg $cfg

    if ($script:FailCount -gt 0) {
        Write-Host "   [FAIL] $($script:FailCount) item(s) failed - the lab is INCOMPLETE (see the [FAIL] lines above)." -ForegroundColor Red
        Write-Host "Fix the cause and run this again: every completed item is skipped." -ForegroundColor Yellow
        Write-Host ""
        try { Stop-Transcript | Out-Null } catch { }
        exit 1
    }

    Write-Host "Lab is ready. Take a VM snapshot now - that snapshot is the reset button." -ForegroundColor Cyan
    Write-Host ""
    try { Stop-Transcript | Out-Null } catch { }
    exit 0
}
catch {
    Write-Host ""
    Write-Fail "$($_.Exception.Message)"
    # No string matching on the message: it is localized on a non-English
    # Server, so the hints are printed unconditionally instead.
    Write-Host ""
    Write-Host "  NOTHING WILL RESUME BY ITSELF after this failure." -ForegroundColor Yellow
    Write-Host "  The startup task is only registered when the script reboots on purpose," -ForegroundColor Yellow
    Write-Host "  so run ad_lab.bat again yourself - every step is idempotent and the" -ForegroundColor Yellow
    Write-Host "  work already done (network, AD DS role) is skipped." -ForegroundColor Yellow
    Write-Host "  If the message above mentions a pending restart or a role change in" -ForegroundColor Yellow
    Write-Host "  progress, reboot first and then run it again." -ForegroundColor Yellow
    Write-Host ""
    Write-Note "$($_.ScriptStackTrace)"
    Write-Host ""
    try { Stop-Transcript | Out-Null } catch { }
    exit 1
}
