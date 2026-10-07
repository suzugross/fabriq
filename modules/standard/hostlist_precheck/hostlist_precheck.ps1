# ========================================
# Hostlist Pre-Check Script
# ========================================
# [PURPOSE]
# Read-only gate. Compares the selected hostlist row (SELECTED_* env vars)
# with this PC's saved configuration: every compared item matches ->
# Success, any mismatch -> Error. Nothing is written to the PC.
#
# [NOTES]
# - The check items are chosen in precheck_list.csv (one row per item).
# - Place this module first in a profile and put __GATE__ right after it:
#   the gate refuses every later module while this one is in Error.
# - An address counts as static when EITHER the saved settings in the
#   registry (Tcpip\Parameters\Interfaces\{GUID}, EnableDHCP = 0) OR the
#   live stack (Manual-origin address, as ipaddress_config tests it) says
#   so. An address leased by DHCP satisfies neither. "Static or DHCP" is
#   never decided from an interface-level flag alone: that judged a PC
#   fabriq itself had configured as DHCP (field reports, 2026-10-07).
# - The gateway and the static DNS servers are read from the saved
#   settings, so they do not depend on the state of the link.
# - DNS1-4 are shared by both adapters. An adapter's DNS item is compared
#   only when the hostlist also gives that adapter an IP address.
# - Deliberate deviations from dev/template:
#     Step 4  No confirmation prompt. A Cancelled result passes __GATE__,
#             so a prompt would let the operator wave the check through.
#     Step 6  New-ModuleResult instead of New-BatchResult. A partly failed
#             check must be Error, never Partial.
# - Adapter selection mirrors ipaddress_config (Get-NetworkAdapter). Keep
#   the two in sync: this module must look at the adapter that
#   ipaddress_config would configure.
# ========================================

Write-Host ""
Show-Separator
Write-Host "Hostlist Pre-Check" -ForegroundColor Cyan
Show-Separator
Write-Host ""


# ========================================
# Helper functions
# ========================================
# The functions up to Get-PrecheckOutcome are pure (no access to the PC)
# and are unit-tested in tests/modules/hostlist_precheck. The two reader
# functions below them are the only code that touches the PC.
# ========================================

function Get-PrecheckKnownItems {
    return @(
        'NewPCName', 'OldPCName',
        'EthernetIP', 'EthernetSubnet', 'EthernetGateway', 'EthernetDNS',
        'WifiIP', 'WifiSubnet', 'WifiGateway', 'WifiDNS'
    )
}

function Resolve-PrecheckItemName {
    # Returns the canonical spelling of a CheckItem, or $null when the name
    # is not a known item. Matching ignores case and surrounding spaces.
    param([string]$Name)

    if ([string]::IsNullOrWhiteSpace($Name)) { return $null }
    $wanted = $Name.Trim()
    foreach ($known in (Get-PrecheckKnownItems)) {
        if ($known -eq $wanted) { return $known }
    }
    return $null
}

function Convert-PrecheckMaskToPrefix {
    # Dotted subnet mask -> prefix length. Returns $null for anything that
    # is not a valid contiguous mask (the caller reports it as unreadable).
    param([string]$SubnetMask)

    if ([string]::IsNullOrWhiteSpace($SubnetMask)) { return $null }
    $octets = $SubnetMask.Trim().Split('.')
    if ($octets.Count -ne 4) { return $null }

    $bits = ''
    foreach ($octet in $octets) {
        if ($octet -notmatch '^\d{1,3}$') { return $null }
        $number = [int]$octet
        if ($number -gt 255) { return $null }
        $bits += [Convert]::ToString($number, 2).PadLeft(8, '0')
    }
    if ($bits -notmatch '^1*0*$') { return $null }
    return ($bits.Length - $bits.Replace('1', '').Length)
}

function Convert-PrecheckPrefixToMask {
    # Prefix length -> dotted subnet mask (display only).
    param([int]$PrefixLength)

    if ($PrefixLength -lt 0 -or $PrefixLength -gt 32) { return '' }
    $bits = ('1' * $PrefixLength).PadRight(32, '0')
    $parts = @()
    for ($i = 0; $i -lt 4; $i++) {
        $parts += [Convert]::ToInt32($bits.Substring($i * 8, 8), 2)
    }
    return ($parts -join '.')
}

function ConvertTo-PrecheckSavedConfig {
    # Turns one adapter's saved IPv4 settings (the registry values under
    # Tcpip\Parameters\Interfaces\{GUID}) into the fields the judge uses.
    # Pure: it is handed the values and reads nothing itself.
    #
    #   Dhcp        'Disabled' (EnableDHCP = 0) / 'Enabled' / $null (unknown)
    #   StaticIPv4  saved static addresses, each with its mask as a prefix
    #               length ($null prefix = the mask could not be interpreted)
    #   Gateways    saved default gateways
    #   StaticDns   saved DNS servers, in order
    #
    # StaticIPv4 is filled only when EnableDHCP is 0. With DHCP on (or the
    # flag unknown) whatever is left in IPAddress is not in effect and must
    # not count as a static address.
    param($RegistryValues)

    $config = @{ Dhcp = $null; StaticIPv4 = $null; Gateways = $null; StaticDns = $null }
    if ($null -eq $RegistryValues) { return $config }

    if ($null -ne $RegistryValues.EnableDHCP) {
        $config.Dhcp = if ([string]$RegistryValues.EnableDHCP -eq '0') { 'Disabled' } else { 'Enabled' }
    }

    $static = @()
    if ($config.Dhcp -eq 'Disabled') {
        $addresses = @($RegistryValues.IPAddress)
        $masks     = @($RegistryValues.SubnetMask)
        for ($i = 0; $i -lt $addresses.Count; $i++) {
            $address = ([string]$addresses[$i]).Trim()
            if ([string]::IsNullOrWhiteSpace($address) -or $address -eq '0.0.0.0') { continue }
            $mask = if ($i -lt $masks.Count) { [string]$masks[$i] } else { '' }
            $static += @{ IPAddress = $address; PrefixLength = (Convert-PrecheckMaskToPrefix -SubnetMask $mask) }
        }
    }
    $config.StaticIPv4 = $static

    $config.Gateways = @(@($RegistryValues.DefaultGateway) |
        ForEach-Object { ([string]$_).Trim() } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) })

    # NameServer is one string; the separator is a comma or a space.
    $config.StaticDns = @(([string]$RegistryValues.NameServer) -split '[,\s]+' |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) })

    return $config
}

function Join-PrecheckStaticAddresses {
    # Merges the two witnesses of "this adapter has a static IPv4 address":
    #   -Saved  addresses in the saved settings (registry, EnableDHCP = 0)
    #   -Live   addresses the live stack reports with a Manual origin
    #           (the test ipaddress_config uses)
    # Either witness alone proves an address is static, and neither can
    # report an address that was handed out by DHCP. Two are needed because
    # each has a blind spot: an interface-level DHCP flag is not reliable on
    # every PC, and the live stack depends on the state of the link.
    # $null = that witness could not be read. Pure.
    #
    # Returns @{ Addresses; SeenIn }:
    #   Addresses  merged list, or $null when BOTH witnesses failed
    #   SeenIn     '<ip>' -> which witness(es) reported it (display only)
    # Entries are merged by IP address. A prefix only one witness knows is
    # taken from that witness; two different prefixes stay as two entries,
    # so an inconsistent PC cannot pass as "one static address".
    param($Saved, $Live)

    $merged = @{ Addresses = $null; SeenIn = @{} }
    if ($null -eq $Saved -and $null -eq $Live) { return $merged }

    $order    = @()
    $prefixes = @{}
    $witness  = @{}
    $sources  = @(
        @{ Name = 'saved settings'; Items = $Saved }
        @{ Name = 'live stack';     Items = $Live }
    )
    foreach ($source in $sources) {
        foreach ($entry in @($source.Items)) {
            if ($null -eq $entry) { continue }
            $ip = ([string]$entry.IPAddress).Trim()
            if ([string]::IsNullOrWhiteSpace($ip)) { continue }

            if ($order -notcontains $ip) {
                $order += $ip
                $prefixes[$ip] = @()
                $witness[$ip]  = @()
            }
            if ($witness[$ip] -notcontains $source.Name) { $witness[$ip] += $source.Name }
            if ($null -ne $entry.PrefixLength) {
                $prefix = [int]$entry.PrefixLength
                if ($prefixes[$ip] -notcontains $prefix) { $prefixes[$ip] += $prefix }
            }
        }
    }

    $addresses = @()
    foreach ($ip in $order) {
        if (@($prefixes[$ip]).Count -eq 0) {
            $addresses += @{ IPAddress = $ip; PrefixLength = $null }
        }
        else {
            foreach ($prefix in $prefixes[$ip]) {
                $addresses += @{ IPAddress = $ip; PrefixLength = $prefix }
            }
        }
        $merged.SeenIn[$ip] = $witness[$ip] -join ' + '
    }
    $merged.Addresses = $addresses
    return $merged
}

function Get-PrecheckItemResult {
    # Judges ONE check item. Pure: everything it needs arrives in $Expected
    # (hostlist values) and $Live (a snapshot of this PC).
    #
    # State is one of:
    #   Match         - compared and equal
    #   Mismatch      - compared and different
    #   Unreadable    - could not be compared (adapter missing, read failed,
    #                   hostlist value not interpretable); counts as failure
    #   NotApplicable - the hostlist has no value for this item; not compared.
    #                   A DNS item is also not applicable when the hostlist
    #                   gives that adapter no IP address (see below).
    param(
        [string]$CheckItem,
        [hashtable]$Expected,
        [hashtable]$Live
    )

    # Starts as Unreadable so that any path that does not reach an explicit
    # verdict counts as a failure (fail-closed).
    $result = [PSCustomObject]@{
        Item     = $CheckItem
        Label    = $CheckItem
        Expected = ''
        Actual   = ''
        State    = 'Unreadable'
        Note     = ''
    }

    # ----------------------------------------
    # Hostname items
    # ----------------------------------------
    if ($CheckItem -eq 'NewPCName' -or $CheckItem -eq 'OldPCName') {
        $result.Label = "Hostname ($CheckItem)"

        $expectedName = [string]$Expected[$CheckItem]
        if ([string]::IsNullOrWhiteSpace($expectedName)) {
            $result.State = 'NotApplicable'
            return $result
        }
        $expectedName = $expectedName.Trim()
        $result.Expected = $expectedName

        $actualName = [string]$Live.ComputerName
        if ([string]::IsNullOrWhiteSpace($actualName)) {
            $result.Actual = '(could not be read)'
            return $result
        }
        $result.Actual = $actualName

        if ([string]::Equals($expectedName, $actualName, [System.StringComparison]::OrdinalIgnoreCase)) {
            $result.State = 'Match'
            return $result
        }

        $result.State = 'Mismatch'
        $pendingName = [string]$Live.PendingComputerName
        if (-not [string]::IsNullOrWhiteSpace($pendingName) -and
            -not [string]::Equals($pendingName, $actualName, [System.StringComparison]::OrdinalIgnoreCase)) {
            $result.Note = "A rename to '$pendingName' is pending and takes effect after a restart"
        }
        return $result
    }

    # ----------------------------------------
    # Adapter items
    # ----------------------------------------
    $parsed = [regex]::Match($CheckItem, '^(Ethernet|Wifi)(IP|Subnet|Gateway|DNS)$')
    if (-not $parsed.Success) {
        $result.Note = 'Unknown check item'
        return $result
    }
    $adapterKey   = $parsed.Groups[1].Value
    $aspect       = $parsed.Groups[2].Value
    $adapterLabel = if ($adapterKey -eq 'Wifi') { 'Wi-Fi' } else { 'Ethernet' }
    $aspectLabel  = switch ($aspect) {
        'IP'      { 'IP address' }
        'Subnet'  { 'subnet mask' }
        'Gateway' { 'default gateway' }
        default   { 'DNS servers' }
    }
    $result.Label = "$adapterLabel $aspectLabel"

    # The hostlist value comes first: a blank value means "nothing to
    # assert", even when the adapter itself does not exist.
    $expectedValue  = ''
    $expectedPrefix = $null
    $expectedDns    = @()
    if ($aspect -eq 'DNS') {
        # DNS1-4 are one set of columns shared by both adapters. They are an
        # expectation for THIS adapter only when the hostlist also gives it
        # an IP address - the condition under which ipaddress_config sets
        # DNS on an adapter. Without it, a hostlist that configures only
        # Ethernet would fail on the DNS servers of a DHCP Wi-Fi adapter.
        if ([string]::IsNullOrWhiteSpace([string]$Expected["${adapterKey}IP"])) {
            $result.State = 'NotApplicable'
            $result.Note = "The hostlist has no $adapterLabel IP address, so the DNS servers of this adapter are not compared"
            return $result
        }

        $expectedDns = @($Expected.Dns |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            ForEach-Object { ([string]$_).Trim() })
        if ($expectedDns.Count -eq 0) {
            $result.State = 'NotApplicable'
            return $result
        }
        $result.Expected = $expectedDns -join ', '
    }
    else {
        $expectedValue = [string]$Expected["$adapterKey$aspect"]
        if ([string]::IsNullOrWhiteSpace($expectedValue)) {
            $result.State = 'NotApplicable'
            return $result
        }
        $expectedValue = $expectedValue.Trim()
        $result.Expected = $expectedValue

        if ($aspect -eq 'Subnet') {
            $expectedPrefix = Convert-PrecheckMaskToPrefix -SubnetMask $expectedValue
            if ($null -eq $expectedPrefix) {
                $result.Note = 'The subnet mask in the hostlist is not a valid mask'
                return $result
            }
            $result.Expected = "$expectedValue (/$expectedPrefix)"
        }
    }

    $adapter = $Live[$adapterKey]
    if ($null -eq $adapter -or -not $adapter.Found) {
        $result.Actual = '(adapter not found)'
        $result.Note = "No $adapterLabel adapter was found on this PC"
        return $result
    }

    # ---- IP address / subnet mask: the single static IPv4 ----
    # StaticIPv4 holds only addresses proven static (see
    # Join-PrecheckStaticAddresses), so an address leased by DHCP is never
    # in it - even when it happens to equal the hostlist value. No
    # interface-level DHCP flag takes part in the verdict.
    if ($aspect -eq 'IP' -or $aspect -eq 'Subnet') {
        if ($null -eq $adapter.StaticIPv4) {
            $result.Actual = '(could not be read)'
            return $result
        }

        $static = @($adapter.StaticIPv4)
        if ($static.Count -eq 0) {
            $result.State = 'Mismatch'
            $result.Actual = '(no static IPv4 address)'
            $liveText = (@($adapter.LiveIPv4) | Where-Object { $_ }) -join ', '
            if ($liveText) {
                $result.Note = "The current address $liveText was obtained automatically (DHCP); it is not a static address"
            }
            return $result
        }
        if ($static.Count -gt 1) {
            $result.State = 'Mismatch'
            $result.Actual = ($static | ForEach-Object { "$($_.IPAddress)/$($_.PrefixLength)" }) -join ', '
            $result.Note = 'More than one static IPv4 address is set on the adapter'
            return $result
        }

        $actualIp = [string]$static[0].IPAddress
        if ($aspect -eq 'IP') {
            $result.Actual = $actualIp
            $result.State = if ($actualIp -eq $expectedValue) { 'Match' } else { 'Mismatch' }
            return $result
        }

        if ($null -eq $static[0].PrefixLength) {
            $result.Actual = '(could not be read)'
            $result.Note = 'The subnet mask saved on the adapter could not be interpreted'
            return $result
        }
        $actualPrefix = [int]$static[0].PrefixLength
        $result.Actual = "$(Convert-PrecheckPrefixToMask -PrefixLength $actualPrefix) (/$actualPrefix)"
        $result.State = if ($actualPrefix -eq $expectedPrefix) { 'Match' } else { 'Mismatch' }
        return $result
    }

    # ---- Default gateway: the single saved default gateway ----
    # A gateway handed out by DHCP is never saved on the adapter, so a DHCP
    # adapter naturally reads as "no static default gateway".
    if ($aspect -eq 'Gateway') {
        if ($null -eq $adapter.Gateways) {
            $result.Actual = '(could not be read)'
            return $result
        }

        $gateways = @($adapter.Gateways | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        if ($gateways.Count -eq 0) {
            $result.State = 'Mismatch'
            $result.Actual = '(no static default gateway)'
            return $result
        }
        $result.Actual = $gateways -join ', '
        if ($gateways.Count -gt 1) {
            $result.State = 'Mismatch'
            $result.Note = 'More than one static default gateway is saved on the adapter'
            return $result
        }
        $result.State = if ([string]$gateways[0] -eq $expectedValue) { 'Match' } else { 'Mismatch' }
        return $result
    }

    # ---- DNS servers: same count, same order ----
    if ($null -eq $adapter.Dns) {
        $result.Actual = '(could not be read)'
        return $result
    }

    $actualDns = @($adapter.Dns |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        ForEach-Object { ([string]$_).Trim() })
    $result.Actual = if ($actualDns.Count -gt 0) { $actualDns -join ', ' } else { '(none)' }

    $sameDns = ($actualDns.Count -eq $expectedDns.Count)
    if ($sameDns) {
        for ($i = 0; $i -lt $expectedDns.Count; $i++) {
            if ($actualDns[$i] -ne $expectedDns[$i]) { $sameDns = $false; break }
        }
    }
    $result.State = if ($sameDns) { 'Match' } else { 'Mismatch' }
    return $result
}

function Get-PrecheckOutcome {
    # Folds the per-item results into the module result. Pure.
    #   nothing compared            -> Skipped, Verified = $null
    #   every compared item matches -> Success, Verified = $true
    #   anything else               -> Error,   Verified = $false
    # There is no Partial: one failed item fails the whole check.
    param([array]$Results)

    $compared = @(@($Results) | Where-Object { $null -ne $_ -and $_.State -ne 'NotApplicable' })
    $failed   = @($compared | Where-Object { $_.State -ne 'Match' })

    if ($compared.Count -eq 0) {
        return @{
            Status   = 'Skipped'
            Verified = $null
            Message  = 'Nothing to compare: the hostlist has no value for any enabled check item'
        }
    }
    if ($failed.Count -gt 0) {
        $failedNames = ($failed | ForEach-Object { $_.Item }) -join ', '
        return @{
            Status   = 'Error'
            Verified = $false
            Message  = "Check failed: $failedNames ($($failed.Count) of $($compared.Count) checked)"
        }
    }
    return @{
        Status   = 'Success'
        Verified = $true
        Message  = "All $($compared.Count) checked items match the hostlist"
    }
}

function Get-PrecheckAdapter {
    # Same rule as ipaddress_config: physical adapters only, Disabled ones
    # excluded, link-down (Disconnected) ones INCLUDED, first match wins.
    # InterfaceDescription is English regardless of the OS language.
    param([string]$Type)   # "Ethernet" or "Wifi"

    $physicalAdapters = @(Get-NetAdapter -Physical -ErrorAction SilentlyContinue |
                          Where-Object { $_.Status -ne 'Disabled' })

    if ($Type -eq 'Wifi') {
        return ($physicalAdapters | Where-Object {
            $_.InterfaceDescription -match 'Wi-Fi|Wireless|WLAN|802\.11'
        } | Select-Object -First 1)
    }
    return ($physicalAdapters | Where-Object {
        $_.InterfaceDescription -notmatch 'Wi-Fi|Wireless|WLAN|802\.11|Bluetooth'
    } | Select-Object -First 1)
}

function Get-PrecheckAdapterState {
    # Snapshot of one adapter's IPv4 configuration.
    #
    # StaticIPv4 / Gateways / Dns hold $null when the value could not be
    # read (judged as a failure) and an empty array when the read worked
    # and nothing is configured.
    #
    # StaticIPv4 is the merge of two witnesses (Join-PrecheckStaticAddresses):
    # the saved settings in the registry and the Manual-origin addresses of
    # the live stack. Do NOT decide "static or DHCP" from an interface-level
    # flag alone: a PC that fabriq itself had configured with a static
    # address was judged as DHCP that way (field reports, 2026-10-07).
    # The gateway and the static DNS servers come from the saved settings.
    # Dhcp (the saved flag), StaticSeenIn and LiveIPv4 are display only.
    param([string]$Type)

    $state = @{
        Found        = $false
        Name         = ''
        Description  = ''
        LinkStatus   = ''
        Dhcp         = $null
        StaticIPv4   = $null
        StaticSeenIn = @{}
        Gateways     = $null
        Dns          = $null
        LiveIPv4     = @()
    }

    $adapter = Get-PrecheckAdapter -Type $Type
    if ($null -eq $adapter) { return $state }

    $ifIndex = $adapter.ifIndex
    $state.Found       = $true
    $state.Name        = [string]$adapter.Name
    $state.Description = [string]$adapter.InterfaceDescription
    $state.LinkStatus  = [string]$adapter.Status

    # Saved IPv4 settings of this adapter. An unreadable key leaves every
    # field at $null (= could not be read).
    $saved = $null
    $interfaceGuid = [string]$adapter.InterfaceGuid
    if (-not [string]::IsNullOrWhiteSpace($interfaceGuid)) {
        try {
            $registryPath = "HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces\$interfaceGuid"
            $saved = ConvertTo-PrecheckSavedConfig -RegistryValues (Get-ItemProperty -LiteralPath $registryPath -ErrorAction Stop)
        }
        catch {
            $saved = $null
        }
    }

    $savedStatic = $null
    if ($null -ne $saved) {
        $state.Dhcp     = $saved.Dhcp
        $state.Gateways = $saved.Gateways
        $savedStatic    = $saved.StaticIPv4
        if (@($saved.StaticDns).Count -gt 0) { $state.Dns = $saved.StaticDns }
    }

    # Live stack: the Manual-origin addresses are the second witness of a
    # static address; the rest (leased by DHCP) is kept for display. An
    # interface without any IPv4 address raises CmdletizationQuery_NotFound*,
    # which is a normal answer ("none"), not a read failure.
    $liveStatic = $null
    try {
        $liveAddresses = @(Get-NetIPAddress -InterfaceIndex $ifIndex -AddressFamily IPv4 -ErrorAction Stop)
        $liveStatic = @($liveAddresses |
            Where-Object { $_.PrefixOrigin -eq 'Manual' -and $_.SuffixOrigin -eq 'Manual' } |
            ForEach-Object { @{ IPAddress = [string]$_.IPAddress; PrefixLength = [int]$_.PrefixLength } })
        $state.LiveIPv4 = @($liveAddresses |
            Where-Object { $_.PrefixOrigin -ne 'Manual' -and $_.PrefixOrigin -ne 'WellKnown' } |
            ForEach-Object { [string]$_.IPAddress })
    }
    catch {
        if ($_.FullyQualifiedErrorId -like 'CmdletizationQuery_NotFound*') { $liveStatic = @() }
        else { $liveStatic = $null }
    }

    $joined = Join-PrecheckStaticAddresses -Saved $savedStatic -Live $liveStatic
    $state.StaticIPv4   = $joined.Addresses
    $state.StaticSeenIn = $joined.SeenIn

    # No DNS server is saved on the adapter: compare against the servers
    # the stack currently uses (the ones handed out by DHCP).
    if ($null -eq $state.Dns) {
        try {
            $dnsEntry = Get-DnsClientServerAddress -InterfaceIndex $ifIndex -AddressFamily IPv4 -ErrorAction Stop
            $state.Dns = @($dnsEntry.ServerAddresses |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                ForEach-Object { [string]$_ })
        }
        catch {
            $state.Dns = $null
        }
    }

    return $state
}


# ========================================
# Step 1: Load CSV
# ========================================
$csvPath = Join-Path $PSScriptRoot "precheck_list.csv"

$enabledItems = Import-ModuleCsv -Path $csvPath -FilterEnabled `
    -RequiredColumns @("Enabled", "CheckItem")

if ($null -eq $enabledItems) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load precheck_list.csv")
}
if (@($enabledItems).Count -eq 0) {
    return (New-ModuleResult -Status "Skipped" -Message "No enabled entries")
}


# ========================================
# Step 2: Prerequisite check (early return)
# ========================================
# Every enabled CheckItem must be a known item. A misspelled name must not
# silently drop a check, so nothing is compared and the module fails.
# ========================================
$checkItems   = @()
$unknownItems = @()

foreach ($row in $enabledItems) {
    $canonical = Resolve-PrecheckItemName -Name $row.CheckItem
    if ($null -eq $canonical) {
        $rawName = ([string]$row.CheckItem).Trim()
        $unknownItems += $(if ($rawName) { $rawName } else { '(blank)' })
        continue
    }
    if ($checkItems -notcontains $canonical) { $checkItems += $canonical }
}

if ($unknownItems.Count -gt 0) {
    $unknownText = $unknownItems -join ', '
    Show-Error "Unknown CheckItem in precheck_list.csv: $unknownText"
    Show-Info "Valid items: $((Get-PrecheckKnownItems) -join ', ')"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Unknown CheckItem in precheck_list.csv: $unknownText")
}


# ========================================
# Step 3: Compare the hostlist with this PC and show the result
# ========================================
# Expected values are the selected hostlist row (SELECTED_* env vars).
# An item is compared whenever its hostlist value is not blank (a DNS item
# additionally needs an IP address for its adapter in the hostlist).
# ========================================
$expected = @{
    NewPCName       = $env:SELECTED_NEW_PCNAME
    OldPCName       = $env:SELECTED_OLD_PCNAME
    EthernetIP      = $env:SELECTED_ETH_IP
    EthernetSubnet  = $env:SELECTED_ETH_SUBNET
    EthernetGateway = $env:SELECTED_ETH_GATEWAY
    WifiIP          = $env:SELECTED_WIFI_IP
    WifiSubnet      = $env:SELECTED_WIFI_SUBNET
    WifiGateway     = $env:SELECTED_WIFI_GATEWAY
    Dns             = @($env:SELECTED_DNS1, $env:SELECTED_DNS2, $env:SELECTED_DNS3, $env:SELECTED_DNS4)
}

$live = @{
    ComputerName        = $env:COMPUTERNAME
    PendingComputerName = ''
    Ethernet            = @{ Found = $false }
    Wifi                = @{ Found = $false }
}

# The pending name is shown as a hint when the hostname does not match
# (a rename that has not been restarted yet). It never decides the verdict.
try {
    $live.PendingComputerName = [string](Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName" -Name "ComputerName" -ErrorAction Stop).ComputerName
}
catch { }

if (@($checkItems | Where-Object { $_ -like 'Ethernet*' }).Count -gt 0) {
    $live.Ethernet = Get-PrecheckAdapterState -Type 'Ethernet'
}
if (@($checkItems | Where-Object { $_ -like 'Wifi*' }).Count -gt 0) {
    $live.Wifi = Get-PrecheckAdapterState -Type 'Wifi'
}

$results = @()
foreach ($checkItem in $checkItems) {
    $results += Get-PrecheckItemResult -CheckItem $checkItem -Expected $expected -Live $live
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "Hostlist vs This PC" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""
Write-Host "  Selected host    : $($env:SELECTED_KANRI_NO)"
foreach ($adapterKey in @('Ethernet', 'Wifi')) {
    $adapterState = $live[$adapterKey]
    if ($adapterState.Found) {
        $adapterTitle = if ($adapterKey -eq 'Wifi') { 'Wi-Fi adapter   ' } else { 'Ethernet adapter' }
        Write-Host "  $adapterTitle : $($adapterState.Name) ($($adapterState.Description)) [Link: $($adapterState.LinkStatus)]"

        # Where each static address was found. This is the evidence behind
        # the IP / subnet verdicts, shown so a wrong verdict can be traced.
        $savedDhcp = if ($adapterState.Dhcp) { $adapterState.Dhcp } else { 'unknown' }
        if ($null -eq $adapterState.StaticIPv4) {
            $staticText = '(could not be read)'
        }
        elseif (@($adapterState.StaticIPv4).Count -eq 0) {
            $staticText = 'none'
        }
        else {
            $staticText = (@($adapterState.StaticIPv4) | ForEach-Object {
                "$($_.IPAddress)/$($_.PrefixLength) [$($adapterState.StaticSeenIn[$_.IPAddress])]"
            }) -join ', '
        }
        Write-Host "      Static IPv4  : $staticText" -ForegroundColor DarkGray
        Write-Host "      Saved DHCP   : $savedDhcp" -ForegroundColor DarkGray
    }
}
Write-Host ""

foreach ($result in $results) {
    $tag      = '[N/A]'
    $tagColor = 'DarkGray'
    switch ($result.State) {
        'Match'      { $tag = '[MATCH]';      $tagColor = 'Green' }
        'Mismatch'   { $tag = '[MISMATCH]';   $tagColor = 'Red' }
        'Unreadable' { $tag = '[UNREADABLE]'; $tagColor = 'Red' }
    }

    Write-Host ("  {0,-13}{1}" -f $tag, $result.Label) -ForegroundColor $tagColor
    if ($result.State -eq 'NotApplicable') {
        # A note carries the reason when it is not simply a blank value.
        $notComparedText = if ($result.Note) { '(not compared)' } else { '(blank in hostlist - not compared)' }
        Write-Host "      Expected : $notComparedText" -ForegroundColor DarkGray
    }
    else {
        Write-Host "      Expected : $($result.Expected)"
        Write-Host "      Actual   : $($result.Actual)"
    }
    if ($result.Note) {
        Write-Host "      Note     : $($result.Note)" -ForegroundColor DarkGray
    }
    Write-Host ""
}

Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""


# ========================================
# Step 4: User confirmation - intentionally omitted
# ========================================
# This module changes nothing, so there is nothing to confirm. More
# importantly a Cancelled result is not blocked by __GATE__: a prompt here
# would let the operator answer N and continue kitting unchecked.
# ========================================


# ========================================
# Step 5: Report each failed item
# ========================================
foreach ($result in $results) {
    if ($result.State -eq 'Mismatch') {
        Show-Error "$($result.Label): expected '$($result.Expected)', actual '$($result.Actual)'"
    }
    elseif ($result.State -eq 'Unreadable') {
        $reason = if ($result.Note) { $result.Note } else { 'the current value could not be read' }
        Show-Error "$($result.Label): could not be checked ($reason)"
    }
}


# ========================================
# Step 6: Aggregate and return result
# ========================================
# The comparison above IS the read-back, so Verified reports it directly:
# $true = everything compared matches, $false = something does not,
# $null = nothing was compared.
# ========================================
$outcome = Get-PrecheckOutcome -Results $results

switch ($outcome.Status) {
    'Success' { Show-Success $outcome.Message }
    'Skipped' { Show-Skip $outcome.Message }
    default   { Show-Error $outcome.Message }
}
Write-Host ""

return (New-ModuleResult -Status $outcome.Status -Message $outcome.Message -Verified $outcome.Verified)
