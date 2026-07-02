# ========================================
# IP Address Configuration Script
# ========================================
# Description: Configures Ethernet and Wi-Fi IP addresses
# based on settings loaded from hostlist.csv
# ========================================

# Check Administrator Privileges
if (-not (Test-AdminPrivilege)) {
    Show-Error "This script requires administrator privileges."
    return (New-ModuleResult -Status "Error" -Message "Administrator privileges required")
}

Write-Host ""
Show-Separator
Write-Host "  IP Address Configuration" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# ========================================
# Load Settings from Environment Variables
# ========================================
Show-Info "Loading configuration..."

$config = @{
    KanriNo = $env:SELECTED_KANRI_NO
    OldPCName = $env:SELECTED_OLD_PCNAME
    NewPCName = $env:SELECTED_NEW_PCNAME

    EthIP = $env:SELECTED_ETH_IP
    EthSubnet = $env:SELECTED_ETH_SUBNET
    EthGateway = $env:SELECTED_ETH_GATEWAY

    WiFiIP = $env:SELECTED_WIFI_IP
    WiFiSubnet = $env:SELECTED_WIFI_SUBNET
    WiFiGateway = $env:SELECTED_WIFI_GATEWAY

    DNS1 = $env:SELECTED_DNS1
    DNS2 = $env:SELECTED_DNS2
    DNS3 = $env:SELECTED_DNS3
    DNS4 = $env:SELECTED_DNS4
}

# Display Configuration
Write-Host ""
Write-Host "[Selected Device Info]" -ForegroundColor Yellow
Write-Host "  Admin ID: $($config.KanriNo)"
Write-Host "  PC Name: $($config.OldPCName) -> $($config.NewPCName)"
Write-Host ""
Write-Host "[Ethernet Settings]" -ForegroundColor Yellow
Write-Host "  IP Address: $($config.EthIP)"
Write-Host "  Subnet Mask: $($config.EthSubnet)"
Write-Host "  Default Gateway: $($config.EthGateway)"
Write-Host ""
Write-Host "[Wi-Fi Settings]" -ForegroundColor Yellow
Write-Host "  IP Address: $($config.WiFiIP)"
Write-Host "  Subnet Mask: $($config.WiFiSubnet)"
Write-Host "  Default Gateway: $($config.WiFiGateway)"
Write-Host ""
Write-Host "[DNS Settings (Common)]" -ForegroundColor Yellow
Write-Host "  DNS1: $($config.DNS1)"
Write-Host "  DNS2: $($config.DNS2)"
Write-Host "  DNS3: $($config.DNS3)"
Write-Host "  DNS4: $($config.DNS4)"
Write-Host ""

# ========================================
# Function: Convert Subnet Mask to Prefix Length
# ========================================
function Convert-SubnetMaskToPrefix {
    param([string]$SubnetMask)

    $octets = $SubnetMask.Split('.')
    $binaryString = ""
    foreach ($octet in $octets) {
        $binaryString += [Convert]::ToString([int]$octet, 2).PadLeft(8, '0')
    }
    return ($binaryString.ToCharArray() | Where-Object { $_ -eq '1' }).Count
}

# ========================================
# Function: Detect Network Adapter
# ========================================
function Get-NetworkAdapter {
    param(
        [string]$Type  # "Ethernet" or "WiFi"
    )

    Show-Info "Detecting ${Type} adapter..."

    # Get physical adapters only (exclude virtual adapters like Hyper-V, VPN)
    # Exclude disabled adapters, but include disconnected ones (cable unplugged)
    $physicalAdapters = @(Get-NetAdapter -Physical -ErrorAction SilentlyContinue |
                          Where-Object { $_.Status -ne 'Disabled' })

    if ($Type -eq "Ethernet") {
        # Negative match: exclude Wi-Fi/Wireless/Bluetooth
        # InterfaceDescription is always English regardless of OS locale
        $adapter = $physicalAdapters | Where-Object {
            $_.InterfaceDescription -notmatch 'Wi-Fi|Wireless|WLAN|802\.11|Bluetooth'
        } | Select-Object -First 1
    }
    elseif ($Type -eq "WiFi") {
        # Positive match for Wi-Fi (InterfaceDescription is always English)
        $adapter = $physicalAdapters | Where-Object {
            $_.InterfaceDescription -match 'Wi-Fi|Wireless|WLAN|802\.11'
        } | Select-Object -First 1
    }

    if ($adapter) {
        Show-Success "${Type} adapter found: $($adapter.Name) ($($adapter.InterfaceDescription)) [Status: $($adapter.Status)]"
        return $adapter
    }
    else {
        Show-Warning "${Type} adapter not found"
        return $null
    }
}

# ========================================
# Function: Set IP Configuration
# ========================================
function Set-IPConfiguration {
    param(
        [object]$Adapter,
        [string]$IPAddress,
        [string]$SubnetMask,
        [string]$Gateway,
        [array]$DNSServers,
        [string]$AdapterType
    )

    Write-Host ""
    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "Configuring ${AdapterType}..." -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor White

    $adapterName = $Adapter.Name

    try {
        # Use netsh for IP configuration
        # netsh writes to the legacy store (TCP/IPv4 properties GUI) and syncs with modern stack
        # Handles DHCP->Static transition automatically, works on disconnected adapters

        # Set static IP address
        if ($Gateway -and $Gateway.Trim() -ne '') {
            Show-Info "Setting IP address: $IPAddress / $SubnetMask / Gateway: $Gateway"
            $output = & netsh interface ip set address name="$adapterName" static $IPAddress $SubnetMask $Gateway 2>&1
        }
        else {
            Show-Info "Setting IP address: $IPAddress / $SubnetMask (no gateway)"
            $output = & netsh interface ip set address name="$adapterName" static $IPAddress $SubnetMask 2>&1
        }

        if ($LASTEXITCODE -ne 0) {
            throw "netsh set address failed: $output"
        }
        Show-Success "IP address set"

        # Set DNS servers using netsh
        $validDNS = @($DNSServers | Where-Object { $_ -and $_.Trim() -ne '' })
        if ($validDNS.Count -gt 0) {
            Show-Info "Setting DNS servers..."

            # Primary DNS
            $output = & netsh interface ip set dns name="$adapterName" static $($validDNS[0]) primary 2>&1
            if ($LASTEXITCODE -ne 0) {
                throw "netsh set dns failed: $output"
            }

            # Additional DNS servers
            for ($i = 1; $i -lt $validDNS.Count; $i++) {
                $output = & netsh interface ip add dns name="$adapterName" $($validDNS[$i]) index=$($i + 1) 2>&1
                if ($LASTEXITCODE -ne 0) {
                    Show-Warning "Failed to add DNS $($validDNS[$i]): $output"
                }
            }

            Show-Success "DNS servers set: $($validDNS -join ', ')"
        }

        # Display configured settings
        Write-Host ""
        Show-Info "Configured settings:"
        Write-Host "  IP Address:      $IPAddress"
        Write-Host "  Subnet Mask:     $SubnetMask"
        if ($Gateway -and $Gateway.Trim() -ne '') {
            Write-Host "  Default Gateway: $Gateway"
        }
        if ($validDNS.Count -gt 0) {
            Write-Host "  DNS Servers:     $($validDNS -join ', ')"
        }

        Write-Host ""
        Show-Success "${AdapterType} configuration completed"

        return $true
    }
    catch {
        Show-Error "Error occurred during ${AdapterType} configuration: $_"
        return $false
    }
}

# ========================================
# Function: Strict idempotency probe (pre-apply skip gate)
# ========================================
# Returns $true ONLY when the adapter's live IPv4 config EXACTLY matches the
# target: a single Manual IPv4 == IP/prefix, a single default gateway == target
# (when a gateway is specified), and the DNS list equal in BOTH count and order.
# Any read failure or ambiguity returns $false (fail-closed) so the caller
# APPLIES rather than wrongly skipping. The asymmetric risk here is false-skip
# (leaving a wrong config), so the predicate is deliberately strict and only
# skips on a provable full match. Step 5.5 verification keeps its own (lenient)
# "the values I set are present" semantics on purpose - see CLAUDE.md notes.
function Test-IpConfigMatch {
    param(
        [object]$Adapter,
        [string]$IPAddress,
        [string]$SubnetMask,
        [string]$Gateway,
        [array]$DNSServers
    )
    try {
        $ifIndex = $Adapter.ifIndex
        $expectedPrefix = Convert-SubnetMaskToPrefix -SubnetMask $SubnetMask

        # (1) exactly one Manual IPv4, equal to target IP + prefix
        $manual = @(Get-NetIPAddress -InterfaceIndex $ifIndex -AddressFamily IPv4 -ErrorAction Stop |
                    Where-Object { $_.PrefixOrigin -eq 'Manual' -and $_.SuffixOrigin -eq 'Manual' })
        if ($manual.Count -ne 1) { return $false }
        if ($manual[0].IPAddress -ne $IPAddress -or $manual[0].PrefixLength -ne $expectedPrefix) { return $false }

        # (2) gateway (when target specifies one): exactly one default GW == target
        if ($Gateway -and $Gateway.Trim() -ne '') {
            $cfg = Get-NetIPConfiguration -InterfaceIndex $ifIndex -ErrorAction Stop
            $gws = @($cfg.IPv4DefaultGateway.NextHop)
            if ($gws.Count -ne 1 -or $gws[0] -ne $Gateway) { return $false }
        }

        # (3) DNS exact ordered equality (when target specifies DNS)
        $validDNS = @($DNSServers | Where-Object { $_ -and $_.Trim() -ne '' })
        if ($validDNS.Count -gt 0) {
            $cur = @((Get-DnsClientServerAddress -InterfaceIndex $ifIndex -AddressFamily IPv4 -ErrorAction Stop).ServerAddresses)
            if ($cur.Count -ne $validDNS.Count) { return $false }
            for ($i = 0; $i -lt $validDNS.Count; $i++) {
                if ($cur[$i] -ne $validDNS[$i]) { return $false }
            }
        }
        return $true
    }
    catch {
        return $false   # fail-closed: any read failure -> not a match -> apply
    }
}

# ========================================
# Main Process
# ========================================

Write-Host "========================================" -ForegroundColor White
Write-Host "Starting Network Configuration" -ForegroundColor White
Write-Host "========================================" -ForegroundColor White
Write-Host ""

$dnsServers = @($config.DNS1, $config.DNS2, $config.DNS3, $config.DNS4) | Where-Object { $_ -and $_.Trim() -ne '' }

$successCount = 0
$skipCount    = 0
$totalAdapters = 0

# Ethernet Configuration
if ($config.EthIP -and $config.EthIP.Trim() -ne '') {
    $totalAdapters++
    $ethAdapter = Get-NetworkAdapter -Type "Ethernet"
    if ($ethAdapter) {
        if (Test-IpConfigMatch -Adapter $ethAdapter -IPAddress $config.EthIP -SubnetMask $config.EthSubnet -Gateway $config.EthGateway -DNSServers $dnsServers) {
            Show-Skip "Ethernet already matches target configuration (no change applied)"
            $skipCount++
        }
        else {
            $result = Set-IPConfiguration -Adapter $ethAdapter -IPAddress $config.EthIP -SubnetMask $config.EthSubnet -Gateway $config.EthGateway -DNSServers $dnsServers -AdapterType "Ethernet"
            if ($result) { $successCount++ }
        }
    }
    else {
        Show-Warning "Ethernet adapter not found. Skipping."
    }
}

# Wi-Fi Configuration
if ($config.WiFiIP -and $config.WiFiIP.Trim() -ne '') {
    $totalAdapters++
    $wifiAdapter = Get-NetworkAdapter -Type "WiFi"
    if ($wifiAdapter) {
        if (Test-IpConfigMatch -Adapter $wifiAdapter -IPAddress $config.WiFiIP -SubnetMask $config.WiFiSubnet -Gateway $config.WiFiGateway -DNSServers $dnsServers) {
            Show-Skip "Wi-Fi already matches target configuration (no change applied)"
            $skipCount++
        }
        else {
            $result = Set-IPConfiguration -Adapter $wifiAdapter -IPAddress $config.WiFiIP -SubnetMask $config.WiFiSubnet -Gateway $config.WiFiGateway -DNSServers $dnsServers -AdapterType "Wi-Fi"
            if ($result) { $successCount++ }
        }
    }
    else {
        Show-Warning "Wi-Fi adapter not found. Skipping."
    }
}

# Result Summary
Write-Host ""
Write-Host "========================================" -ForegroundColor White
Write-Host "Configuration Completed" -ForegroundColor White
Write-Host "========================================" -ForegroundColor White
Write-Host ""

# A skipped adapter is already in the desired state -> counts as configured (not a failure).
$configured = $successCount + $skipCount
if ($totalAdapters -gt 0 -and $configured -eq $totalAdapters) {
    Show-Success "All network adapters in desired state (applied: $successCount, already-matched: $skipCount / $totalAdapters)"
}
elseif ($configured -gt 0) {
    Show-Warning "Some network adapters configured (applied: $successCount, already-matched: $skipCount / $totalAdapters)"
}
else {
    Show-Error "Failed to configure network adapters"
}

Write-Host ""

# ========================================
# Step 5.5: Post-Apply Verification
# ========================================
Show-Info "Verifying network configuration..."
Write-Host ""

$verifyPass = 0
$verifyFail = 0

# Verify each configured adapter
$adapterChecks = @()
if ($config.EthIP -and $config.EthIP.Trim() -ne '' -and $ethAdapter) {
    $adapterChecks += @{
        Name    = "Ethernet"
        Adapter = $ethAdapter
        IP      = $config.EthIP
        Subnet  = $config.EthSubnet
        Gateway = $config.EthGateway
    }
}
if ($config.WiFiIP -and $config.WiFiIP.Trim() -ne '' -and $wifiAdapter) {
    $adapterChecks += @{
        Name    = "Wi-Fi"
        Adapter = $wifiAdapter
        IP      = $config.WiFiIP
        Subnet  = $config.WiFiSubnet
        Gateway = $config.WiFiGateway
    }
}

# A configured target whose adapter was never found must count as a
# verification failure, not silently vanish from the check list (which
# yielded Verified=true alongside Status=Partial). Status already carries
# the failure, so this adds no new __GATE__ blocking - it only stops the
# Verified flag from over-claiming.
$missingTargets = 0
if ($config.EthIP -and $config.EthIP.Trim() -ne '' -and -not $ethAdapter) {
    Write-Host "  [VERIFY FAILED] Ethernet: target IP configured but no adapter was found" -ForegroundColor Red
    $verifyFail++
    $missingTargets++
}
if ($config.WiFiIP -and $config.WiFiIP.Trim() -ne '' -and -not $wifiAdapter) {
    Write-Host "  [VERIFY FAILED] Wi-Fi: target IP configured but no adapter was found" -ForegroundColor Red
    $verifyFail++
    $missingTargets++
}

foreach ($check in $adapterChecks) {
    $ifIndex = $check.Adapter.ifIndex
    $adapterLabel = $check.Name

    # IP Address + Prefix verification
    $expectedPrefix = Convert-SubnetMaskToPrefix -SubnetMask $check.Subnet
    $currentIP = Get-NetIPAddress -InterfaceIndex $ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue |
        Where-Object { $_.IPAddress -eq $check.IP -and $_.PrefixLength -eq $expectedPrefix }

    if ($currentIP) {
        Write-Host "  [VERIFIED] $adapterLabel IP: $($check.IP)/$expectedPrefix" -ForegroundColor Green
        $verifyPass++
    } else {
        $actualIP = (Get-NetIPAddress -InterfaceIndex $ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue |
            Select-Object -First 1)
        $actualDisplay = if ($actualIP) { "$($actualIP.IPAddress)/$($actualIP.PrefixLength)" } else { "none" }
        Write-Host "  [VERIFY FAILED] $adapterLabel IP: expected $($check.IP)/$expectedPrefix, actual $actualDisplay" -ForegroundColor Red
        $verifyFail++
    }

    # Gateway verification
    if ($check.Gateway -and $check.Gateway.Trim() -ne '') {
        $netConfig = Get-NetIPConfiguration -InterfaceIndex $ifIndex -ErrorAction SilentlyContinue
        $gwMatch = $netConfig.IPv4DefaultGateway | Where-Object { $_.NextHop -eq $check.Gateway }

        if ($gwMatch) {
            Write-Host "  [VERIFIED] $adapterLabel Gateway: $($check.Gateway)" -ForegroundColor Green
            $verifyPass++
        } else {
            $actualGw = if ($netConfig.IPv4DefaultGateway) { $netConfig.IPv4DefaultGateway.NextHop -join ', ' } else { "none" }
            Write-Host "  [VERIFY FAILED] $adapterLabel Gateway: expected $($check.Gateway), actual $actualGw" -ForegroundColor Red
            $verifyFail++
        }
    }

    # DNS verification
    $validDNS = @($dnsServers)
    if ($validDNS.Count -gt 0) {
        $currentDNS = @((Get-DnsClientServerAddress -InterfaceIndex $ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue).ServerAddresses)
        $dnsMatch = $true
        for ($i = 0; $i -lt $validDNS.Count; $i++) {
            if ($i -ge $currentDNS.Count -or $currentDNS[$i] -ne $validDNS[$i]) {
                $dnsMatch = $false
                break
            }
        }

        if ($dnsMatch) {
            Write-Host "  [VERIFIED] $adapterLabel DNS: $($validDNS -join ', ')" -ForegroundColor Green
            $verifyPass++
        } else {
            $actualDNS = if ($currentDNS.Count -gt 0) { $currentDNS -join ', ' } else { "none" }
            Write-Host "  [VERIFY FAILED] $adapterLabel DNS: expected $($validDNS -join ', '), actual $actualDNS" -ForegroundColor Red
            $verifyFail++
        }
    }
}

Write-Host ""
# $null only when there was nothing to verify at all (no reachable checks
# AND no configured-but-missing targets); missing targets force $false.
$verified = if ($adapterChecks.Count -eq 0 -and $missingTargets -eq 0) { $null } else { $verifyFail -eq 0 }

# Return ModuleResult
# Skipped adapters (already matching) count as configured, not failures.
$overallStatus = if ($totalAdapters -eq 0) { "Skipped" }
    elseif ($configured -eq $totalAdapters -and $successCount -eq 0) { "Skipped" }   # all already in desired state
    elseif ($configured -eq $totalAdapters) { "Success" }                            # applied + already-matched, none failed
    elseif ($configured -gt 0) { "Partial" }                                         # some configured, some failed/not-found
    else { "Error" }
return (New-ModuleResult -Status $overallStatus -Message "Applied: $successCount, Skipped(match): $skipCount / $totalAdapters adapters" -Verified $verified)