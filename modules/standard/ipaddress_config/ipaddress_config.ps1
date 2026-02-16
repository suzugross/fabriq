# ========================================
# IP Address Configuration Script
# ========================================
# Description: Configures Ethernet and Wi-Fi IP addresses
# based on settings loaded from hostlist.csv
# ========================================

# Check Administrator Privileges
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
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
# Main Process
# ========================================

Write-Host "========================================" -ForegroundColor White
Write-Host "Starting Network Configuration" -ForegroundColor White
Write-Host "========================================" -ForegroundColor White
Write-Host ""

$dnsServers = @($config.DNS1, $config.DNS2, $config.DNS3, $config.DNS4) | Where-Object { $_ -and $_.Trim() -ne '' }

$successCount = 0
$totalAdapters = 0

# Ethernet Configuration
if ($config.EthIP -and $config.EthIP.Trim() -ne '') {
    $totalAdapters++
    $ethAdapter = Get-NetworkAdapter -Type "Ethernet"
    if ($ethAdapter) {
        $result = Set-IPConfiguration -Adapter $ethAdapter -IPAddress $config.EthIP -SubnetMask $config.EthSubnet -Gateway $config.EthGateway -DNSServers $dnsServers -AdapterType "Ethernet"
        if ($result) { $successCount++ }
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
        $result = Set-IPConfiguration -Adapter $wifiAdapter -IPAddress $config.WiFiIP -SubnetMask $config.WiFiSubnet -Gateway $config.WiFiGateway -DNSServers $dnsServers -AdapterType "Wi-Fi"
        if ($result) { $successCount++ }
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

if ($successCount -eq $totalAdapters -and $totalAdapters -gt 0) {
    Show-Success "All network adapters configured successfully ($successCount/$totalAdapters)"
}
elseif ($successCount -gt 0) {
    Show-Warning "Some network adapters configured successfully ($successCount/$totalAdapters)"
}
else {
    Show-Error "Failed to configure network adapters"
}

Write-Host ""

# Return ModuleResult
$overallStatus = if ($totalAdapters -eq 0) { "Skipped" }
    elseif ($successCount -eq $totalAdapters) { "Success" }
    elseif ($successCount -gt 0) { "Partial" }
    else { "Error" }
return (New-ModuleResult -Status $overallStatus -Message "Success: $successCount/$totalAdapters adapters")