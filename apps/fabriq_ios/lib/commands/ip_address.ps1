# ip address * command implementations.

function Get-IpAddressCompletionFromHostlist {
    param([hashtable]$State)
    $candidates = @('from-hostlist')
    if ($env:SELECTED_ETH_IP) { $candidates += $env:SELECTED_ETH_IP }
    if ($env:SELECTED_WIFI_IP -and $env:SELECTED_WIFI_IP -ne $env:SELECTED_ETH_IP) {
        $candidates += $env:SELECTED_WIFI_IP
    }
    return $candidates
}

function Invoke-IpAddressFromHostlist {
    param([hashtable]$State)
    if ($State.Mode -ne 'InterfaceConfig') {
        Write-Host "% 'ip address' is only available in interface configuration mode." -ForegroundColor Red
        return
    }
    if (-not $env:SELECTED_NEW_PCNAME) {
        Write-Host "% No host context. Run 'hostname <NewName>' in (config)# first to bind a host."
        return
    }
    if (-not $env:SELECTED_ETH_IP) {
        Write-Host "% Selected host has no EthernetIP; nothing to apply."
        return
    }

    $modulePath = Join-Path $script:FabriqRoot 'modules\standard\ipaddress_config\ipaddress_config.ps1'
    $previousPass = $global:FabriqMasterPassphrase
    if ($State.Passphrase) { $global:FabriqMasterPassphrase = $State.Passphrase }
    try {
        $result = Invoke-FabriqIosModule -ScriptPath $modulePath
    } finally {
        $global:FabriqMasterPassphrase = $previousPass
    }

    if (-not $result) {
        Write-Host "% Module returned no ModuleResult." -ForegroundColor Red
        return
    }

    switch ($result.Status) {
        'Success' {
            $prefix = ConvertFrom-SubnetMaskToPrefix $env:SELECTED_ETH_SUBNET
            Write-FabriqIosSyslog -Severity 6 -Mnemonic 'IPADDR' -Key 'whispered' `
                -Placeholders @{
                    Ip        = $env:SELECTED_ETH_IP
                    Prefix    = $prefix
                    Interface = $State.CurrentInterface
                }
            if ($env:SELECTED_ETH_GATEWAY) {
                Write-FabriqIosSyslog -Severity 6 -Mnemonic 'IPADDR' -Key 'gateway' `
                    -Placeholders @{ Gateway = $env:SELECTED_ETH_GATEWAY }
            }
            $dns = @($env:SELECTED_DNS1, $env:SELECTED_DNS2, $env:SELECTED_DNS3, $env:SELECTED_DNS4) |
                   Where-Object { $_ }
            if ($dns.Count -gt 0) {
                Write-FabriqIosSyslog -Severity 6 -Mnemonic 'IPADDR' -Key 'dns' `
                    -Placeholders @{ DnsList = ($dns -join ', ') }
            }
        }
        'Partial' {
            Write-Host ("% Partial: {0}" -f $result.Message) -ForegroundColor Yellow
        }
        'Skipped' {
            Write-Host ("% Skipped: {0}" -f $result.Message) -ForegroundColor Yellow
        }
        'Cancelled' {
            Write-Host ("% Cancelled: {0}" -f $result.Message) -ForegroundColor Yellow
        }
        default {
            Write-Host ("% IP address configuration failed: {0}" -f $result.Message) -ForegroundColor Red
        }
    }
}

function Invoke-IpAddressManual {
    param(
        [string]$Ip,
        [string]$Mask,
        [hashtable]$State
    )
    if ($State.Mode -ne 'InterfaceConfig') {
        Write-Host "% 'ip address' is only available in interface configuration mode." -ForegroundColor Red
        return
    }
    if (-not $env:SELECTED_NEW_PCNAME) {
        Write-Host "% No host context. Run 'hostname <NewName>' in (config)# first to bind a host."
        return
    }

    # Override Ethernet IP and subnet only; gateway / DNS are kept
    # from the previously bound host so the module can still set them.
    $prevIp     = $env:SELECTED_ETH_IP
    $prevSubnet = $env:SELECTED_ETH_SUBNET
    $env:SELECTED_ETH_IP     = $Ip
    $env:SELECTED_ETH_SUBNET = $Mask
    try {
        Invoke-IpAddressFromHostlist -State $State
    } finally {
        $env:SELECTED_ETH_IP     = $prevIp
        $env:SELECTED_ETH_SUBNET = $prevSubnet
    }
}
