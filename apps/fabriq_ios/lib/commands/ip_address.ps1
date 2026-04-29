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
    # Ad-hoc IP configuration with optional Gateway / DNS overrides.
    # Each parameter beyond Ip+Mask is optional: when omitted, the
    # bound host's value (from `(config)# hostname <NewName>`) is
    # used; when no host is bound and the operator omits Gateway/DNS,
    # those env vars stay empty and the underlying module skips
    # those settings (or warns, depending on module behaviour).
    #
    # When no host is bound at all we seed an `(adhoc)` identity so
    # ipaddress_config still has SELECTED_NEW_PCNAME / OldPCName /
    # KanriNo populated for its display lines. All env mutations are
    # restored in finally regardless of success.
    param(
        [string]$Ip,
        [string]$Mask,
        [string]$Gateway,
        [string]$Dns1,
        [string]$Dns2,
        [hashtable]$State
    )
    if ($State.Mode -ne 'InterfaceConfig') {
        Write-Host "% 'ip address' is only available in interface configuration mode." -ForegroundColor Red
        return
    }

    $prev = @{
        NewName = $env:SELECTED_NEW_PCNAME
        OldName = $env:SELECTED_OLD_PCNAME
        Kanri   = $env:SELECTED_KANRI_NO
        Ip      = $env:SELECTED_ETH_IP
        Subnet  = $env:SELECTED_ETH_SUBNET
        Gateway = $env:SELECTED_ETH_GATEWAY
        Dns1    = $env:SELECTED_DNS1
        Dns2    = $env:SELECTED_DNS2
    }

    if ([string]::IsNullOrWhiteSpace($prev.NewName)) {
        $env:SELECTED_NEW_PCNAME = '(adhoc)'
        $env:SELECTED_OLD_PCNAME = $env:COMPUTERNAME
        $env:SELECTED_KANRI_NO   = '0'
    }
    $env:SELECTED_ETH_IP     = $Ip
    $env:SELECTED_ETH_SUBNET = $Mask
    if (-not [string]::IsNullOrWhiteSpace($Gateway)) { $env:SELECTED_ETH_GATEWAY = $Gateway }
    if (-not [string]::IsNullOrWhiteSpace($Dns1))    { $env:SELECTED_DNS1        = $Dns1 }
    if (-not [string]::IsNullOrWhiteSpace($Dns2))    { $env:SELECTED_DNS2        = $Dns2 }

    try {
        Invoke-IpAddressFromHostlist -State $State
    } finally {
        $env:SELECTED_NEW_PCNAME  = $prev.NewName
        $env:SELECTED_OLD_PCNAME  = $prev.OldName
        $env:SELECTED_KANRI_NO    = $prev.Kanri
        $env:SELECTED_ETH_IP      = $prev.Ip
        $env:SELECTED_ETH_SUBNET  = $prev.Subnet
        $env:SELECTED_ETH_GATEWAY = $prev.Gateway
        $env:SELECTED_DNS1        = $prev.Dns1
        $env:SELECTED_DNS2        = $prev.Dns2
    }
}
