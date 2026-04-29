# ip address * command implementations.
# Phase 8: hostlist coupling removed. The only command form is
# positional:
#   ip address <ip> <mask>                                    (IP+Mask only)
#   ip address <ip> <mask> <gw>                               (+ Gateway)
#   ip address <ip> <mask> <gw> <dns1>                        (+ DNS1)
#   ip address <ip> <mask> <gw> <dns1> <dns2>                 (+ DNS2)
# `from-hostlist` is gone; standalone Ad-hoc identity is provided
# automatically when no `(config)# hostname <Name>` has been run.

function ConvertFrom-SubnetMaskToPrefix {
    param([string]$Mask)
    if ([string]::IsNullOrWhiteSpace($Mask)) { return '' }
    try {
        $octets = $Mask.Split('.')
        if ($octets.Count -ne 4) { return '' }
        $bin = ''
        foreach ($o in $octets) {
            $bin += [Convert]::ToString([int]$o, 2).PadLeft(8, '0')
        }
        return (($bin.ToCharArray() | Where-Object { $_ -eq '1' }).Count)
    } catch {
        return ''
    }
}

function Invoke-IpAddressManual {
    # Ad-hoc IP configuration. All env-var mutations are
    # save/restored in finally; bound-host state survives the call.
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
    if ([string]::IsNullOrWhiteSpace($Ip) -or [string]::IsNullOrWhiteSpace($Mask)) {
        Write-Host "% Usage: 'ip address <ip> <mask> [<gw> [<dns1> [<dns2>]]]'" -ForegroundColor Red
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

    $modulePath = Join-Path $script:FabriqRoot 'modules\standard\ipaddress_config\ipaddress_config.ps1'
    $previousPass = $global:FabriqMasterPassphrase
    if ($State.Passphrase) { $global:FabriqMasterPassphrase = $State.Passphrase }
    try {
        $result = Invoke-FabriqIosModule -ScriptPath $modulePath
    } finally {
        $global:FabriqMasterPassphrase = $previousPass
        $env:SELECTED_NEW_PCNAME  = $prev.NewName
        $env:SELECTED_OLD_PCNAME  = $prev.OldName
        $env:SELECTED_KANRI_NO    = $prev.Kanri
        $env:SELECTED_ETH_IP      = $prev.Ip
        $env:SELECTED_ETH_SUBNET  = $prev.Subnet
        $env:SELECTED_ETH_GATEWAY = $prev.Gateway
        $env:SELECTED_DNS1        = $prev.Dns1
        $env:SELECTED_DNS2        = $prev.Dns2
    }

    if (-not $result) {
        Write-Host "% Module returned no ModuleResult." -ForegroundColor Red
        return
    }

    switch ($result.Status) {
        'Success' {
            $prefix = ConvertFrom-SubnetMaskToPrefix $Mask
            Write-FabriqIosSyslog -Severity 6 -Mnemonic 'IPADDR' -Key 'whispered' `
                -Placeholders @{
                    Ip        = $Ip
                    Prefix    = $prefix
                    Interface = $State.CurrentInterface
                }
            if (-not [string]::IsNullOrWhiteSpace($Gateway)) {
                Write-FabriqIosSyslog -Severity 6 -Mnemonic 'IPADDR' -Key 'gateway' `
                    -Placeholders @{ Gateway = $Gateway }
            }
            $dns = @($Dns1, $Dns2) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
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
