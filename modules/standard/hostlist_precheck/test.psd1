# Test descriptor (C2 contract, schema 1). Non-shipping test metadata (Deploy-excluded). Category A (+fixture).
#
# hostlist_precheck is read-only: it changes nothing on the VM, so there is no
# resulting state for an oracle to read back. What is asserted is the verdict
# for a controlled SELECTED_* input, hence oracle type 'none' on every scenario.
#
# SELECTED_* live in the rig's remote session and survive from one scenario to
# the next, so every fixture first clears all of them and then sets only what
# the scenario needs; every teardown clears them again.
#
# The 'match' fixture fills the expectation from the live stack. The module
# accepts a static address proven by the saved settings (registry) or by the
# live stack, and reads the gateway and static DNS from the saved settings.
@{
    schema   = 1
    module   = 'hostlist_precheck'
    category = 'A'

    scenarios = @(
        @{
            name      = 'match'
            script    = 'hostlist_precheck.ps1'
            context   = 'noninteractive'
            winrmSafe = $true
            reboot    = $false
            secrets   = $false
            envelope  = @{ autopilot = $true; selected = @{}; passphrase = '' }
            fixture   = @(
                @{ type = 'stage-asset'; from = 'fixtures\precheck_list.test.csv'; to = 'precheck_list.csv' }
                @{ type = 'expr'; run = '"NEW_PCNAME","OLD_PCNAME","ETH_IP","ETH_SUBNET","ETH_GATEWAY","WIFI_IP","WIFI_SUBNET","WIFI_GATEWAY","DNS1","DNS2","DNS3","DNS4" | ForEach-Object { [Environment]::SetEnvironmentVariable("SELECTED_$_", $null) }' }
                @{ type = 'expr'; run = '$a = Get-NetAdapter -Physical | Where-Object { $_.Status -ne "Disabled" -and $_.InterfaceDescription -notmatch "Wi-Fi|Wireless|WLAN|802\.11|Bluetooth" } | Select-Object -First 1; $ip = Get-NetIPAddress -InterfaceIndex $a.ifIndex -AddressFamily IPv4 | Where-Object { $_.PrefixOrigin -eq "Manual" } | Select-Object -First 1; $m = [uint32]([math]::Pow(2, 32) - [math]::Pow(2, 32 - $ip.PrefixLength)); $env:SELECTED_NEW_PCNAME = $env:COMPUTERNAME; $env:SELECTED_ETH_IP = $ip.IPAddress; $env:SELECTED_ETH_SUBNET = (3..0 | ForEach-Object { ($m -shr (8 * $_)) -band 255 }) -join "."; $env:SELECTED_ETH_GATEWAY = (Get-NetIPConfiguration -InterfaceIndex $a.ifIndex).IPv4DefaultGateway.NextHop; $dns = @((Get-DnsClientServerAddress -InterfaceIndex $a.ifIndex -AddressFamily IPv4).ServerAddresses); for ($i = 0; $i -lt 4 -and $i -lt $dns.Count; $i++) { [Environment]::SetEnvironmentVariable("SELECTED_DNS" + ($i + 1), $dns[$i]) }' }
            )
            expect    = @{ status = @('Success'); verified = $true }
            oracle    = @{ type = 'none'; reason = 'read-only check module: the asserted status for a controlled SELECTED_* input is the verdict; nothing is changed that could be read back' }
            idempotent = @{ secondRun = 'Success' }      # read-only: a second run gives the same verdict
            cleanup   = 'undo'
            teardown  = @(
                @{ type = 'expr'; run = '"NEW_PCNAME","OLD_PCNAME","ETH_IP","ETH_SUBNET","ETH_GATEWAY","WIFI_IP","WIFI_SUBNET","WIFI_GATEWAY","DNS1","DNS2","DNS3","DNS4" | ForEach-Object { [Environment]::SetEnvironmentVariable("SELECTED_$_", $null) }' }
                @{ type = 'restore-asset'; path = 'precheck_list.csv' }
            )
            notes     = 'Needs a static IPv4 on the first physical Ethernet adapter (the rig VM has one). Hostname, IP, subnet, gateway and DNS all compared.'
        }
        @{
            name      = 'mismatch-hostname'
            script    = 'hostlist_precheck.ps1'
            context   = 'noninteractive'
            winrmSafe = $true
            reboot    = $false
            secrets   = $false
            envelope  = @{ autopilot = $true; selected = @{}; passphrase = '' }
            fixture   = @(
                @{ type = 'stage-asset'; from = 'fixtures\precheck_list.test.csv'; to = 'precheck_list.csv' }
                @{ type = 'expr'; run = '"NEW_PCNAME","OLD_PCNAME","ETH_IP","ETH_SUBNET","ETH_GATEWAY","WIFI_IP","WIFI_SUBNET","WIFI_GATEWAY","DNS1","DNS2","DNS3","DNS4" | ForEach-Object { [Environment]::SetEnvironmentVariable("SELECTED_$_", $null) }; $env:SELECTED_NEW_PCNAME = "PRECHECK-NOMATCH"' }
            )
            expect    = @{ status = @('Error'); verified = $false }
            oracle    = @{ type = 'none'; reason = 'read-only check module: the asserted status for a controlled SELECTED_* input is the verdict' }
            idempotent = @{ secondRun = $null }
            cleanup   = 'undo'
            teardown  = @(
                @{ type = 'expr'; run = '"NEW_PCNAME","OLD_PCNAME","ETH_IP","ETH_SUBNET","ETH_GATEWAY","WIFI_IP","WIFI_SUBNET","WIFI_GATEWAY","DNS1","DNS2","DNS3","DNS4" | ForEach-Object { [Environment]::SetEnvironmentVariable("SELECTED_$_", $null) }' }
                @{ type = 'restore-asset'; path = 'precheck_list.csv' }
            )
            notes     = 'Only NewPCName has an expectation, and it is wrong: one compared item, one failure.'
        }
        @{
            name      = 'mismatch-dns'
            script    = 'hostlist_precheck.ps1'
            context   = 'noninteractive'
            winrmSafe = $true
            reboot    = $false
            secrets   = $false
            envelope  = @{ autopilot = $true; selected = @{}; passphrase = '' }
            fixture   = @(
                @{ type = 'stage-asset'; from = 'fixtures\precheck_list.test.csv'; to = 'precheck_list.csv' }
                @{ type = 'expr'; run = '"NEW_PCNAME","OLD_PCNAME","ETH_IP","ETH_SUBNET","ETH_GATEWAY","WIFI_IP","WIFI_SUBNET","WIFI_GATEWAY","DNS1","DNS2","DNS3","DNS4" | ForEach-Object { [Environment]::SetEnvironmentVariable("SELECTED_$_", $null) }; $env:SELECTED_DNS1 = "203.0.113.53"' }
            )
            expect    = @{ status = @('Error'); verified = $false }
            oracle    = @{ type = 'none'; reason = 'read-only check module: the asserted status for a controlled SELECTED_* input is the verdict' }
            idempotent = @{ secondRun = $null }
            cleanup   = 'undo'
            teardown  = @(
                @{ type = 'expr'; run = '"NEW_PCNAME","OLD_PCNAME","ETH_IP","ETH_SUBNET","ETH_GATEWAY","WIFI_IP","WIFI_SUBNET","WIFI_GATEWAY","DNS1","DNS2","DNS3","DNS4" | ForEach-Object { [Environment]::SetEnvironmentVariable("SELECTED_$_", $null) }' }
                @{ type = 'restore-asset'; path = 'precheck_list.csv' }
            )
            notes     = '203.0.113.53 is a documentation address (RFC 5737), so it is never the DNS server of the VM.'
        }
        @{
            name      = 'nothing-to-compare'
            script    = 'hostlist_precheck.ps1'
            context   = 'noninteractive'
            winrmSafe = $true
            reboot    = $false
            secrets   = $false
            envelope  = @{ autopilot = $true; selected = @{}; passphrase = '' }
            fixture   = @(
                @{ type = 'stage-asset'; from = 'fixtures\precheck_list.test.csv'; to = 'precheck_list.csv' }
                @{ type = 'expr'; run = '"NEW_PCNAME","OLD_PCNAME","ETH_IP","ETH_SUBNET","ETH_GATEWAY","WIFI_IP","WIFI_SUBNET","WIFI_GATEWAY","DNS1","DNS2","DNS3","DNS4" | ForEach-Object { [Environment]::SetEnvironmentVariable("SELECTED_$_", $null) }' }
            )
            expect    = @{ status = @('Skipped'); verified = $null }
            oracle    = @{ type = 'none'; reason = 'read-only check module: the asserted status for a controlled SELECTED_* input is the verdict' }
            idempotent = @{ secondRun = 'Skipped' }
            cleanup   = 'undo'
            teardown  = @(
                @{ type = 'restore-asset'; path = 'precheck_list.csv' }
            )
            notes     = 'Every hostlist value is blank: nothing is compared, so the result is Skipped and never Success.'
        }
        @{
            name      = 'unknown-item'
            script    = 'hostlist_precheck.ps1'
            context   = 'noninteractive'
            winrmSafe = $true
            reboot    = $false
            secrets   = $false
            envelope  = @{ autopilot = $true; selected = @{}; passphrase = '' }
            fixture   = @(
                @{ type = 'stage-asset'; from = 'fixtures\precheck_list.unknown.test.csv'; to = 'precheck_list.csv' }
                @{ type = 'expr'; run = '"NEW_PCNAME","OLD_PCNAME","ETH_IP","ETH_SUBNET","ETH_GATEWAY","WIFI_IP","WIFI_SUBNET","WIFI_GATEWAY","DNS1","DNS2","DNS3","DNS4" | ForEach-Object { [Environment]::SetEnvironmentVariable("SELECTED_$_", $null) }; $env:SELECTED_NEW_PCNAME = $env:COMPUTERNAME' }
            )
            expect    = @{ status = @('Error'); verified = $null }
            oracle    = @{ type = 'none'; reason = 'read-only check module: the asserted status for a controlled SELECTED_* input is the verdict' }
            idempotent = @{ secondRun = $null }
            cleanup   = 'undo'
            teardown  = @(
                @{ type = 'expr'; run = '"NEW_PCNAME","OLD_PCNAME","ETH_IP","ETH_SUBNET","ETH_GATEWAY","WIFI_IP","WIFI_SUBNET","WIFI_GATEWAY","DNS1","DNS2","DNS3","DNS4" | ForEach-Object { [Environment]::SetEnvironmentVariable("SELECTED_$_", $null) }' }
                @{ type = 'restore-asset'; path = 'precheck_list.csv' }
            )
            notes     = 'A misspelled CheckItem fails the whole check before anything is compared (Verified stays null), even though the hostname would have matched.'
        }
    )
}
