# ========================================
# Pester v5 unit tests for hostlist_precheck's verdict logic
# ========================================
# Functions: modules/standard/hostlist_precheck/hostlist_precheck.ps1 ::
#            Resolve-PrecheckItemName / Convert-PrecheckMaskToPrefix /
#            Convert-PrecheckPrefixToMask / ConvertTo-PrecheckSavedConfig /
#            Join-PrecheckStaticAddresses / Get-PrecheckItemResult /
#            Get-PrecheckOutcome
# Run      : powershell.exe -File ./dev/run_tests.ps1
#
# The module body runs on import, so the function definitions are lifted out
# with the AST and evaluated on their own (same approach as
# tests/modules/local_user_config). Nothing outside the function bodies runs.
#
# Only the PURE functions are exercised. They receive the hostlist values and
# a snapshot of the PC as arguments, so no test reads this machine's hostname,
# adapters or registry: the result is the same on a developer PC and on CI.
#
# What is pinned is the direction every doubtful case falls in. This module is
# a gate, so the expensive mistake is a false PASS: a value that could not be
# read, an adapter that is missing, an address that is only there by DHCP and
# a misspelled check item must each fail the check, and "nothing compared"
# must never be reported as Success.
# ========================================

BeforeAll {
    . "$PSScriptRoot\..\..\_helpers\test_state.ps1"
    $script:RepoRoot = Get-FabriqRepoRoot
    $script:ModuleDir = Join-Path $script:RepoRoot 'modules\standard\hostlist_precheck'
    $script:ModuleScript = Join-Path $script:ModuleDir 'hostlist_precheck.ps1'

    $tokens = $null
    $parseErrors = $null
    $script:Ast = [System.Management.Automation.Language.Parser]::ParseFile(
        $script:ModuleScript, [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors -and $parseErrors.Count -gt 0) {
        throw "hostlist_precheck.ps1 failed to parse: $($parseErrors[0].Message)"
    }

    $functions = $script:Ast.FindAll(
        { param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] },
        $false)
    foreach ($fn in $functions) {
        . ([scriptblock]::Create($fn.Extent.Text))
    }

    # Every command the module calls, by name (comments and strings excluded).
    $script:CommandNames = @($script:Ast.FindAll(
        { param($node) $node -is [System.Management.Automation.Language.CommandAst] },
        $true) | ForEach-Object { $_.GetCommandName() } | Where-Object { $_ })

    function New-TestAdapter {
        param(
            $Dhcp = 'Disabled',
            $StaticIPv4 = @(@{ IPAddress = '192.168.1.100'; PrefixLength = 24 }),
            $Gateways = @('192.168.1.1'),
            $Dns = @('8.8.8.8', '8.8.4.4'),
            $LiveIPv4 = @('192.168.1.100')
        )
        return @{
            Found = $true; Name = 'Ethernet'; Description = 'Test NIC'; LinkStatus = 'Up'
            Dhcp = $Dhcp; StaticIPv4 = $StaticIPv4; Gateways = $Gateways; Dns = $Dns; LiveIPv4 = $LiveIPv4
        }
    }

    function New-TestLive {
        param(
            $Ethernet = (New-TestAdapter),
            $Wifi = @{ Found = $false },
            [string]$ComputerName = 'PC-0001',
            [string]$Pending = 'PC-0001'
        )
        return @{ ComputerName = $ComputerName; PendingComputerName = $Pending; Ethernet = $Ethernet; Wifi = $Wifi }
    }

    function New-TestExpected {
        param([hashtable]$Override = @{})
        $expected = @{
            NewPCName = 'PC-0001'; OldPCName = 'OLD-0001'
            EthernetIP = '192.168.1.100'; EthernetSubnet = '255.255.255.0'; EthernetGateway = '192.168.1.1'
            WifiIP = ''; WifiSubnet = ''; WifiGateway = ''
            Dns = @('8.8.8.8', '8.8.4.4', '', '')
        }
        foreach ($key in $Override.Keys) { $expected[$key] = $Override[$key] }
        return $expected
    }
}

Describe 'Resolve-PrecheckItemName' {

    It 'returns the canonical spelling for a known item' {
        Resolve-PrecheckItemName -Name 'EthernetDNS' | Should -BeExactly 'EthernetDNS'
    }

    It 'ignores case and surrounding spaces' {
        Resolve-PrecheckItemName -Name '  ethernetdns ' | Should -BeExactly 'EthernetDNS'
        Resolve-PrecheckItemName -Name 'NEWPCNAME' | Should -BeExactly 'NewPCName'
    }

    It 'returns null for a name that is not a check item' {
        Resolve-PrecheckItemName -Name 'EthernetDnss' | Should -Be $null
        Resolve-PrecheckItemName -Name 'DNS' | Should -Be $null
        Resolve-PrecheckItemName -Name 'Pin' | Should -Be $null
    }

    It 'returns null for a blank name' {
        Resolve-PrecheckItemName -Name '' | Should -Be $null
        Resolve-PrecheckItemName -Name '   ' | Should -Be $null
    }
}

Describe 'Convert-PrecheckMaskToPrefix' {

    It 'converts <Mask> to /<Prefix>' -TestCases @(
        @{ Mask = '255.255.255.0';   Prefix = 24 }
        @{ Mask = '255.255.255.192'; Prefix = 26 }
        @{ Mask = '255.255.254.0';   Prefix = 23 }
        @{ Mask = '255.0.0.0';       Prefix = 8 }
        @{ Mask = '255.255.255.255'; Prefix = 32 }
        @{ Mask = '0.0.0.0';         Prefix = 0 }
        @{ Mask = ' 255.255.255.0 '; Prefix = 24 }
    ) {
        param($Mask, $Prefix)
        Convert-PrecheckMaskToPrefix -SubnetMask $Mask | Should -Be $Prefix
    }

    It 'returns null for <Mask> (not a valid mask)' -TestCases @(
        @{ Mask = '255.0.255.0' }
        @{ Mask = '255.255.255' }
        @{ Mask = '255.255.255.0.0' }
        @{ Mask = '256.0.0.0' }
        @{ Mask = '24' }
        @{ Mask = '/24' }
        @{ Mask = 'abc' }
        @{ Mask = '255.255.255.x' }
        @{ Mask = '' }
    ) {
        param($Mask)
        Convert-PrecheckMaskToPrefix -SubnetMask $Mask | Should -Be $null
    }

    It 'keeps 0.0.0.0 apart from an invalid mask (0 is a prefix, not "no answer")' {
        $prefix = Convert-PrecheckMaskToPrefix -SubnetMask '0.0.0.0'
        ($null -eq $prefix) | Should -BeFalse
    }
}

Describe 'Convert-PrecheckPrefixToMask' {

    It 'converts /<Prefix> to <Mask>' -TestCases @(
        @{ Prefix = 24; Mask = '255.255.255.0' }
        @{ Prefix = 26; Mask = '255.255.255.192' }
        @{ Prefix = 0;  Mask = '0.0.0.0' }
        @{ Prefix = 32; Mask = '255.255.255.255' }
    ) {
        param($Prefix, $Mask)
        Convert-PrecheckPrefixToMask -PrefixLength $Prefix | Should -Be $Mask
    }

    It 'returns an empty string for a prefix outside 0-32' {
        Convert-PrecheckPrefixToMask -PrefixLength 33 | Should -Be ''
        Convert-PrecheckPrefixToMask -PrefixLength -1 | Should -Be ''
    }
}

Describe 'ConvertTo-PrecheckSavedConfig' {

    # The input mirrors what Get-ItemProperty returns for
    # Tcpip\Parameters\Interfaces\{GUID}: EnableDHCP is a number, the address
    # values are string arrays and NameServer is one string.

    It 'reads a static adapter' {
        $c = ConvertTo-PrecheckSavedConfig -RegistryValues ([PSCustomObject]@{
            EnableDHCP = 0
            IPAddress = [string[]]@('192.168.1.100'); SubnetMask = [string[]]@('255.255.255.0')
            DefaultGateway = [string[]]@('192.168.1.1'); NameServer = '8.8.8.8,8.8.4.4'
        })
        $c.Dhcp | Should -Be 'Disabled'
        @($c.StaticIPv4).Count | Should -Be 1
        $c.StaticIPv4[0].IPAddress | Should -Be '192.168.1.100'
        $c.StaticIPv4[0].PrefixLength | Should -Be 24
        ($c.Gateways -join '|') | Should -Be '192.168.1.1'
        ($c.StaticDns -join '|') | Should -Be '8.8.8.8|8.8.4.4'
    }

    It 'reads a DHCP adapter as: no static address, no gateway, no static DNS' {
        $c = ConvertTo-PrecheckSavedConfig -RegistryValues ([PSCustomObject]@{
            EnableDHCP = 1
            IPAddress = [string[]]@('0.0.0.0'); SubnetMask = [string[]]@('0.0.0.0')
            DefaultGateway = [string[]]@(); NameServer = ''
        })
        $c.Dhcp | Should -Be 'Enabled'
        ($null -eq $c.StaticIPv4) | Should -BeFalse
        @($c.StaticIPv4).Count | Should -Be 0
        @($c.Gateways).Count | Should -Be 0
        @($c.StaticDns).Count | Should -Be 0
    }

    It 'treats values that are missing altogether as "nothing saved", not as unreadable' {
        $c = ConvertTo-PrecheckSavedConfig -RegistryValues ([PSCustomObject]@{ EnableDHCP = 1 })
        ($null -eq $c.StaticIPv4) | Should -BeFalse
        @($c.StaticIPv4).Count | Should -Be 0
        @($c.Gateways).Count | Should -Be 0
        @($c.StaticDns).Count | Should -Be 0
    }

    It 'leaves Dhcp unknown when EnableDHCP is missing, and then reports no static address' {
        $c = ConvertTo-PrecheckSavedConfig -RegistryValues ([PSCustomObject]@{
            IPAddress = [string[]]@('192.168.1.100'); SubnetMask = [string[]]@('255.255.255.0')
        })
        ($null -eq $c.Dhcp) | Should -BeTrue
        @($c.StaticIPv4).Count | Should -Be 0
    }

    It 'does not report an address left behind in IPAddress as static when EnableDHCP is 1' {
        $c = ConvertTo-PrecheckSavedConfig -RegistryValues ([PSCustomObject]@{
            EnableDHCP = 1
            IPAddress = [string[]]@('192.168.1.100'); SubnetMask = [string[]]@('255.255.255.0')
        })
        $c.Dhcp | Should -Be 'Enabled'
        @($c.StaticIPv4).Count | Should -Be 0
    }

    It 'returns every field as unknown when the key could not be read' {
        $c = ConvertTo-PrecheckSavedConfig -RegistryValues $null
        ($null -eq $c.Dhcp) | Should -BeTrue
        ($null -eq $c.StaticIPv4) | Should -BeTrue
        ($null -eq $c.Gateways) | Should -BeTrue
        ($null -eq $c.StaticDns) | Should -BeTrue
    }

    It 'pairs each address with its own mask' {
        $c = ConvertTo-PrecheckSavedConfig -RegistryValues ([PSCustomObject]@{
            EnableDHCP = 0
            IPAddress = [string[]]@('192.168.1.100', '10.0.0.5'); SubnetMask = [string[]]@('255.255.255.0', '255.0.0.0')
        })
        @($c.StaticIPv4).Count | Should -Be 2
        $c.StaticIPv4[1].IPAddress | Should -Be '10.0.0.5'
        $c.StaticIPv4[1].PrefixLength | Should -Be 8
    }

    It 'keeps the address and marks the prefix unknown when the saved mask is not a mask' {
        $c = ConvertTo-PrecheckSavedConfig -RegistryValues ([PSCustomObject]@{
            EnableDHCP = 0; IPAddress = [string[]]@('192.168.1.100'); SubnetMask = [string[]]@('garbage')
        })
        @($c.StaticIPv4).Count | Should -Be 1
        ($null -eq $c.StaticIPv4[0].PrefixLength) | Should -BeTrue
    }

    It 'splits NameServer on <Name>' -TestCases @(
        @{ Name = 'commas';            Value = '8.8.8.8,8.8.4.4,1.1.1.1' }
        @{ Name = 'spaces';            Value = '8.8.8.8 8.8.4.4 1.1.1.1' }
        @{ Name = 'commas and spaces'; Value = ' 8.8.8.8, 8.8.4.4 ,1.1.1.1 ' }
    ) {
        param($Name, $Value)
        $c = ConvertTo-PrecheckSavedConfig -RegistryValues ([PSCustomObject]@{ EnableDHCP = 0; NameServer = $Value })
        ($c.StaticDns -join '|') | Should -Be '8.8.8.8|8.8.4.4|1.1.1.1'
    }
}

Describe 'Join-PrecheckStaticAddresses' {

    BeforeAll {
        $script:Addr24 = @{ IPAddress = '192.168.1.100'; PrefixLength = 24 }
    }

    It 'is unknown only when BOTH witnesses could not be read' {
        $j = Join-PrecheckStaticAddresses -Saved $null -Live $null
        ($null -eq $j.Addresses) | Should -BeTrue
    }

    It 'is an empty list (not unknown) when both witnesses were read and found nothing' {
        $j = Join-PrecheckStaticAddresses -Saved @() -Live @()
        ($null -eq $j.Addresses) | Should -BeFalse
        @($j.Addresses).Count | Should -Be 0
    }

    It 'is an empty list when one witness failed and the other found nothing' {
        $j = Join-PrecheckStaticAddresses -Saved $null -Live @()
        ($null -eq $j.Addresses) | Should -BeFalse
        @($j.Addresses).Count | Should -Be 0
    }

    It 'accepts an address only the saved settings report' {
        $j = Join-PrecheckStaticAddresses -Saved @($script:Addr24) -Live @()
        @($j.Addresses).Count | Should -Be 1
        $j.SeenIn['192.168.1.100'] | Should -Be 'saved settings'
    }

    It 'accepts an address only the live stack reports (the saved settings say DHCP)' {
        $j = Join-PrecheckStaticAddresses -Saved @() -Live @($script:Addr24)
        @($j.Addresses).Count | Should -Be 1
        $j.Addresses[0].IPAddress | Should -Be '192.168.1.100'
        $j.Addresses[0].PrefixLength | Should -Be 24
        $j.SeenIn['192.168.1.100'] | Should -Be 'live stack'
    }

    It 'accepts an address from the live stack when the saved settings could not be read' {
        $j = Join-PrecheckStaticAddresses -Saved $null -Live @($script:Addr24)
        @($j.Addresses).Count | Should -Be 1
    }

    It 'counts an address both witnesses report once' {
        $j = Join-PrecheckStaticAddresses -Saved @($script:Addr24) -Live @(@{ IPAddress = '192.168.1.100'; PrefixLength = 24 })
        @($j.Addresses).Count | Should -Be 1
        $j.SeenIn['192.168.1.100'] | Should -Be 'saved settings + live stack'
    }

    It 'takes the prefix from the witness that knows it' {
        $j = Join-PrecheckStaticAddresses -Saved @(@{ IPAddress = '192.168.1.100'; PrefixLength = $null }) -Live @($script:Addr24)
        @($j.Addresses).Count | Should -Be 1
        $j.Addresses[0].PrefixLength | Should -Be 24
    }

    It 'keeps two entries when the witnesses disagree on the prefix' {
        $j = Join-PrecheckStaticAddresses -Saved @($script:Addr24) -Live @(@{ IPAddress = '192.168.1.100'; PrefixLength = 16 })
        @($j.Addresses).Count | Should -Be 2
    }

    It 'keeps two entries when the witnesses report different addresses' {
        $j = Join-PrecheckStaticAddresses -Saved @($script:Addr24) -Live @(@{ IPAddress = '10.0.0.5'; PrefixLength = 8 })
        @($j.Addresses).Count | Should -Be 2
    }
}

Describe 'Get-PrecheckItemResult' {

    Context 'Static address proven by the live stack only (regression: judged as DHCP)' {

        # The reported case: fabriq had set a static address, yet the saved
        # DHCP flag did not say "static". The Manual-origin address in the
        # live stack is proof enough on its own.
        BeforeAll {
            $saved = ConvertTo-PrecheckSavedConfig -RegistryValues ([PSCustomObject]@{
                EnableDHCP = 1
                IPAddress = [string[]]@('192.168.1.100'); SubnetMask = [string[]]@('255.255.255.0')
                DefaultGateway = [string[]]@('192.168.1.1'); NameServer = '8.8.8.8,8.8.4.4'
            })
            $joined = Join-PrecheckStaticAddresses -Saved $saved.StaticIPv4 -Live @(@{ IPAddress = '192.168.1.100'; PrefixLength = 24 })
            $script:LiveOnlyLive = New-TestLive -Ethernet @{
                Found = $true; Name = 'Ethernet'; Description = 'Test NIC'; LinkStatus = 'Up'
                Dhcp = $saved.Dhcp; StaticIPv4 = $joined.Addresses; Gateways = $saved.Gateways
                Dns = $saved.StaticDns; LiveIPv4 = @()
            }
        }

        It 'matches <Item>' -TestCases @(
            @{ Item = 'EthernetIP' }
            @{ Item = 'EthernetSubnet' }
            @{ Item = 'EthernetGateway' }
            @{ Item = 'EthernetDNS' }
        ) {
            param($Item)
            $r = Get-PrecheckItemResult -CheckItem $Item -Expected (New-TestExpected) -Live $script:LiveOnlyLive
            $r.State | Should -Be 'Match'
        }
    }

    Context 'Link-down adapter (regression: static address was reported as DHCP)' {

        # The verdict must come from the saved settings alone. A link-down
        # adapter has no live address, and judging DHCP from the live stack
        # reported a statically configured adapter as DHCP.
        BeforeAll {
            $saved = ConvertTo-PrecheckSavedConfig -RegistryValues ([PSCustomObject]@{
                EnableDHCP = 0
                IPAddress = [string[]]@('192.168.1.100'); SubnetMask = [string[]]@('255.255.255.0')
                DefaultGateway = [string[]]@('192.168.1.1'); NameServer = '8.8.8.8,8.8.4.4'
            })
            $script:LinkDownLive = New-TestLive -Ethernet @{
                Found = $true; Name = 'Ethernet'; Description = 'Test NIC'; LinkStatus = 'Disconnected'
                Dhcp = $saved.Dhcp; StaticIPv4 = $saved.StaticIPv4; Gateways = $saved.Gateways
                Dns = $saved.StaticDns; LiveIPv4 = @()
            }
        }

        It 'matches <Item>' -TestCases @(
            @{ Item = 'EthernetIP' }
            @{ Item = 'EthernetSubnet' }
            @{ Item = 'EthernetGateway' }
            @{ Item = 'EthernetDNS' }
        ) {
            param($Item)
            $r = Get-PrecheckItemResult -CheckItem $Item -Expected (New-TestExpected) -Live $script:LinkDownLive
            $r.State | Should -Be 'Match'
        }
    }

    Context 'Hostname' {

        It 'matches NewPCName against the computer name' {
            $r = Get-PrecheckItemResult -CheckItem 'NewPCName' -Expected (New-TestExpected) -Live (New-TestLive)
            $r.State | Should -Be 'Match'
            $r.Expected | Should -Be 'PC-0001'
            $r.Actual | Should -Be 'PC-0001'
        }

        It 'ignores case' {
            $r = Get-PrecheckItemResult -CheckItem 'NewPCName' -Expected (New-TestExpected @{ NewPCName = 'pc-0001' }) -Live (New-TestLive)
            $r.State | Should -Be 'Match'
        }

        It 'reports a mismatch when the name differs' {
            $r = Get-PrecheckItemResult -CheckItem 'NewPCName' -Expected (New-TestExpected @{ NewPCName = 'PC-0002' }) -Live (New-TestLive)
            $r.State | Should -Be 'Mismatch'
            $r.Note | Should -BeNullOrEmpty
        }

        It 'checks OldPCName against the same computer name' {
            $match = Get-PrecheckItemResult -CheckItem 'OldPCName' -Expected (New-TestExpected) -Live (New-TestLive -ComputerName 'OLD-0001' -Pending 'OLD-0001')
            $match.State | Should -Be 'Match'
            $other = Get-PrecheckItemResult -CheckItem 'OldPCName' -Expected (New-TestExpected) -Live (New-TestLive)
            $other.State | Should -Be 'Mismatch'
        }

        It 'is not applicable when the hostlist value is blank' {
            $r = Get-PrecheckItemResult -CheckItem 'NewPCName' -Expected (New-TestExpected @{ NewPCName = '' }) -Live (New-TestLive)
            $r.State | Should -Be 'NotApplicable'
        }

        It 'is unreadable (not a match) when the computer name could not be read' {
            $r = Get-PrecheckItemResult -CheckItem 'NewPCName' -Expected (New-TestExpected) -Live (New-TestLive -ComputerName '')
            $r.State | Should -Be 'Unreadable'
        }

        It 'still fails a rename that is pending, and says so in the note' {
            $live = New-TestLive -ComputerName 'OLD-0001' -Pending 'PC-0001'
            $r = Get-PrecheckItemResult -CheckItem 'NewPCName' -Expected (New-TestExpected) -Live $live
            $r.State | Should -Be 'Mismatch'
            $r.Note | Should -Match 'pending'
        }
    }

    Context 'IP address' {

        It 'matches the single saved static address' {
            $r = Get-PrecheckItemResult -CheckItem 'EthernetIP' -Expected (New-TestExpected) -Live (New-TestLive)
            $r.State | Should -Be 'Match'
            $r.Label | Should -Be 'Ethernet IP address'
        }

        It 'reports a mismatch when the address differs' {
            $r = Get-PrecheckItemResult -CheckItem 'EthernetIP' -Expected (New-TestExpected @{ EthernetIP = '192.168.1.101' }) -Live (New-TestLive)
            $r.State | Should -Be 'Mismatch'
            $r.Actual | Should -Be '192.168.1.100'
        }

        It 'does not accept a DHCP adapter, even when the leased address is the expected one' {
            $adapter = New-TestAdapter -Dhcp 'Enabled' -StaticIPv4 @() -LiveIPv4 @('192.168.1.100')
            $r = Get-PrecheckItemResult -CheckItem 'EthernetIP' -Expected (New-TestExpected) -Live (New-TestLive -Ethernet $adapter)
            $r.State | Should -Be 'Mismatch'
            $r.Actual | Should -Be '(no static IPv4 address)'
            $r.Note | Should -Match 'DHCP'
            $r.Note | Should -Match '192\.168\.1\.100'
        }

        It 'decides from the proven static addresses alone, whatever the saved DHCP flag says' {
            foreach ($flag in @('Enabled', 'Disabled', $null)) {
                $adapter = New-TestAdapter -Dhcp $flag
                $r = Get-PrecheckItemResult -CheckItem 'EthernetIP' -Expected (New-TestExpected) -Live (New-TestLive -Ethernet $adapter)
                $r.State | Should -Be 'Match'
            }
        }

        It 'reports a mismatch when no static address is saved' {
            $adapter = New-TestAdapter -StaticIPv4 @()
            $r = Get-PrecheckItemResult -CheckItem 'EthernetIP' -Expected (New-TestExpected) -Live (New-TestLive -Ethernet $adapter)
            $r.State | Should -Be 'Mismatch'
        }

        It 'reports a mismatch when a second static address is saved next to the expected one' {
            $adapter = New-TestAdapter -StaticIPv4 @(
                @{ IPAddress = '192.168.1.100'; PrefixLength = 24 }
                @{ IPAddress = '10.0.0.5'; PrefixLength = 8 }
            )
            $r = Get-PrecheckItemResult -CheckItem 'EthernetIP' -Expected (New-TestExpected) -Live (New-TestLive -Ethernet $adapter)
            $r.State | Should -Be 'Mismatch'
            $r.Note | Should -Match 'More than one'
        }

        It 'is unreadable when the adapter is missing' {
            $r = Get-PrecheckItemResult -CheckItem 'EthernetIP' -Expected (New-TestExpected) -Live (New-TestLive -Ethernet @{ Found = $false })
            $r.State | Should -Be 'Unreadable'
            $r.Note | Should -Match 'No Ethernet adapter'
        }

        It 'is unreadable when the static addresses could not be read' {
            $adapter = New-TestAdapter -StaticIPv4 $null
            $r = Get-PrecheckItemResult -CheckItem 'EthernetIP' -Expected (New-TestExpected) -Live (New-TestLive -Ethernet $adapter)
            $r.State | Should -Be 'Unreadable'
        }

        It 'is not applicable when the hostlist value is blank, even without an adapter' {
            $r = Get-PrecheckItemResult -CheckItem 'WifiIP' -Expected (New-TestExpected) -Live (New-TestLive)
            $r.State | Should -Be 'NotApplicable'
        }

        It 'fails a Wi-Fi item with a hostlist value when the PC has no Wi-Fi adapter' {
            $r = Get-PrecheckItemResult -CheckItem 'WifiIP' -Expected (New-TestExpected @{ WifiIP = '192.168.2.100' }) -Live (New-TestLive)
            $r.State | Should -Be 'Unreadable'
            $r.Label | Should -Be 'Wi-Fi IP address'
        }
    }

    Context 'Subnet mask' {

        It 'matches the mask against the saved prefix length' {
            $r = Get-PrecheckItemResult -CheckItem 'EthernetSubnet' -Expected (New-TestExpected) -Live (New-TestLive)
            $r.State | Should -Be 'Match'
            $r.Expected | Should -Be '255.255.255.0 (/24)'
            $r.Actual | Should -Be '255.255.255.0 (/24)'
        }

        It 'reports a mismatch when the prefix length differs' {
            $r = Get-PrecheckItemResult -CheckItem 'EthernetSubnet' -Expected (New-TestExpected @{ EthernetSubnet = '255.255.255.192' }) -Live (New-TestLive)
            $r.State | Should -Be 'Mismatch'
        }

        It 'is unreadable when the hostlist mask is not a valid mask' {
            $r = Get-PrecheckItemResult -CheckItem 'EthernetSubnet' -Expected (New-TestExpected @{ EthernetSubnet = '24' }) -Live (New-TestLive)
            $r.State | Should -Be 'Unreadable'
            $r.Note | Should -Match 'not a valid mask'
        }

        It 'is unreadable when the saved mask could not be interpreted, while the IP item still compares' {
            $adapter = New-TestAdapter -StaticIPv4 @(@{ IPAddress = '192.168.1.100'; PrefixLength = $null })
            $live = New-TestLive -Ethernet $adapter
            $subnet = Get-PrecheckItemResult -CheckItem 'EthernetSubnet' -Expected (New-TestExpected) -Live $live
            $subnet.State | Should -Be 'Unreadable'
            $ip = Get-PrecheckItemResult -CheckItem 'EthernetIP' -Expected (New-TestExpected) -Live $live
            $ip.State | Should -Be 'Match'
        }

        It 'is compared on its own when the IP value is blank' {
            $r = Get-PrecheckItemResult -CheckItem 'EthernetSubnet' -Expected (New-TestExpected @{ EthernetIP = '' }) -Live (New-TestLive)
            $r.State | Should -Be 'Match'
        }

        It 'does not accept a DHCP adapter' {
            $adapter = New-TestAdapter -Dhcp 'Enabled' -StaticIPv4 @()
            $r = Get-PrecheckItemResult -CheckItem 'EthernetSubnet' -Expected (New-TestExpected) -Live (New-TestLive -Ethernet $adapter)
            $r.State | Should -Be 'Mismatch'
        }
    }

    Context 'Default gateway' {

        It 'matches the single saved default gateway' {
            $r = Get-PrecheckItemResult -CheckItem 'EthernetGateway' -Expected (New-TestExpected) -Live (New-TestLive)
            $r.State | Should -Be 'Match'
        }

        It 'reports a mismatch when the gateway differs' {
            $r = Get-PrecheckItemResult -CheckItem 'EthernetGateway' -Expected (New-TestExpected @{ EthernetGateway = '192.168.1.254' }) -Live (New-TestLive)
            $r.State | Should -Be 'Mismatch'
            $r.Actual | Should -Be '192.168.1.1'
        }

        It 'reports a mismatch when no default gateway is saved' {
            $adapter = New-TestAdapter -Gateways @()
            $r = Get-PrecheckItemResult -CheckItem 'EthernetGateway' -Expected (New-TestExpected) -Live (New-TestLive -Ethernet $adapter)
            $r.State | Should -Be 'Mismatch'
        }

        It 'reports a mismatch when a second default gateway is saved next to the expected one' {
            $adapter = New-TestAdapter -Gateways @('192.168.1.1', '192.168.1.254')
            $r = Get-PrecheckItemResult -CheckItem 'EthernetGateway' -Expected (New-TestExpected) -Live (New-TestLive -Ethernet $adapter)
            $r.State | Should -Be 'Mismatch'
        }

        It 'is unreadable when the saved routes could not be read' {
            $adapter = New-TestAdapter -Gateways $null
            $r = Get-PrecheckItemResult -CheckItem 'EthernetGateway' -Expected (New-TestExpected) -Live (New-TestLive -Ethernet $adapter)
            $r.State | Should -Be 'Unreadable'
        }
    }

    Context 'DNS servers' {

        It 'matches the same servers in the same order' {
            $r = Get-PrecheckItemResult -CheckItem 'EthernetDNS' -Expected (New-TestExpected) -Live (New-TestLive)
            $r.State | Should -Be 'Match'
            $r.Expected | Should -Be '8.8.8.8, 8.8.4.4'
        }

        It 'reports a mismatch when the order is swapped' {
            $r = Get-PrecheckItemResult -CheckItem 'EthernetDNS' -Expected (New-TestExpected @{ Dns = @('8.8.4.4', '8.8.8.8', '', '') }) -Live (New-TestLive)
            $r.State | Should -Be 'Mismatch'
        }

        It 'reports a mismatch when the PC has an extra server' {
            $adapter = New-TestAdapter -Dns @('8.8.8.8', '8.8.4.4', '1.1.1.1')
            $r = Get-PrecheckItemResult -CheckItem 'EthernetDNS' -Expected (New-TestExpected) -Live (New-TestLive -Ethernet $adapter)
            $r.State | Should -Be 'Mismatch'
        }

        It 'reports a mismatch when the PC has fewer servers' {
            $adapter = New-TestAdapter -Dns @('8.8.8.8')
            $r = Get-PrecheckItemResult -CheckItem 'EthernetDNS' -Expected (New-TestExpected) -Live (New-TestLive -Ethernet $adapter)
            $r.State | Should -Be 'Mismatch'
        }

        It 'reports a mismatch when the PC has no server at all' {
            $adapter = New-TestAdapter -Dns @()
            $r = Get-PrecheckItemResult -CheckItem 'EthernetDNS' -Expected (New-TestExpected) -Live (New-TestLive -Ethernet $adapter)
            $r.State | Should -Be 'Mismatch'
            $r.Actual | Should -Be '(none)'
        }

        It 'skips blank DNS slots in the hostlist (DNS1 and DNS3 filled)' {
            $r = Get-PrecheckItemResult -CheckItem 'EthernetDNS' -Expected (New-TestExpected @{ Dns = @('8.8.8.8', '', '8.8.4.4', '') }) -Live (New-TestLive)
            $r.State | Should -Be 'Match'
        }

        It 'is not applicable when every DNS slot in the hostlist is blank' {
            $r = Get-PrecheckItemResult -CheckItem 'EthernetDNS' -Expected (New-TestExpected @{ Dns = @('', '', '', '') }) -Live (New-TestLive)
            $r.State | Should -Be 'NotApplicable'
        }

        It 'is unreadable when the DNS servers could not be read' {
            $adapter = New-TestAdapter -Dns $null
            $r = Get-PrecheckItemResult -CheckItem 'EthernetDNS' -Expected (New-TestExpected) -Live (New-TestLive -Ethernet $adapter)
            $r.State | Should -Be 'Unreadable'
        }

        It 'judges EthernetDNS on the Ethernet adapter only (a DHCP Wi-Fi with other servers does not matter)' {
            $wifi = New-TestAdapter -Dhcp 'Enabled' -StaticIPv4 @() -Gateways @() -Dns @('10.99.0.1') -LiveIPv4 @('10.99.0.57')
            $live = New-TestLive -Wifi $wifi
            $r = Get-PrecheckItemResult -CheckItem 'EthernetDNS' -Expected (New-TestExpected @{ WifiIP = '192.168.2.100' }) -Live $live
            $r.State | Should -Be 'Match'
        }

        It 'judges WifiDNS on the Wi-Fi adapter only' {
            $wifi = New-TestAdapter -Dhcp 'Enabled' -StaticIPv4 @() -Gateways @() -Dns @('10.99.0.1') -LiveIPv4 @('10.99.0.57')
            $live = New-TestLive -Wifi $wifi
            $r = Get-PrecheckItemResult -CheckItem 'WifiDNS' -Expected (New-TestExpected @{ WifiIP = '192.168.2.100' }) -Live $live
            $r.State | Should -Be 'Mismatch'
            $r.Label | Should -Be 'Wi-Fi DNS servers'
            $r.Actual | Should -Be '10.99.0.1'
        }

        It 'does not compare WifiDNS when the hostlist has no Wi-Fi IP (Ethernet-only hostlist, DHCP Wi-Fi)' {
            $wifi = New-TestAdapter -Dhcp 'Enabled' -StaticIPv4 @() -Gateways @() -Dns @('10.99.0.1') -LiveIPv4 @('10.99.0.57')
            $live = New-TestLive -Wifi $wifi
            $r = Get-PrecheckItemResult -CheckItem 'WifiDNS' -Expected (New-TestExpected) -Live $live
            $r.State | Should -Be 'NotApplicable'
            $r.Note | Should -Match 'no Wi-Fi IP address'
        }

        It 'does not compare WifiDNS without a Wi-Fi IP even when the PC has no Wi-Fi adapter' {
            $r = Get-PrecheckItemResult -CheckItem 'WifiDNS' -Expected (New-TestExpected) -Live (New-TestLive)
            $r.State | Should -Be 'NotApplicable'
        }

        It 'does not compare EthernetDNS when the hostlist has no Ethernet IP' {
            $r = Get-PrecheckItemResult -CheckItem 'EthernetDNS' -Expected (New-TestExpected @{ EthernetIP = '' }) -Live (New-TestLive)
            $r.State | Should -Be 'NotApplicable'
        }
    }

    Context 'Unknown item' {

        It 'is unreadable (a failure), never a match' {
            $r = Get-PrecheckItemResult -CheckItem 'EthernetDnss' -Expected (New-TestExpected) -Live (New-TestLive)
            $r.State | Should -Be 'Unreadable'
        }
    }
}

Describe 'Get-PrecheckOutcome' {

    BeforeAll {
        function New-TestResult {
            param([string]$Item, [string]$State)
            return [PSCustomObject]@{ Item = $Item; Label = $Item; Expected = 'e'; Actual = 'a'; State = $State; Note = '' }
        }
    }

    It 'is Success with Verified true when every compared item matches' {
        $o = Get-PrecheckOutcome -Results @(
            (New-TestResult 'NewPCName' 'Match'), (New-TestResult 'EthernetIP' 'Match'))
        $o.Status | Should -Be 'Success'
        $o.Verified | Should -BeTrue
        $o.Message | Should -Match 'All 2 checked items'
    }

    It 'ignores items that are not applicable when the rest matches' {
        $o = Get-PrecheckOutcome -Results @(
            (New-TestResult 'NewPCName' 'Match'), (New-TestResult 'WifiIP' 'NotApplicable'))
        $o.Status | Should -Be 'Success'
        $o.Message | Should -Match 'All 1 checked items'
    }

    It 'is Error with Verified false when one item does not match (never Partial)' {
        $o = Get-PrecheckOutcome -Results @(
            (New-TestResult 'NewPCName' 'Match'), (New-TestResult 'EthernetIP' 'Mismatch'), (New-TestResult 'EthernetDNS' 'Match'))
        $o.Status | Should -Be 'Error'
        $o.Verified | Should -BeFalse
        ($null -eq $o.Verified) | Should -BeFalse
        $o.Message | Should -Match 'EthernetIP'
        $o.Message | Should -Match '1 of 3 checked'
    }

    It 'is Error when an item is unreadable' {
        $o = Get-PrecheckOutcome -Results @(
            (New-TestResult 'NewPCName' 'Match'), (New-TestResult 'WifiDNS' 'Unreadable'))
        $o.Status | Should -Be 'Error'
        $o.Verified | Should -BeFalse
        $o.Message | Should -Match 'WifiDNS'
    }

    It 'is Error for a state it does not know (fail-closed)' {
        $o = Get-PrecheckOutcome -Results @((New-TestResult 'NewPCName' 'Something'))
        $o.Status | Should -Be 'Error'
    }

    It 'is Skipped with Verified null when nothing was compared' {
        $o = Get-PrecheckOutcome -Results @(
            (New-TestResult 'NewPCName' 'NotApplicable'), (New-TestResult 'EthernetIP' 'NotApplicable'))
        $o.Status | Should -Be 'Skipped'
        ($null -eq $o.Verified) | Should -BeTrue
    }

    It 'is Skipped for an empty result list' {
        $o = Get-PrecheckOutcome -Results @()
        $o.Status | Should -Be 'Skipped'
        ($null -eq $o.Verified) | Should -BeTrue
    }
}

Describe 'hostlist_precheck design pins' {

    It 'never asks for confirmation (a Cancelled result would pass __GATE__)' {
        $script:CommandNames | Should -Not -Contain 'Confirm-ModuleExecution'
        $script:CommandNames | Should -Not -Contain 'Confirm-Execution'
        $script:CommandNames | Should -Not -Contain 'Read-Host'
    }

    It 'never builds its result with New-BatchResult (which can return Partial)' {
        $script:CommandNames | Should -Not -Contain 'New-BatchResult'
        $script:CommandNames | Should -Contain 'New-ModuleResult'
    }

    It 'calls nothing that changes the PC' {
        $writers = @($script:CommandNames | Where-Object {
            $_ -match '^(Set|New|Remove|Rename|Enable|Disable|Restart|Stop|Start|Add|Clear|Copy|Move|Invoke|Export|Out)-' -or
            $_ -match '^(netsh|reg|sc|wmic)(\.exe)?$'
        } | Where-Object { $_ -ne 'New-ModuleResult' } | Sort-Object -Unique)
        $writers | Should -BeNullOrEmpty
    }

    It 'ships a CSV whose check items are exactly the known items, once each' {
        $rows = @(Import-Csv -LiteralPath (Join-Path $script:ModuleDir 'precheck_list.csv'))
        $resolved = @($rows | ForEach-Object { Resolve-PrecheckItemName -Name $_.CheckItem })
        @($resolved | Where-Object { $null -eq $_ }).Count | Should -Be 0
        ($resolved | Sort-Object) -join ',' | Should -Be ((Get-PrecheckKnownItems | Sort-Object) -join ',')
    }

    It 'ships with every Wi-Fi item and OldPCName disabled' {
        $rows = @(Import-Csv -LiteralPath (Join-Path $script:ModuleDir 'precheck_list.csv'))
        $enabled = @($rows | Where-Object { $_.Enabled -eq '1' } | ForEach-Object { $_.CheckItem })
        @($enabled | Where-Object { $_ -like 'Wifi*' }).Count | Should -Be 0
        $enabled | Should -Not -Contain 'OldPCName'
        $enabled | Should -Contain 'EthernetDNS'
    }
}
