# ========================================
# Pester v5 unit tests for Set-SelectedHostEnvironment
# ========================================
# Function: kernel/main.ps1 :: Set-SelectedHostEnvironment
# Run    : pwsh ./dev/run_tests.ps1
#
# Pins the SELECTED_* env-var contract (KERNEL_API.md §3.1) that every
# downstream module reads. Set-SelectedHostEnvironment is the single
# bridge from a hostlist row PSCustomObject to ~47 environment variables
# (3 identity + 3 ethernet + 3 wifi + 4 DNS + 1 PIN + 30 printer slots).
# A regression in any of these mappings - a column rename, an indexing
# off-by-one in the printer loop, a passphrase-less decrypt branch leak -
# silently corrupts every module that reads $env:SELECTED_*, observable
# only at integration time. Pin all four routes here:
#   1. CSV column -> env var name mapping (regression: rename / swap)
#   2. Printer enumeration 1..10 x NAME/DRIVER/PORT (regression: off-by-one)
#   3. Empty/null value passthrough (regression: branch flip)
#   4. ENC: decryption routing (regression: leaking ciphertext when
#      passphrase is unset / silent decrypt without passphrase guard)
#
# main.ps1 is NOT dot-sourceable from tests because it has top-level
# startup side effects (sleep suppression, console size, GUI load,
# exit 1 if no GUI). We extract Set-SelectedHostEnvironment via AST
# (Get-FabriqMainFunctionScriptBlock helper) and dot-source the
# isolated function definition into the test scope.
# ========================================

BeforeAll {
    . "$PSScriptRoot\..\_helpers\test_state.ps1"
    $script:RepoRoot = Get-FabriqRepoRoot
    . (Join-Path $script:RepoRoot 'kernel\common.ps1')

    # Inject Set-SelectedHostEnvironment into the test scope by parsing
    # main.ps1's AST and dot-sourcing only the isolated function body.
    $sb = Get-FabriqMainFunctionScriptBlock -Name 'Set-SelectedHostEnvironment'
    . $sb

    # Save originals for AfterAll restore. SELECTED_* env vars may be
    # populated by a real fabriq session that ran before the test
    # process; do not destroy them.
    $script:_origSelectedEnv = @{}
    Get-ChildItem env: | Where-Object { $_.Name -like 'SELECTED_*' } |
        ForEach-Object { $script:_origSelectedEnv[$_.Name] = $_.Value }
    $script:_origPassphrase = $global:FabriqMasterPassphrase
}

AfterAll {
    Get-ChildItem env: | Where-Object { $_.Name -like 'SELECTED_*' } |
        ForEach-Object { Remove-Item "env:$($_.Name)" -ErrorAction SilentlyContinue }
    foreach ($k in $script:_origSelectedEnv.Keys) {
        Set-Item -Path "env:$k" -Value $script:_origSelectedEnv[$k] -ErrorAction SilentlyContinue
    }
    $global:FabriqMasterPassphrase = $script:_origPassphrase
}

Describe 'Set-SelectedHostEnvironment' {

    BeforeEach {
        # Wipe any SELECTED_* leftovers from prior tests so missing-key
        # assertions (empty env on missing CSV column) are not masked
        # by the previous case's writes.
        Get-ChildItem env: | Where-Object { $_.Name -like 'SELECTED_*' } |
            ForEach-Object { Remove-Item "env:$($_.Name)" -ErrorAction SilentlyContinue }
        $global:FabriqMasterPassphrase = $null
    }

    Context 'Plain identity / network env mapping (KERNEL_API.md §3.1)' {

        It 'maps identity trio (AdminID -> KANRI_NO, OldPCName, NewPCName)' {
            $h = [PSCustomObject]@{
                AdminID    = '42'
                OldPCName  = 'OLD-01'
                NewPCName  = 'NEW-01'
            }
            Set-SelectedHostEnvironment -SelectedHost $h
            $env:SELECTED_KANRI_NO   | Should -Be '42'
            $env:SELECTED_OLD_PCNAME | Should -Be 'OLD-01'
            $env:SELECTED_NEW_PCNAME | Should -Be 'NEW-01'
        }

        It 'maps Ethernet trio (IP / Subnet / Gateway)' {
            $h = [PSCustomObject]@{
                EthernetIP      = '192.168.1.10'
                EthernetSubnet  = '255.255.255.0'
                EthernetGateway = '192.168.1.1'
            }
            Set-SelectedHostEnvironment -SelectedHost $h
            $env:SELECTED_ETH_IP      | Should -Be '192.168.1.10'
            $env:SELECTED_ETH_SUBNET  | Should -Be '255.255.255.0'
            $env:SELECTED_ETH_GATEWAY | Should -Be '192.168.1.1'
        }

        It 'maps Wifi trio (IP / Subnet / Gateway)' {
            $h = [PSCustomObject]@{
                WifiIP      = '10.0.0.5'
                WifiSubnet  = '255.255.0.0'
                WifiGateway = '10.0.0.1'
            }
            Set-SelectedHostEnvironment -SelectedHost $h
            $env:SELECTED_WIFI_IP      | Should -Be '10.0.0.5'
            $env:SELECTED_WIFI_SUBNET  | Should -Be '255.255.0.0'
            $env:SELECTED_WIFI_GATEWAY | Should -Be '10.0.0.1'
        }

        It 'maps all 4 DNS slots (DNS1..DNS4)' {
            $h = [PSCustomObject]@{
                DNS1 = '8.8.8.8'
                DNS2 = '8.8.4.4'
                DNS3 = '1.1.1.1'
                DNS4 = '1.0.0.1'
            }
            Set-SelectedHostEnvironment -SelectedHost $h
            $env:SELECTED_DNS1 | Should -Be '8.8.8.8'
            $env:SELECTED_DNS2 | Should -Be '8.8.4.4'
            $env:SELECTED_DNS3 | Should -Be '1.1.1.1'
            $env:SELECTED_DNS4 | Should -Be '1.0.0.1'
        }

        It 'maps Pin column to SELECTED_PIN' {
            $h = [PSCustomObject]@{ Pin = '1234' }
            Set-SelectedHostEnvironment -SelectedHost $h
            $env:SELECTED_PIN | Should -Be '1234'
        }
    }

    Context 'Printer slot enumeration (1..10 x NAME/DRIVER/PORT)' {

        It 'writes SELECTED_PRINTER_1_* trio (first slot)' {
            $h = [PSCustomObject]@{
                Printer1Name   = 'PRN-A'
                Printer1Driver = 'Driver A'
                Printer1Port   = 'IP_192.168.1.50'
            }
            Set-SelectedHostEnvironment -SelectedHost $h
            $env:SELECTED_PRINTER_1_NAME   | Should -Be 'PRN-A'
            $env:SELECTED_PRINTER_1_DRIVER | Should -Be 'Driver A'
            $env:SELECTED_PRINTER_1_PORT   | Should -Be 'IP_192.168.1.50'
        }

        It 'writes SELECTED_PRINTER_10_* trio (last slot, off-by-one guard)' {
            # Pin the upper bound: the loop is `for ($i = 1; $i -le 10; ...)`
            # so slot 10 must be populated. A regression to `-lt 10` would
            # silently drop the final printer slot.
            $h = [PSCustomObject]@{
                Printer10Name   = 'PRN-J'
                Printer10Driver = 'Driver J'
                Printer10Port   = 'USB001'
            }
            Set-SelectedHostEnvironment -SelectedHost $h
            $env:SELECTED_PRINTER_10_NAME   | Should -Be 'PRN-J'
            $env:SELECTED_PRINTER_10_DRIVER | Should -Be 'Driver J'
            $env:SELECTED_PRINTER_10_PORT   | Should -Be 'USB001'
        }

        It 'writes all 30 printer slot keys (1..10 x 3 suffixes)' {
            # Build a hostlist row populating every printer slot with a
            # value that encodes its slot index and suffix, so any
            # cross-slot or cross-suffix wiring error surfaces.
            $props = [ordered]@{}
            for ($i = 1; $i -le 10; $i++) {
                $props["Printer${i}Name"]   = "name-$i"
                $props["Printer${i}Driver"] = "driver-$i"
                $props["Printer${i}Port"]   = "port-$i"
            }
            $h = [PSCustomObject]$props
            Set-SelectedHostEnvironment -SelectedHost $h

            for ($i = 1; $i -le 10; $i++) {
                (Get-Item -Path "env:SELECTED_PRINTER_${i}_NAME").Value   | Should -Be "name-$i"
                (Get-Item -Path "env:SELECTED_PRINTER_${i}_DRIVER").Value | Should -Be "driver-$i"
                (Get-Item -Path "env:SELECTED_PRINTER_${i}_PORT").Value   | Should -Be "port-$i"
            }
        }

        It 'tolerates missing printer slot columns without throwing' {
            # Only Printer1 fields populated; slots 2..10 absent from
            # the input row. The function uses Set-Item with
            # -ErrorAction SilentlyContinue precisely so the loop can
            # walk all 10 slots regardless of which columns the CSV
            # actually carries. Pin the no-throw + populated-slot
            # behavior; do NOT assert on the unpopulated slot env
            # values, since `Set-Item -Value $null` is a $null-bind
            # error silenced by the production code, leaving slot
            # carryover semantics intentionally unspecified.
            $h = [PSCustomObject]@{ Printer1Name = 'only-one' }
            { Set-SelectedHostEnvironment -SelectedHost $h } | Should -Not -Throw
            $env:SELECTED_PRINTER_1_NAME | Should -Be 'only-one'
        }
    }

    Context 'Empty / null value handling' {

        It 'preserves empty string without attempting decryption' {
            $h = [PSCustomObject]@{ AdminID = '' }
            # Mock guard: empty values must short-circuit before
            # Unprotect-FabriqValue is even considered, regardless of
            # whether a passphrase is set.
            $global:FabriqMasterPassphrase = 'whatever'
            Mock Unprotect-FabriqValue { throw 'should not be called for empty value' }
            Set-SelectedHostEnvironment -SelectedHost $h
            $env:SELECTED_KANRI_NO | Should -BeNullOrEmpty
            Should -Invoke Unprotect-FabriqValue -Times 0 -Exactly
        }

        It 'maps a column missing from the input PSCustomObject to empty env' {
            # PSCustomObject dynamic property access on a missing column
            # returns $null; Resolve-HostValue's IsNullOrEmpty branch
            # passes it through. PowerShell direct env-var assignment
            # of $null unsets the variable (read-back yields $null on
            # PS 5.1, '' on PS Core); BeNullOrEmpty covers both.
            $h = [PSCustomObject]@{ NewPCName = 'NEW-X' }   # AdminID absent
            Set-SelectedHostEnvironment -SelectedHost $h
            $env:SELECTED_NEW_PCNAME | Should -Be 'NEW-X'
            $env:SELECTED_KANRI_NO   | Should -BeNullOrEmpty
        }
    }

    Context 'ENC: decryption routing (KERNEL_API.md §1.5)' {
        # Set-SelectedHostEnvironment owns the routing decision: when
        # to invoke Unprotect-FabriqValue and when to pass the value
        # through. The Unprotect-FabriqValue algorithm itself
        # (PBKDF2 + AES-256-CBC) is pinned by its own test file; here
        # we only verify the routing branches via Mock.

        It 'passes plain (non-ENC:) values straight through even with passphrase set' {
            $global:FabriqMasterPassphrase = 'master-pass'
            Mock Unprotect-FabriqValue { throw 'should not decrypt plain values' }
            $h = [PSCustomObject]@{ AdminID = '42' }
            Set-SelectedHostEnvironment -SelectedHost $h
            $env:SELECTED_KANRI_NO | Should -Be '42'
            Should -Invoke Unprotect-FabriqValue -Times 0 -Exactly
        }

        It 'leaves ENC: values untouched when passphrase is unset (no silent decrypt)' {
            # Critical anti-leak guard: with no passphrase loaded, the
            # function must NOT attempt decryption (which would surface
            # an algorithm error or worse, succeed with a default key).
            # The raw ENC:... ciphertext should pass through to the
            # env var, which is observable but not an information leak.
            $global:FabriqMasterPassphrase = $null
            Mock Unprotect-FabriqValue { throw 'should not decrypt without passphrase' }
            $h = [PSCustomObject]@{ AdminID = 'ENC:dGVzdA==' }
            Set-SelectedHostEnvironment -SelectedHost $h
            $env:SELECTED_KANRI_NO | Should -Be 'ENC:dGVzdA=='
            Should -Invoke Unprotect-FabriqValue -Times 0 -Exactly
        }

        It 'leaves ENC: values untouched when passphrase is whitespace-only' {
            # IsNullOrWhiteSpace guards against accidentally-blank
            # passphrase from a malformed prompt.
            $global:FabriqMasterPassphrase = "   `t"
            Mock Unprotect-FabriqValue { throw 'should not decrypt with blank passphrase' }
            $h = [PSCustomObject]@{ AdminID = 'ENC:abc' }
            Set-SelectedHostEnvironment -SelectedHost $h
            $env:SELECTED_KANRI_NO | Should -Be 'ENC:abc'
            Should -Invoke Unprotect-FabriqValue -Times 0 -Exactly
        }

        It 'invokes Unprotect-FabriqValue when an ENC: value meets a real passphrase' {
            $global:FabriqMasterPassphrase = 'master-pass'
            Mock Unprotect-FabriqValue {
                param($EncryptedValue, $Passphrase)
                # Sanity: the function passed the right passphrase through.
                $Passphrase | Should -Be 'master-pass'
                return 'decrypted-plaintext'
            }
            $h = [PSCustomObject]@{ AdminID = 'ENC:opaque' }
            Set-SelectedHostEnvironment -SelectedHost $h
            $env:SELECTED_KANRI_NO | Should -Be 'decrypted-plaintext'
            Should -Invoke Unprotect-FabriqValue -Times 1 -Exactly
        }

        It 'falls back to the raw ENC: value when decryption throws (graceful degradation)' {
            # Algorithm mismatch / wrong passphrase / corrupted ciphertext
            # must not crash the session - the function logs a warning
            # and the raw ENC: value bleeds into the env var. The module
            # downstream can then notice the leftover ENC: prefix and
            # fail loudly rather than silently configure a host with
            # garbled data.
            $global:FabriqMasterPassphrase = 'master-pass'
            Mock Unprotect-FabriqValue { throw 'simulated decryption failure' }
            $h = [PSCustomObject]@{ AdminID = 'ENC:bad-cipher' }
            Set-SelectedHostEnvironment -SelectedHost $h
            $env:SELECTED_KANRI_NO | Should -Be 'ENC:bad-cipher'
        }
    }

    Context 'Self-referencing mode (__SELF__ -> live PC value)' {
        # A cell value of __SELF__ resolves to THIS PC's live value for the
        # column's semantic, sourced once from Get-CurrentPCInfo. Resolution
        # is column-contextual (the same token means the hostname in a
        # PC-name column and the live IP in an IP column). Unresolvable
        # (no self-source column, or blank live value) -> empty + warning.
        # Get-CurrentPCInfo is mocked so these tests are hardware-independent.

        BeforeEach {
            Mock Get-CurrentPCInfo {
                @{
                    ComputerName    = 'SELF-PC'
                    EthernetIP      = '10.1.1.10'
                    EthernetSubnet  = '255.255.255.0'
                    EthernetGateway = '10.1.1.1'
                    WifiIP          = '10.2.2.20'
                    WifiSubnet      = '255.255.0.0'
                    WifiGateway     = '10.2.2.1'
                    DNS             = @('9.9.9.9', '149.112.112.112')
                    Printers        = @()
                }
            }
        }

        It 'resolves __SELF__ in OldPCName/NewPCName to the live computer name' {
            $h = [PSCustomObject]@{ OldPCName = '__SELF__'; NewPCName = '__SELF__' }
            Set-SelectedHostEnvironment -SelectedHost $h
            $env:SELECTED_OLD_PCNAME | Should -Be 'SELF-PC'
            $env:SELECTED_NEW_PCNAME | Should -Be 'SELF-PC'
        }

        It 'resolves __SELF__ in the Ethernet trio to live adapter values' {
            $h = [PSCustomObject]@{ EthernetIP = '__SELF__'; EthernetSubnet = '__SELF__'; EthernetGateway = '__SELF__' }
            Set-SelectedHostEnvironment -SelectedHost $h
            $env:SELECTED_ETH_IP      | Should -Be '10.1.1.10'
            $env:SELECTED_ETH_SUBNET  | Should -Be '255.255.255.0'
            $env:SELECTED_ETH_GATEWAY | Should -Be '10.1.1.1'
        }

        It 'resolves __SELF__ in the Wifi trio to live adapter values' {
            $h = [PSCustomObject]@{ WifiIP = '__SELF__'; WifiSubnet = '__SELF__'; WifiGateway = '__SELF__' }
            Set-SelectedHostEnvironment -SelectedHost $h
            $env:SELECTED_WIFI_IP      | Should -Be '10.2.2.20'
            $env:SELECTED_WIFI_SUBNET  | Should -Be '255.255.0.0'
            $env:SELECTED_WIFI_GATEWAY | Should -Be '10.2.2.1'
        }

        It 'maps __SELF__ DNS slots to live DNS servers by index' {
            $h = [PSCustomObject]@{ DNS1 = '__SELF__'; DNS2 = '__SELF__' }
            Set-SelectedHostEnvironment -SelectedHost $h
            $env:SELECTED_DNS1 | Should -Be '9.9.9.9'
            $env:SELECTED_DNS2 | Should -Be '149.112.112.112'
        }

        It 'leaves a DNS slot empty (with warning) when fewer live DNS servers exist than the slot index' {
            Mock Show-Warning {}
            $h = [PSCustomObject]@{ DNS3 = '__SELF__' }   # only 2 live DNS servers in the mock
            Set-SelectedHostEnvironment -SelectedHost $h
            $env:SELECTED_DNS3 | Should -BeNullOrEmpty
            Should -Invoke Show-Warning -Times 1 -Exactly
        }

        It 'leaves an unresolvable column empty (with warning) when the live value is blank (e.g. no Wi-Fi)' {
            Mock Get-CurrentPCInfo {
                @{ ComputerName = 'SELF-PC'; EthernetIP = '10.1.1.10'; EthernetSubnet = ''; EthernetGateway = '';
                   WifiIP = ''; WifiSubnet = ''; WifiGateway = ''; DNS = @(); Printers = @() }
            }
            Mock Show-Warning {}
            $h = [PSCustomObject]@{ WifiIP = '__SELF__' }
            Set-SelectedHostEnvironment -SelectedHost $h
            $env:SELECTED_WIFI_IP | Should -BeNullOrEmpty
            Should -Invoke Show-Warning -Times 1 -Exactly
        }

        It 'leaves __SELF__ empty (with warning) for a column with no self-source (PIN)' {
            Mock Show-Warning {}
            $h = [PSCustomObject]@{ Pin = '__SELF__' }
            Set-SelectedHostEnvironment -SelectedHost $h
            $env:SELECTED_PIN | Should -BeNullOrEmpty
            Should -Invoke Show-Warning -Times 1 -Exactly
        }

        It 'does NOT query live PC info when no __SELF__ token is present (zero overhead for plain rows)' {
            $h = [PSCustomObject]@{ OldPCName = 'OLD-01'; EthernetIP = '192.168.1.10' }
            Set-SelectedHostEnvironment -SelectedHost $h
            $env:SELECTED_OLD_PCNAME | Should -Be 'OLD-01'
            Should -Invoke Get-CurrentPCInfo -Times 0 -Exactly
        }

        It 'queries live PC info exactly once even with multiple __SELF__ tokens (single cached snapshot)' {
            $h = [PSCustomObject]@{ OldPCName = '__SELF__'; NewPCName = '__SELF__'; EthernetIP = '__SELF__'; DNS1 = '__SELF__' }
            Set-SelectedHostEnvironment -SelectedHost $h
            Should -Invoke Get-CurrentPCInfo -Times 1 -Exactly
        }

        It 'resolves __SELF__ and ENC: independently in the same row (no cross-interference)' {
            $global:FabriqMasterPassphrase = 'master-pass'
            Mock Unprotect-FabriqValue { return 'DECRYPTED' }
            $h = [PSCustomObject]@{ OldPCName = '__SELF__'; NewPCName = 'ENC:opaque' }
            Set-SelectedHostEnvironment -SelectedHost $h
            $env:SELECTED_OLD_PCNAME | Should -Be 'SELF-PC'
            $env:SELECTED_NEW_PCNAME | Should -Be 'DECRYPTED'
        }
    }
}
