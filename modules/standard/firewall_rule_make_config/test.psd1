# Test descriptor (C2 contract, schema 1). Non-shipping test metadata (Deploy-excluded). Category A (+fixture).
@{
    schema = 1
    module = 'firewall_rule_make_config'
    category = 'A'
    scenarios = @(
        @{
            name = 'create'; script = 'firewall_rule_make_config.ps1'
            context = 'noninteractive'; winrmSafe = $true; reboot = $false; secrets = $false
            envelope = @{ autopilot = $true; selected = @{}; passphrase = '' }
            # ships 6 rows all Enabled=0 -> stage a 1-rule test CSV
            fixture = @(
                @{ type = 'stage-asset'; from = 'fixtures\firewall_rule_make_list.test.csv'; to = 'firewall_rule_make_list.csv' }
            )
            expect = @{ status = @('Success'); verified = 'any' }
            oracle = @{ type = 'command'; run = "[bool](Get-NetFirewallRule -DisplayName 'FabriqRigTestRule' -ErrorAction SilentlyContinue)"; equals = 'True' }
            idempotent = @{ secondRun = 'Skipped' }
            cleanup = 'undo'
            teardown = @(
                @{ type = 'expr'; run = "Remove-NetFirewallRule -DisplayName 'FabriqRigTestRule' -ErrorAction SilentlyContinue" }
                @{ type = 'restore-asset'; path = 'firewall_rule_make_list.csv' }
            )
            notes = 'Stages a single inbound TCP/65000 test rule; oracle confirms creation; teardown removes the rule and restores the shipped CSV.'
        }
    )
}
