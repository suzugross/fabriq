# Test descriptor (C2 contract, schema 1). Non-shipping test metadata (Deploy-excluded). Category A (export).
@{
    schema = 1
    module = 'firewall_rule_config'
    category = 'A'
    scenarios = @(
        @{
            name = 'export'; script = 'firewall_rule_export.ps1'
            context = 'noninteractive'; winrmSafe = $true; reboot = $false; secrets = $false
            envelope = @{ autopilot = $true; selected = @{}; passphrase = '' }
            fixture = @()   # ships one Enabled=1 Export row (import rows are ack-gated -> Skipped)
            expect = @{ status = @('Success'); verified = 'any' }
            # C6: independent check that a policy.wfw snapshot was written under backup\<timestamp>\
            oracle = @{ type = 'command'
                        run = '[bool](Get-ChildItem "C:\fabriq\modules\standard\firewall_rule_config\backup" -Recurse -Filter policy.wfw -ErrorAction SilentlyContinue)'
                        equals = 'True' }
            idempotent = @{ secondRun = 'Success' }
            cleanup = 'none'   # export appends an Import row to its CSV; -SyncRepo restores the shipped CSV each run
            notes = 'Exports the firewall policy snapshot (.wfw). The import variant is ack-gated and Skips by default (would be D / WinRM-risky).'
        }
    )
}
