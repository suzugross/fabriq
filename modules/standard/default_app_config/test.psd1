# Test descriptor (C2 contract, schema 1). Non-shipping. Category B.
@{
    schema = 1
    module = 'default_app_config'
    category = 'B'
    scenarios = @(
        @{
            name = 'apply'; script = 'default_app_config.ps1'
            context = 'noninteractive'; winrmSafe = $true; reboot = $false; secrets = $false
            envelope = @{ autopilot = $true; selected = @{}; passphrase = ''; segment = '' }
            fixture = @()   # needs xml\AppAssoc.xml (shipped); effect is deferred to new-user logon
            expect = @{ status = @('Success','Skipped'); verified = 'any' }
            oracle = @{ type = 'self-verified' }   # DISM ingest of default app associations; deferred effect => weak deterministic oracle
            idempotent = @{ secondRun = 'Success' }
            cleanup = 'none'
            notes = 'DISM default app associations. Skipped if xml\AppAssoc.xml absent.'
        }
    )
}
