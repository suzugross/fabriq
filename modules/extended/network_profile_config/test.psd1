# Test descriptor (C2 contract, schema 1). Non-shipping test metadata (Deploy-excluded). Category A.
@{
    schema = 1
    module = 'network_profile_config'
    category = 'A'
    scenarios = @(
        @{
            name = 'apply'; script = 'network_profile_config.ps1'
            context = 'noninteractive'; winrmSafe = $true; reboot = $false; secrets = $false
            envelope = @{ autopilot = $true; selected = @{}; passphrase = '' }
            fixture = @()
            expect = @{ status = @('Success','Skipped'); verified = 'any' }
            oracle = @{ type = 'self-verified' }   # module has its own Step 5.5 read-back of NetworkList Category
            idempotent = @{ secondRun = 'Skipped' }
            cleanup = 'none'
            notes = 'Sets NetworkList\Profiles Category. Skipped if no saved profile matches.'
        }
    )
}
