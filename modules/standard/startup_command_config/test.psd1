# Test descriptor (C2 contract, schema 1). Non-shipping test metadata (Deploy-excluded). Category A.
@{
    schema = 1
    module = 'startup_command_config'
    category = 'A'
    scenarios = @(
        @{
            name = 'apply'; script = 'startup_command_config.ps1'
            context = 'noninteractive'; winrmSafe = $true; reboot = $false; secrets = $false
            envelope = @{ autopilot = $true; selected = @{}; passphrase = '' }
            fixture = @()
            expect = @{ status = @('Success','Skipped'); verified = 'any' }
            oracle = @{ type = 'self-verified' }   # deploys an apply script + Active Setup/Startup trigger; only runs at new-user logon
            idempotent = @{ secondRun = 'Success' }
            cleanup = 'none'
            notes = 'Deploys (does not execute) startup commands to C:\ProgramData\fabriq + registry trigger.'
        }
    )
}
