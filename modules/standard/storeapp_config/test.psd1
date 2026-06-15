# Test descriptor (C2 contract, schema 1). Non-shipping test metadata (Deploy-excluded). Category A/D.
@{
    schema = 1
    module = 'storeapp_config'
    category = 'A'
    scenarios = @(
        @{
            name = 'apply'; script = 'storeapp_config.ps1'
            context = 'noninteractive'; winrmSafe = $true; reboot = $false; secrets = $false
            envelope = @{ autopilot = $true; selected = @{}; passphrase = '' }
            fixture = @()
            expect = @{ status = @('Success','Skipped'); verified = 'any' }
            oracle = @{ type = 'self-verified' }   # C6 can synthesize a Get-AppxPackage absence oracle for the removed set
            idempotent = @{ secondRun = 'Skipped' }
            cleanup = 'snapshot'   # removes provisioned Store apps -> not in-guest undoable; MANUAL REVERT after
            notes = 'Removes ~27 Store apps (current-user + provisioned). Destructive; revert clean-base-v2 after.'
        }
    )
}
