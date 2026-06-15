# Test descriptor (C2 contract, schema 1). Non-shipping test metadata (Deploy-excluded). Category A.
@{
    schema = 1
    module = 'taskbar_config'
    category = 'A'
    scenarios = @(
        @{
            name = 'apply'; script = 'taskbar_config.ps1'
            context = 'noninteractive'; winrmSafe = $true; reboot = $false; secrets = $false
            envelope = @{ autopilot = $true; selected = @{}; passphrase = '' }
            fixture = @()
            expect = @{ status = @('Success','Skipped'); verified = 'any' }
            oracle = @{ type = 'self-verified' }   # deploys LayoutModification.xml; C6 can synthesize a file-exists oracle on the deployed path
            idempotent = @{ secondRun = 'Success' }
            cleanup = 'none'
            notes = 'Deploys taskbar LayoutModification.xml to the Default profile + sysprep source.'
        }
    )
}
