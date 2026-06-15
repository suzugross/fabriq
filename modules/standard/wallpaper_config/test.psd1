# Test descriptor (C2 contract, schema 1). Non-shipping. Category B.
@{
    schema = 1
    module = 'wallpaper_config'
    category = 'B'
    scenarios = @(
        @{
            name = 'apply'; script = 'wallpaper_config.ps1'
            context = 'noninteractive'; winrmSafe = $true; reboot = $false; secrets = $false
            envelope = @{ autopilot = $true; selected = @{}; passphrase = ''; segment = '' }
            fixture = @()   # needs an image under wallpaper\ (shipped)
            expect = @{ status = @('Success','Skipped'); verified = 'any' }
            oracle = @{ type = 'self-verified' }   # registry + image copy + Active Setup; C6 can synthesize a registry/file oracle
            idempotent = @{ secondRun = 'Success' }
            cleanup = 'none'
            notes = 'Desktop wallpaper registry + image copy + Spotlight disable. Skipped if no image.'
        }
    )
}
