# Test descriptor (C2 contract, schema 1). Non-shipping test metadata (Deploy-excluded). Category A (interactive/HW).
@{
    schema = 1
    module = 'resolution_api_config'
    category = 'A-interactive'
    scenarios = @(
        @{
            name = 'apply'; script = 'resolution_api_config.ps1'
            context = 'interactive'; winrmSafe = $true; reboot = $false; secrets = $false
            envelope = @{ autopilot = $true; selected = @{}; passphrase = '' }
            fixture = @()
            expect = @{ status = @('Success','Skipped'); verified = 'any' }
            oracle = @{ type = 'self-verified' }   # ChangeDisplaySettings; VM display may pin to a fixed mode
            idempotent = @{ secondRun = 'Skipped' }
            cleanup = 'none'
            notes = 'Display resolution via ChangeDisplaySettings. VM may already match / pin the mode => Skipped.'
        }
    )
}
