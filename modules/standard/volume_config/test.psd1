# Test descriptor (C2 contract, schema 1). Non-shipping test metadata (Deploy-excluded). Category A (interactive/HW).
@{
    schema = 1
    module = 'volume_config'
    category = 'A-interactive'
    scenarios = @(
        @{
            name = 'apply'; script = 'volume_config.ps1'
            context = 'interactive'; winrmSafe = $true; reboot = $false; secrets = $false
            envelope = @{ autopilot = $true; selected = @{}; passphrase = '' }
            fixture = @()
            expect = @{ status = @('Success','Skipped'); verified = 'any' }
            oracle = @{ type = 'self-verified' }   # Core Audio COM; no audio endpoint on a bare VM => Skipped
            idempotent = @{ secondRun = 'Skipped' }
            cleanup = 'none'
            notes = 'Master volume/mute via Core Audio. VM may lack an audio device => Skipped.'
        }
    )
}
