# Test descriptor (C2 contract, schema 1). Non-shipping test metadata (Deploy-excluded). Category A (interactive/HW).
@{
    schema = 1
    module = 'dpi_api_config'
    category = 'A-interactive'
    scenarios = @(
        @{
            name = 'apply'; script = 'dpi_api_config.ps1'
            context = 'interactive'; winrmSafe = $true; reboot = $false; secrets = $false
            envelope = @{ autopilot = $true; selected = @{}; passphrase = '' }
            fixture = @()
            expect = @{ status = @('Success','Skipped'); verified = 'any' }
            oracle = @{ type = 'self-verified' }   # live DisplayConfig DPI P/Invoke; needs an active display path
            idempotent = @{ secondRun = 'Skipped' }
            cleanup = 'none'
            notes = 'Live per-monitor DPI via DisplayConfigSetDeviceInfo. Headless/limited VM display may Skip/error.'
        }
    )
}
