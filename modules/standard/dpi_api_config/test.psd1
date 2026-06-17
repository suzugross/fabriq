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
            notes = 'Live per-monitor DPI via DisplayConfigSetDeviceInfo. NOT VM-verifiable: the test VM basic/virtual display exposes no active DisplayConfig path (measured 2026-06-17: GetMonitorCount=0, GetCurrentDpi(0)=-1, SetDpi returns "Failed to query display configuration"), so apply legitimately Errors and -Verified is False. The verify logic (read back GetCurrentDpi==scale after a Success) mirrors the validated dpi_config module and is exercised only on real display hardware.'
        }
    )
}
