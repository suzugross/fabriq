# Test descriptor (C2 contract, schema 1). Non-shipping test metadata (Deploy-excluded). Category A.
@{
    schema = 1
    module = 'process_killer'
    category = 'A'
    scenarios = @(
        @{
            name = 'apply'; script = 'process_killer.ps1'
            context = 'noninteractive'; winrmSafe = $true; reboot = $false; secrets = $false
            envelope = @{ autopilot = $true; selected = @{}; passphrase = '' }
            fixture = @()
            expect = @{ status = @('Success','Skipped'); verified = 'any' }
            oracle = @{ type = 'self-verified' }   # target processes typically absent on a bare VM => Skipped
            idempotent = @{ secondRun = 'Skipped' }
            cleanup = 'none'
            notes = 'Stops target processes. Skipped when none are running (expected on a bare VM). A future fixture could launch a target process to exercise the kill path.'
        }
    )
}
