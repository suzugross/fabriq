# Test descriptor (C2 contract, schema 1). Non-shipping test metadata (Deploy-excluded). Category A.
@{
    schema = 1
    module = 'restore_point'
    category = 'A'
    scenarios = @(
        @{
            name = 'apply'; script = 'restore_point.ps1'
            context = 'noninteractive'; winrmSafe = $true; reboot = $false; secrets = $false
            envelope = @{ autopilot = $true; selected = @{}; passphrase = '' }
            fixture = @()
            expect = @{ status = @('Success','Skipped','Partial'); verified = 'any' }
            oracle = @{ type = 'self-verified' }   # Checkpoint-Computer may be throttled (24h) => Skipped
            idempotent = @{ secondRun = 'Skipped' }
            cleanup = 'none'
            notes = 'Enables System Restore + creates a checkpoint. 24h throttle can yield Skipped.'
        }
    )
}
