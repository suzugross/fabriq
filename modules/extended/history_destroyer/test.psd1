# Test descriptor (C2 contract, schema 1). Non-shipping test metadata (Deploy-excluded). Category D.
@{
    schema = 1
    module = 'history_destroyer'
    category = 'D'
    scenarios = @(
        @{
            name = 'apply'; script = 'history_destroyer.ps1'
            context = 'noninteractive'; winrmSafe = $true; reboot = $false; secrets = $false
            envelope = @{ autopilot = $true; selected = @{}; passphrase = '' }
            fixture = @()
            expect = @{ status = @('Success','Skipped','Partial'); verified = 'any' }
            oracle = @{ type = 'self-verified' }
            idempotent = @{ secondRun = 'Success' }
            cleanup = 'snapshot'   # clears event logs / caches / recycle bin etc.; stops Explorer+WSearch -> revert after
            notes = 'Clears event logs, caches, recycle bin, Wi-Fi profiles, etc. + arbitrary CSV Commands. Restarts Explorer/WSearch.'
        }
    )
}
