# Test descriptor (C2 contract, schema 1). Non-shipping test metadata (Deploy-excluded). Category A.
@{
    schema = 1
    module = 'time_sync_config'
    category = 'A'
    scenarios = @(
        @{
            name = 'apply'; script = 'time_sync_config.ps1'
            context = 'noninteractive'; winrmSafe = $true; reboot = $false; secrets = $false
            envelope = @{ autopilot = $true; selected = @{}; passphrase = '' }
            fixture = @()
            expect = @{ status = @('Success','Skipped','Partial'); verified = 'any' }
            oracle = @{ type = 'command'; run = "(Get-Service W32Time -EA SilentlyContinue).Status.ToString()"; equals = 'Running' }
            idempotent = @{ secondRun = 'Success' }
            cleanup = 'none'
            notes = 'W32Time service + NTP peerlist. Returns Partial if clock still local after retries (NTP reachability optional).'
        }
    )
}
