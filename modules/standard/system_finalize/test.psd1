# Test descriptor (C2 contract, schema 1). Non-shipping test metadata (Deploy-excluded). Category A.
@{
    schema = 1
    module = 'system_finalize'
    category = 'A'
    scenarios = @(
        @{
            name = 'apply'; script = 'system_finalize.ps1'
            context = 'noninteractive'; winrmSafe = $true; reboot = $false; secrets = $false
            envelope = @{ autopilot = $true; selected = @{}; passphrase = '' }
            fixture = @()
            expect = @{ status = @('Success','Skipped','Partial'); verified = 'any' }
            oracle = @{ type = 'self-verified' }
            idempotent = @{ secondRun = 'Success' }
            cleanup = 'none'
            notes = 'regsvr32 + icon/thumbnail cache clear + Explorer restart. Not a reboot despite the name.'
        }
    )
}
