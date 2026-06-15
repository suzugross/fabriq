# Test descriptor (C2 contract, schema 1). Non-shipping. SCAFFOLD (not a real kitting module).
@{
    schema = 1
    module = 'test_harness_config'
    category = 'A'
    scenarios = @(
        @{
            name = 'simulate'; script = 'test_harness_config.ps1'
            context = 'noninteractive'; winrmSafe = $true; reboot = $false; secrets = $false
            envelope = @{ autopilot = $true; selected = @{}; passphrase = ''; segment = '' }
            fixture = @()
            expect = @{ status = @('Success','Partial','Skipped'); verified = 'any' }   # data-driven mix of simulated statuses
            oracle = @{ type = 'self-verified' }
            idempotent = @{ secondRun = $null }
            cleanup = 'none'
            notes = 'Framework test scaffold (data-driven Status/Verified/ErrorMode simulation). Doubles as a harness self-test.'
        }
    )
}
