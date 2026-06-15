# Test descriptor (C2 contract, schema 1). Non-shipping. SCAFFOLD (not a real kitting module).
@{
    schema = 1
    module = 'test_error_module'
    category = 'A'
    scenarios = @(
        @{
            name = 'simulate-error'; script = 'test_error.ps1'
            context = 'noninteractive'; winrmSafe = $true; reboot = $false; secrets = $false
            envelope = @{ autopilot = $true; selected = @{}; passphrase = ''; segment = '' }
            fixture = @()
            expect = @{ status = @('Error'); verified = 'any' }   # by design returns Error to exercise AutoPilot ErrorMode
            oracle = @{ type = 'none'; reason = 'scaffold: success == deliberately returning Error; nothing to read back' }
            idempotent = @{ secondRun = 'Error' }
            cleanup = 'none'
            notes = 'Framework test scaffold. Doubles as a harness self-test (proves the rig correctly handles an Error result).'
        }
    )
}
