# Test descriptor (C2 contract, schema 1). Non-shipping. Category B.
@{
    schema = 1
    module = 'robocopy_config'
    category = 'B'
    scenarios = @(
        @{
            name = 'copy'; script = 'robocopy_config.ps1'
            context = 'noninteractive'; winrmSafe = $true; reboot = $false; secrets = $false
            envelope = @{ autopilot = $true; selected = @{}; passphrase = ''; segment = '' }
            fixture = @()   # local-source rows copy; UNC rows need a share. Skipped if source absent.
            expect = @{ status = @('Success','Skipped'); verified = 'any' }
            oracle = @{ type = 'self-verified' }   # C6 can synthesize a dest file-exists oracle for a staged local source
            idempotent = @{ secondRun = 'Success' }
            cleanup = 'none'
            notes = 'robocopy copy/mirror. Has empty-source + protected-path guards. Skipped if source missing.'
        }
    )
}
