# Test descriptor (C2 contract, schema 1). Non-shipping test metadata (Deploy-excluded). Category D.
@{
    schema = 1
    module = 'directory_cleaner'
    category = 'D'
    scenarios = @(
        @{
            name = 'apply'; script = 'directory_cleaner.ps1'
            context = 'noninteractive'; winrmSafe = $true; reboot = $false; secrets = $false
            envelope = @{ autopilot = $true; selected = @{}; passphrase = '' }
            fixture = @()
            expect = @{ status = @('Success','Skipped'); verified = 'any' }
            oracle = @{ type = 'self-verified' }   # Test-Path residue per TargetPath; C6 can synthesize an absence oracle
            idempotent = @{ secondRun = 'Skipped' }
            cleanup = 'snapshot'   # recursive delete -> revert clean-base-v2 after (has forbidden-path + depth guards)
            notes = 'Bulk recursive cleanup of TargetPaths. Skipped when targets absent.'
        }
    )
}
