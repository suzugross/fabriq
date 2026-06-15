# Test descriptor (C2 contract, schema 1). Non-shipping test metadata (Deploy-excluded). Category D.
@{
    schema = 1
    module = 'bloatware_remove'
    category = 'D'
    scenarios = @(
        @{
            name = 'apply'; script = 'bloatware_remove.ps1'
            context = 'noninteractive'; winrmSafe = $true; reboot = $false; secrets = $false
            envelope = @{ autopilot = $true; selected = @{}; passphrase = '' }
            fixture = @()
            expect = @{ status = @('Success','Skipped'); verified = 'any' }
            oracle = @{ type = 'self-verified' }   # re-scans Uninstall\*; C6 can synthesize an absence oracle for the target set
            idempotent = @{ secondRun = 'Skipped' }
            cleanup = 'snapshot'   # uninstalls apps -> revert clean-base-v2 after
            notes = 'Runs target UninstallString + removes uninstall keys. Skipped when targets absent.'
        }
    )
}
