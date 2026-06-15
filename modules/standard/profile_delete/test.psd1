# Test descriptor (C2 contract, schema 1). Non-shipping. Category B.
@{
    schema = 1
    module = 'profile_delete'
    category = 'B'
    scenarios = @(
        @{
            name = 'delete'; script = 'profile_delete.ps1'
            context = 'noninteractive'; winrmSafe = $true; reboot = $false; secrets = $false
            envelope = @{ autopilot = $true; selected = @{}; passphrase = ''; segment = '' }
            fixture = @()   # needs the target user profile to exist (a created local user has no profile until first logon) => usually Skipped
            expect = @{ status = @('Success','Skipped'); verified = 'any' }
            oracle = @{ type = 'self-verified' }   # Test-Path folder + WMI absence; has C:\Users\ path guard
            idempotent = @{ secondRun = 'Skipped' }
            cleanup = 'snapshot'   # deletes a user profile -> revert after
            notes = 'Skipped when the target profile is absent (common on a bare VM).'
        }
    )
}
