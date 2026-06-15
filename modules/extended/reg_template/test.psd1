# Test descriptor (C2 contract, schema 1). Non-shipping. Category B (backup variant).
@{
    schema = 1
    module = 'reg_template'
    category = 'B'
    scenarios = @(
        @{
            name = 'backup'; script = 'reg_backup.ps1'
            context = 'noninteractive'; winrmSafe = $true; reboot = $false; secrets = $false
            envelope = @{ autopilot = $true; selected = @{}; passphrase = ''; segment = '' }
            fixture = @()   # exports the CSV-listed registry keys to .reg (self-contained); import needs that .reg
            expect = @{ status = @('Success','Skipped'); verified = 'any' }
            oracle = @{ type = 'self-verified' }   # C6 can synthesize a file-exists oracle on the backup .reg
            idempotent = @{ secondRun = 'Success' }
            cleanup = 'none'
            notes = 'Tests reg_backup (export). reg_import needs a .reg fixture.'
        }
    )
}
