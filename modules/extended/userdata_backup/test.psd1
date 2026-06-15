# Test descriptor (C2 contract, schema 1). Non-shipping. Category B (backup variant).
@{
    schema = 1
    module = 'userdata_backup'
    category = 'B'
    scenarios = @(
        @{
            name = 'backup'; script = 'userdata_backup.ps1'
            context = 'noninteractive'; winrmSafe = $true; reboot = $false; secrets = $false
            envelope = @{ autopilot = $true; selected = @{}; passphrase = ''; segment = '' }
            fixture = @()   # robocopy /B of user profile paths (Documents etc.) to backup\<PC>\<ts>\ ; source exists for FabriqTest
            expect = @{ status = @('Success','Skipped','Partial'); verified = 'any' }
            oracle = @{ type = 'self-verified' }   # C6 can synthesize a backup-tree + manifest.json oracle
            idempotent = @{ secondRun = $null }
            cleanup = 'none'
            notes = 'Tests userdata_backup (local file copy). userdata_restore needs the backup. Non-trivial coverage needs a populated profile.'
        }
    )
}
