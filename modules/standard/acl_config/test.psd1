# Test descriptor (C2 contract, schema 1). Non-shipping. Category B (backup variant tested).
@{
    schema = 1
    module = 'acl_config'
    category = 'B'
    scenarios = @(
        @{
            name = 'backup'; script = 'acl_backup.ps1'
            context = 'noninteractive'; winrmSafe = $true; reboot = $false; secrets = $false
            envelope = @{ autopilot = $true; selected = @{}; passphrase = ''; segment = '' }
            fixture = @()   # backup reads ACLs of CSV target paths and writes to backup\ (self-contained); restore needs that backup
            expect = @{ status = @('Success','Skipped'); verified = 'any' }
            oracle = @{ type = 'self-verified' }   # on the verification-exclusion list (icacls restore false-PASS risk)
            idempotent = @{ secondRun = 'Success' }
            cleanup = 'none'
            notes = 'Tests acl_backup (self-contained). acl_restore needs a backup fixture + is oracle-hard.'
        }
    )
}
