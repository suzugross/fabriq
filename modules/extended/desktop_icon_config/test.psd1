# Test descriptor (C2 contract, schema 1). Non-shipping. Category B (interactive, backup variant).
@{
    schema = 1
    module = 'desktop_icon_config'
    category = 'B'
    scenarios = @(
        @{
            name = 'backup'; script = 'desktop_icon_backup.ps1'
            context = 'interactive'; winrmSafe = $true; reboot = $false; secrets = $false
            envelope = @{ autopilot = $true; selected = @{}; passphrase = ''; segment = '' }
            fixture = @()   # reads HKCU Shell\Bags\1\Desktop -> needs a logged-on user (session-b); restore needs the backup
            expect = @{ status = @('Success','Skipped'); verified = 'any' }
            oracle = @{ type = 'self-verified' }
            idempotent = @{ secondRun = 'Success' }
            cleanup = 'none'
            notes = 'Tests desktop_icon_backup in the interactive session (HKCU). Skipped if the Bags key is absent.'
        }
    )
}
