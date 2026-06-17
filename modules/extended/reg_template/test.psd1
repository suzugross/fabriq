# Test descriptor (C2 contract, schema 1). Non-shipping test metadata (Deploy-excluded). Category A (+fixture, backup variant).
@{
    schema = 1
    module = 'reg_template'
    category = 'A'
    scenarios = @(
        @{
            name = 'backup'; script = 'reg_backup.ps1'
            context = 'noninteractive'; winrmSafe = $true; reboot = $false; secrets = $false
            envelope = @{ autopilot = $true; selected = @{}; passphrase = ''; segment = '' }
            # Ships rows disabled -> stage an enabled row exporting an always-present key with data,
            # so the export is non-hollow and the verify '[key]' section check passes.
            fixture = @(
                @{ type = 'stage-asset'; from = 'fixtures\reg_list.test.csv'; to = 'reg_list.csv' }
            )
            expect = @{ status = @('Success'); verified = 'True' }   # export file exists + has a [key] section -> Verified True
            oracle = @{ type = 'self-verified' }   # module re-reads the export: file exists AND contains a registry key section
            idempotent = @{ secondRun = 'Success' }   # re-exports (new timestamp file) -> Success again
            cleanup = 'undo'
            teardown = @(
                @{ type = 'restore-asset'; path = 'reg_list.csv' }
                @{ type = 'expr'; run = 'Get-ChildItem (Join-Path $ModuleDirVM "backup") -Filter "*.reg" -EA SilentlyContinue | Remove-Item -Force -EA SilentlyContinue' }
            )
            notes = 'Stages an enabled CSV row exporting HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion (always present, has values -> non-hollow). Module exports to backup\; -Verified asserts file exists AND has a [key] section. Teardown restores CSV + clears exported .reg. reg_import needs a separate .reg fixture (deferred).'
        }
    )
}
