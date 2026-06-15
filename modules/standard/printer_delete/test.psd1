# Test descriptor (C2 contract, schema 1). Non-shipping test metadata (Deploy-excluded). Category A.
@{
    schema = 1
    module = 'printer_delete'
    category = 'A'
    scenarios = @(
        @{
            name = 'apply'; script = 'printer_delete.ps1'
            context = 'noninteractive'; winrmSafe = $true; reboot = $false; secrets = $false
            envelope = @{ autopilot = $true; selected = @{}; passphrase = '' }
            fixture = @()
            expect = @{ status = @('Success','Skipped'); verified = 'any' }
            oracle = @{ type = 'self-verified' }   # Get-Printer read-back; targets may be absent on a bare VM => Skipped
            idempotent = @{ secondRun = 'Skipped' }
            cleanup = 'snapshot'   # may remove printers (e.g. Print to PDF) -> revert clean-base-v2 after
            notes = 'Removes printers per KeepList/Explicit/Manual mode. Skipped when targets absent.'
        }
    )
}
