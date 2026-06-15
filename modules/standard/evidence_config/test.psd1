# Test descriptor (C2 contract, schema 1). Non-shipping test metadata. Category A (read-only).
@{
    schema = 1
    module = 'evidence_config'
    category = 'A'
    scenarios = @(
        @{
            name = 'collect'; script = 'evidence_config.ps1'
            context = 'noninteractive'; winrmSafe = $true; reboot = $false; secrets = $false
            envelope = @{ autopilot = $true; selected = @{}; passphrase = ''; segment = '' }
            fixture = @()
            expect = @{ status = @('Success','Partial'); verified = 'any' }   # Partial on a bare VM (WiFi/Credential sections need interactive session)
            oracle = @{ type = 'command'; run = '[bool](Get-ChildItem "C:\fabriq_test\evidence" -Recurse -File -ErrorAction SilentlyContinue)'; equals = 'True' }
            idempotent = @{ secondRun = $null }
            cleanup = 'none'
            notes = 'Read-only evidence collection; oracle confirms report files were written.'
        }
    )
}
