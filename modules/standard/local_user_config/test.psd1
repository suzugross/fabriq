# Test descriptor (C2 contract, schema 1). Non-shipping test metadata (Deploy-excluded). Category A.
@{
    schema = 1
    module = 'local_user_config'
    category = 'A'
    scenarios = @(
        @{
            name = 'apply'; script = 'local_user_config.ps1'
            context = 'noninteractive'; winrmSafe = $true; reboot = $false; secrets = $false
            envelope = @{ autopilot = $true; selected = @{}; passphrase = '' }
            fixture = @()   # ships 3 enabled sample users: admin_user / standard_user / remote_user
            expect = @{ status = @('Success','Skipped'); verified = 'any' }
            oracle = @{ type = 'command'; run = "[bool](Get-LocalUser -Name admin_user -ErrorAction SilentlyContinue)"; equals = 'True' }
            idempotent = @{ secondRun = 'Skipped' }
            cleanup = 'undo'
            teardown = @(
                @{ type = 'delete-localuser'; name = 'admin_user' }
                @{ type = 'delete-localuser'; name = 'standard_user' }
                @{ type = 'delete-localuser'; name = 'remote_user' }
            )
            notes = 'Creates sample local users (plaintext pw in CSV); oracle confirms admin_user exists; teardown removes all three.'
        }
    )
}
