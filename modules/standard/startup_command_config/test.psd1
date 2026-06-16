# Test descriptor (C2 contract, schema 1). Non-shipping test metadata (Deploy-excluded). Category A (+fixture).
@{
    schema = 1
    module = 'startup_command_config'
    category = 'A'
    scenarios = @(
        @{
            name = 'apply'; script = 'startup_command_config.ps1'
            context = 'noninteractive'; winrmSafe = $true; reboot = $false; secrets = $false
            envelope = @{ autopilot = $true; selected = @{}; passphrase = '' }
            # ships all rows Enabled=0 -> stage a 1-command CSV to force a deploy
            fixture = @(
                @{ type = 'stage-asset'; from = 'fixtures\startup_command_list.test.csv'; to = 'startup_command_list.csv' }
            )
            expect = @{ status = @('Success'); verified = 'any' }
            oracle = @{ type = 'self-verified' }   # module Test-Paths the 3 deployed artifacts + helper booleans
            idempotent = @{ secondRun = 'Success' }
            cleanup = 'undo'
            teardown = @(
                @{ type = 'restore-asset'; path = 'startup_command_list.csv' }
                @{ type = 'expr'; run = 'Remove-Item (Join-Path $env:ProgramData "fabriqpply_startup_commands.ps1") -Force -ErrorAction SilentlyContinue' }
            )
            notes = 'Stages 1 startup command to force deploy; verifies the 3 deployed artifacts; teardown restores CSV + removes apply script.'
        }
    )
}
