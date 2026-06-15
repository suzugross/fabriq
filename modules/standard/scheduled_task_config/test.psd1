# Test descriptor (C2 contract, schema 1). Non-shipping test metadata (Deploy-excluded). Category B.
@{
    schema = 1
    module = 'scheduled_task_config'
    category = 'B'
    scenarios = @(
        @{
            name = 'disable'; script = 'scheduled_task_disable_config.ps1'
            context = 'noninteractive'; winrmSafe = $true; reboot = $false; secrets = $false
            envelope = @{ autopilot = $true; selected = @{}; passphrase = '' }
            fixture = @()   # ships 1 enabled row targeting a built-in task (AppxDeploymentClient\Pre-staged app cleanup)
            expect = @{ status = @('Success'); verified = 'any' }
            # C6: independent read of the target task state (single-quoted to avoid psd1 expansion)
            oracle = @{ type = 'state-query'
                        query = '(Get-ScheduledTask -TaskPath "\Microsoft\Windows\AppxDeploymentClient\" -TaskName "Pre-staged app cleanup" -ErrorAction SilentlyContinue).State'
                        expect = @{ value = 'Disabled' } }
            idempotent = @{ secondRun = 'Skipped' }
            cleanup = 'none'   # toggling a built-in task state is low-impact; reverted at next clean-base
            notes = 'Enables/disables a built-in scheduled task. Skipped if the target task is absent.'
        }
    )
}
