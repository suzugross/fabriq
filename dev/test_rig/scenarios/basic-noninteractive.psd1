# Scenario (Layer 2) - ordered modules under one shared envelope. Non-shipping test metadata.
@{
    schema      = 1
    name        = 'basic-noninteractive'
    description = 'reg_hklm -> reg_hkcu -> browser_addon -> time_sync in profile order; combined end-state checked.'
    envelope    = @{ autopilot = $true; passphrase = ''; selected = @{}; segment = '' }
    steps = @(
        @{ module = 'reg_hklm_config';      scenario = 'apply' }
        @{ module = 'reg_hkcu_config';      scenario = 'apply' }
        @{ module = 'browser_addon_config'; scenario = 'apply' }
        @{ module = 'time_sync_config';     scenario = 'apply' }
    )
    # combined end-state after the whole sequence (reg_hklm DisableCAD + time_sync W32Time both took effect)
    finalOracle = @{
        type   = 'command'
        run    = '((Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -Name DisableCAD -ErrorAction SilentlyContinue).DisableCAD -eq 1) -and ((Get-Service W32Time -ErrorAction SilentlyContinue).Status -eq "Running")'
        equals = 'True'
    }
}
