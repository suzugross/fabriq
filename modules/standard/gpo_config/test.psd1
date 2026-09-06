# Test descriptor (C2 contract, schema 1). Non-shipping test metadata (Deploy-excluded).
# Archetype: category A - headless-testable as-is (Registry.pol + gpupdate work over WinRM Session 0).
@{
    schema   = 1
    module   = 'gpo_config'
    category = 'A'
    scenarios = @(
        @{
            name      = 'apply'
            script    = 'gpo_config.ps1'
            context   = 'noninteractive'
            winrmSafe = $true
            reboot    = $false
            secrets   = $false
            envelope  = @{ autopilot = $true; selected = @{}; passphrase = ''; segment = '' }
            fixture   = @()
            expect    = @{ status = @('Success','Skipped'); verified = $true }
            # Independent oracle: the shipped gpo_list.csv disables Automatic Updates via local policy.
            oracle    = @{ type = 'state-query'; query = '(Get-ItemProperty "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" -ErrorAction SilentlyContinue).NoAutoUpdate'; expect = @{ value = '1' } }
            idempotent = @{ secondRun = 'Skipped' }     # rows already present in Registry.pol -> all Skip
            cleanup   = 'none'                          # harmless WU policy on a test VM; reverted at next clean-base
            notes     = 'User-scope row runs gpupdate /target:user inside the WinRM session (FabriqTest HKCU). Dev-machine harness 7/7 (2026-09-05); VM rig PASS incl. idempotency (2026-09-06).'
        },
        @{
            name      = 'backup'
            script    = 'gpo_backup.ps1'
            context   = 'noninteractive'
            winrmSafe = $true
            reboot    = $false
            secrets   = $false
            envelope  = @{ autopilot = $true; selected = @{}; passphrase = ''; segment = '' }
            fixture   = @()
            expect    = @{ status = @('Success','Skipped'); verified = 'any' }
            oracle    = @{ type = 'command'; run = '(Get-ChildItem "C:\fabriq\modules\standard\gpo_config\backup" -Recurse -Filter gpo_list_backup.csv | Measure-Object).Count -ge 1'; equals = 'True' }
            idempotent = $null
            cleanup   = 'none'
            notes     = 'Skipped when the VM has no Registry.pol yet; run after the apply scenario.'
        }
    )
}
