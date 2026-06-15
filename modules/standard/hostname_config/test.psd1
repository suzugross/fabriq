# Test descriptor (C2 contract, schema 1). Non-shipping test metadata (Deploy-excluded).
# Archetype: category E (__RESTART__ straddle) + WinRM-unsafe (host returns under a new name).
@{
    schema   = 1
    module   = 'hostname_config'
    category = 'E'
    scenarios = @(
        @{
            name      = 'rename'
            script    = 'hostname_config.ps1'
            context   = 'noninteractive'
            winrmSafe = $false                           # rename + mandatory reboot; reconnect by IP (10.1.10.8), not by name
            reboot    = $true
            secrets   = $false
            envelope  = @{ autopilot = $true; selected = @{ NEW_PCNAME = 'WIN-TEST01' }; passphrase = '' }
            fixture   = @()
            # Phase 1 (pre-reboot): pending name written to the ComputerName registry.
            expect    = @{ status = @('Success'); verified = $true }
            oracle    = @{ type = 'state-query'
                           query  = "(Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName').ComputerName"
                           expect = @{ value = 'WIN-TEST01' } }
            idempotent = @{ secondRun = 'Skipped' }      # already the target name
            cleanup   = 'snapshot'
            notes     = 'Verify pending name pre-reboot; reconnect by IP, not name. Full convergence (active name post-reboot) spans a reboot -> C7 two-phase.'
        }
    )
}
