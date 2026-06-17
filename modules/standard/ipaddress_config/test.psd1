# Test descriptor (C2 contract, schema 1). Non-shipping test metadata (Deploy-excluded).
# Category C: mutates the adapter that carries the WinRM transport -> NOT rig-safe.
@{
    schema   = 1
    module   = 'ipaddress_config'
    category = 'C'
    scenarios = @(
        @{
            name      = 'skip-on-exact-match'
            script    = 'ipaddress_config.ps1'
            context   = 'noninteractive'
            winrmSafe = $false   # apply changes the mgmt NIC IP -> drops the rig's own WinRM session
            reboot    = $false
            secrets   = $false
            envelope  = @{ autopilot = $true; selected = @{}; passphrase = '' }
            fixture   = @()
            expect    = @{ status = @('Skipped','Success'); verified = 'any' }
            oracle    = @{ type = 'none'
                           reason = 'IP/GW/DNS are the WinRM transport state; applying on the single mgmt NIC would disconnect the rig. The strict skip predicate (Test-IpConfigMatch) and skip-on-exact-match are validated out-of-band by a read-only predicate dry-run plus a controlled skip-on-match run (target = live config, so no netsh executes, connection stays up). apply->skip full cycle needs a dedicated isolated NIC.' }
            idempotent = @{ secondRun = 'Skipped' }   # re-run with the same target hits the strict match -> Skip
            cleanup   = 'none'
            notes     = 'ENV-driven (SELECTED_ETH_IP/.. ), not CSV. winrmSafe=false: the rig must NOT run the apply path. v1.1.0 adds a strict pre-apply skip (exact IP/prefix/gateway/DNS match, fail-closed -> apply on any doubt). Step 5.5 verification keeps its existing lenient semantics on purpose.'
        }
    )
}
