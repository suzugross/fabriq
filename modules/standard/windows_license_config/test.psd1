# Test descriptor (C2 contract, schema 1). Non-shipping test metadata (Deploy-excluded).
#
# Category C: installing and activating a product key needs a real key and an
# activation service, and the OEM script needs a firmware key that no virtual
# machine has. Only the one path a lone VM can reach is described here: the
# OEM script on a PC WITHOUT a firmware key.
#
# NOT covered (no scenario on purpose - do not fake one):
#   windows_license_install.ps1      needs a real product key
#   windows_license_auth.ps1         needs an activation service
#   windows_license_oem_install.ps1  Success / Skipped paths need a physical
#                                    OEM PC; its decision logic is unit-tested
#                                    in tests/modules/windows_license_config
@{
    schema   = 1
    module   = 'windows_license_config'
    category = 'C'

    scenarios = @(
        @{
            name      = 'oem-install-no-firmware-key'
            script    = 'windows_license_oem_install.ps1'
            context   = 'noninteractive'
            winrmSafe = $true
            reboot    = $false
            secrets   = $false
            envelope  = @{ autopilot = $true; selected = @{}; passphrase = '' }
            fixture   = @()
            expect    = @{ status = @('Error'); verified = $null }
            # Independent read of the precondition: with no firmware key, Error is the
            # correct verdict and the script must have stopped before InstallProductKey.
            oracle    = @{ type = 'command'; run = '[string]::IsNullOrWhiteSpace((Get-CimInstance -ClassName SoftwareLicensingService).OA3xOriginalProductKey)'; equals = 'True' }
            idempotent = @{ secondRun = $null }      # the rig asserts a 2nd run only after Success / Skipped / Partial
            cleanup   = 'none'
            notes     = 'The rig VM has no firmware key, so the script must end in Error without installing anything and without falling back to license_key.csv or manual input.'
        }
    )
}
