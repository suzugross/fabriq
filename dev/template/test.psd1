# ============================================================================
# Module Test Descriptor (C2 contract, schema 1) - TEMPLATE
# ============================================================================
# Co-located with the module. Tells the test rig (dev/test_rig/run_module_tests.ps1)
# how to verify this module on an isolated VM over WinRM.
#
# NON-SHIPPING test metadata. Does NOT affect the module's behaviour, VERSION,
# REQUIRES_KERNEL, or the kernel public API. Do not bump versions for editing it.
#
# When creating a new module: fill this in alongside the module. If the module is
# structurally NOT verifiable on a lone VM (needs AD / Azure AD / a CA / vendor
# print drivers / internet repos / licensing servers / a GUI session), do NOT
# fake an oracle: set category='C' and oracle type='none' with a reason, or delete
# this file and let the rig report the module as UNCOVERED (honest, not silent).
#
# Run:  powershell.exe -File dev\test_rig\run_module_tests.ps1 -Password <pw> -SyncRepo -Module <name>
# ============================================================================
@{
    schema   = 1
    module   = '<module_name>'              # = the module folder name
    category = 'A'                          # A=headless-ok / B=needs fixture / C=external(not VM) / D=destructive(revert-safe) / E=__RESTART__ straddle

    scenarios = @(
        @{
            name      = 'apply'             # a module may have several (apply/delete/backup/restore/export/import...)
            script    = '<entry>.ps1'       # the entry .ps1 -- take it from module.csv 'Script' column (do NOT assume <name>.ps1)

            context   = 'noninteractive'    # 'noninteractive' = WinRM Session 0 (registry/service/file/network)
                                            # 'interactive'    = needs a desktop session (user32/SPI/display/dpi/credential) -> routed to session-b
            winrmSafe = $true               # $false = would drop the mgmt transport (NIC reconfig / rename+reboot / firewall-off) -> rig SKIPs it
            reboot    = $false              # $true = __RESTART__ straddle (E); rig SKIPs until C7 two-phase support
            secrets   = $false              # $true = consumes ENC: CSV values -> needs a TEST passphrase + test-encrypted CSV (never production secrets)

            envelope  = @{
                autopilot  = $true          # almost always true (suppresses Confirm)
                passphrase = ''             # '' = plaintext-CSV path
                selected   = @{}            # e.g. @{ NEW_PCNAME = 'WIN-TEST01' } -> sets $env:SELECTED_NEW_PCNAME
                segment    = ''             # '' = default segment; set to match a segmented CSV (FABRIQ_SEGMENT). Empty-Segment rows match ''.
            }

            # Precondition setup (declarative). Each step is revert-safe under snapshot.
            #   @{ type='create-files';     paths=@('C:\fabriq_test\scratch\a.txt') }
            #   @{ type='stage-asset';      from='fixtures\xxx.test.csv'; to='xxx.csv' }   # auto-backs up the original to .harnessbak
            #   @{ type='create-localuser'; name='tmpfix' }
            #   @{ type='expr';             run='<single-quoted PowerShell, no $ expansion at parse>' }
            fixture   = @()

            # Acceptable outcome. With an INDEPENDENT oracle, keep status permissive (the oracle is authoritative).
            expect    = @{ status = @('Success','Skipped'); verified = 'any' }   # verified: $true | $false | $null | 'any'

            # How the rig INDEPENDENTLY verifies the resulting state (prefer this over self-verified):
            #   @{ type='registry-csv'; csv='<name>_list*.csv' }                 # compares each KeyPath\KeyName to Value
            #   @{ type='file-exists';  mode='present'|'absent'; paths=@('...') }
            #   @{ type='state-query';  query='<single-quoted expr>'; expect=@{ value='<expected>' } }
            #   @{ type='command';      run='<single-quoted expr>'; equals='True' }   # or match='<regex>'
            #   @{ type='self-verified' }                                        # LOWEST tier: trusts module .Verified (flagged in report). Use only when an independent read is infeasible.
            #   @{ type='none'; reason='...' }                                   # oracle infeasible by design (e.g. credential cmdkey)
            # NOTE: query/run strings MUST be single-quoted so PSD1 does not expand $_ / $true; use double quotes for inner string literals.
            oracle    = @{ type = 'self-verified' }

            idempotent = @{ secondRun = 'Skipped' }    # expected status on a 2nd run (C7); $null = not asserted
            cleanup    = 'none'                         # 'none' = idempotent/isolated, leave / 'undo' = run teardown / 'snapshot' = NOT auto-revertable -> rig flags MANUAL REVERT REQUIRED

            # In-guest undo (only when cleanup='undo'). Mirrors fixture types:
            #   @{ type='delete-files'; paths=@('...') } / @{ type='delete-localuser'; name='...' }
            #   @{ type='reg-delete';   path='HKLM:\...'; name='...' } / @{ type='restore-asset'; path='xxx.csv' }
            #   @{ type='expr';         run='<single-quoted expr>' }
            teardown   = @()

            notes      = ''
        }
    )
}
