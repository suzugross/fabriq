# Test descriptor (C2 contract, schema 1). Non-shipping test metadata (Deploy-excluded).
# Archetype: category A - headless-testable as-is.
@{
    schema   = 1
    module   = 'reg_hklm_config'
    category = 'A'
    scenarios = @(
        @{
            name      = 'apply'
            script    = 'reg_hklm_config.ps1'
            context   = 'noninteractive'                 # noninteractive WinRM is sufficient
            winrmSafe = $true
            reboot    = $false
            secrets   = $false
            envelope  = @{ autopilot = $true; selected = @{}; passphrase = '' }
            fixture   = @()                              # no precondition needed
            expect    = @{ status = @('Success'); verified = $true }
            oracle    = @{ type = 'registry-csv'; csv = 'reg_hklm_list*.csv' }
            idempotent = @{ secondRun = 'Success' }      # FORCE_OVERWRITE reapplies deterministically
            cleanup   = 'none'                           # idempotent + isolated HKLM policy keys; harmless to leave between tests (reverted at next clean-base)
            notes     = 'PoC-proven 6/6 (2026-06-15); independent registry oracle confirmed.'
        }
        @{
            # Profile data overlay Phase 2 / W2 (plan section 12.3): the multi-CSV
            # ENUMERATION is module-level all-or-nothing. The PDF holds one
            # reg_hklm_list*.csv writing a unique marker; the shipped CSV writes
            # DisableCAD. The oracle asserts marker present AND DisableCAD absent
            # -> the shipped file was not merged in.
            name      = 'overlay-enumeration'
            script    = 'reg_hklm_config.ps1'
            context   = 'noninteractive'
            winrmSafe = $true
            reboot    = $false
            secrets   = $false
            envelope  = @{ autopilot = $true; selected = @{}; passphrase = ''; segment = ''; profileDataDir = 'C:\fabriq\profiles\_test_p2' }
            fixture   = @(
                @{ type = 'expr'; run = 'New-Item -ItemType Directory -Force -Path "C:\fabriq\profiles\_test_p2\modules\reg_hklm_config" | Out-Null; [System.IO.File]::WriteAllLines("C:\fabriq\profiles\_test_p2\modules\reg_hklm_config\reg_hklm_list_overlay.csv", @("Enabled,AdminID,SettingTitle,KeyPath,KeyName,Type,Value,Segment", "1,1,W2 overlay marker,HKEY_LOCAL_MACHINE\SOFTWARE\FabriqRigOverlay,W2Marker,REG_DWORD,42,"), [System.Text.Encoding]::ASCII); Remove-Item "HKLM:\SOFTWARE\FabriqRigOverlay" -Recurse -Force -EA SilentlyContinue; Remove-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -Name "DisableCAD" -Force -EA SilentlyContinue' }
            )
            expect    = @{ status = @('Success'); verified = 'any' }
            oracle    = @{ type = 'command'; run = '$m = (Get-ItemProperty "HKLM:\SOFTWARE\FabriqRigOverlay" -Name W2Marker -EA SilentlyContinue).W2Marker; $s = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -Name DisableCAD -EA SilentlyContinue); "$($m -eq 42)/$($null -eq $s)"'; equals = 'True/True' }
            idempotent = @{ secondRun = 'Success' }
            cleanup   = 'undo'
            teardown  = @(
                @{ type = 'expr'; run = 'Remove-Item "HKLM:\SOFTWARE\FabriqRigOverlay" -Recurse -Force -EA SilentlyContinue' }
                @{ type = 'expr'; run = 'if (Test-Path "C:\fabriq\profiles\_test_p2") { Remove-Item "C:\fabriq\profiles\_test_p2" -Recurse -Force -EA SilentlyContinue }' }
            )
            notes     = 'W2 enumeration overlay: PDF-only marker applied AND shipped-only DisableCAD absent = no merge.'
        }
    )
}
