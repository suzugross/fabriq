# Test descriptor (C2 contract, schema 1). Non-shipping test metadata (Deploy-excluded). Category A (+fixture).
# Non-destructive: encrypts a throwaway 512MB scratch VHD mounted as T: (NOT the OS drive),
# using the non-PIN path (RecoveryPassword only -> no TPM needed), UsedSpaceOnly for speed.
@{
    schema = 1
    module = 'bitlocker_config'
    category = 'A'
    scenarios = @(
        @{
            name = 'apply'; script = 'bitlocker_config.ps1'
            context = 'noninteractive'; winrmSafe = $true; reboot = $false; secrets = $false
            envelope = @{ autopilot = $true; selected = @{}; passphrase = ''; segment = '' }
            # Stage a scratch VHD as T: + an enabled CSV row targeting it (Pin empty -> RecoveryPassword path).
            fixture = @(
                @{ type = 'expr'; run = '$vhd="C:\fabriq_test\rig_bl.vhd"; New-Item -ItemType Directory -Path C:\fabriq_test -Force | Out-Null; if (Test-Path $vhd) { Remove-Item $vhd -Force -EA SilentlyContinue }; $f="$env:TEMP\rig_bl_c.txt"; Set-Content -Path $f -Encoding ASCII -Value @("create vdisk file=$vhd maximum=512 type=expandable","select vdisk file=$vhd","attach vdisk","create partition primary","format fs=ntfs quick label=RIGBL","assign letter=T"); diskpart /s $f | Out-Null; Remove-Item $f -Force -EA SilentlyContinue' }
                @{ type = 'stage-asset'; from = 'fixtures\bitlocker_list.test.csv'; to = 'bitlocker_list.csv' }
            )
            expect = @{ status = @('Success'); verified = 'True' }
            # Independent oracle: T: has a key protector AND conversion is underway/done (no ProtectionStatus reliance).
            oracle = @{ type = 'state-query'
                        query = '$v = Get-BitLockerVolume -MountPoint "T:" -ErrorAction SilentlyContinue; [bool]($v -and $v.KeyProtector.Count -gt 0 -and ($v.VolumeStatus -in @("EncryptionInProgress","FullyEncrypted")))'
                        expect = @{ value = 'True' } }
            cleanup = 'undo'
            teardown = @(
                @{ type = 'expr'; run = 'Disable-BitLocker -MountPoint T: -EA SilentlyContinue | Out-Null; $vhd="C:\fabriq_test\rig_bl.vhd"; $f="$env:TEMP\rig_bl_d.txt"; Set-Content -Path $f -Encoding ASCII -Value @("select vdisk file=$vhd","detach vdisk"); diskpart /s $f | Out-Null; Remove-Item $f -Force -EA SilentlyContinue; Remove-Item $vhd -Force -EA SilentlyContinue' }
                @{ type = 'restore-asset'; path = 'bitlocker_list.csv' }
            )
            notes = 'Non-destructive: diskpart creates a 512MB expandable VHD -> T: (NTFS), module encrypts T: with RecoveryPassword + UsedSpaceOnly (no TPM, no OS drive touched). Independent oracle confirms protector+conversion; -Verified asserts the enable-acceptance signature. Teardown decrypts (best-effort), detaches + deletes the VHD, restores CSV. C: is never touched; VM revert is the ultimate safety net.'
        }
    )
}
