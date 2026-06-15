# ========================================
# Fabriq test VM provisioning (session-b + power) -- run from the dev machine over WinRM
# ========================================
# Idempotent. Sets up everything the test rig needs ON the VM so that, after a revert,
# re-running this restores a test-ready baseline:
#   - never-sleep power settings (a sleeping VM drops WinRM and becomes unreachable)
#   - autologon (recreates the interactive Session 1 after a reboot -> session-b)
#   - the interactive runner + the 'FabriqRigInteractive' scheduled task (session-b)
#
# Typical clean rebuild:  revert to clean-base -> run this -> take a fresh snapshot (with memory).
#
# Usage: powershell.exe -File dev\test_rig\provision_test_vm.ps1 -Password <pw>
# ========================================
param(
    [string]$ComputerName = '10.1.10.8',
    [string]$User = 'FabriqTest',
    [string]$Password = 'P@ssw0rd1251'
)
$ErrorActionPreference = 'Stop'
$cred = New-Object pscredential($User, (ConvertTo-SecureString $Password -AsPlainText -Force))
$runnerSrc = Join-Path $PSScriptRoot '_rig_interactive_runner.ps1'
$s = New-PSSession -ComputerName $ComputerName -Credential $cred
try {
    Invoke-Command -Session $s { New-Item -ItemType Directory -Force 'C:\fabriq_test\rig' | Out-Null }
    Copy-Item -ToSession $s $runnerSrc 'C:\fabriq_test\rig\_rig_interactive_runner.ps1' -Force
    Invoke-Command -Session $s -ArgumentList $User, $Password {
        param($u, $p)

        # 1) never sleep / never turn off display or disk (a sleeping VM = unreachable over WinRM)
        powercfg /change standby-timeout-ac 0
        powercfg /change standby-timeout-dc 0
        powercfg /change monitor-timeout-ac 0
        powercfg /change monitor-timeout-dc 0
        powercfg /change disk-timeout-ac 0
        powercfg /change disk-timeout-dc 0
        powercfg /change hibernate-timeout-ac 0
        powercfg /change hibernate-timeout-dc 0
        powercfg /hibernate off 2>$null

        # 2) autologon -> recreates the interactive Session 1 after a reboot
        $wl = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon'
        Set-ItemProperty $wl AutoAdminLogon '1'
        Set-ItemProperty $wl DefaultUserName $u
        Set-ItemProperty $wl DefaultPassword $p
        Set-ItemProperty $wl DefaultDomainName $env:COMPUTERNAME

        # 3) on-demand scheduled task that runs the runner in the interactive desktop session (session-b)
        $action    = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument '-NoProfile -ExecutionPolicy Bypass -File C:\fabriq_test\rig\_rig_interactive_runner.ps1'
        $principal = New-ScheduledTaskPrincipal -UserId $u -LogonType Interactive -RunLevel Highest
        Register-ScheduledTask -TaskName 'FabriqRigInteractive' -Action $action -Principal $principal -Force | Out-Null

        'provisioned: never-sleep + autologon + session-b task'
    }
}
finally { Remove-PSSession $s }
