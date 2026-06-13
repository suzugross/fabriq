# Cisco IOS `reload` for Fabriq IOS: a real reboot wrapped in a
# Surkittinist farewell sequence, with a one-shot post-reboot relaunch
# of Fabriq_IOS.exe via HKLM RunOnce.
#
# Flow: confirm -> arm RunOnce (fail-closed) -> farewell theatre -> reboot.
# IOS is stateless, so there is no resume_state to restore; the shell
# simply relaunches fresh at the next interactive logon. Reached from the
# privileged EXEC `reload` command and from `do reload` in config modes.

function Register-FabriqIosRunOnce {
    # Arm a one-shot relaunch of Fabriq_IOS.exe at the next interactive
    # logon (HKLM RunOnce auto-removes the entry once it fires). The value
    # name is deliberately distinct from the kernel's 'FabriqAutoStart' so
    # an IOS reload never clobbers a kernel resume entry. Requires admin;
    # the IOS launcher runs requireAdministrator. Returns $true on success.
    $exe = Join-Path $script:FabriqRoot 'Fabriq_IOS.exe'
    if (-not (Test-Path $exe)) {
        Show-Error ("Fabriq_IOS.exe not found, cannot arm post-reboot relaunch: {0}" -f $exe)
        return $false
    }
    $runOncePath = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce'
    try {
        if (-not (Test-Path $runOncePath)) {
            New-Item -Path $runOncePath -Force | Out-Null
        }
        New-ItemProperty -Path $runOncePath -Name 'FabriqIosAutoStart' `
            -Value ('"{0}"' -f $exe) -PropertyType String -Force -ErrorAction Stop | Out-Null
        return $true
    } catch {
        Show-Error ("Failed to register RunOnce: {0}" -f $_)
        return $false
    }
}

function Show-FabriqIosReloadTheatre {
    # Stream the farewell sequence: Cisco-style RESTART syslog lines woven
    # with a manifesto quote, paced by short delays. Purely cosmetic; runs
    # only after the operator has confirmed (the confirmation is the commit
    # point, so the sequence is intentionally not interruptible).
    param([int]$StepMs = 700)
    $steps = @(
        @{ Severity = 5; Mnemonic = 'RESTART';   Key = 'reload_requested' }
        @{ Severity = 6; Mnemonic = 'RESTART';   Key = 'saving_dreams' }
        @{ Severity = 4; Mnemonic = 'RESTART';   Key = 'reload_dream' }
        @{ Severity = 7; Mnemonic = 'MANIFESTO'; Key = 'quote' }
        @{ Severity = 5; Mnemonic = 'RESTART';   Key = 'reload_now' }
    )
    foreach ($s in $steps) {
        Write-FabriqIosSyslog -Severity $s.Severity -Mnemonic $s.Mnemonic -Key $s.Key -Placeholders @{}
        Start-Sleep -Milliseconds $StepMs
    }
}

function Invoke-FabriqIosReload {
    # Confirm, arm the relaunch (fail-closed), run the farewell, then reboot.
    param([hashtable]$State)

    # Cisco-style confirmation. Enter or y/yes proceeds; anything else
    # aborts. Guards against an accidental `do reload` rebooting the host.
    $answer = Read-Host 'Proceed with reload? [confirm]'
    if (-not ([string]::IsNullOrEmpty($answer) -or $answer -match '^(y|yes)$')) {
        Write-Host '% Reload aborted.' -ForegroundColor DarkGray
        return
    }

    # Fail-closed: if the shell cannot arm its own return, do NOT reboot -
    # rebooting into a state where IOS never comes back defeats the point.
    if (-not (Register-FabriqIosRunOnce)) {
        Show-Error 'Reload aborted: the shell could not arm its own return.'
        return
    }

    Show-FabriqIosReloadTheatre
    Restart-Computer -Force
}
