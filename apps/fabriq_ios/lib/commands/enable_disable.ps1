# enable / disable command implementations.

function Invoke-Enable {
    param([hashtable]$State)

    if ($State.Mode -ne 'UserExec') {
        Write-Host "% Already in privileged mode."
        return
    }

    $verifyPath = Join-Path $script:FabriqRoot 'kernel\txt\passphrase_verify.txt'
    if (-not (Test-Path $verifyPath)) {
        Write-Host "% Passphrase verification token not found at:"
        Write-Host ("    {0}" -f $verifyPath)
        Write-Host "% Initialise the passphrase via Fabriq Studio first."
        return
    }

    $secure = Read-Host -Prompt 'Passphrase' -AsSecureString
    $bstr   = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
    try {
        $plain = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
    } finally {
        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
    }

    if (Test-MasterPassphrase -Passphrase $plain -VerifyTokenPath $verifyPath) {
        $State.Passphrase = $plain
        Set-ShellMode -State $State -NewMode 'PrivilegedExec'
        Write-FabriqIosSyslog -Severity 5 -Mnemonic 'ENABLE' -Key 'seance_begins' -Placeholders @{}
    } else {
        Write-FabriqIosSyslog -Severity 3 -Mnemonic 'ENABLE' -Key 'passphrase_refused' -Placeholders @{}
    }
}

function Invoke-Disable {
    param([hashtable]$State)
    if ($State.Mode -ne 'PrivilegedExec') {
        Write-Host "% 'disable' is only available in privileged EXEC mode."
        return
    }
    $State.Passphrase = $null
    Set-ShellMode -State $State -NewMode 'UserExec'
    Write-FabriqIosSyslog -Severity 5 -Mnemonic 'DISABLE' -Key 'disrobed' -Placeholders @{}
}
