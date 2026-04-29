# Local re-implementation of a minimal kernel module dispatcher.
# fabriq_ios cannot dot-source kernel/main.ps1 (it would launch a
# second fabriq_operator GUI inside our isolated subprocess), so
# Invoke-KittingScript is reproduced here in stripped-down form.
# Coupling acknowledged: KERNEL_API.md section 6 (internal) plus the
# ModuleResult contract from section 5. Re-validate after any kernel
# bump that touches main.ps1.

function Invoke-FabriqIosModule {
    param([string]$ScriptPath)

    if (-not (Test-Path $ScriptPath)) {
        Write-Host ("% Module script not found: {0}" -f $ScriptPath) -ForegroundColor Red
        return $null
    }

    $name = Split-Path $ScriptPath -Leaf
    Write-Host ""
    Write-Host ("--- Dispatching: {0} ---" -f $name) -ForegroundColor DarkGray
    Write-Host ""

    $global:_LastModuleResult = $null
    $output = $null
    try {
        $output = & $ScriptPath
    } catch {
        Write-Host ""
        Write-Host ("% Module raised an exception: {0}" -f $_.Exception.Message) -ForegroundColor Red
        return [pscustomobject]@{
            _IsModuleResult = $true
            Status          = 'Error'
            Message         = $_.Exception.Message
            Details         = @()
            Verified        = $null
            Timestamp       = (Get-Date)
        }
    }

    # Match Invoke-KittingScript's ModuleResult discovery: scan the
    # pipeline output for a tagged PSCustomObject, then fall back to
    # the global captured by New-ModuleResult.
    $moduleResult = $null
    if ($null -ne $output) {
        foreach ($item in @($output)) {
            if ($item -is [PSCustomObject] -and $item._IsModuleResult -eq $true) {
                $moduleResult = $item
            }
        }
    }
    if (-not $moduleResult -and $null -ne $global:_LastModuleResult) {
        $moduleResult = $global:_LastModuleResult
    }
    $global:_LastModuleResult = $null

    Write-Host ""
    Write-Host "--- End dispatch ---" -ForegroundColor DarkGray
    Write-Host ""

    return $moduleResult
}

function Set-FabriqIosHostEnvironment {
    # Mirrors kernel/main.ps1 Set-SelectedHostEnvironment but stays
    # local to fabriq_ios. Decryption already handled by Import-ModuleCsv
    # before we receive $Row, so plaintext values flow straight through.
    param([object]$Row)
    if (-not $Row) { return }

    $env:SELECTED_KANRI_NO    = "$($Row.AdminID)"
    $env:SELECTED_OLD_PCNAME  = "$($Row.OldPCName)"
    $env:SELECTED_NEW_PCNAME  = "$($Row.NewPCName)"
    $env:SELECTED_ETH_IP      = "$($Row.EthernetIP)"
    $env:SELECTED_ETH_SUBNET  = "$($Row.EthernetSubnet)"
    $env:SELECTED_ETH_GATEWAY = "$($Row.EthernetGateway)"
    $env:SELECTED_WIFI_IP     = "$($Row.WifiIP)"
    $env:SELECTED_WIFI_SUBNET = "$($Row.WifiSubnet)"
    $env:SELECTED_WIFI_GATEWAY = "$($Row.WifiGateway)"
    $env:SELECTED_DNS1        = "$($Row.DNS1)"
    $env:SELECTED_DNS2        = "$($Row.DNS2)"
    $env:SELECTED_DNS3        = "$($Row.DNS3)"
    $env:SELECTED_DNS4        = "$($Row.DNS4)"
    $env:SELECTED_PIN         = "$($Row.Pin)"
}

function Find-HostlistRowByNewName {
    param(
        [string]$NewName,
        [hashtable]$State
    )
    $csvPath = Join-Path $script:FabriqRoot 'kernel\csv\hostlist.csv'
    if (-not (Test-Path $csvPath)) { return $null }
    $previousPass = $global:FabriqMasterPassphrase
    if ($State.Passphrase) { $global:FabriqMasterPassphrase = $State.Passphrase }
    try {
        $rows = @(Import-ModuleCsv -Path $csvPath)
    } finally {
        $global:FabriqMasterPassphrase = $previousPass
    }
    if (-not $rows) { return $null }
    return ($rows | Where-Object { $_.NewPCName -eq $NewName } | Select-Object -First 1)
}

function ConvertFrom-SubnetMaskToPrefix {
    param([string]$Mask)
    if ([string]::IsNullOrWhiteSpace($Mask)) { return '' }
    try {
        $octets = $Mask.Split('.')
        if ($octets.Count -ne 4) { return '' }
        $bin = ''
        foreach ($o in $octets) {
            $bin += [Convert]::ToString([int]$o, 2).PadLeft(8, '0')
        }
        return (($bin.ToCharArray() | Where-Object { $_ -eq '1' }).Count)
    } catch {
        return ''
    }
}
