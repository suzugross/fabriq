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
