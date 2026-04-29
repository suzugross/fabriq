# ShellState schema and mode transition helpers.

function New-ShellState {
    # Returns a fresh shell state hashtable.
    return @{
        Mode             = 'UserExec'   # UserExec / PrivilegedExec / GlobalConfig / InterfaceConfig
        CurrentInterface = $null        # set in InterfaceConfig
        Passphrase       = $null        # set on enable; consumed by Import-ModuleCsv
        SelectedHost     = $null        # PSCustomObject row from hostlist.csv
        ShouldExit       = $false
        ProcessId        = $PID
    }
}

function Set-ShellMode {
    param(
        [hashtable]$State,
        [string]$NewMode
    )
    $valid = @('UserExec','PrivilegedExec','GlobalConfig','InterfaceConfig')
    if ($NewMode -notin $valid) {
        throw "Invalid mode: $NewMode"
    }
    $allowed = @{
        'UserExec'        = @('PrivilegedExec')
        'PrivilegedExec'  = @('UserExec','GlobalConfig')
        'GlobalConfig'    = @('PrivilegedExec','InterfaceConfig')
        'InterfaceConfig' = @('GlobalConfig','PrivilegedExec')
    }
    if ($NewMode -notin $allowed[$State.Mode]) {
        throw "Disallowed transition: $($State.Mode) -> $NewMode"
    }
    $State.Mode = $NewMode
}
