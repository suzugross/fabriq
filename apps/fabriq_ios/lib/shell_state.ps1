# ShellState schema and mode transition helpers.

function New-ShellState {
    return @{
        Mode               = 'UserExec'   # UserExec / PrivilegedExec / GlobalConfig / InterfaceConfig / ModuleConfig
        CurrentInterface   = $null
        Passphrase         = $null
        SelectedHost       = $null
        ShouldExit         = $false
        ProcessId          = $PID
        ConfigModuleName   = $null        # set when entering ModuleConfig
        ConfigModuleSchema = $null        # set when entering ModuleConfig
    }
}

function Set-ShellMode {
    param(
        [hashtable]$State,
        [string]$NewMode
    )
    $valid = @('UserExec','PrivilegedExec','GlobalConfig','InterfaceConfig','ModuleConfig')
    if ($NewMode -notin $valid) {
        throw "Invalid mode: $NewMode"
    }
    $allowed = @{
        'UserExec'        = @('PrivilegedExec')
        'PrivilegedExec'  = @('UserExec','GlobalConfig')
        'GlobalConfig'    = @('PrivilegedExec','InterfaceConfig','ModuleConfig')
        'InterfaceConfig' = @('GlobalConfig','PrivilegedExec')
        'ModuleConfig'    = @('GlobalConfig','PrivilegedExec')
    }
    if ($NewMode -notin $allowed[$State.Mode]) {
        throw "Disallowed transition: $($State.Mode) -> $NewMode"
    }
    $State.Mode = $NewMode
}
