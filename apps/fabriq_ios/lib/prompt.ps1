# Mode-aware prompt builder.

function Get-FabriqIosPrompt {
    param([hashtable]$State)
    switch ($State.Mode) {
        'UserExec'        { return 'fabriq>' }
        'PrivilegedExec'  { return 'fabriq#' }
        'GlobalConfig'    { return 'fabriq(config)#' }
        'InterfaceConfig' { return 'fabriq(config-if)#' }
        default           { return 'fabriq?' }
    }
}
