# Mode-aware prompt builder.

function Get-FabriqIosPrompt {
    param([hashtable]$State)
    switch ($State.Mode) {
        'UserExec'        { return 'fabriq>' }
        'PrivilegedExec'  { return 'fabriq#' }
        'GlobalConfig'    { return 'fabriq(config)#' }
        'InterfaceConfig' { return 'fabriq(config-if)#' }
        'ModuleConfig' {
            # Prompt suffix depends on which category (verb) was used
            # to enter the mode. Defaults to config-mod when category
            # context is missing for any reason.
            $suffix = 'config-mod'
            if ($State.CurrentCategoryId) {
                $cat = Get-FabriqIosCategoryById -Id $State.CurrentCategoryId
                if ($cat -and $cat.promptSuffix) { $suffix = $cat.promptSuffix }
            }
            return ('fabriq({0})#' -f $suffix)
        }
        default           { return 'fabriq?' }
    }
}
