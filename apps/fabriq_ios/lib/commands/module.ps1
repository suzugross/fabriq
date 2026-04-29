# `module <name>` command implementation - run an arbitrary fabriq
# module (single-execution, no profile features). Lives in
# GlobalConfig mode. Reuses Phase 4's Invoke-FabriqIosModule to
# dispatch the underlying module script and capture the ModuleResult.

# Modules deliberately hidden from `module ?` and `module <Tab>`:
# - GUI-only / scaffolding (no normal use case from the joke shell)
# - Modules that already have a dedicated fabriq_ios command
$script:FabriqIosExcludedModules = @(
    'windows_update'        # GUI Dashboard button only; no module.csv
    'test_error_module'     # Test scaffolding
    'test_harness_config'   # Test scaffolding
    'fabriq_app_launcher'   # Recursive launcher
    'hostname_config'       # Use (config)# `hostname <NewName>`
    'ipaddress_config'      # Use (config-if)# `ip address ...`
)

$script:FabriqIosModuleCache = $null

function Get-ModuleCompletionFromFilesystem {
    # Auto-discovers modules under modules/standard and modules/extended
    # following the `<dir>/<dir>.ps1` convention used by the operator's
    # apps_dialog. Result is cached for the session - new modules added
    # mid-session require restart of fabriq_ios.
    if ($null -ne $script:FabriqIosModuleCache) {
        return $script:FabriqIosModuleCache
    }
    $names = New-Object System.Collections.Generic.List[string]
    foreach ($subdir in @('standard', 'extended')) {
        $dir = Join-Path $script:FabriqRoot ('modules\{0}' -f $subdir)
        if (-not (Test-Path $dir)) { continue }
        $children = @(Get-ChildItem -Path $dir -Directory -ErrorAction SilentlyContinue)
        foreach ($child in $children) {
            $name = $child.Name
            if ($name -in $script:FabriqIosExcludedModules) { continue }
            $entry = Join-Path $child.FullName ('{0}.ps1' -f $name)
            if (Test-Path $entry) {
                $names.Add($name)
            }
        }
    }
    $script:FabriqIosModuleCache = @($names | Sort-Object)
    return $script:FabriqIosModuleCache
}

function Find-ModulePath {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return $null }
    if ($Name -in $script:FabriqIosExcludedModules) { return $null }
    foreach ($subdir in @('standard', 'extended')) {
        $candidate = Join-Path $script:FabriqRoot ('modules\{0}\{1}\{1}.ps1' -f $subdir, $Name)
        if (Test-Path $candidate) { return $candidate }
    }
    return $null
}

function Invoke-ModuleByName {
    param(
        [string]$Name,
        [hashtable]$State
    )

    if ($State.Mode -ne 'GlobalConfig') {
        Write-Host "% 'module' is only available in global configuration mode." -ForegroundColor Red
        return
    }
    if ([string]::IsNullOrWhiteSpace($Name)) {
        Write-Host "% Incomplete: 'module <name>' (try 'module ?')" -ForegroundColor Red
        return
    }

    $modulePath = Find-ModulePath -Name $Name
    if (-not $modulePath) {
        Write-Host ("% Module not found or excluded: {0}" -f $Name) -ForegroundColor Red
        return
    }

    # Set master passphrase for the duration of the run so that any
    # Import-ModuleCsv calls inside the module decrypt ENC: cells.
    $previousPass = $global:FabriqMasterPassphrase
    if ($State.Passphrase) { $global:FabriqMasterPassphrase = $State.Passphrase }
    try {
        $result = Invoke-FabriqIosModule -ScriptPath $modulePath
    } finally {
        $global:FabriqMasterPassphrase = $previousPass
    }

    if (-not $result) {
        Write-FabriqIosSyslog -Severity 3 -Mnemonic 'MODULE' -Key 'error' `
                              -Placeholders @{ Name = $Name; Detail = '(no result)' }
        return
    }

    $detail = if ($result.Message) { $result.Message } else { '' }
    switch ($result.Status) {
        'Success' {
            Write-FabriqIosSyslog -Severity 5 -Mnemonic 'MODULE' -Key 'success' `
                                  -Placeholders @{ Name = $Name }
        }
        'Partial' {
            Write-FabriqIosSyslog -Severity 4 -Mnemonic 'MODULE' -Key 'partial' `
                                  -Placeholders @{ Name = $Name; Detail = $detail }
        }
        'Skipped' {
            Write-FabriqIosSyslog -Severity 6 -Mnemonic 'MODULE' -Key 'skipped' `
                                  -Placeholders @{ Name = $Name; Detail = $detail }
        }
        'Cancelled' {
            Write-FabriqIosSyslog -Severity 5 -Mnemonic 'MODULE' -Key 'cancelled' `
                                  -Placeholders @{ Name = $Name }
        }
        default {
            Write-FabriqIosSyslog -Severity 3 -Mnemonic 'MODULE' -Key 'error' `
                                  -Placeholders @{ Name = $Name; Detail = $detail }
        }
    }
}
