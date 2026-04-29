# `module <name>` command implementation - enter ModuleConfig mode
# for ephemeral configuration of a module. Cisco IOS semantics:
# `(config)# module <name>` enters (config-mod)#, where each
# `set <col1> <val1> <col2> <val2> ...` (or `add ...` alias)
# IMMEDIATELY runs the module with that single ephemeral row -
# the existing CSV file on disk is never touched. Path-matched
# Import-ModuleCsv override redirects only that module's CSV
# read; other reads (e.g. hostlist) pass through.

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

function Get-ModuleCsvSchema {
    # Returns @{ Columns; Enums; CsvFileName; CsvPath } or $null.
    # The list-CSV naming convention is "<module-name-minus-suffix>_list.csv"
    # rather than always "<module-name>_list.csv" - e.g. reg_hklm_config
    # uses reg_hklm_list.csv, app_config uses app_list.csv, bloatware_remove
    # uses bloatware_list.csv. We discover by globbing *_list.csv inside
    # the module directory rather than guessing the prefix.
    # Reads the header and optionally preset.csv for enum hints.
    # $null indicates the module has no *_list.csv - ephemeral
    # configuration is not possible.
    param([string]$Name)

    $modulePath = Find-ModulePath -Name $Name
    if (-not $modulePath) { return $null }

    $moduleDir = Split-Path $modulePath -Parent
    $csvFiles = @(Get-ChildItem -Path $moduleDir -Filter '*_list.csv' -File -ErrorAction SilentlyContinue)
    if ($csvFiles.Count -eq 0) { return $null }

    # Pick the first matching file. Most modules have exactly one
    # *_list.csv. If a module has more than one (rare), the operator
    # should configure via fabriq_studio.
    $csvFile = $csvFiles[0]
    $csvPath = $csvFile.FullName

    $headerLine = ''
    try {
        $headerLine = Get-Content -Path $csvPath -TotalCount 1 -Encoding UTF8
    } catch {
        return $null
    }
    if ($null -eq $headerLine) { return $null }

    # Strip BOM and split.
    $headerLine = $headerLine -replace '^﻿', ''
    $columns = @($headerLine -split ',' | ForEach-Object { $_.Trim().Trim('"') } | Where-Object { $_ })

    $enums = @{}
    $presetPath = Join-Path $moduleDir 'preset.csv'
    if (Test-Path $presetPath) {
        try {
            $rows = @(Import-Csv -Path $presetPath -Encoding UTF8)
            foreach ($r in $rows) {
                $col = $r.Column
                if ([string]::IsNullOrWhiteSpace($col)) { continue }
                if (-not $enums.ContainsKey($col)) {
                    $enums[$col] = New-Object System.Collections.Generic.List[string]
                }
                [void]$enums[$col].Add($r.Value)
            }
        } catch {}
    }

    return @{
        Columns     = $columns
        Enums       = $enums
        CsvFileName = $csvFile.Name
        CsvPath     = $csvPath
    }
}

function Invoke-VerbModeEntry {
    # Helper invoked by the GlobalConfig dispatcher for each
    # category verb (module / cleanup / copy / install / script).
    # Validates Args.Count and delegates to Enter-CategoryConfigMode.
    param(
        [string]$Verb,
        [hashtable]$Resolved,
        [hashtable]$State
    )
    if ($Resolved.Args.Count -lt 1) {
        Write-Host ("% Incomplete: '{0} <name>' (try '{0} ?')" -f $Verb) -ForegroundColor Red
        return
    }
    Enter-CategoryConfigMode -Verb $Verb -Name $Resolved.Args[0] -State $State
}

function Enter-CategoryConfigMode {
    # Validates verb -> category, module-in-category, schema, and
    # transitions to ModuleConfig. The category id is recorded in
    # State.CurrentCategoryId so the prompt renders the appropriate
    # suffix (config-mod / config-clean / config-copy / etc.).
    param(
        [string]$Verb,
        [string]$Name,
        [hashtable]$State
    )
    if ($State.Mode -ne 'GlobalConfig') {
        Write-Host ("% '{0}' is only available in global configuration mode." -f $Verb) -ForegroundColor Red
        return
    }
    if ([string]::IsNullOrWhiteSpace($Name)) {
        Write-Host ("% Incomplete: '{0} <name>' (try '{0} ?')" -f $Verb) -ForegroundColor Red
        return
    }

    $cat = Get-FabriqIosCategoryByVerb -Verb $Verb
    if (-not $cat) {
        Write-Host ("% Unknown category verb: {0}" -f $Verb) -ForegroundColor Red
        return
    }

    if (-not (Test-ModuleInCategory -CategoryId $cat.id -ModuleName $Name)) {
        Write-Host ("% Module '{0}' is not in the '{1}' category. Try '{2} ?'." -f $Name, $cat.id, $Verb) -ForegroundColor Red
        return
    }

    $modulePath = Find-ModulePath -Name $Name
    if (-not $modulePath) {
        Write-Host ("% Module not found or excluded: {0}" -f $Name) -ForegroundColor Red
        return
    }

    $schema = Get-ModuleCsvSchema -Name $Name
    if (-not $schema) {
        Write-Host ("% Module '{0}' has no '*_list.csv'; ephemeral configuration is not available." -f $Name) -ForegroundColor Red
        return
    }

    $State.ConfigModuleName    = $Name
    $State.ConfigModuleSchema  = $schema
    $State.CurrentCategoryId   = $cat.id
    Set-ShellMode -State $State -NewMode 'ModuleConfig'

    Write-Host $cat.intro
    Write-Host ("Configuring module {0}. Type 'show' for schema, 'set <col> <val> [<col> <val>...]' to configure-and-run, 'exit' to leave." -f $Name)
}

function Show-ModuleConfigSchema {
    # Renders the schema in a Cisco-ish layout for `show` inside (config-mod)#.
    param([hashtable]$State)
    if (-not $State.ConfigModuleSchema -or -not $State.ConfigModuleName) {
        Write-Host "% No module configuration context."
        return
    }
    $schema = $State.ConfigModuleSchema
    Write-Host ""
    Write-Host ("  Module: {0}" -f $State.ConfigModuleName)
    Write-Host ""
    Write-Host "  Columns:"
    foreach ($col in $schema.Columns) {
        $hint = ''
        if ($schema.Enums.ContainsKey($col)) {
            $hint = '  [' + ($schema.Enums[$col] -join '|') + ']'
        }
        Write-Host ("    {0}{1}" -f $col, $hint)
    }
    Write-Host ""
    Write-Host "  Usage:"
    Write-Host "    set <col> <val> [<col> <val> ...]   immediately run module with one ephemeral row"
    Write-Host "    add <col> <val> [<col> <val> ...]   alias for set"
    Write-Host "    exit                                 return to (config)#"
    Write-Host "    end                                  return to # (privileged EXEC)"
    Write-Host ""
    Write-Host "  Notes:"
    Write-Host "    - Existing <name>_list.csv on disk is never modified."
    Write-Host "    - The Enabled column defaults to '1' if you omit it."
    Write-Host "    - ENC: encrypted columns cannot be set ephemerally."
    Write-Host ""
}

function Invoke-ModuleEphemeralRun {
    # Builds an ephemeral PSCustomObject from the column/value pairs,
    # overrides Import-ModuleCsv path-matched, and dispatches the
    # module via Phase 4's Invoke-FabriqIosModule. The override is
    # restored in finally regardless of success.
    param(
        [string[]]$PairArgs,
        [hashtable]$State
    )

    if (-not $State.ConfigModuleName -or -not $State.ConfigModuleSchema) {
        Write-Host "% No module configuration context."
        return
    }
    if (-not $PairArgs -or $PairArgs.Count -eq 0) {
        Write-Host "% Usage: set <col> <val> [<col> <val> ...]" -ForegroundColor Red
        return
    }
    if ($PairArgs.Count % 2 -ne 0) {
        Write-Host "% Odd argument count. Pairs of <col> <val> required." -ForegroundColor Red
        return
    }

    $name   = $State.ConfigModuleName
    $schema = $State.ConfigModuleSchema

    # Build the row hashtable from pairs. Reject unknown columns.
    $row = @{}
    for ($i = 0; $i -lt $PairArgs.Count; $i += 2) {
        $col = $PairArgs[$i]
        $val = $PairArgs[$i + 1]
        if ($col -notin $schema.Columns) {
            Write-Host ("% Unknown column '{0}'. Valid columns: {1}" -f $col, ($schema.Columns -join ', ')) -ForegroundColor Red
            return
        }
        $row[$col] = $val
    }

    # Default Enabled=1 so most modules don't filter the row out.
    if ('Enabled' -in $schema.Columns -and -not $row.ContainsKey('Enabled')) {
        $row['Enabled'] = '1'
    }

    # Materialise as PSCustomObject in CSV column order, with empty
    # strings for columns the user did not set.
    $orderedProps = [ordered]@{}
    foreach ($col in $schema.Columns) {
        if ($row.ContainsKey($col)) {
            $orderedProps[$col] = "$($row[$col])"
        } else {
            $orderedProps[$col] = ''
        }
    }
    $ephemeralRow = [PSCustomObject]$orderedProps

    $modulePath = Find-ModulePath -Name $name
    if (-not $modulePath) {
        Write-Host ("% Module not found: {0}" -f $name) -ForegroundColor Red
        return
    }

    # Path-matched override of Import-ModuleCsv: only the module's
    # own *_list.csv read is intercepted; reads of other CSVs (e.g.
    # hostlist.csv) pass through to the real implementation. The
    # CSV filename is taken from the schema (auto-discovered via glob)
    # because the naming convention drops the action suffix
    # (reg_hklm_config -> reg_hklm_list.csv, etc.).
    $expectedFile = $schema.CsvFileName
    $original = Get-Item Function:Import-ModuleCsv -ErrorAction SilentlyContinue
    if (-not $original) {
        Write-Host "% Import-ModuleCsv is not loaded; cannot ephemeral-dispatch." -ForegroundColor Red
        return
    }
    $originalScript = $original.ScriptBlock

    $override = {
        param(
            [Parameter(Mandatory = $true)][string]$Path,
            [switch]$FilterEnabled,
            [string[]]$RequiredColumns,
            [string]$Segment = $env:FABRIQ_SEGMENT
        )
        if ([System.IO.Path]::GetFileName($Path) -eq $expectedFile) {
            # Honour FilterEnabled so modules that rely on it still
            # see the same semantics; ignore Segment (ephemeral row
            # intent overrides segment scoping).
            if ($FilterEnabled) {
                return @($ephemeralRow | Where-Object { $_.Enabled -eq '1' })
            }
            return @($ephemeralRow)
        }
        return & $originalScript @PSBoundParameters
    }.GetNewClosure()

    $previousPass = $global:FabriqMasterPassphrase
    if ($State.Passphrase) { $global:FabriqMasterPassphrase = $State.Passphrase }

    # Set the override at the global function scope so that scripts
    # invoked via `& $modulePath` (which run in their own scope but
    # resolve unqualified function calls via the global scope chain)
    # see our shadow. Default `Set-Item Function:` is current-scope
    # only and would not be visible to module-side calls.
    Set-Item Function:Global:Import-ModuleCsv -Value $override
    try {
        $result = Invoke-FabriqIosModule -ScriptPath $modulePath
    } finally {
        Set-Item Function:Global:Import-ModuleCsv -Value $originalScript
        $global:FabriqMasterPassphrase = $previousPass
    }

    if (-not $result) {
        Write-FabriqIosSyslog -Severity 3 -Mnemonic 'MODULE' -Key 'error' `
                              -Placeholders @{ Name = $name; Detail = '(no result)' }
        return
    }

    $detail = if ($result.Message) { $result.Message } else { '' }
    switch ($result.Status) {
        'Success' {
            Write-FabriqIosSyslog -Severity 5 -Mnemonic 'MODULE' -Key 'success' `
                                  -Placeholders @{ Name = $name }
        }
        'Partial' {
            Write-FabriqIosSyslog -Severity 4 -Mnemonic 'MODULE' -Key 'partial' `
                                  -Placeholders @{ Name = $name; Detail = $detail }
        }
        'Skipped' {
            Write-FabriqIosSyslog -Severity 6 -Mnemonic 'MODULE' -Key 'skipped' `
                                  -Placeholders @{ Name = $name; Detail = $detail }
        }
        'Cancelled' {
            Write-FabriqIosSyslog -Severity 5 -Mnemonic 'MODULE' -Key 'cancelled' `
                                  -Placeholders @{ Name = $name }
        }
        default {
            Write-FabriqIosSyslog -Severity 3 -Mnemonic 'MODULE' -Key 'error' `
                                  -Placeholders @{ Name = $name; Detail = $detail }
        }
    }
}
