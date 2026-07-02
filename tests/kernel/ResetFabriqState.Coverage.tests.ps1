# ========================================
# Pester v5 static-coverage tests for Reset-FabriqState
# ========================================
# Function: kernel/common.ps1 :: Reset-FabriqState
# Run     : powershell.exe -File ./dev/run_tests.ps1
#
# Bug class being pinned: a new $global: variable or SELECTED_*/FABRIQ_*
# environment variable carrying per-session (per-customer) state is added
# to the kernel, but Reset-FabriqState is not updated, so the value leaks
# into the next kitting session on the same process. This has happened
# twice (SELECTED_PIN plaintext residue, TM t-0022; master passphrase
# residue, TM t-0052), so coverage is enforced mechanically:
#
#   every global the kernel assigns, and every session env var the kernel
#   sets, must be either (a) reassigned/cleared inside Reset-FabriqState,
#   or (b) listed in the reasoned allowlist below.
#
# The check is static (AST) - nothing is executed - so it also protects
# code paths a runtime test could not reach (GUI loops, restart paths).
#
# Known blind spots (accepted):
#   - Set-Variable -Scope Global with a dynamic name (single site:
#     Invoke-SafeCommandAsync's child-runspace injection; that scope dies
#     with the runspace and never touches main-process session state).
#   - Restore-HostEnvironment's dynamic Set-Item env: loop (it restores
#     the same SELECTED_* names that are covered explicitly).
#   - apps/fabriq_operator/lib scripts run in their own runspaces; the
#     one global they touch (FabriqEvidenceBasePath) is covered here.
# ========================================

BeforeAll {
    . "$PSScriptRoot\..\_helpers\test_state.ps1"
    $script:RepoRoot = Get-FabriqRepoRoot

    $parseErrors = $null
    $tokens = $null
    $script:KernelAsts = @(
        (Join-Path $script:RepoRoot 'kernel\common.ps1'),
        (Join-Path $script:RepoRoot 'kernel\main.ps1')
    ) | ForEach-Object {
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($_, [ref]$tokens, [ref]$parseErrors)
        if ($parseErrors -and $parseErrors.Count -gt 0) {
            throw "Parse errors in ${_}: $($parseErrors[0].Message)"
        }
        $ast
    }

    # ---- Reset-FabriqState body (common.ps1) ----
    $script:ResetAst = $script:KernelAsts[0].FindAll({
        param($node)
        $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
        $node.Name -eq 'Reset-FabriqState'
    }, $true) | Select-Object -First 1
    if (-not $script:ResetAst) {
        throw 'Reset-FabriqState not found in kernel/common.ps1 (renamed or removed?)'
    }

    # ---- Collector: $global:X = ... assignment targets in an AST ----
    function Get-AssignedGlobalNames {
        param($Ast)
        $names = @()
        foreach ($stmt in $Ast.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.AssignmentStatementAst]
        }, $true)) {
            $lhs = $stmt.Left
            if ($lhs -is [System.Management.Automation.Language.AttributedExpressionAst]) { $lhs = $lhs.Child }
            if ($lhs -is [System.Management.Automation.Language.VariableExpressionAst] -and
                $lhs.VariablePath.UserPath -match '^global:(.+)$') {
                $names += $Matches[1]
            }
        }
        return $names
    }

    # ---- Collector: session env vars (SELECTED_* / FABRIQ_*) SET in an AST ----
    # Returns exact names plus literal name-prefixes for dynamic names like
    # "env:SELECTED_PRINTER_$($i)_NAME" (leading literal up to the first '$').
    function Get-SetSessionEnvNames {
        param($Ast)
        $exact = @()
        $prefixes = @()

        # Form 1: $env:NAME = <non-null>
        foreach ($stmt in $Ast.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.AssignmentStatementAst]
        }, $true)) {
            $lhs = $stmt.Left
            if ($lhs -is [System.Management.Automation.Language.VariableExpressionAst] -and
                $lhs.VariablePath.UserPath -match '^env:((SELECTED_|FABRIQ_)[A-Za-z0-9_]+)$' -and
                $stmt.Right.Extent.Text.Trim() -ne '$null') {
                $exact += $Matches[1]
            }
        }

        # Form 2: Set-Item -Path "env:NAME..." / Set-Item "env:NAME..."
        foreach ($cmd in $Ast.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.CommandAst] -and
            $node.GetCommandName() -eq 'Set-Item'
        }, $true)) {
            foreach ($el in $cmd.CommandElements) {
                $text = $el.Extent.Text.Trim('"', "'")
                if ($text -match '^env:((SELECTED_|FABRIQ_)[A-Za-z0-9_$({}]*)') {
                    $name = $Matches[1]
                    if ($name.Contains('$')) { $prefixes += $name.Substring(0, $name.IndexOf('$')) }
                    else { $exact += $name }
                }
            }
        }

        # Form 3: [Environment]::SetEnvironmentVariable("NAME", <non-null>, ...)
        foreach ($inv in $Ast.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and
            "$($node.Member)" -eq 'SetEnvironmentVariable'
        }, $true)) {
            if ($inv.Arguments.Count -lt 2) { continue }
            if ($inv.Arguments[1].Extent.Text.Trim() -eq '$null') { continue }
            $text = $inv.Arguments[0].Extent.Text.Trim('"', "'")
            if ($text -match '^((SELECTED_|FABRIQ_)[A-Za-z0-9_$({}]*)') {
                $name = $Matches[1]
                if ($name.Contains('$')) { $prefixes += $name.Substring(0, $name.IndexOf('$')) }
                else { $exact += $name }
            }
        }

        return @{ Exact = $exact; Prefixes = $prefixes }
    }

    # ---- Collector: env names CLEARED by Reset-FabriqState ----
    # Any SELECTED_*/FABRIQ_* string constant inside the Reset body counts
    # as covered (the $envKeys array, the SetEnvironmentVariable clears);
    # expandable strings contribute their leading literal as a prefix
    # ("SELECTED_PRINTER_${i}_${suffix}" -> SELECTED_PRINTER_).
    function Get-ResetClearedEnvNames {
        param($ResetAst)
        $exact = @()
        $prefixes = @()
        foreach ($s in $ResetAst.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.StringConstantExpressionAst]
        }, $true)) {
            if ($s.Value -match '^(SELECTED_|FABRIQ_)[A-Za-z0-9_]+$') { $exact += $s.Value }
        }
        foreach ($s in $ResetAst.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.ExpandableStringExpressionAst]
        }, $true)) {
            $text = $s.Extent.Text.Trim('"')
            if ($text -match '^((SELECTED_|FABRIQ_)[A-Za-z0-9_]*)\$') { $prefixes += $Matches[1] }
        }
        return @{ Exact = $exact; Prefixes = $prefixes }
    }

    $script:AssignedGlobals = @($script:KernelAsts | ForEach-Object { Get-AssignedGlobalNames -Ast $_ }) |
        Sort-Object -Unique
    $script:ResetGlobals = @(Get-AssignedGlobalNames -Ast $script:ResetAst) | Sort-Object -Unique

    $envSets = @($script:KernelAsts | ForEach-Object { Get-SetSessionEnvNames -Ast $_ })
    $script:SetEnvExact    = @($envSets | ForEach-Object { $_.Exact })    | Sort-Object -Unique
    $script:SetEnvPrefixes = @($envSets | ForEach-Object { $_.Prefixes }) | Sort-Object -Unique

    $resetEnv = Get-ResetClearedEnvNames -ResetAst $script:ResetAst
    $script:ResetEnvExact    = @($resetEnv.Exact)    | Sort-Object -Unique
    $script:ResetEnvPrefixes = @($resetEnv.Prefixes) | Sort-Object -Unique

    # ---- Allowlist: globals that deliberately survive Reset-FabriqState ----
    # Every entry needs a reason. An entry that is no longer assigned by the
    # kernel, or that Reset started covering, is flagged as stale by the
    # hygiene test below - the list cannot rot silently.
    $script:GlobalAllowlist = @{
        'FabriqUniqueId'               = 'hardware ID; explicitly session-independent (Reset reuses it for the new log name)'
        'ArtPulseFilePath'             = 'ArtPulse liveness signal; machine/process-scoped, no customer data'
        'ArtPulseCounter'              = 'ArtPulse liveness counter; no customer data'
        '_FabriqExitCalled'            = 'process-exit latch; only meaningful while the process is shutting down'
        'FabriqVerboseCaptureActive'   = 'deployment config flag (verbose_capture.flag); installation-scoped'
        'AutoConfirmMode'              = 'per-run transient; set/reset in try/finally around Flex single-module runs (main.ps1)'
        '_FabriqCurrentProfileContext' = 'per-batch transient; cleared in Invoke-BatchExecution finally'
        '_CurrentModuleTelemetry'      = 'per-module telemetry envelope; nulled on module completion paths'
        '_TelemetryWriting'            = 'telemetry reentrancy guard; transient within a single write'
        '_TelemetrySalt'               = 'telemetry pseudonymization salt; installation-scoped by design'
        '_TelemetrySaltDigest'         = 'telemetry salt digest cache; installation-scoped'
        '_TelemetryUtf8NoBom'          = 'cached UTF8-no-BOM encoder instance; content-free'
        '_TelemetryHostInfo'           = 'machine-scoped host info cache (OS/hardware/PS version); no session data'
        '_TelemetryModuleSeq'          = 'monotonic in-process sequence counter; ordering only, no payload'
        'PSDefaultParameterValues'     = 'PowerShell built-in preference dictionary; verbose capture toggles *:Verbose around module execution and restores it in finally'
    }

    # Session env vars that deliberately survive Reset (none today).
    $script:EnvAllowlist = @{}
}

Describe 'Reset-FabriqState global-variable coverage' {

    It 'kernel assigns at least the well-known globals (collector sanity)' {
        # If the AST collector silently broke, coverage would pass vacuously.
        foreach ($known in @('FabriqMasterPassphrase', 'AutoPilotMode', '_LastModuleResult', 'FabriqEvidenceBasePath')) {
            $script:AssignedGlobals -contains $known | Should -BeTrue -Because "collector must see `$global:$known"
        }
        $script:ResetGlobals -contains 'FabriqMasterPassphrase' | Should -BeTrue -Because 'Reset must clear the master passphrase (TM t-0052)'
    }

    It 'every global assigned by the kernel is reset or allowlisted' {
        $uncovered = @($script:AssignedGlobals | Where-Object {
            $script:ResetGlobals -notcontains $_ -and
            -not $script:GlobalAllowlist.ContainsKey($_)
        })
        $uncovered.Count | Should -Be 0 -Because (
            'these globals survive Reset-FabriqState without a documented reason ' +
            '(clear them in Reset or add a reasoned allowlist entry): ' +
            ($uncovered -join ', ')
        )
    }

    It 'the allowlist carries no stale or redundant entries' {
        $stale = @($script:GlobalAllowlist.Keys | Where-Object { $script:AssignedGlobals -notcontains $_ })
        $stale.Count | Should -Be 0 -Because (
            'allowlist entries no longer assigned anywhere in the kernel (remove them): ' +
            ($stale -join ', ')
        )
        $redundant = @($script:GlobalAllowlist.Keys | Where-Object { $script:ResetGlobals -contains $_ })
        $redundant.Count | Should -Be 0 -Because (
            'allowlist entries now covered by Reset itself (remove them): ' +
            ($redundant -join ', ')
        )
    }
}

Describe 'Reset-FabriqState session env-var coverage' {

    It 'kernel sets at least the well-known session env vars (collector sanity)' {
        foreach ($known in @('SELECTED_KANRI_NO', 'SELECTED_PIN', 'FABRIQ_WORKER_NAME')) {
            $script:SetEnvExact -contains $known | Should -BeTrue -Because "collector must see `$env:$known setter"
        }
        $script:SetEnvPrefixes -contains 'SELECTED_PRINTER_' | Should -BeTrue -Because 'collector must see the printer env loop'
        $script:ResetEnvExact -contains 'SELECTED_PIN' | Should -BeTrue -Because 'Reset must clear SELECTED_PIN (TM t-0022)'
    }

    It 'every session env var set by the kernel is cleared by Reset or allowlisted' {
        $isCovered = {
            param($name)
            if ($script:ResetEnvExact -contains $name) { return $true }
            if ($script:EnvAllowlist.ContainsKey($name)) { return $true }
            foreach ($p in $script:ResetEnvPrefixes) {
                if ($name.StartsWith($p, [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
            }
            return $false
        }
        $uncovered = @($script:SetEnvExact | Where-Object { -not (& $isCovered $_) })
        # Dynamic setters (name prefixes) are covered when Reset clears the
        # same prefix (e.g. the SELECTED_PRINTER_ loop on both sides).
        $uncovered += @($script:SetEnvPrefixes | Where-Object {
            $sp = $_
            -not (@($script:ResetEnvPrefixes | Where-Object {
                $_.StartsWith($sp, [System.StringComparison]::OrdinalIgnoreCase) -or
                $sp.StartsWith($_, [System.StringComparison]::OrdinalIgnoreCase)
            }).Count -gt 0)
        })
        $uncovered.Count | Should -Be 0 -Because (
            'these session env vars survive Reset-FabriqState ' +
            '(add them to the $envKeys clear list in Reset or to the env allowlist): ' +
            ($uncovered -join ', ')
        )
    }
}
