# ========================================
#
# Fabriq ver3.6 - Manifeste du Surkitinisme -
#
# ========================================

param(
    [switch]$NoGui
)

# ========================================
# Load Common Function Library
# ========================================
$commonPath = ".\kernel\common.ps1"
if (Test-Path $commonPath) {
    . $commonPath
}

# Load Manifesto GUI function
$manifestoPath = ".\kernel\ps1\manifesto.ps1"
if (Test-Path $manifestoPath) {
    . $manifestoPath
}

# Load Fabriq Operator GUI
# GUI is mandatory for menu navigation. -NoGui is deprecated and will exit.
$operatorPath = ".\apps\fabriq_operator\fabriq_operator.ps1"
if (Test-Path $operatorPath) {
    . $operatorPath
}

if (-not $script:UseGui) {
    Show-Error "Failed to load fabriq operator GUI."
    Show-Error "Ensure fabriq_operator files exist and System.Windows.Forms is available."
    exit 1
}

# Enable sleep suppression while Fabriq is running
Enable-SleepSuppression

# Disable QuickEdit to prevent accidental freeze on console click
Disable-QuickEditMode

# Set compact console window size
Set-ConsoleSize -Columns 75 -Lines 35

# ========================================
# Constants
# ========================================
$HOSTLIST_CSV = ".\kernel\csv\hostlist.csv"
$APPS_DIR = ".\apps"
$script:AutoPilotMaxRetry = 5
# Passphrase verification token (Studio-generated). Defined here at script
# scope so EVERY Show-SessionSetupForm call site receives a real path -
# defining it only inside the fresh-start branch left it $null on the
# resume-entry path, and session_form silently SKIPS verification when the
# token path is empty (any wrong passphrase was accepted via NewSession).
$verifyTokenPath = Join-Path $PSScriptRoot "txt\passphrase_verify.txt"

# ========================================
# Function: Load hostlist.csv
# ========================================
function Load-HostList {
    if (-not (Test-Path $HOSTLIST_CSV)) {
        Show-Error "hostlist.csv not found: $HOSTLIST_CSV"
        return $null
    }

    try {
        $hostList = Import-Csv -Path $HOSTLIST_CSV -Encoding Default
        Show-Success "Loaded hostlist.csv ($(($hostList | Measure-Object).Count) items)"
        return $hostList
    }
    catch {
        Show-Error "Failed to load hostlist.csv: $_"
        return $null
    }
}

# (Load-Categories, Load-ModulesFromDirectory, Load-AllModules, Build-CategoryMenu
#  have been consolidated into Initialize-ModuleSystem and Build-CategoryMenu in common.ps1)

# ========================================
# Function: Set Environment Variables
# ========================================
function Set-SelectedHostEnvironment {
    param([object]$SelectedHost)

    # Self-referencing mode: a cell whose value is the literal __SELF__
    # token resolves to THIS PC's live value for that column's semantic
    # ("always correct" - the current machine is always the answer). Used
    # for hostlist-free kitting / terminal investigation, e.g. evidence
    # file names that carry the actual PC name. Resolution happens once
    # here (kitting-screen entry, after passphrase). The resolved concrete
    # value is baked into the SELECTED_* env var, which Save-ResumeState
    # snapshots and Restore-HostEnvironment restores verbatim across a
    # __RESTART__ - so a later hostname/IP change does NOT re-follow (by
    # design; the token never reaches the resume state).
    $SELF_TOKEN = '__SELF__'

    # Query live PC info only when at least one __SELF__ token is present,
    # so plain hostlists pay zero overhead (Get-CurrentPCInfo enumerates
    # all network adapters and printers).
    $hasSelf = $false
    foreach ($prop in $SelectedHost.PSObject.Properties) {
        if ($prop.Value -is [string] -and $prop.Value -eq $SELF_TOKEN) { $hasSelf = $true; break }
    }
    $live = $null
    if ($hasSelf) {
        try { $live = Get-CurrentPCInfo } catch { $live = $null }
    }

    # Helper: resolve a single hostlist cell.
    #   1. __SELF__  -> this PC's live value for $SelfKind. Empty + warning
    #                   if the column has no self-source (e.g. PIN/printer)
    #                   or the live value is unavailable (e.g. no Wi-Fi
    #                   adapter, fewer DNS servers than the slot index).
    #   2. ENC:...   -> transparent decryption when a passphrase is loaded.
    #   3. otherwise -> literal passthrough.
    function Resolve-HostValue {
        param(
            [string]$Value,
            [object]$Live = $null,
            [string]$SelfKind = ''
        )
        if ([string]::IsNullOrEmpty($Value)) { return $Value }

        if ($Value -eq '__SELF__') {
            if ([string]::IsNullOrEmpty($SelfKind) -or $null -eq $Live) {
                Show-Warning "Self-reference (__SELF__) is not supported for this field; leaving empty"
                return ""
            }
            $resolved = switch ($SelfKind) {
                'ComputerName'    { $Live.ComputerName }
                'EthernetIP'      { $Live.EthernetIP }
                'EthernetSubnet'  { $Live.EthernetSubnet }
                'EthernetGateway' { $Live.EthernetGateway }
                'WifiIP'          { $Live.WifiIP }
                'WifiSubnet'      { $Live.WifiSubnet }
                'WifiGateway'     { $Live.WifiGateway }
                'DNS1'            { if (@($Live.DNS).Count -ge 1) { @($Live.DNS)[0] } else { "" } }
                'DNS2'            { if (@($Live.DNS).Count -ge 2) { @($Live.DNS)[1] } else { "" } }
                'DNS3'            { if (@($Live.DNS).Count -ge 3) { @($Live.DNS)[2] } else { "" } }
                'DNS4'            { if (@($Live.DNS).Count -ge 4) { @($Live.DNS)[3] } else { "" } }
                default           { "" }
            }
            if ([string]::IsNullOrEmpty($resolved)) {
                Show-Warning "Self-reference (__SELF__) could not be resolved for $SelfKind on this PC; leaving empty"
                return ""
            }
            return $resolved
        }

        if ($Value.StartsWith('ENC:') -and -not [string]::IsNullOrWhiteSpace($global:FabriqMasterPassphrase)) {
            try {
                return (Unprotect-FabriqValue -EncryptedValue $Value -Passphrase $global:FabriqMasterPassphrase)
            }
            catch {
                Show-Warning "Failed to decrypt host value: $_"
                return $Value
            }
        }
        return $Value
    }

    # Note: These keys must match your CSV headers
    $env:SELECTED_KANRI_NO   = Resolve-HostValue $SelectedHost.'AdminID'
    $env:SELECTED_OLD_PCNAME = Resolve-HostValue $SelectedHost.'OldPCName' -Live $live -SelfKind 'ComputerName'
    $env:SELECTED_NEW_PCNAME = Resolve-HostValue $SelectedHost.'NewPCName' -Live $live -SelfKind 'ComputerName'

    $env:SELECTED_ETH_IP      = Resolve-HostValue $SelectedHost.'EthernetIP'      -Live $live -SelfKind 'EthernetIP'
    $env:SELECTED_ETH_SUBNET  = Resolve-HostValue $SelectedHost.'EthernetSubnet'  -Live $live -SelfKind 'EthernetSubnet'
    $env:SELECTED_ETH_GATEWAY = Resolve-HostValue $SelectedHost.'EthernetGateway' -Live $live -SelfKind 'EthernetGateway'

    $env:SELECTED_WIFI_IP      = Resolve-HostValue $SelectedHost.'WifiIP'      -Live $live -SelfKind 'WifiIP'
    $env:SELECTED_WIFI_SUBNET  = Resolve-HostValue $SelectedHost.'WifiSubnet'  -Live $live -SelfKind 'WifiSubnet'
    $env:SELECTED_WIFI_GATEWAY = Resolve-HostValue $SelectedHost.'WifiGateway' -Live $live -SelfKind 'WifiGateway'

    $env:SELECTED_DNS1 = Resolve-HostValue $SelectedHost.'DNS1' -Live $live -SelfKind 'DNS1'
    $env:SELECTED_DNS2 = Resolve-HostValue $SelectedHost.'DNS2' -Live $live -SelfKind 'DNS2'
    $env:SELECTED_DNS3 = Resolve-HostValue $SelectedHost.'DNS3' -Live $live -SelfKind 'DNS3'
    $env:SELECTED_DNS4 = Resolve-HostValue $SelectedHost.'DNS4' -Live $live -SelfKind 'DNS4'

    $env:SELECTED_PIN = Resolve-HostValue $SelectedHost.'Pin'

    for ($i = 1; $i -le 10; $i++) {
        # CSV headers like: Printer1Name, Printer1Driver, Printer1Port
        $nameKey   = "Printer$($i)Name"
        $driverKey = "Printer$($i)Driver"
        $portKey   = "Printer$($i)Port"

        Set-Item -Path "env:SELECTED_PRINTER_$($i)_NAME"   -Value (Resolve-HostValue $SelectedHost.$nameKey)   -ErrorAction SilentlyContinue
        Set-Item -Path "env:SELECTED_PRINTER_$($i)_DRIVER" -Value (Resolve-HostValue $SelectedHost.$driverKey) -ErrorAction SilentlyContinue
        Set-Item -Path "env:SELECTED_PRINTER_$($i)_PORT"   -Value (Resolve-HostValue $SelectedHost.$portKey)   -ErrorAction SilentlyContinue
    }

    Show-Info "Environment variables set."

    # Push the just-set SELECTED_* values to the execution toolbar's
    # PC Info pane. No-op when the toolbar is not yet running (the
    # toolbar self-pushes from the same env vars on its own startup,
    # so the fresh-start / resume cases are covered there).
    if (Get-Command Update-ExecutionToolbar -ErrorAction SilentlyContinue) {
        try { Update-ExecutionToolbar -TargetHostInfo (Get-FabriqHostInfoFromEnv) } catch { }
    }
}

# ========================================
# Function: Execute Script (With History)
# ========================================
function Invoke-KittingScript {
    param(
        [string]$ScriptPath,
        [string]$ModuleName = "",
        [string]$Category = ""
    )

    if (-not (Test-Path $ScriptPath)) {
        Show-Error "Script not found: $ScriptPath"
        Add-ExecutionResult -Operation $ModuleName -Status "Error" -Message "Script undetected"
        $null = Write-ExecutionHistory -ModuleName $ModuleName -Category $Category -Status "Error" -Message "Script undetected"
        return $false
    }

    Show-Info "Executing script: $ScriptPath"
    Write-Host ""

    try {
        # Clear the global fallback before invocation
        $global:_LastModuleResult = $null

        $output = & $ScriptPath

        # Find a ModuleResult in the pipeline output
        $moduleResult = $null
        if ($null -ne $output) {
            foreach ($item in @($output)) {
                if ($item -is [PSCustomObject] -and $item._IsModuleResult -eq $true) {
                    $moduleResult = $item
                }
            }
        }

        # Fallback: pipeline capture failed, recover from the global
        if (-not $moduleResult -and $null -ne $global:_LastModuleResult) {
            $moduleResult = $global:_LastModuleResult
        }
        $global:_LastModuleResult = $null

        if ($moduleResult) {
            # Use the result the module returned
            $status = $moduleResult.Status
            $message = $moduleResult.Message

            switch ($status) {
                "Success"   { Write-Host ""; Show-Success "Script execution completed" }
                "Error"     { Write-Host ""; Show-Error "Script reported error: $message" }
                "Cancelled" { Write-Host ""; Show-Info "Script was cancelled by user" }
                "Skipped"   { Write-Host ""; Show-Skip "Script was skipped: $message" }
                "Partial"   { Write-Host ""; Show-Warning "Script completed with partial results: $message" }
            }

            # Verified field (Post-Apply Verification)
            $verifiedStr = ""
            if ($null -ne $moduleResult.Verified) {
                $verifiedStr = if ($moduleResult.Verified) { "True" } else { "False" }
            }

            Add-ExecutionResult -Operation $ModuleName -Status $status -Message $message -Verified $moduleResult.Verified
            $null = Write-ExecutionHistory -ModuleName $ModuleName -Category $Category -Status $status -Message $message -Verified $verifiedStr
            Capture-ScreenEvidence -ModuleName $ModuleName -Status $status
            return ($status -eq "Success")
        }
        else {
            # Fail-closed: no ModuleResult returned despite normal completion.
            # Record Error instead of assuming success (module result contract
            # violation; mirrors Invoke-SafeCommand / Invoke-SafeCommandAsync).
            Write-Host ""
            Show-Warning "No ModuleResult returned - recording Error (module result contract violation)"
            Add-ExecutionResult -Operation $ModuleName -Status "Error" -Message "No ModuleResult returned (module result contract violation)"
            $null = Write-ExecutionHistory -ModuleName $ModuleName -Category $Category -Status "Error" -Message "No ModuleResult returned (module result contract violation)"
            Capture-ScreenEvidence -ModuleName $ModuleName -Status "Error"
            return $false
        }
    }
    catch {
        Write-Host ""
        Show-Error "Error occurred during script execution: $_"
        Add-ExecutionResult -Operation $ModuleName -Status "Error" -Message $_.Exception.Message
        $null = Write-ExecutionHistory -ModuleName $ModuleName -Category $Category -Status "Error" -Message $_.Exception.Message
        Capture-ScreenEvidence -ModuleName $ModuleName -Status "Error"
        return $false
    }
}

# ========================================
# Function: Batch Execution (With History)
# ========================================
function Invoke-BatchExecution {
    param(
        [array]$SelectedModules,
        [switch]$AutoPilot,
        [int]$AutoPilotWaitSec = 3,
        # Profile restart support (optional)
        [string]$ProfilePath = "",
        [string]$ProfileName = "",
        # Full profile module list for checklist (covers pre-restart modules in resume)
        # If omitted, $SelectedModules is used as-is
        [array]$FullProfileModules = $null,
        # Absolute start timestamp of the profile run. On a fresh invocation
        # this defaults to "now"; on resume after __RESTART__ the caller
        # passes the original start timestamp restored from resume_state.json
        # so the displayed elapsed time spans the entire wall-clock duration
        # (including reboot/login/startup gaps).
        [datetime]$ProfileStartTime = (Get-Date),
        # Whether to fire the post-profile finalize pipeline
        # (Complete-ProfileExecution) on natural completion. Default $true
        # preserves Linear behavior. FlexProfile passes :$false for batch
        # / single-execution runs that should leave finalize to the
        # operator-driven [Complete] button. Has no effect when
        # __RESTART__ exits the loop early (resume path will set this
        # again on its second-leg call).
        # [bool] (not [switch]) is intentional so the default can be $true
        # without tripping PSAvoidDefaultValueSwitchParameter; callers
        # use -FinalizeOnComplete:$false to opt out.
        [bool]$FinalizeOnComplete = $true,
        # FlexProfile pass-through params for resume_state.json v2.
        # When ExecutionMode='Linear' (default) the internal Save-ResumeState
        # call writes v1 format (byte-for-byte compatible). When 'Flex',
        # SelectedOrders / ModuleStates are persisted into the v2 resume
        # state so the post-reboot dashboard can rebuild the operator's
        # checkbox subset accurately.
        [ValidateSet('Linear','Flex')][string]$ExecutionMode = 'Linear',
        [int[]]$SelectedOrders = @(),
        [hashtable]$ModuleStates = @{}
    )

    # AutoPilot: set global flag (Profile scope, reset in finally)
    if ($AutoPilot) {
        $global:AutoPilotMode = $true
        $global:AutoPilotWaitSec = $AutoPilotWaitSec
    }

    # Per-module results accumulator. Declared at function scope (not
    # inside try) so the finally block can publish a final snapshot to
    # $script:LastBatchResults regardless of how the batch ends
    # (cancel / completion / mid-throw / __RESTART__ early exit).
    $completedResults = @()
    $script:LastBatchResults = @()

    try {

    # Confirm execution
    if (-not (Show-BatchConfirmation -SelectedModules $SelectedModules)) {
        Show-Info "Batch execution canceled"
        return
    }

    Clear-ExecutionResults

    # Execution toolbar: signal batch start (enables Skip / Gyotaq
    # buttons; per-module label updates happen inside the foreach
    # loop below).
    try { Update-ExecutionToolbar -ExecutionState 'Running' } catch { }

    # Reload current-session history into IsRestored entries so the
    # previous batch's results survive Clear-ExecutionResults's wipe of
    # non-IsRestored entries between batches in the same session.
    # Without this, FlexProfile single re-runs caused HTML checklist /
    # dashboard to show all other modules as "NotRun" because their
    # fresh entries (added by Add-ExecutionResult during the prior batch)
    # were dropped on the next Clear. SessionID filter prevents
    # cross-session entries from leaking in.
    Restore-ExecutionHistory -SessionIDFilter $script:SessionID

    # Telemetry kernel event: profile/batch start
    try {
        Write-KernelTelemetryEvent -Type "profile.start" -Data ([ordered]@{
            profileName    = $ProfileName
            profilePath    = $ProfilePath
            executionMode  = $ExecutionMode
            moduleCount    = $SelectedModules.Count
            autoPilot      = [bool]$AutoPilot
            selectedOrders = @($SelectedOrders)
        })
    } catch { }

    $total = $SelectedModules.Count
    $current = 0

    # Cross-module dependency tracking for envelope.start enrichment.
    $prevModuleName   = ""
    $prevModuleStatus = ""

    # __GATE__ admission control (TM t-0073). $gateRows is the FULL ordered
    # profile (so gate positions / window membership are known even when the
    # operator runs a sparse Flex selection); $statusByOrder is the live
    # per-Order status used to evaluate the barrier dynamically. It is seeded
    # from the current session's history (captures prior runs AND modules
    # completed before a __RESTART__) and updated as each module finishes.
    $gateRows = if ($null -ne $FullProfileModules) { $FullProfileModules } else { $SelectedModules }

    # Order -> MenuName for the CURRENT profile. execution_history.csv has no
    # profile column and Order is per-profile-local, so the seed below only
    # accepts a history row whose ModuleName matches the module occupying
    # that Order in this profile. Without this, a same-Order module from a
    # DIFFERENT profile run earlier in the same session (master-image flow
    # runs Pre/Config stages back-to-back with overlapping Orders) could seed
    # a false barrier and silently leave legitimate modules Pending.
    $gateMenuByOrder = @{}
    foreach ($gr in $gateRows) {
        $gro = $null
        try { $gro = [int]$gr.Order } catch { $gro = $null }
        if ($null -ne $gro -and -not $gateMenuByOrder.ContainsKey($gro)) {
            $gateMenuByOrder[$gro] = "$($gr.MenuName)"
        }
    }

    $statusByOrder = @{}
    $verifiedByOrder = @{}
    if (-not [string]::IsNullOrEmpty($env:SELECTED_KANRI_NO)) {
        $gateHist = @(Import-ExecutionHistory -FilterKanriNo $env:SELECTED_KANRI_NO | Where-Object { $_.SessionID -eq $script:SessionID })
        foreach ($gh in $gateHist) {
            $gho = $null
            try { $gho = [int]$gh.Order } catch { $gho = $null }
            if ($null -eq $gho -or $gho -le 0 -or $statusByOrder.ContainsKey($gho)) { continue }
            # Import-ExecutionHistory is descending by Timestamp (second
            # resolution), so the first occurrence per Order is the latest;
            # the live in-run update below corrects any same-second tie once
            # the Order is re-executed this batch. Bind the row to this
            # profile by MenuName before trusting it.
            if ($gateMenuByOrder.ContainsKey($gho) -and "$($gh.ModuleName)" -eq $gateMenuByOrder[$gho]) {
                $statusByOrder[$gho] = $gh.Status
                # History Verified column is "True"/"False"/"" -> bool/$null.
                $verifiedByOrder[$gho] = if ($gh.Verified -eq 'True') { $true } elseif ($gh.Verified -eq 'False') { $false } else { $null }
            }
        }
    }
    $gateBlockedOrders = @()

    foreach ($module in $SelectedModules) {
        $current++

        # Set per-module profile context (consumed by Start-ModuleTelemetry).
        # Cleared in finally below so a finished/failed module doesn't leak
        # context into the next batch.
        $global:_FabriqCurrentProfileContext = @{
            ProfileName      = $ProfileName
            ProfileOrder     = [int]$module.Order
            ExecutionMode    = $ExecutionMode
            PrevModuleName   = $prevModuleName
            PrevModuleStatus = $prevModuleStatus
        }

        # __GATE__ marker: a checkpoint, not an executable module. It never
        # runs anything; Get-FabriqGateBarrier accounts for its position
        # when computing the barrier below.
        if ($module._IsGate) {
            Show-BatchProgress -Current $current -Total $total -ItemName "[GATE]"
            Show-Info "[GATE] checkpoint at Order $($module.Order)"
            continue
        }

        # __GATE__ admission control (TM t-0073): refuse any Order at or
        # beyond the first unsatisfied gate. Evaluated dynamically against
        # the live status map, so a failure earlier in THIS run (or a prior
        # run / pre-__RESTART__ run via history) blocks everything past the
        # gate until the operator clears it and re-runs. Blocked modules are
        # left untouched (Pending), not recorded as a new status.
        $gateBarrier = Get-FabriqGateBarrier -Rows $gateRows -StatusMap $statusByOrder -VerifiedMap $verifiedByOrder
        if ($null -ne $gateBarrier -and [int]$module.Order -ge $gateBarrier) {
            Show-BatchProgress -Current $current -Total $total -ItemName $module.MenuName
            Show-Warning "[GATE] Blocked: Order $($module.Order) ($($module.MenuName)) is beyond the unsatisfied gate at Order $gateBarrier. Resolve the failure(s) above and re-run."
            $gateBlockedOrders += [int]$module.Order
            continue
        }

        # __RESTART__ marker handling
        if ($module._IsRestart) {
            if ([string]::IsNullOrEmpty($ProfilePath)) {
                # Non-profile batch run: skip restart markers
                continue
            }

            Show-BatchProgress -Current $current -Total $total -ItemName "[RESTART]"
            Write-Host ""
            Write-Host "========================================" -ForegroundColor Yellow
            Write-Host "  Profile Restart Phase" -ForegroundColor Yellow
            Write-Host "========================================" -ForegroundColor Yellow
            Write-Host ""
            Write-Host "  Progress: $($completedResults.Count) modules completed" -ForegroundColor White
            Write-Host "  Remaining: $($total - $current) modules after restart" -ForegroundColor White
            Write-Host ""

            # Save resume state. ProfileStartTime is the absolute origin
            # (not elapsed) so the post-resume display covers the full
            # wall clock including reboot/login/startup gaps.
            # When ExecutionMode='Linear' (default) Save-ResumeState writes
            # v1 format; when 'Flex', it writes schemaVersion=2 with
            # SelectedOrders / ModuleStates pre-populated from the caller.
            Save-ResumeState -ProfilePath $ProfilePath `
                             -ProfileName $ProfileName `
                             -ResumeAfterOrder $module.Order `
                             -CompletedModules $completedResults `
                             -ProfileStartTime $ProfileStartTime `
                             -ExecutionMode $ExecutionMode `
                             -SelectedOrders $SelectedOrders `
                             -ModuleStates $ModuleStates

            # Register RunOnce
            if (-not (Register-FabriqRunOnce)) {
                # Fail-closed: the reboot-and-resume leg cannot be armed, so the
                # remaining modules (scheduled to run AFTER the reboot) must NOT
                # execute in this un-rebooted session. Running them on stale
                # state (e.g. a domain join before the hostname change takes
                # effect) causes misconfiguration. Stop the batch instead of
                # continuing past the restart marker (TM t-0045).
                Show-Error "RunOnce registration failed - restart aborted; remaining modules skipped"
                Remove-ResumeState
                Add-ExecutionResult -Operation "[RESTART]" -Status "Error" -Message "RunOnce registration failed (restart aborted, remaining modules skipped)" -Order $module.Order
                $null = Write-ExecutionHistory -ModuleName "[RESTART]" -Category "System" -Status "Error" -Message "RunOnce registration failed - restart aborted, remaining modules skipped (fail-closed)" -Order $module.Order
                break
            }

            # Record in execution history
            Add-ExecutionResult -Operation "[RESTART]" -Status "Success" -Message "Restarting..." -Order $module.Order
            $null = Write-ExecutionHistory -ModuleName "[RESTART]" -Category "System" -Status "Success" -Message "Profile restart (ResumeAfter: $($module.Order))" -Order $module.Order

            # Hide toolbar cleanly before reboot kills the process
            try { Hide-ExecutionToolbar } catch { }

            Invoke-CountdownRestart
            return
        }

        # __REEXPLORER__ marker handling
        if ($module._IsReexplorer) {
            Show-BatchProgress -Current $current -Total $total -ItemName "[REEXPLORER]"
            Show-Info "Restarting Explorer..."
            # Capture actual outcome so $completedResults reflects the
            # try/catch result instead of a hardcoded "Success" (the
            # latter caused resume display to mislabel failed reexplorer
            # markers as successful).
            $reexplorerStatus  = "Success"
            $reexplorerMessage = "Explorer restarted"
            try {
                Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
                $maxWait = 15; $interval = 1; $elapsed = 0; $restarted = $false
                while ($elapsed -lt $maxWait) {
                    Start-Sleep -Seconds $interval
                    $elapsed += $interval
                    if (@(Get-Process -Name explorer -ErrorAction SilentlyContinue).Count -gt 0) {
                        $restarted = $true; break
                    }
                }
                # Only force-start when Windows did not auto-revive Explorer in time
                if (-not $restarted) { Start-Process explorer.exe }
                Add-ExecutionResult -Operation "[REEXPLORER]" -Status "Success" -Message "Explorer restarted" -Order $module.Order
                $null = Write-ExecutionHistory -ModuleName "[REEXPLORER]" -Category "System" -Status "Success" -Message "Explorer restarted" -Order $module.Order
            }
            catch {
                $reexplorerStatus  = "Error"
                $reexplorerMessage = $_.Exception.Message
                Add-ExecutionResult -Operation "[REEXPLORER]" -Status "Error" -Message $_.Exception.Message -Order $module.Order
                $null = Write-ExecutionHistory -ModuleName "[REEXPLORER]" -Category "System" -Status "Error" -Message $_.Exception.Message -Order $module.Order
            }
            $completedResults += @{
                Order    = $module.Order
                MenuName = "[REEXPLORER]"
                Status   = $reexplorerStatus
                Verified = $null
                Message  = $reexplorerMessage
            }
            $statusByOrder[[int]$module.Order] = $reexplorerStatus
            $verifiedByOrder[[int]$module.Order] = $null
            continue
        }

        # Normal module execution
        Show-BatchProgress -Current $current -Total $total -ItemName $module.MenuName

        # Execution toolbar: update label to current module name
        try { Update-ExecutionToolbar -ExecutionState 'Running' -ModuleName $module.MenuName } catch { }

        # AutoPilot: inter-module wait
        if ($global:AutoPilotMode -and $current -gt 1) {
            Write-Host "[AUTOPILOT] Next module in $($global:AutoPilotWaitSec)s..." -ForegroundColor Magenta
            Start-Sleep -Seconds $global:AutoPilotWaitSec
        }

        # Retry loop for Error/Partial handling in AutoPilot mode
        $retryModule = $false
        $autoRetryCount = 0
        do {
            if ($retryModule) {
                Show-Info "Retrying: $($module.MenuName)"
                Write-Host ""
                $retryModule = $false
            }

            # __AUTO_to_<User>__ parameter passing via environment variable
            if ($module._AutoLogonUser) {
                $env:FABRIQ_AUTOLOGON_USER = $module._AutoLogonUser
            }

            # Segment parameter passing via environment variable
            if ($module._Segment) {
                $env:FABRIQ_SEGMENT = $module._Segment
            }

            # Dispatch: async (runspace + skip/timeout monitoring) or sync
            if ($module._IsAsync -and (Get-FabriqAsyncConfig).Enabled) {
                Write-Host "[ASYNC] Running in monitored runspace (Skip available via Status Monitor)" -ForegroundColor DarkCyan
                $result = Invoke-SafeCommandAsync -OperationName $module.MenuName `
                    -ScriptPath $module.Script -ContinueOnError
            }
            else {
                $result = Invoke-SafeCommand -OperationName $module.MenuName -ScriptBlock {
                    & $module.Script
                } -ContinueOnError
            }

            # Clean up AutoLogon environment variable
            if ($module._AutoLogonUser) {
                $env:FABRIQ_AUTOLOGON_USER = $null
            }

            # Clean up Segment environment variable
            if ($module._Segment) {
                $env:FABRIQ_SEGMENT = $null
            }

            # Error/Partial notification and AutoPilot retry dialog
            if ($result.Status -eq "Error" -or $result.Status -eq "Partial") {
                Invoke-ErrorNotification -ModuleName $module.MenuName -Status $result.Status

                if ($global:AutoPilotMode) {
                    $errorMode = if ($module._ErrorMode) { $module._ErrorMode.ToLower() } else { "" }

                    if ($errorMode -eq "skip") {
                        Show-Warning "[AUTOPILOT] ErrorMode=Skip -> recording $($result.Status) and continuing"
                        # fall through: record original Error/Partial status
                    }
                    elseif ($errorMode -eq "retry") {
                        if ($autoRetryCount -lt $script:AutoPilotMaxRetry) {
                            $autoRetryCount++
                            Show-Warning "[AUTOPILOT] ErrorMode=Retry -> auto-retry ($autoRetryCount/$script:AutoPilotMaxRetry)"
                            Start-Sleep -Seconds $global:AutoPilotWaitSec
                            $retryModule = $true
                        }
                        else {
                            Show-Error "[AUTOPILOT] ErrorMode=Retry -> max retry ($script:AutoPilotMaxRetry) reached, recording $($result.Status)"
                        }
                    }
                    else {
                        # Ask / empty = legacy interactive dialog
                        $dialogChoice = Show-AutoPilotErrorDialog `
                            -ModuleName $module.MenuName `
                            -Status $result.Status `
                            -Message $result.Message
                        if ($dialogChoice -eq "Retry") {
                            $retryModule = $true
                        }
                        # Skip: fall through to record original Error/Partial status
                    }
                }
            }
        } while ($retryModule)

        # Verified field (Post-Apply Verification)
        $verifiedStr = ""
        if ($null -ne $result.Verified) {
            $verifiedStr = if ($result.Verified) { "True" } else { "False" }
        }

        Add-ExecutionResult -Operation $module.MenuName -Status $result.Status -Message $result.Message -Verified $result.Verified -Order $module.Order
        $null = Write-ExecutionHistory -ModuleName $module.MenuName -Category $module.Category -Status $result.Status -Message $result.Message -Verified $verifiedStr -Order $module.Order
        Capture-ScreenEvidence -ModuleName $module.MenuName -Status $result.Status

        # Track completed results for resume state and FlexProfile feedback.
        # Save-ResumeState reshapes to {MenuName, Status} on serialization;
        # extra keys are dropped. Flex dashboard reads the full hashtable
        # via $script:LastBatchResults (published in finally below).
        $completedResults += @{
            Order    = $module.Order
            MenuName = $module.MenuName
            Status   = $result.Status
            Verified = $result.Verified
            Message  = $result.Message
        }

        # Feed this result into the live gate status map so a downstream
        # __GATE__ sees it on the next iteration (dynamic evaluation).
        # Verified (PASS/FAIL/$null) is tracked too: a Verified=FAIL also
        # trips the gate even when Status is Success.
        $statusByOrder[[int]$module.Order] = $result.Status
        $verifiedByOrder[[int]$module.Order] = $result.Verified

        # Update cross-module dependency tracker for next iteration's
        # envelope.start (consumed via $global:_FabriqCurrentProfileContext).
        $prevModuleName   = $module.MenuName
        $prevModuleStatus = $result.Status
    }

    # __GATE__: surface a single summary when downstream modules were
    # blocked, so the operator knows execution stopped at a gate rather
    # than completing the whole selection.
    if ($gateBlockedOrders.Count -gt 0) {
        Show-Warning "[GATE] $($gateBlockedOrders.Count) module(s) blocked by an unsatisfied gate (Orders: $($gateBlockedOrders -join ', ')). Resolve the upstream failure(s) and re-run."
    }

    # All modules completed (no restart, or all restarts done)
    Remove-ResumeState

    # Telemetry kernel event: profile/batch end (natural completion only;
    # __RESTART__ early-exit takes a different path before this point).
    try {
        $okCnt   = @($completedResults | Where-Object { $_.Status -eq 'Success'   }).Count
        $errCnt  = @($completedResults | Where-Object { $_.Status -eq 'Error'     }).Count
        $skipCnt = @($completedResults | Where-Object { $_.Status -eq 'Skipped'   }).Count
        $partCnt = @($completedResults | Where-Object { $_.Status -eq 'Partial'   }).Count
        $cancCnt = @($completedResults | Where-Object { $_.Status -eq 'Cancelled' }).Count
        Write-KernelTelemetryEvent -Type "profile.end" -Data ([ordered]@{
            profileName    = $ProfileName
            executionMode  = $ExecutionMode
            modulesRun     = $completedResults.Count
            successCount   = $okCnt
            errorCount     = $errCnt
            skippedCount   = $skipCnt
            partialCount   = $partCnt
            cancelledCount = $cancCnt
            outcome        = if ($errCnt -gt 0) { 'WithErrors' } elseif ($partCnt -gt 0) { 'Partial' } else { 'Success' }
        })
    } catch { }

    # Calculate elapsed time as a single subtraction from the absolute
    # profile start timestamp. Naturally includes reboot/login/startup
    # gaps for profiles with __RESTART__ markers.
    $batchElapsed = (Get-Date) - $ProfileStartTime

    Show-ExecutionSummary -ElapsedTime $batchElapsed

    # Auto-export evidence if this is a profile execution.
    # Pipeline (export history / HTML checklist / log_uploader / viewer)
    # is centralized in Complete-ProfileExecution; -Mode 'Auto' preserves
    # Linear finalize behavior (silent upload, viewer last).
    # FlexProfile callers pass -FinalizeOnComplete:$false to skip finalize
    # and let the operator commit explicitly via the [Complete] button.
    if (-not [string]::IsNullOrEmpty($ProfileName) -and $FinalizeOnComplete) {
        $checklistModules = if ($null -ne $FullProfileModules) { $FullProfileModules } else { $SelectedModules }
        $null = Complete-ProfileExecution `
            -ProfileName    $ProfileName `
            -ProfilePath    $ProfilePath `
            -DefinedModules $checklistModules `
            -ElapsedTime    $batchElapsed `
            -Mode           'Auto'
    }

    } # end try
    finally {
        # AutoPilot: always reset (Profile scope guarantee)
        $global:AutoPilotMode = $false
        $global:AutoPilotWaitSec = 3
        # Execution toolbar: return to Idle (disables Skip / Gyotaq).
        # Guarded - a failed toolbar update must never break finalize.
        try { Update-ExecutionToolbar -ExecutionState 'Idle' } catch { }
        # Publish per-module results for FlexProfile dashboard polling.
        # Always set in finally so cancel / mid-throw paths produce a
        # consistent snapshot (possibly partial, possibly empty).
        # Note: __RESTART__ early-exit kills the process before finally
        # runs; the post-restart resume call rewrites this with the
        # post-restart segment only, which is the intended semantic.
        $script:LastBatchResults = $completedResults
        # Clear telemetry profile context so it doesn't leak into ad-hoc
        # module runs invoked outside Invoke-BatchExecution.
        $global:_FabriqCurrentProfileContext = $null
    }
}

# ========================================
# Function: FlexProfile Inner Loop
# ========================================
# Drives the FlexProfile dashboard sub-loop. Called from two entry
# points:
#   (a) main loop "FlexProfile" action  (operator opens from main dashboard)
#   (b) Flex resume bootstrap            (post-reboot path, schemaVersion=2)
#
# The dashboard returns an action intent on each iteration; this function
# dispatches the intent (Run / Complete / RestartNow / ResetState) and
# loops back. Returns when the operator presses Back / closes the form.
#
# Resume state lifecycle inside this loop:
#   - Open with resume_state.json possibly present (Flex resume entry).
#   - "RestartNow" rewrites resume_state.json with fresh ModuleStates.
#   - "Close" (Back) removes resume_state.json defensively.
#   - Anything else leaves resume_state.json alone so a crash mid-session
#     can re-enter the loop on the next launch.
# ========================================
function Invoke-FlexProfileLoop {
    param(
        [Parameter(Mandatory)][string]$ProfilePath,
        [Parameter(Mandatory)][string]$ProfileName,
        [Parameter(Mandatory)][array]$AllModules,
        # When set, the dashboard opens with the "PENDING FINALIZE" badge
        # already visible. Used by the Flex resume entry point: when
        # auto-continue ran a batch on resume, the operator returns to a
        # dashboard that already needs [Complete] pressed.
        [bool]$InitialPendingFinalize = $false
    )

    if (-not (Test-Path $ProfilePath)) {
        Show-Error "FlexProfile: profile not found: $ProfilePath"
        Wait-KeyPress
        return
    }

    $lastFinalizedAt = ""
    # Tracks "batch has run since last [Complete]". Flipped to $true by
    # RunBatch / RunSingle / ResetState and back to $false by Complete.
    # Drives the dashboard's PENDING FINALIZE badge and the Back / X
    # close-confirmation dialog.
    $pendingFinalize = $InitialPendingFinalize
    $exitInner = $false

    while (-not $exitInner) {
        Write-StatusFile -Phase "idle"
        Hide-ConsoleWindow

        $flex = Show-FlexDashboard `
            -ProfilePath      $ProfilePath `
            -ProfileName      $ProfileName `
            -AllModules       $AllModules `
            -LastBatchResults $script:LastBatchResults `
            -LastFinalizedAt  $lastFinalizedAt `
            -PendingFinalize  $pendingFinalize `
            -HostName         $env:SELECTED_NEW_PCNAME `
            -WorkerName       $env:FABRIQ_WORKER_NAME

        Show-ConsoleWindow
        Clear-Host

        switch ($flex.Action) {

            "RunSingle" {
                $resolved = Resolve-ProfileModules -ProfileCsvPath $ProfilePath -AllModules $AllModules -IncludeDisabled
                $tgt = @($resolved.ValidModules | Where-Object { [int]$_.Order -eq [int]$flex.TargetOrder })[0]
                if ($null -eq $tgt) {
                    Show-Error "FlexProfile: target order $($flex.TargetOrder) not found"
                    Wait-KeyPress
                    continue
                }
                Show-Info "Running single module (Order $($tgt.Order)): $($tgt.MenuName)"
                Write-Host ""

                # AutoConfirm so Y/N + Press-Enter are skipped for one-click run.
                # Resets in finally even if Invoke-BatchExecution throws.
                $global:AutoConfirmMode = $true
                try {
                    Invoke-BatchExecution -SelectedModules @($tgt) `
                        -ProfilePath         $ProfilePath `
                        -ProfileName         $ProfileName `
                        -FinalizeOnComplete:$false `
                        -ExecutionMode       'Flex' `
                        -SelectedOrders      @([int]$tgt.Order) `
                        -FullProfileModules  $resolved.ValidModules
                } finally {
                    $global:AutoConfirmMode = $false
                }
                $pendingFinalize = $true
                Write-Host ""
            }

            "RunGroup" {
                # Group execution: same Invoke-BatchExecution pipeline as
                # RunBatch, but SelectedOrders comes from the dashboard's
                # group filter (one-click [Run: <Group>] button) instead
                # of the checkbox state. Per the literal-Group contract,
                # rows whose Group value differs from the clicked group
                # are excluded - including __RESTART__ markers in other
                # groups (operator must put RESTART in the desired group
                # to be honored). FinalizeOnComplete:$false keeps
                # completion as an explicit operator action.
                $resolved = Resolve-ProfileModules -ProfileCsvPath $ProfilePath -AllModules $AllModules -IncludeDisabled
                $batch = @($resolved.ValidModules | Where-Object { [int]$_.Order -in $flex.SelectedOrders } | Sort-Object { [int]$_.Order })
                if ($batch.Count -eq 0) {
                    Show-Warning "FlexProfile: no modules matched group '$($flex.TargetGroup)'"
                    Wait-KeyPress
                    continue
                }
                Show-Info "Running group [$($flex.TargetGroup)] ($($batch.Count) modules)..."
                Write-Host ""

                Invoke-BatchExecution -SelectedModules $batch `
                    -AutoPilot:         $true `
                    -AutoPilotWaitSec   $flex.AutoPilotWaitSec `
                    -ProfilePath        $ProfilePath `
                    -ProfileName        $ProfileName `
                    -FinalizeOnComplete:$false `
                    -ExecutionMode      'Flex' `
                    -SelectedOrders     @($flex.SelectedOrders) `
                    -FullProfileModules $resolved.ValidModules

                $pendingFinalize = $true
                Write-Host ""
            }

            "RunBatch" {
                $resolved = Resolve-ProfileModules -ProfileCsvPath $ProfilePath -AllModules $AllModules -IncludeDisabled
                $batch = @($resolved.ValidModules | Where-Object { [int]$_.Order -in $flex.SelectedOrders } | Sort-Object { [int]$_.Order })
                if ($batch.Count -eq 0) {
                    Show-Warning "FlexProfile: no modules matched the selected orders"
                    Wait-KeyPress
                    continue
                }
                Show-Info "Running batch ($($batch.Count) modules)..."
                Write-Host ""

                # 3.1.5 onward: execution is unconditionally AutoPilot
                # (unattended batch with ErrorMode dispatch / inter-module
                # wait). Finalize is unconditionally manual - operator
                # presses [Complete] when ready. This decouples
                # "execution mode" from "completion declaration".
                Invoke-BatchExecution -SelectedModules $batch `
                    -AutoPilot:         $true `
                    -AutoPilotWaitSec   $flex.AutoPilotWaitSec `
                    -ProfilePath        $ProfilePath `
                    -ProfileName        $ProfileName `
                    -FinalizeOnComplete:$false `
                    -ExecutionMode      'Flex' `
                    -SelectedOrders     @($flex.SelectedOrders) `
                    -FullProfileModules $resolved.ValidModules

                $pendingFinalize = $true
                Write-Host ""
            }

            "Complete" {
                $resolved = Resolve-ProfileModules -ProfileCsvPath $ProfilePath -AllModules $AllModules -IncludeDisabled
                $null = Complete-ProfileExecution `
                    -ProfileName    $ProfileName `
                    -ProfilePath    $ProfilePath `
                    -DefinedModules $resolved.ValidModules `
                    -Mode           'Manual'
                $lastFinalizedAt = (Get-Date).ToString("HH:mm:ss")
                $pendingFinalize = $false
                Wait-KeyPress
            }

            "RestartNow" {
                $resolved = Resolve-ProfileModules -ProfileCsvPath $ProfilePath -AllModules $AllModules -IncludeDisabled

                # Build current ModuleStates snapshot from execution_history.csv
                # (latest entry per ModuleName, current SessionID).
                $moduleStates = @{}
                $hist = @()
                if (-not [string]::IsNullOrEmpty($env:SELECTED_KANRI_NO)) {
                    $hist = @(Import-ExecutionHistory -FilterKanriNo $env:SELECTED_KANRI_NO | Where-Object { $_.SessionID -eq $script:SessionID })
                }
                $byMenu = @{}
                foreach ($h in $hist) {
                    if (-not $byMenu.ContainsKey($h.ModuleName)) { $byMenu[$h.ModuleName] = $h }
                }
                foreach ($r in $resolved.ValidModules) {
                    $key = "$($r.Order)"
                    if ($byMenu.ContainsKey($r.MenuName)) {
                        $h = $byMenu[$r.MenuName]
                        $moduleStates[$key] = @{
                            Status   = $h.Status
                            Verified = $h.Verified
                            Message  = $h.Message
                        }
                    } else {
                        $moduleStates[$key] = @{
                            Status   = 'Pending'
                            Verified = ''
                            Message  = ''
                        }
                    }
                }

                # ResumeAfterOrder=-1 sentinel: post-reboot path simply
                # reopens the FlexProfile dashboard (no Linear-style
                # auto-continuation of "Order > N" remaining modules).
                Save-ResumeState `
                    -ProfilePath        $ProfilePath `
                    -ProfileName        $ProfileName `
                    -ResumeAfterOrder   -1 `
                    -CompletedModules   @() `
                    -ExecutionMode      'Flex' `
                    -SelectedOrders     @($flex.SelectedOrders) `
                    -ModuleStates       $moduleStates

                if (-not (Register-FabriqRunOnce)) {
                    Show-Error "Failed to register RunOnce; restart aborted"
                    Remove-ResumeState
                    Wait-KeyPress
                    continue
                }

                # [RESTART NOW] is profile-external - Order=0 marks
                # "no Profile row association" (CSV cell stays empty).
                Add-ExecutionResult -Operation "[RESTART NOW]" -Status "Success" -Message "FlexProfile Restart Now" -Order 0
                $null = Write-ExecutionHistory -ModuleName "[RESTART NOW]" -Category "System" -Status "Success" -Message "FlexProfile Restart Now (dashboard reopen on resume)" -Order 0

                # Hide toolbar cleanly before reboot kills the process
                try { Hide-ExecutionToolbar } catch { }

                Invoke-CountdownRestart
                # Process exits here (Restart-Computer in Invoke-CountdownRestart).
                return
            }

            "ResetState" {
                $resolved = Resolve-ProfileModules -ProfileCsvPath $ProfilePath -AllModules $AllModules -IncludeDisabled
                $tgt = @($resolved.ValidModules | Where-Object { [int]$_.Order -eq [int]$flex.ResetTargetOrder })[0]
                if ($null -ne $tgt) {
                    # Order is critical here so per-row state matching
                    # picks up the Pending entry against the correct
                    # Profile row (not its MenuName twin).
                    Add-ExecutionResult -Operation $tgt.MenuName -Status "Pending" -Message "Reset by operator" -Order $tgt.Order
                    $null = Write-ExecutionHistory -ModuleName $tgt.MenuName -Category $tgt.Category -Status "Pending" -Message "Reset by operator" -Order $tgt.Order

                    # The dashboard overlays $script:LastBatchResults LAST
                    # (highest precedence), so a stale entry from the prior
                    # batch would mask this reset until the next run
                    # republishes it. Update the matching entry in place so
                    # the reset is reflected immediately (TM t-0080).
                    foreach ($lbr in $script:LastBatchResults) {
                        if ([int]$lbr.Order -eq [int]$tgt.Order) {
                            $lbr.Status   = "Pending"
                            $lbr.Verified = $null
                            $lbr.Message  = "Reset by operator"
                        }
                    }

                    Show-Info "Reset state: Order $($tgt.Order) ($($tgt.MenuName)) -> Pending"
                    # State changed since last [Complete] (if any) -
                    # mark pending so the operator is reminded to
                    # regenerate the checklist.
                    $pendingFinalize = $true
                }
            }

            "Close" {
                # Defensive cleanup: any lingering Flex resume state from
                # entry into this loop is consumed at this point. Linear
                # paths never have resume_state alive at this moment, so
                # this is a no-op for them (Remove-ResumeState is guarded
                # by Test-Path).
                Remove-ResumeState
                $exitInner = $true
            }

            default {
                $exitInner = $true
            }
        }
    }
}

# ========================================
# Function: Windows Update AutoLogon Helper
# ========================================
# Sets AutoLogon with enough count to survive WU reboot loops.
# CBS finalization after WU install consumes 1 AutoLogonCount per reboot,
# so we need MaxLoops * 2 (1 for CBS + 1 for actual user logon per loop).
# Credentials are cleaned up by Clear-WindowsUpdateAutoLogon when WU completes.
function Set-WindowsUpdateAutoLogon {
    param([int]$MaxLoops = 5)

    $alCsvPath = Join-Path (Resolve-Path ".").Path "modules\standard\autologon_config\autologon_list.csv"
    if (-not (Test-Path $alCsvPath)) {
        Show-Warning "AutoLogon: autologon_list.csv not found"
        return
    }
    try {
        $alEntries = Import-ModuleCsv -Path $alCsvPath -RequiredColumns @("Enabled", "No", "User", "Password")
        $alEnabled = @($alEntries | Where-Object { $_.Enabled -eq "1" })
        $currentUser = $env:USERNAME
        $alTarget = $alEnabled | Where-Object { $_.User -eq $currentUser } | Select-Object -First 1

        if ($alTarget) {
            $winlogonPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
            $logonCount = $MaxLoops * 2
            Set-ItemProperty -Path $winlogonPath -Name "AutoAdminLogon" -Value "1" -Type String -Force -ErrorAction Stop
            Set-ItemProperty -Path $winlogonPath -Name "DefaultUserName" -Value $alTarget.User -Type String -Force -ErrorAction Stop
            Set-ItemProperty -Path $winlogonPath -Name "DefaultPassword" -Value $alTarget.Password -Type String -Force -ErrorAction Stop
            Set-ItemProperty -Path $winlogonPath -Name "AutoLogonCount" -Value $logonCount -Type DWord -Force -ErrorAction Stop
            if (-not [string]::IsNullOrWhiteSpace($alTarget.Domain)) {
                Set-ItemProperty -Path $winlogonPath -Name "DefaultDomainName" -Value $alTarget.Domain -Type String -Force -ErrorAction Stop
            }
            Show-Success "AutoLogon configured for '$currentUser' (count=$logonCount)"
        }
        else {
            Show-Warning "AutoLogon: no matching entry for '$currentUser' in autologon_list.csv"
        }
    }
    catch {
        Show-Warning "AutoLogon configuration failed (non-fatal): $_"
    }
}

# ========================================
# Function: Clear Windows Update AutoLogon
# ========================================
# Removes AutoLogon registry values after WU loop completes.
# Ensures no credentials remain in the registry.
function Clear-WindowsUpdateAutoLogon {
    $winlogonPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
    try {
        Set-ItemProperty -Path $winlogonPath -Name "AutoAdminLogon" -Value "0" -Type String -Force -ErrorAction SilentlyContinue
        Remove-ItemProperty -Path $winlogonPath -Name "DefaultPassword" -Force -ErrorAction SilentlyContinue
        Remove-ItemProperty -Path $winlogonPath -Name "AutoLogonCount" -Force -ErrorAction SilentlyContinue
        Show-Info "AutoLogon credentials cleared"
    }
    catch {
        Show-Warning "Failed to clear AutoLogon (non-fatal): $_"
    }
}

# ========================================
# Function: Windows Update Orchestration Loop
# ========================================
# Manages the WU reboot loop using the same pattern as Profile __RESTART__:
#   Register-FabriqRunOnce -> AutoLogon -> Invoke-CountdownRestart
# WU module (windows_update.ps1) performs a single pass per call.
function Invoke-WindowsUpdateLoop {
    $wuScript = Join-Path (Resolve-Path ".").Path "modules\standard\windows_update\windows_update.ps1"
    $wuStatePath = Join-Path (Resolve-Path ".").Path "modules\standard\windows_update\wu_state.json"
    $wuConfigPath = Join-Path (Resolve-Path ".").Path "modules\standard\windows_update\windows_update_list.csv"

    if (-not (Test-Path $wuScript)) {
        Show-Error "windows_update.ps1 not found: $wuScript"
        return
    }

    # --- Load or initialize WU state ---
    if (Test-Path $wuStatePath) {
        try {
            $stateRaw = Get-Content $wuStatePath -Raw -Encoding UTF8 | ConvertFrom-Json
            $wuState = @{
                LoopCount    = [int]$stateRaw.LoopCount
                MaxLoops     = [int]$stateRaw.MaxLoops
                InstalledKBs = @($stateRaw.InstalledKBs)
                FailedKBs    = @($stateRaw.FailedKBs)
                StartTime    = $stateRaw.StartTime
                RebootSec    = if ($stateRaw.RebootSec) { [int]$stateRaw.RebootSec } else { 15 }
                AutoLogon    = if ($null -ne $stateRaw.AutoLogon) { [bool]$stateRaw.AutoLogon } else { $true }
            }
            Show-Info "Resumed WU loop (Loop $($wuState.LoopCount) of $($wuState.MaxLoops))"
        }
        catch {
            Show-Warning "Failed to load wu_state.json, starting fresh: $_"
            Remove-Item $wuStatePath -Force -ErrorAction SilentlyContinue
            $wuState = $null
        }
    }

    if ($null -eq $wuState) {
        # Read config from CSV
        $maxLoops = 5
        $rebootSec = 15
        $autoLogon = $true
        if (Test-Path $wuConfigPath) {
            try {
                $cfgItems = Import-ModuleCsv -Path $wuConfigPath -FilterEnabled -RequiredColumns @("Enabled", "SettingName", "Value")
                $cfgMap = @{}
                foreach ($c in $cfgItems) { $cfgMap[$c.SettingName] = $c.Value }
                if ($cfgMap["MaxRebootLoops"]) { $maxLoops = [int]$cfgMap["MaxRebootLoops"] }
                if ($cfgMap["RebootCountdownSeconds"]) { $rebootSec = [int]$cfgMap["RebootCountdownSeconds"] }
                if ($cfgMap["AutoLogonEnabled"] -eq "0") { $autoLogon = $false }
            }
            catch { }
        }
        $wuState = @{
            LoopCount    = 0
            MaxLoops     = $maxLoops
            InstalledKBs = @()
            FailedKBs    = @()
            StartTime    = (Get-Date).ToString("o")
            RebootSec    = $rebootSec
            AutoLogon    = $autoLogon
        }
    }

    # --- Safety valve ---
    if ($wuState.LoopCount -ge $wuState.MaxLoops) {
        Show-Warning "Max WU loop limit reached ($($wuState.MaxLoops)). Stopping."
        Write-Host ""
        if ($wuState.AutoLogon) { Clear-WindowsUpdateAutoLogon }
        Show-WindowsUpdateSummary -WuState $wuState
        Remove-Item $wuStatePath -Force -ErrorAction SilentlyContinue
        return
    }

    # --- Build skip list: KBs installed 2+ times (phantom re-appearing updates) ---
    $kbCounts = @{}
    foreach ($ik in $wuState.InstalledKBs) {
        $k = if ($ik.KB) { $ik.KB } else { continue }
        if ($kbCounts.ContainsKey($k)) { $kbCounts[$k]++ } else { $kbCounts[$k] = 1 }
    }
    $skipKBs = @($kbCounts.GetEnumerator() | Where-Object { $_.Value -ge 3 } | ForEach-Object { $_.Key })

    if ($skipKBs.Count -gt 0) {
        Show-Warning "Skipping $($skipKBs.Count) re-appearing KB(s): $($skipKBs -join ', ')"
        Write-Host ""
    }

    # --- Run WU single pass ---
    $autoConfirm = ($wuState.LoopCount -gt 0)  # Auto-confirm on resumed loops
    $result = & $wuScript -SkipKBs $skipKBs -AutoConfirm:$autoConfirm

    if ($null -eq $result -or $result.Status -eq "Cancelled") {
        # Operator aborted the loop. On a resumed loop (LoopCount > 0) the
        # previous reboot leg configured AutoLogon via
        # Set-WindowsUpdateAutoLogon, so clear those credentials here; on a
        # fresh first pass (LoopCount = 0) WU has written nothing yet and
        # the registry must stay untouched (autologon_config may own it).
        if ($wuState.AutoLogon -and $wuState.LoopCount -gt 0) {
            Clear-WindowsUpdateAutoLogon
        }
        Remove-Item $wuStatePath -Force -ErrorAction SilentlyContinue
        return
    }

    # --- Track results ---
    $wuState.LoopCount++
    $wuState.InstalledKBs = @($wuState.InstalledKBs) + @($result.InstalledKBs | ForEach-Object {
        @{ KB = $_.KB; Title = $_.Title; Loop = $wuState.LoopCount }
    })
    $wuState.FailedKBs = @($wuState.FailedKBs) + @($result.FailedKBs | ForEach-Object {
        @{ KB = $_.KB; Title = $_.Title; HResult = $_.HResult; Loop = $wuState.LoopCount }
    })

    # Record in execution history
    $loopMsg = "Loop $($wuState.LoopCount): $($result.InstalledCount) installed, $($result.FailedCount) failed"
    Add-ExecutionResult -Operation "Windows Update" -Status $result.Status -Message $loopMsg
    $null = Write-ExecutionHistory -ModuleName "Windows Update" -Category "Maintenance" -Status $result.Status -Message $loopMsg

    # --- Reboot decision (same as Profile __RESTART__) ---
    if ($result.RebootRequired -and $result.InstalledCount -gt 0 -and $wuState.LoopCount -lt $wuState.MaxLoops) {

        # Save WU state
        $wuState | ConvertTo-Json -Depth 5 | Out-File -FilePath $wuStatePath -Encoding UTF8 -Force
        Show-Info "WU state saved: Loop $($wuState.LoopCount), Total KBs: $($wuState.InstalledKBs.Count)"

        # AutoLogon (same pattern as Profile, count covers CBS consumption)
        if ($wuState.AutoLogon) {
            Set-WindowsUpdateAutoLogon -MaxLoops $wuState.MaxLoops
        }

        # Register RunOnce (same function as Profile __RESTART__)
        if (-not (Register-FabriqRunOnce)) {
            # Set-WindowsUpdateAutoLogon ran just above under the same
            # AutoLogon condition; without RunOnce there is no resume leg
            # to consume or clear the credentials, so clear them before
            # bailing out (otherwise DefaultPassword stays in the registry
            # with AutoLogonCount intact).
            if ($wuState.AutoLogon) {
                Clear-WindowsUpdateAutoLogon
            }
            Remove-Item $wuStatePath -Force -ErrorAction SilentlyContinue
            Show-Error "Failed to register RunOnce for WU resume"
            return
        }

        # Hide toolbar cleanly before reboot kills the process
        try { Hide-ExecutionToolbar } catch { }

        # Reboot (same function as Profile __RESTART__)
        Invoke-CountdownRestart -Seconds $wuState.RebootSec
        return  # Process ends here
    }

    # --- All done (no reboot needed, or all updates complete) ---

    # Clean up AutoLogon credentials (remove remaining count + password)
    if ($wuState.AutoLogon) {
        Clear-WindowsUpdateAutoLogon
    }

    Show-WindowsUpdateSummary -WuState $wuState
    Remove-Item $wuStatePath -Force -ErrorAction SilentlyContinue

    # Save wu_completed.json for session import
    $completedPath = Join-Path (Resolve-Path ".").Path "modules\standard\windows_update\wu_completed.json"
    $completedData = [PSCustomObject]@{
        TotalInstalled = ($wuState.InstalledKBs | Measure-Object).Count
        TotalFailed    = ($wuState.FailedKBs | Measure-Object).Count
        TotalLoops     = $wuState.LoopCount
        ElapsedMinutes = [math]::Round(((Get-Date) - [datetime]$wuState.StartTime).TotalMinutes, 1)
        InstalledKBs   = @($wuState.InstalledKBs | ForEach-Object { $_.KB })
        FailedKBs      = @($wuState.FailedKBs)
        CompletedAt    = (Get-Date).ToString("o")
    } | ConvertTo-Json -Depth 5
    $completedData | Out-File -FilePath $completedPath -Encoding UTF8 -Force
    Show-Info "Completion results saved: wu_completed.json"
    Write-Host ""
}

# ========================================
# Function: Windows Update Summary Display
# ========================================
function Show-WindowsUpdateSummary {
    param([hashtable]$WuState)

    $allKBs = @($WuState.InstalledKBs)
    $allFailed = @($WuState.FailedKBs)
    $elapsed = (Get-Date) - [datetime]$WuState.StartTime
    $elapsedMinutes = [math]::Round($elapsed.TotalMinutes, 1)

    Show-Separator
    Write-Host "Windows Update Complete" -ForegroundColor Cyan
    Show-Separator
    Write-Host ""
    Write-Host "  Total loops:     $($WuState.LoopCount)" -ForegroundColor White
    Write-Host "  Total installed: $($allKBs.Count) updates" -ForegroundColor Green
    Write-Host "  Total failed:    $($allFailed.Count) updates" -ForegroundColor $(if ($allFailed.Count -gt 0) { "Red" } else { "White" })
    Write-Host "  Elapsed time:    ${elapsedMinutes} minutes" -ForegroundColor White
    Write-Host ""

    if ($WuState.LoopCount -ge $WuState.MaxLoops) {
        Show-Warning "Max loop limit reached. Additional updates may remain."
        Write-Host ""
    }

    if ($allKBs.Count -gt 0) {
        Write-Host "  Installed updates:" -ForegroundColor DarkGray
        foreach ($kb in $allKBs) {
            $kbId = if ($kb.KB) { $kb.KB } else { "N/A" }
            $kbTitle = if ($kb.Title) { $kb.Title } else { "Unknown" }
            Write-Host "    [Loop $($kb.Loop)] $kbId - $kbTitle" -ForegroundColor DarkGray
        }
        Write-Host ""
    }

    if ($allFailed.Count -gt 0) {
        Write-Host "  Failed updates:" -ForegroundColor Red
        foreach ($fk in $allFailed) {
            $fkTitle = if ($fk.Title) { $fk.Title } else { "Unknown" }
            $fkHR = if ($fk.HResult) { $fk.HResult } else { "N/A" }
            Write-Host "    $($fk.KB) - $fkTitle" -ForegroundColor Red
            Write-Host "      HResult: $fkHR (Loop $($fk.Loop))" -ForegroundColor DarkGray
        }
        Write-Host ""
    }

    Show-Separator
    Write-Host ""
}

# ========================================
# Main Process
# ========================================

# Start logging
$logDir = ".\logs"
if (-not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
}
$timestamp  = Get-Date -Format "yyyy_MM_dd_HHmmss"
$uniqueId   = Get-HardwareUniqueId
$hostname   = $env:COMPUTERNAME
$logFile    = Join-Path $logDir "${timestamp}_${uniqueId}_${hostname}.log"
$global:FabriqTranscriptPath   = $logFile
$global:FabriqUniqueId         = $uniqueId   # hardware unique ID (BIOS SN or MAC)
$global:FabriqSessionTimestamp = $timestamp  # session start time (yyyy_MM_dd_HHmmss)
Start-Transcript -Path $logFile -Append | Out-Null

Write-Host ""
Show-Separator
Write-Host "Fabriq ver3.6 - Manifeste du Surkitinisme - " -ForegroundColor Green
Show-Separator
Write-Host ""
Show-Info "Log file: $logFile"
Write-Host ""

# ========================================
# Resume Detection (must run BEFORE passphrase prompt)
# ========================================
$global:FabriqMasterPassphrase = $null
$isResuming = $false
$isWuResuming = $false
$resumeState = Load-ResumeState

# Windows Update resume detection (wu_state.json presence = mid-loop reboot)
# WU resume runs immediately - no passphrase, worker, or host selection needed.
# Uses Invoke-WindowsUpdateLoop which follows the Profile __RESTART__ pattern.
# After WU completes (all loops done), main.ps1 continues normal startup.
$wuStatePath = Join-Path $PSScriptRoot "..\modules\standard\windows_update\wu_state.json"
if (Test-Path $wuStatePath) {
    $isWuResuming = $true

    Initialize-ExecutionHistory

    Invoke-WindowsUpdateLoop
    # If WU rebooted, execution does not reach here (process ends at Invoke-CountdownRestart).
    # If WU completed all loops, continue to normal startup below.
    $isWuResuming = $false
    Write-Host ""
}

$isFlexResuming = $false
$flexAutoContinue = $false
if ($null -ne $resumeState) {
    # Detect resume_state.json schemaVersion. Absent / pre-P4 = legacy v1
    # (Linear). schemaVersion>=2 + ExecutionMode='Flex' = FlexProfile resume.
    $resumeSchemaVer = if ($null -ne $resumeState.schemaVersion) { [int]$resumeState.schemaVersion } else { 1 }
    $isFlexResuming = ($resumeSchemaVer -ge 2 -and $resumeState.ExecutionMode -eq 'Flex')

    Write-Host ""
    Write-Host "========================================" -ForegroundColor Yellow
    if ($isFlexResuming) {
        Write-Host "  FlexProfile Resume Detected" -ForegroundColor Yellow
    } else {
        Write-Host "  Profile Resume Detected" -ForegroundColor Yellow
    }
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  Profile:  $($resumeState.ProfileName)" -ForegroundColor White
    Write-Host "  PC:       $($resumeState.HostEnvironment.SELECTED_NEW_PCNAME)" -ForegroundColor White

    $completedCount = @($resumeState.CompletedModules).Count
    Write-Host "  Progress: $completedCount modules completed" -ForegroundColor White
    Write-Host ""

    $resumeIsAutoPilot = ($resumeState.AutoPilot -eq $true)

    if ($isFlexResuming) {
        # FlexProfile resume branches on (AutoPilot, ResumeAfterOrder):
        #   - AutoPilot=true  + ResumeAfterOrder>=0 (mid-batch __RESTART__)
        #     -> honor unattended contract: countdown + auto-continue
        #     execution (Linear-symmetric); fall through to Flex
        #     auto-continue execution block below.
        #   - AutoPilot=false (manual mode mid-batch) OR
        #     ResumeAfterOrder=-1 ([Restart Now] sentinel)
        #     -> reopen the FlexProfile dashboard so the operator decides
        #     what to run next.
        $flexHasMidBatch = ([int]$resumeState.ResumeAfterOrder -ge 0)
        if ($resumeIsAutoPilot -and $flexHasMidBatch) {
            Wait-SystemReady
            $flexAutoContinue = Invoke-AutoResumeCountdown -Seconds 60
            # If countdown is aborted, $flexAutoContinue is $false and
            # we fall back to dashboard-reopen behavior. Either way the
            # environment must be restored, so $shouldResume stays true.
            $shouldResume = $true
        }
        else {
            Show-Info "FlexProfile resume: dashboard will reopen after environment restore"
            $shouldResume = $true
        }
    }
    elseif ($resumeIsAutoPilot) {
        # AutoPilot resume: wait for system stability, then countdown
        Wait-SystemReady
        $shouldResume = Invoke-AutoResumeCountdown -Seconds 60
    }
    else {
        # Manual profile: keep existing Y/N prompt
        $shouldResume = Confirm-Execution -Message "Resume profile execution?"
    }

    if ($shouldResume) {
        $isResuming = $true
        Restore-HostEnvironment -HostEnv $resumeState.HostEnvironment
        Show-Success "Environment restored for: $($resumeState.HostEnvironment.SELECTED_NEW_PCNAME)"

        # Restore the host PIN from its DPAPI-protected field. New-format
        # resume states exclude the PIN from the HostEnvironment snapshot;
        # old files with an inline plaintext PIN were restored verbatim
        # above (backward compatible).
        if (-not [string]::IsNullOrWhiteSpace($resumeState.ProtectedPin)) {
            try {
                $env:SELECTED_PIN = Unprotect-PassphraseFromResume -ProtectedBase64 $resumeState.ProtectedPin
            }
            catch {
                Show-Warning "Failed to restore PIN from DPAPI: $_ (PIN unavailable this session)"
            }
        }
        $script:SessionID = $resumeState.SessionID

        # Restore evidence base path from resume state (or fallback to new generation)
        if (-not [string]::IsNullOrWhiteSpace($resumeState.EvidenceBasePath)) {
            $global:FabriqEvidenceBasePath = $resumeState.EvidenceBasePath
            $global:FabriqEvidenceRootPath = Split-Path $resumeState.EvidenceBasePath -Parent
            $env:FABRIQ_EVIDENCE_BASE     = $resumeState.EvidenceBasePath
            Show-Success "Evidence base path restored: $($resumeState.EvidenceBasePath)"
        }
        else {
            Initialize-EvidenceBasePath
        }

        # Restore master passphrase from DPAPI-protected resume state
        if (-not [string]::IsNullOrWhiteSpace($resumeState.ProtectedPassphrase)) {
            try {
                $global:FabriqMasterPassphrase = Unprotect-PassphraseFromResume -ProtectedBase64 $resumeState.ProtectedPassphrase
                Show-Success "Master passphrase restored from resume state"
            }
            catch {
                Show-Warning "Failed to restore passphrase from DPAPI: $_"
                Show-Info "Manual passphrase entry required to continue."
                Write-Host ""

                # Fallback: prompt for manual passphrase entry
                # ($verifyTokenPath is defined in the Constants section)
                if (-not (Test-Path $verifyTokenPath)) {
                    Show-Error "Passphrase verification token not found: $verifyTokenPath"
                    Show-Error "Cannot verify passphrase. Aborting."
                    exit 1
                }

                $maxAttempts = 3
                $passphraseAccepted = $false

                for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
                    $ppInput = Read-Host -Prompt "Master Passphrase"

                    if ([string]::IsNullOrWhiteSpace($ppInput)) {
                        $remaining = $maxAttempts - $attempt
                        if ($remaining -gt 0) {
                            Show-Warning "Passphrase is required. $remaining attempt(s) remaining."
                        }
                        continue
                    }

                    if (Test-MasterPassphrase -Passphrase $ppInput -VerifyTokenPath $verifyTokenPath) {
                        $global:FabriqMasterPassphrase = $ppInput
                        Show-Success "Master passphrase verified and set for this session"
                        $passphraseAccepted = $true
                        break
                    }
                    $remaining = $maxAttempts - $attempt
                    if ($remaining -gt 0) {
                        Show-Warning "Passphrase verification failed. $remaining attempt(s) remaining."
                    }
                }
                if (-not $passphraseAccepted) {
                    Show-Error "Passphrase verification failed $maxAttempts times. Aborting."
                    exit 1
                }
            }
        }
    }
    else {
        Remove-ResumeState
        Show-Info "Resume state cleared. Starting normally."
    }
    Write-Host ""
}

# ========================================
# Master Passphrase + Worker + Host Selection
# ========================================
if (-not $isResuming) {
    # ($verifyTokenPath is defined in the Constants section)
    if (-not (Test-Path $verifyTokenPath)) {
        Show-Error "Passphrase verification token not found: $verifyTokenPath"
        Show-Error "Run Fabriq Studio to generate the verification token."
        exit 1
    }

    # --- GUI Session Setup ---
    $hostList = Load-HostList
    if (-not $hostList) {
        Show-Error "Aborting process"
        Exit-Fabriq
        exit 1
    }

    $workers = @()
    if (Test-Path $script:WorkersCsvPath) {
        try { $workers = @(Import-Csv -Path $script:WorkersCsvPath -Encoding Default) } catch { }
    }

    Hide-ConsoleWindow

    $sessionSetup = Show-SessionSetupForm `
        -Workers $workers `
        -HostList $hostList `
        -VerifyTokenPath $verifyTokenPath `
        -CurrentPCName $env:COMPUTERNAME

    Show-ConsoleWindow

    if ($sessionSetup.Cancelled) {
        Exit-Fabriq
        exit 0
    }

    # Apply session results
    $global:FabriqMasterPassphrase = $sessionSetup.MasterPassphrase
    Show-Success "Master passphrase verified and set for this session"

    # Build and save session info (mirrors Initialize-Session logic)
    $mediaSerial = ""
    if (Test-Path $script:SourceMediaIdPath) {
        try { $mediaSerial = (Get-Content $script:SourceMediaIdPath -Raw -ErrorAction Stop).Trim() } catch { }
    }
    if ([string]::IsNullOrWhiteSpace($mediaSerial)) {
        $currentDrive = (Resolve-Path ".").Drive.Name + ":"
        $mediaSerial = Get-VolumeSerial -DriveLetter $currentDrive
    }

    $script:SessionInfo = [PSCustomObject]@{
        WorkerID     = $sessionSetup.WorkerID
        WorkerName   = $sessionSetup.WorkerName
        MediaSerial  = $mediaSerial
        StartTime    = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        WindowsUser  = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
        ComputerName = $env:COMPUTERNAME
    }
    try {
        $script:SessionInfo | ConvertTo-Json -Depth 3 | Out-File -FilePath $script:SessionFilePath -Encoding UTF8 -Force
    } catch { Show-Warning "Failed to save session.json: $_" }
    $env:FABRIQ_WORKER_NAME = $sessionSetup.WorkerName
    Show-Success "Session initialized: Worker=$($sessionSetup.WorkerName)"

    # Set host environment
    $selectedHost = $sessionSetup.SelectedHost
    Set-SelectedHostEnvironment -SelectedHost $selectedHost
    Initialize-EvidenceBasePath
    Write-Host ""
}

# Clear history CSV on fresh startup
if (-not $isResuming) {
    $historyFile = ".\logs\history\execution_history.csv"
    if (Test-Path $historyFile) {
        Remove-Item $historyFile -Force -ErrorAction SilentlyContinue
    }
    if (Test-Path "$historyFile.bak") {
        Remove-Item "$historyFile.bak" -Force -ErrorAction SilentlyContinue
    }
}

# Initialize history (Create backup) - both paths need this
if (-not $isResuming) {
    Initialize-ExecutionHistory
}

# Initialize session for resume path
if ($isResuming) {
    $sessionResult = Initialize-Session
    if ($sessionResult -eq $false) {
        Show-Warning "Session info unavailable (session.json not found)"
    }
}

# Restore execution history for the selected PC
Restore-ExecutionHistory
Write-Host ""

# Import Windows Update completion results if available
$wuCompletedPath = Join-Path $PSScriptRoot "..\modules\standard\windows_update\wu_completed.json"
if (Test-Path $wuCompletedPath) {
    try {
        $wuResult = Get-Content $wuCompletedPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $wuMsg = "$($wuResult.TotalInstalled) updates installed ($($wuResult.TotalLoops) loops, $($wuResult.ElapsedMinutes)min)"
        Write-ExecutionHistory -ModuleName "Windows Update (Pre-session)" -Category "Maintenance" -Status "Success" -Message $wuMsg
        Remove-Item $wuCompletedPath -Force -ErrorAction SilentlyContinue
        Show-Success "Windows Update results imported: $wuMsg"
        Write-Host ""
    }
    catch {
        Show-Warning "Failed to import Windows Update results: $_"
        Write-Host ""
    }
}

# ========================================
# Module System Initialization
# ========================================
Clear-Host
$moduleSystem = Initialize-ModuleSystem
if ($null -eq $moduleSystem) {
    if ($isResuming) { Remove-ResumeState }
    Stop-Transcript | Out-Null
    exit 1
}
$allModules = $moduleSystem.AllModules
$groupedModules = $moduleSystem.GroupedModules

# ========================================
# Verbose Stream Capture (standard deployment, default ON via kernel/json/verbose_capture.flag)
# ========================================
# Flag is shipped present in git so standard deployments capture cmdlet.verbose
# events automatically. Deleting the flag is the documented opt-out path.
# Invoke-SafeCommand wraps each module run with $VerbosePreference='Continue'
# + $PSDefaultParameterValues['*:Verbose']=$true + 4>&1 stream redirect so
# cmdlet verbose output is routed to telemetry as cmdlet.verbose events.
$null = Enable-FabriqVerboseCapture

# ========================================
# Execution Toolbar
# ========================================
# In-process toolbar on a dedicated STA Runspace within the kernel
# powershell.exe. Replaced the out-of-process Status Monitor in 3.4.0
# (Defender / ASR heuristics blocked its hidden-child spawn pattern);
# the deprecated monitor (status_monitor.ps1, Start/Stop-StatusMonitor)
# was removed in 3.5.0 as scheduled in KERNEL_API.md section 6.
try {
    Show-ExecutionToolbar
}
catch {
    Show-Warning "Failed to show execution toolbar: $($_.Exception.Message)"
}


# ========================================
# Resume Execution (Linear path)
# ========================================
# Runs only for the Linear resume case. FlexProfile resume runs in its
# own block below ($isFlexResuming branch).
if ($isResuming -and -not $isFlexResuming) {
    $validation = Resolve-ProfileModules -ProfileCsvPath $resumeState.ProfilePath -AllModules $allModules
    $remainingModules = @($validation.ValidModules | Where-Object {
        $_.Order -gt $resumeState.ResumeAfterOrder
    })

    if ($remainingModules.Count -eq 0) {
        Show-Info "No remaining modules to execute"
        Remove-ResumeState
    }
    else {
        Show-Info "Resuming profile: $($resumeState.ProfileName)"
        Show-Info "Remaining modules: $($remainingModules.Count)"

        # Restore AutoPilot mode from resume state
        $resumeAutoPilot = ($resumeState.AutoPilot -eq $true)
        $resumeAutoPilotWaitSec = if ($resumeState.AutoPilotWaitSec) { [int]$resumeState.AutoPilotWaitSec } else { 3 }
        if ($resumeAutoPilot) {
            Show-Info "AutoPilot mode: ON (restored from resume state)"
        }
        Write-Host ""

        foreach ($cm in $resumeState.CompletedModules) {
            $cmOrder = if ($null -ne $cm.Order) { [int]$cm.Order } else { 0 }
            Add-ExecutionResult -Operation $cm.MenuName -Status $cm.Status -Message "(completed before restart)" -Order $cmOrder
        }
        $restartOrder = if ($null -ne $resumeState.ResumeAfterOrder) { [int]$resumeState.ResumeAfterOrder } else { 0 }
        Add-ExecutionResult -Operation "[RESTART]" -Status "Success" -Message "Resumed after restart" -Order $restartOrder

        # Restore the absolute profile start timestamp so the post-resume
        # elapsed display covers the full wall-clock duration of the
        # profile (pre-restart + reboot + post-restart). Falls back to
        # "now" if the resume_state.json is from a legacy version that
        # only persisted the obsolete ElapsedSeconds field.
        $resumedProfileStart = $null
        if ($resumeState.ProfileStartTime) {
            try {
                $resumedProfileStart = [datetime]::Parse($resumeState.ProfileStartTime)
            }
            catch {
                Show-Warning "Failed to parse resume_state.json ProfileStartTime ('$($resumeState.ProfileStartTime)'): $_"
            }
        }
        if ($null -eq $resumedProfileStart) {
            Show-Warning "Legacy resume_state.json without ProfileStartTime - elapsed time will be measured from now (pre-restart duration not included)."
            $resumedProfileStart = Get-Date
        }

        Invoke-BatchExecution -SelectedModules $remainingModules `
            -AutoPilot:$resumeAutoPilot `
            -AutoPilotWaitSec $resumeAutoPilotWaitSec `
            -ProfilePath $resumeState.ProfilePath `
            -ProfileName $resumeState.ProfileName `
            -FullProfileModules $validation.ValidModules `
            -ProfileStartTime $resumedProfileStart

        Remove-ResumeState

        # Post-profile completion (same logic as [P] handler)
        Write-Host ""
        Show-Separator

        $hasErrors = @($script:ExecutionResults | Where-Object {
            $_.Status -eq "Error" -and -not $_.IsRestored
        }).Count -gt 0

        if ($hasErrors) {
            Write-Host "Profile Completed with Errors" -ForegroundColor Yellow
            Show-Separator
            Write-Host ""
            Show-Warning "Some modules had errors. You can re-run them from the Script Menu [S]."
            Show-Info "Checklist can be regenerated from [cl] after fixing."
        } else {
            Write-Host "Profile Execution Completed" -ForegroundColor Green
            Show-Separator
        }
        Write-Host ""
        Wait-KeyPress
        Clear-Host
    }
}

# ========================================
# Resume Execution (FlexProfile path)
# ========================================
# Reached when resume_state.json had schemaVersion=2 + ExecutionMode='Flex'.
# Two sub-paths, both ending at the FlexProfile dashboard so the
# operator can press [Complete] manually (3.1.5: finalize is always
# operator-driven, never auto-fired by execution paths):
#   - $flexAutoContinue=true (AutoPilot mid-batch __RESTART__, countdown
#     accepted): re-run the remaining checked subset unattended (preserves
#     the unattended-execution contract for Profile-internal RESTART), then
#     drop into the FlexProfile dashboard with PendingFinalize=$true so the
#     operator returns to a flagged state.
#   - $flexAutoContinue=false (manual mode mid-batch / countdown aborted /
#     [Restart Now] sentinel): no re-execution, dashboard reopens directly.
if ($isFlexResuming) {
    $flexInitialPending = $false

    if ($flexAutoContinue) {
        # Auto-continue: build remaining set as
        # (Order > ResumeAfterOrder) intersect SelectedOrders so we re-run only
        # the operator's previously-checked subset (not the entire
        # post-restart Profile tail).
        $flexResolved = Resolve-ProfileModules -ProfileCsvPath $resumeState.ProfilePath -AllModules $allModules -IncludeDisabled
        $flexSelectedSet = @{}
        foreach ($o in @($resumeState.SelectedOrders)) { $flexSelectedSet[[int]$o] = $true }
        $flexRemaining = @($flexResolved.ValidModules | Where-Object {
            $flexSelectedSet.ContainsKey([int]$_.Order) -and [int]$_.Order -gt [int]$resumeState.ResumeAfterOrder
        } | Sort-Object { [int]$_.Order })

        if ($flexRemaining.Count -eq 0) {
            Show-Info "FlexProfile auto-continue: no remaining checked modules"
            Remove-ResumeState
        }
        else {
            Show-Info "FlexProfile auto-continue: $($flexRemaining.Count) remaining checked modules"

            # Mirror Linear: replay CompletedModules into ExecutionResults
            # for status display continuity across the restart boundary.
            # Order propagated through resume_state v2's CompletedModules
            # reshape (Save-ResumeState now persists Order per entry).
            foreach ($cm in $resumeState.CompletedModules) {
                $cmOrder = if ($null -ne $cm.Order) { [int]$cm.Order } else { 0 }
                Add-ExecutionResult -Operation $cm.MenuName -Status $cm.Status -Message "(completed before restart)" -Order $cmOrder
            }
            $flexRestartOrder = if ($null -ne $resumeState.ResumeAfterOrder) { [int]$resumeState.ResumeAfterOrder } else { 0 }
            Add-ExecutionResult -Operation "[RESTART]" -Status "Success" -Message "Resumed after restart" -Order $flexRestartOrder

            # Restore absolute profile start timestamp for accurate elapsed display
            $flexResumedStart = $null
            if ($resumeState.ProfileStartTime) {
                try {
                    $flexResumedStart = [datetime]::Parse($resumeState.ProfileStartTime)
                }
                catch {
                    Show-Warning "Failed to parse FlexProfile resume_state ProfileStartTime ('$($resumeState.ProfileStartTime)'): $_"
                }
            }
            if ($null -eq $flexResumedStart) { $flexResumedStart = Get-Date }

            $flexResumeWaitSec = if ($resumeState.AutoPilotWaitSec) { [int]$resumeState.AutoPilotWaitSec } else { 3 }

            # 3.1.5: -FinalizeOnComplete:$false. Operator presses [Complete]
            # on the dashboard reopen below.
            Invoke-BatchExecution -SelectedModules $flexRemaining `
                -AutoPilot:$true `
                -AutoPilotWaitSec   $flexResumeWaitSec `
                -ProfilePath        $resumeState.ProfilePath `
                -ProfileName        $resumeState.ProfileName `
                -FullProfileModules $flexResolved.ValidModules `
                -ProfileStartTime   $flexResumedStart `
                -FinalizeOnComplete:$false `
                -ExecutionMode      'Flex' `
                -SelectedOrders     @($resumeState.SelectedOrders)
            # Invoke-BatchExecution removes resume_state on natural completion.

            # Operator returns to a dashboard already flagged for finalize
            $flexInitialPending = $true

            Write-Host ""
            Show-Separator
            $hasErrors = @($script:ExecutionResults | Where-Object {
                $_.Status -eq "Error" -and -not $_.IsRestored
            }).Count -gt 0
            if ($hasErrors) {
                Write-Host "FlexProfile Auto-Continue Completed with Errors" -ForegroundColor Yellow
                Show-Separator
                Write-Host ""
                Show-Warning "Some modules had errors. Press [Complete] on the dashboard once review is done."
            }
            else {
                Write-Host "FlexProfile Auto-Continue Completed" -ForegroundColor Green
                Show-Separator
                Write-Host ""
                Show-Info "Press [Complete] on the dashboard to generate the HTML checklist and upload evidence."
            }
            Write-Host ""
        }
    }

    # Both branches converge here: open the FlexProfile dashboard so the
    # operator can review state and trigger finalize manually. Manual
    # mode / countdown abort / [Restart Now] sentinel start with
    # PendingFinalize=$false; auto-continue starts with $true.
    Invoke-FlexProfileLoop `
        -ProfilePath            $resumeState.ProfilePath `
        -ProfileName            $resumeState.ProfileName `
        -AllModules             $allModules `
        -InitialPendingFinalize $flexInitialPending

    Clear-Host
}

# ========================================
# Main Loop
# ========================================

# Track last execution result summary for GUI status bar
$script:lastGuiResultSummary = ""

# ========================================
# GUI Main Loop (Fabriq Operator Dashboard)
# ========================================
$script:guiExitRequested = $false
    while (-not $script:guiExitRequested) {
        Write-StatusFile -Phase "idle"
        Hide-ConsoleWindow

        $guiSelection = Show-OperatorDashboard `
            -AllModules $allModules `
            -GroupedModules $groupedModules `
            -HostName $env:SELECTED_NEW_PCNAME `
            -WorkerName $env:FABRIQ_WORKER_NAME `
            -LastResultSummary $script:lastGuiResultSummary

        Show-ConsoleWindow
        Clear-Host

        switch ($guiSelection.Action) {

            "FlexProfile" {
                # FlexProfile dashboard sub-loop. Returns when the operator
                # presses Back. P7 wires the entry button on main dashboard.
                Invoke-FlexProfileLoop `
                    -ProfilePath $guiSelection.ProfilePath `
                    -ProfileName $guiSelection.ProfileName `
                    -AllModules  $allModules
            }

            "ExecuteProfile" {
                # Resolve profile modules and execute via existing engine
                $validation = Resolve-ProfileModules -ProfileCsvPath $guiSelection.ProfilePath -AllModules $allModules

                if ($validation.ValidModules.Count -eq 0) {
                    Show-Error "No executable modules found in profile"
                    Wait-KeyPress
                    continue
                }

                Show-Info "Executing profile: $($guiSelection.ProfileName)"
                Write-Host ""

                Invoke-BatchExecution -SelectedModules $validation.ValidModules `
                    -AutoPilot:$guiSelection.AutoPilot `
                    -AutoPilotWaitSec $guiSelection.AutoPilotWaitSec `
                    -ProfilePath $guiSelection.ProfilePath `
                    -ProfileName $guiSelection.ProfileName `
                    -FullProfileModules $validation.ValidModules

                # Build result summary for GUI status bar
                $okCount = @($script:ExecutionResults | Where-Object { $_.Status -eq "Success" -and -not $_.IsRestored }).Count
                $errCount = @($script:ExecutionResults | Where-Object { $_.Status -eq "Error" -and -not $_.IsRestored }).Count
                $skipCount = @($script:ExecutionResults | Where-Object { $_.Status -eq "Skipped" -and -not $_.IsRestored }).Count
                $script:lastGuiResultSummary = "Last: $($guiSelection.ProfileName) - $okCount OK  $errCount NG  $skipCount Skip"

                # Post-profile error check
                Write-Host ""
                Show-Separator
                if ($errCount -gt 0) {
                    Write-Host "Profile Completed with Errors" -ForegroundColor Yellow
                    Show-Separator
                    Write-Host ""
                    Show-Warning "Some modules had errors."
                } else {
                    Write-Host "Profile Execution Completed" -ForegroundColor Green
                    Show-Separator
                }
                Write-Host ""
                Wait-KeyPress
            }

            "ExecuteModules" {
                Show-Info "Executing $($guiSelection.SelectedModules.Count) module(s)..."
                Write-Host ""

                Invoke-BatchExecution -SelectedModules $guiSelection.SelectedModules

                $okCount = @($script:ExecutionResults | Where-Object { $_.Status -eq "Success" -and -not $_.IsRestored }).Count
                $errCount = @($script:ExecutionResults | Where-Object { $_.Status -eq "Error" -and -not $_.IsRestored }).Count
                $script:lastGuiResultSummary = "Last batch: $okCount OK  $errCount NG"

                Wait-KeyPress
            }

            "NewSession" {
                # Fail-closed: never open the session form without a
                # verification token - session_form refuses to verify
                # against an empty/missing token, and an unverified
                # passphrase silently breaks every ENC: decryption later.
                if (-not (Test-Path $verifyTokenPath)) {
                    Show-Error "Passphrase verification token not found: $verifyTokenPath"
                    Show-Error "Cannot start a new session without it. Run Fabriq Studio to generate the token."
                    Wait-KeyPress
                    continue
                }

                # Full state reset (transcript, session, history, env vars)
                Show-ConsoleWindow
                Reset-FabriqState

                $hostList = Load-HostList
                if (-not $hostList) {
                    Show-Error "Failed to load host list"
                    Wait-KeyPress
                    continue
                }

                $workers = @()
                if (Test-Path $script:WorkersCsvPath) {
                    try { $workers = @(Import-Csv -Path $script:WorkersCsvPath -Encoding Default) } catch { }
                }

                Hide-ConsoleWindow

                $sessionSetup = Show-SessionSetupForm `
                    -Workers $workers `
                    -HostList $hostList `
                    -VerifyTokenPath $verifyTokenPath `
                    -CurrentPCName $env:COMPUTERNAME

                Show-ConsoleWindow

                if ($sessionSetup.Cancelled) {
                    # User cancelled - return to dashboard
                    continue
                }

                # Apply new session results
                $global:FabriqMasterPassphrase = $sessionSetup.MasterPassphrase
                Show-Success "Master passphrase verified and set for this session"

                $mediaSerial = ""
                if (Test-Path $script:SourceMediaIdPath) {
                    try { $mediaSerial = (Get-Content $script:SourceMediaIdPath -Raw -ErrorAction Stop).Trim() } catch { }
                }
                if ([string]::IsNullOrWhiteSpace($mediaSerial)) {
                    $currentDrive = (Resolve-Path ".").Drive.Name + ":"
                    $mediaSerial = Get-VolumeSerial -DriveLetter $currentDrive
                }

                $script:SessionInfo = [PSCustomObject]@{
                    WorkerID     = $sessionSetup.WorkerID
                    WorkerName   = $sessionSetup.WorkerName
                    MediaSerial  = $mediaSerial
                    StartTime    = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                    WindowsUser  = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
                    ComputerName = $env:COMPUTERNAME
                }
                try {
                    $script:SessionInfo | ConvertTo-Json -Depth 3 | Out-File -FilePath $script:SessionFilePath -Encoding UTF8 -Force
                } catch { Show-Warning "Failed to save session.json: $_" }
                $env:FABRIQ_WORKER_NAME = $sessionSetup.WorkerName
                Show-Success "Session initialized: Worker=$($sessionSetup.WorkerName)"

                $selectedHost = $sessionSetup.SelectedHost
                Set-SelectedHostEnvironment -SelectedHost $selectedHost
                Initialize-EvidenceBasePath

                # Restore previous execution history for this PC
                Restore-ExecutionHistory

                $script:lastGuiResultSummary = ""
            }

            "OpenCsvEditor" {
                $csvEditorScript = ".\apps\csv_editor\csv_editor.ps1"
                if (Test-Path $csvEditorScript) {
                    & $csvEditorScript
                }
                else {
                    Show-Error "CSV Editor not found"
                    Wait-KeyPress
                }
            }

            "SystemLauncher" {
                $launcherScript = ".\apps\system_launcher\system_launcher.ps1"
                if (Test-Path $launcherScript) {
                    & $launcherScript
                }
                else {
                    Show-Error "System Launcher not found"
                    Wait-KeyPress
                }
            }

            "AppsMode" {
                $appResult = Show-AppsDialog -AppsDir $APPS_DIR
                if ($appResult.Action -eq "Launch" -and $appResult.AppPath) {
                    Clear-Host
                    Show-Info "Launching [$($appResult.AppName)]"
                    Write-Host ""
                    try {
                        & $appResult.AppPath
                        Write-Host ""
                        Show-Success "App closed"
                    }
                    catch {
                        Write-Host ""
                        Show-Error "Error launching app: $_"
                    }
                    Wait-KeyPress
                }
            }

            "LaunchApp" {
                # Direct shortcut path used by Quick Actions buttons that
                # bypass the FabriqApps picker (e.g. the [Fabriq IOS]
                # button). Resolves apps/<AppName>/<AppName>.ps1 the
                # same way Show-AppsDialog does and runs it in-process,
                # mirroring the AppsMode handler above.
                $appName = $guiSelection.AppName
                if ([string]::IsNullOrWhiteSpace($appName)) {
                    Show-Warning "LaunchApp received no AppName"
                    Wait-KeyPress
                }
                else {
                    $appPath = Join-Path $APPS_DIR ("{0}\{0}.ps1" -f $appName)
                    if (-not (Test-Path $appPath)) {
                        Show-Error "App not found: $appPath"
                        Wait-KeyPress
                    }
                    else {
                        Clear-Host
                        Show-Info "Launching [$appName]"
                        Write-Host ""
                        try {
                            & $appPath
                            Write-Host ""
                            Show-Success "App closed"
                        }
                        catch {
                            Write-Host ""
                            Show-Error "Error launching app: $_"
                        }
                        Wait-KeyPress
                    }
                }
            }

            "HistoryExport" {
                $null = Export-ExecutionHistory
                Wait-KeyPress
            }

            "RegenerateChecklist" {
                if ([string]::IsNullOrEmpty($global:FabriqLastProfileName)) {
                    Show-Warning "No profile has been executed in this session"
                    Wait-KeyPress
                    continue
                }
                # Manual mode: viewer-before-upload + record "Log Upload (cl)"
                # in execution history. Behavior matches the previous inline
                # implementation byte-for-byte.
                $null = Complete-ProfileExecution `
                    -ProfileName    $global:FabriqLastProfileName `
                    -ProfilePath    $global:FabriqLastProfilePath `
                    -DefinedModules $global:FabriqLastProfileModules `
                    -Mode           'Manual'
                Wait-KeyPress
            }

            "WindowsUpdate" {
                Invoke-WindowsUpdateLoop
                Wait-KeyPress
            }

            "Restart" {
                Show-Separator
                Write-Host "Restart with AutoRun" -ForegroundColor Yellow
                Show-Separator
                Write-Host ""
                if (Confirm-Execution -Message "Register RunOnce and restart the computer?") {
                    if (Register-FabriqRunOnce) {
                        # Hide toolbar cleanly before reboot kills the process
                        try { Hide-ExecutionToolbar } catch { }
                        Invoke-CountdownRestart
                    }
                }
                else {
                    Show-Info "Canceled"
                }
                Wait-KeyPress
            }

            "Refabriq" {
                Show-Info "Restarting Fabriq..."
                try { Hide-ExecutionToolbar } catch { }
                Remove-StatusFile
                $fabriqRoot = (Resolve-Path ".").Path
                $fabriqExe = Join-Path $fabriqRoot "Fabriq.exe"
                if (Test-Path $fabriqExe) {
                    Start-Process $fabriqExe -WorkingDirectory $fabriqRoot
                } else {
                    Show-Error "Fabriq.exe not found: $fabriqExe"
                    Wait-KeyPress
                    continue
                }
                try { Stop-Transcript | Out-Null } catch { }
                exit 0
            }

            "Manifesto" {
                Show-Manifesto
            }

            "Quit" {
                Exit-Fabriq
                $script:guiExitRequested = $true
            }

            default {
                Exit-Fabriq
                $script:guiExitRequested = $true
            }
        }
    }

# Safety net: Exit-Fabriq is idempotent, safe to call even if already called
Exit-Fabriq