# ========================================
#
# Fabriq ver2.2 - Manifeste du Surkitinisme -
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

    # Helper: Decrypt ENC: prefixed values if master passphrase is available
    function Resolve-HostValue {
        param([string]$Value)
        if ([string]::IsNullOrEmpty($Value)) { return $Value }
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
    $env:SELECTED_OLD_PCNAME = Resolve-HostValue $SelectedHost.'OldPCName'
    $env:SELECTED_NEW_PCNAME = Resolve-HostValue $SelectedHost.'NewPCName'

    $env:SELECTED_ETH_IP      = Resolve-HostValue $SelectedHost.'EthernetIP'
    $env:SELECTED_ETH_SUBNET  = Resolve-HostValue $SelectedHost.'EthernetSubnet'
    $env:SELECTED_ETH_GATEWAY = Resolve-HostValue $SelectedHost.'EthernetGateway'

    $env:SELECTED_WIFI_IP      = Resolve-HostValue $SelectedHost.'WifiIP'
    $env:SELECTED_WIFI_SUBNET  = Resolve-HostValue $SelectedHost.'WifiSubnet'
    $env:SELECTED_WIFI_GATEWAY = Resolve-HostValue $SelectedHost.'WifiGateway'

    $env:SELECTED_DNS1 = Resolve-HostValue $SelectedHost.'DNS1'
    $env:SELECTED_DNS2 = Resolve-HostValue $SelectedHost.'DNS2'
    $env:SELECTED_DNS3 = Resolve-HostValue $SelectedHost.'DNS3'
    $env:SELECTED_DNS4 = Resolve-HostValue $SelectedHost.'DNS4'

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
        # グローバルフォールバック変数をクリア
        $global:_LastModuleResult = $null

        $output = & $ScriptPath

        # ModuleResult を検出（パイプライン出力から）
        $moduleResult = $null
        if ($null -ne $output) {
            foreach ($item in @($output)) {
                if ($item -is [PSCustomObject] -and $item._IsModuleResult -eq $true) {
                    $moduleResult = $item
                }
            }
        }

        # フォールバック: パイプラインキャプチャ失敗時にグローバル変数から取得
        if (-not $moduleResult -and $null -ne $global:_LastModuleResult) {
            $moduleResult = $global:_LastModuleResult
        }
        $global:_LastModuleResult = $null

        if ($moduleResult) {
            # モジュールが返却した結果を使用
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
            # レガシーパス: ModuleResult 未返却（全モジュール移行済み）
            Write-Host ""
            Write-Verbose "ModuleResult not returned from: $ScriptPath"
            Show-Warning "Script execution completed (status unverified)"
            Add-ExecutionResult -Operation $ModuleName -Status "Success" -Message "(legacy - unverified)"
            $null = Write-ExecutionHistory -ModuleName $ModuleName -Category $Category -Status "Success" -Message "(legacy - unverified)"
            Capture-ScreenEvidence -ModuleName $ModuleName -Status "Success"
            return $true
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
        [datetime]$ProfileStartTime = (Get-Date)
    )

    # AutoPilot: set global flag (Profile scope, reset in finally)
    if ($AutoPilot) {
        $global:AutoPilotMode = $true
        $global:AutoPilotWaitSec = $AutoPilotWaitSec
    }

    try {

    # Confirm execution
    if (-not (Show-BatchConfirmation -SelectedModules $SelectedModules)) {
        Show-Info "Batch execution canceled"
        return
    }

    Clear-ExecutionResults
    $total = $SelectedModules.Count
    $current = 0
    $completedResults = @()

    foreach ($module in $SelectedModules) {
        $current++

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
            Save-ResumeState -ProfilePath $ProfilePath `
                             -ProfileName $ProfileName `
                             -ResumeAfterOrder $module.Order `
                             -CompletedModules $completedResults `
                             -ProfileStartTime $ProfileStartTime

            # Register RunOnce
            if (-not (Register-FabriqRunOnce)) {
                Remove-ResumeState
                Add-ExecutionResult -Operation "[RESTART]" -Status "Error" -Message "RunOnce registration failed"
                $null = Write-ExecutionHistory -ModuleName "[RESTART]" -Category "System" -Status "Error" -Message "RunOnce registration failed"
                continue
            }

            # Record in execution history
            Add-ExecutionResult -Operation "[RESTART]" -Status "Success" -Message "Restarting..."
            $null = Write-ExecutionHistory -ModuleName "[RESTART]" -Category "System" -Status "Success" -Message "Profile restart (ResumeAfter: $($module.Order))"

            Invoke-CountdownRestart
            return
        }

        # __REEXPLORER__ marker handling
        if ($module._IsReexplorer) {
            Show-BatchProgress -Current $current -Total $total -ItemName "[REEXPLORER]"
            Show-Info "Restarting Explorer..."
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
                # Windowsの自動再起動が間に合わなかった場合のみ明示的に起動
                if (-not $restarted) { Start-Process explorer.exe }
                Add-ExecutionResult -Operation "[REEXPLORER]" -Status "Success" -Message "Explorer restarted"
                $null = Write-ExecutionHistory -ModuleName "[REEXPLORER]" -Category "System" -Status "Success" -Message "Explorer restarted"
            }
            catch {
                Add-ExecutionResult -Operation "[REEXPLORER]" -Status "Error" -Message $_.Exception.Message
                $null = Write-ExecutionHistory -ModuleName "[REEXPLORER]" -Category "System" -Status "Error" -Message $_.Exception.Message
            }
            $completedResults += @{ MenuName = "[REEXPLORER]"; Status = "Success" }
            continue
        }

        # Normal module execution
        Show-BatchProgress -Current $current -Total $total -ItemName $module.MenuName

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

        Add-ExecutionResult -Operation $module.MenuName -Status $result.Status -Message $result.Message -Verified $result.Verified
        $null = Write-ExecutionHistory -ModuleName $module.MenuName -Category $module.Category -Status $result.Status -Message $result.Message -Verified $verifiedStr
        Capture-ScreenEvidence -ModuleName $module.MenuName -Status $result.Status

        # Track completed results for resume state
        $completedResults += @{ MenuName = $module.MenuName; Status = $result.Status }
    }

    # All modules completed (no restart, or all restarts done)
    Remove-ResumeState

    # Calculate elapsed time as a single subtraction from the absolute
    # profile start timestamp. Naturally includes reboot/login/startup
    # gaps for profiles with __RESTART__ markers.
    $batchElapsed = (Get-Date) - $ProfileStartTime

    Show-ExecutionSummary -ElapsedTime $batchElapsed

    # Auto-export evidence if this is a profile execution
    if (-not [string]::IsNullOrEmpty($ProfileName)) {
        Write-Host ""
        Write-Host "[INFO] Auto-exporting execution history as evidence..." -ForegroundColor Cyan
        $null = Export-ExecutionHistory

        # HTML checklist
        Write-Host "[INFO] Generating HTML checklist..." -ForegroundColor Cyan
        $checklistModules = if ($null -ne $FullProfileModules) { $FullProfileModules } else { $SelectedModules }
        $checklistPath = Export-HtmlChecklist `
            -ProfileName      $ProfileName `
            -ProfilePath      $ProfilePath `
            -DefinedModules   $checklistModules `
            -ExecutionResults $script:ExecutionResults `
            -ElapsedTime      $batchElapsed

        # Retain profile info for checklist regeneration from menu [cl]
        $global:FabriqLastProfileName    = $ProfileName
        $global:FabriqLastProfilePath    = $ProfilePath
        $global:FabriqLastProfileModules = $checklistModules

        # Auto-run log upload
        $logUploaderScript = ".\modules\extended\log_uploader\log_uploader.ps1"
        if (Test-Path $logUploaderScript) {
            $destConfig = ".\kernel\csv\log_destinations.csv"
            $hasDestinations = $false
            if (Test-Path $destConfig) {
                try {
                    $dests = @(Import-Csv -Path $destConfig -Encoding Default | Where-Object { $_.Enabled -eq "1" })
                    $hasDestinations = ($dests.Count -gt 0)
                }
                catch { }
            }

            if ($hasDestinations) {
                Write-Host ""
                Write-Host "[INFO] Auto-uploading logs and evidence..." -ForegroundColor Cyan
                try {
                    $null = & $logUploaderScript
                }
                catch {
                    Show-Warning "Log upload failed: $($_.Exception.Message)"
                }
            }
        }

        # Launch HTML checklist viewer
        if (-not [string]::IsNullOrEmpty($checklistPath) -and (Test-Path $checklistPath)) {
            Write-Host ""
            Write-Host "[INFO] Opening HTML checklist viewer..." -ForegroundColor Cyan
            $viewerScript = ".\kernel\ps1\view_report.ps1"
            if (Test-Path $viewerScript) {
                try {
                    & $viewerScript -HtmlPath $checklistPath
                }
                catch {
                    Show-Warning "Failed to open report viewer: $($_.Exception.Message)"
                }
            }
        }
    }

    } # end try
    finally {
        # AutoPilot: always reset (Profile scope guarantee)
        $global:AutoPilotMode = $false
        $global:AutoPilotWaitSec = 3
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
            Remove-Item $wuStatePath -Force -ErrorAction SilentlyContinue
            Show-Error "Failed to register RunOnce for WU resume"
            return
        }

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
Write-Host "Fabriq ver2.1 - Manifeste du Surkitinisme - " -ForegroundColor Green
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
# WU resume runs immediately — no passphrase, worker, or host selection needed.
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

if ($null -ne $resumeState) {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host "  Profile Resume Detected" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  Profile:  $($resumeState.ProfileName)" -ForegroundColor White
    Write-Host "  PC:       $($resumeState.HostEnvironment.SELECTED_NEW_PCNAME)" -ForegroundColor White

    $completedCount = @($resumeState.CompletedModules).Count
    Write-Host "  Progress: $completedCount modules completed" -ForegroundColor White
    Write-Host ""

    $resumeIsAutoPilot = ($resumeState.AutoPilot -eq $true)

    if ($resumeIsAutoPilot) {
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
                $verifyTokenPath = Join-Path $PSScriptRoot "txt\passphrase_verify.txt"
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
    $verifyTokenPath = Join-Path $PSScriptRoot "txt\passphrase_verify.txt"

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
# Status Monitor
# ========================================
$global:FabriqStatusMonitorProcess = Start-StatusMonitor


# ========================================
# Resume Execution (if resuming)
# ========================================
if ($isResuming) {
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
            Add-ExecutionResult -Operation $cm.MenuName -Status $cm.Status -Message "(completed before restart)"
        }
        Add-ExecutionResult -Operation "[RESTART]" -Status "Success" -Message "Resumed after restart"

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
            Show-Warning "Legacy resume_state.json without ProfileStartTime — elapsed time will be measured from now (pre-restart duration not included)."
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
                    # User cancelled — return to dashboard
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
                Show-Info "Regenerating checklist..."
                $null = Export-ExecutionHistory
                $checklistPath = Export-HtmlChecklist `
                    -ProfileName      $global:FabriqLastProfileName `
                    -ProfilePath      $global:FabriqLastProfilePath `
                    -DefinedModules   $global:FabriqLastProfileModules `
                    -ExecutionResults  $script:ExecutionResults
                if (-not [string]::IsNullOrEmpty($checklistPath) -and (Test-Path $checklistPath)) {
                    $viewerScript = ".\kernel\ps1\view_report.ps1"
                    if (Test-Path $viewerScript) {
                        & $viewerScript -HtmlPath $checklistPath
                    }
                }

                # Auto-upload to log destinations
                $uploaderScript = ".\modules\extended\log_uploader\log_uploader.ps1"
                if (Test-Path $uploaderScript) {
                    Show-Info "Uploading updated evidence..."
                    $uploadResult = & $uploaderScript
                    if ($null -ne $uploadResult -and $uploadResult._IsModuleResult) {
                        Add-ExecutionResult -Operation "Log Upload (cl)" -Status $uploadResult.Status -Message $uploadResult.Message
                        $null = Write-ExecutionHistory -ModuleName "Log Upload (cl)" -Category "System" -Status $uploadResult.Status -Message $uploadResult.Message
                    }
                }
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
                Stop-StatusMonitor -MonitorProcess $global:FabriqStatusMonitorProcess
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