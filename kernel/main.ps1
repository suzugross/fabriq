# ========================================
#
# Fabriq ver2.1 - Manifeste du Surkitinisme -
#
# ========================================

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

# Enable sleep suppression while Fabriq is running
Enable-SleepSuppression

# Set compact console window size
Set-ConsoleSize -Columns 75 -Lines 35

# ========================================
# Constants
# ========================================
$HOSTLIST_CSV = ".\kernel\csv\hostlist.csv"
$COMMANDS_DIR = ".\commands"
$APPS_DIR = ".\apps"

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
# Function: Select Host
# ========================================
function Select-Host {
    param([array]$HostList)

    $currentPCName = $env:COMPUTERNAME
    Show-Info "Current PC Name: $currentPCName"
    Write-Host ""

    # Check auto-selection
    # Note: CSV header 'NewPCName' is required
    $autoSelected = $HostList | Where-Object { $_.'NewPCName' -eq $currentPCName }

    if ($autoSelected) {
        # Note: CSV header 'AdminID', 'OldPCName', 'NewPCName' are required
        Show-Success "Auto-selected: ID [$($autoSelected.'AdminID')] - $($autoSelected.'OldPCName') -> $($autoSelected.'NewPCName')"
        return $autoSelected
    }

    # Manual selection
    Show-Info "Please select the target device."
    Write-Host ""
    Show-Separator
    Write-Host "Target Device List" -ForegroundColor Cyan
    Show-Separator

    foreach ($pc in $HostList) {
        Write-Host "[$($pc.'AdminID')] $($pc.'OldPCName') -> $($pc.'NewPCName')"
    }

    Write-Host "  [Q] Quit" -ForegroundColor DarkGray
    Show-Separator
    Write-Host ""

    while ($true) {
        Write-Host -NoNewline "Please enter the ID: "
        $userInput = Read-Host

        if ($userInput -eq 'q' -or $userInput -eq 'Q') {
            return $null
        }

        $selected = $HostList | Where-Object { $_.'AdminID' -eq $userInput }

        if ($selected) {
            Write-Host ""
            Show-Success "Selected: ID [$($selected.'AdminID')] - $($selected.'OldPCName') -> $($selected.'NewPCName')"
            return $selected
        }
        else {
            Write-Host ""
            Show-Error "Invalid ID. Please try again."
            Write-Host ""
        }
    }
}

# ========================================
# Function: Set Environment Variables
# ========================================
function Set-SelectedHostEnvironment {
    param([object]$SelectedHost)

    # Note: These keys must match your CSV headers
    $env:SELECTED_KANRI_NO = $SelectedHost.'AdminID'
    $env:SELECTED_OLD_PCNAME = $SelectedHost.'OldPCName'
    $env:SELECTED_NEW_PCNAME = $SelectedHost.'NewPCName'

    $env:SELECTED_ETH_IP = $SelectedHost.'EthernetIP'
    $env:SELECTED_ETH_SUBNET = $SelectedHost.'EthernetSubnet'
    $env:SELECTED_ETH_GATEWAY = $SelectedHost.'EthernetGateway'

    $env:SELECTED_WIFI_IP = $SelectedHost.'WifiIP'
    $env:SELECTED_WIFI_SUBNET = $SelectedHost.'WifiSubnet'
    $env:SELECTED_WIFI_GATEWAY = $SelectedHost.'WifiGateway'

    $env:SELECTED_DNS1 = $SelectedHost.'DNS1'
    $env:SELECTED_DNS2 = $SelectedHost.'DNS2'
    $env:SELECTED_DNS3 = $SelectedHost.'DNS3'
    $env:SELECTED_DNS4 = $SelectedHost.'DNS4'

    for ($i = 1; $i -le 10; $i++) {
        # CSV headers like: Printer1Name, Printer1Driver, Printer1Port
        $nameKey = "Printer$($i)Name"
        $driverKey = "Printer$($i)Driver"
        $portKey = "Printer$($i)Port"

        Set-Item -Path "env:SELECTED_PRINTER_$($i)_NAME" -Value $SelectedHost.$nameKey -ErrorAction SilentlyContinue
        Set-Item -Path "env:SELECTED_PRINTER_$($i)_DRIVER" -Value $SelectedHost.$driverKey -ErrorAction SilentlyContinue
        Set-Item -Path "env:SELECTED_PRINTER_$($i)_PORT" -Value $SelectedHost.$portKey -ErrorAction SilentlyContinue
    }

    Show-Info "Environment variables set."
}

# ========================================
# Function: New Kitting Session
# ========================================
function Invoke-NewKittingSession {
    Write-Host ""
    Show-Separator
    Write-Host "New Kitting Session" -ForegroundColor Magenta
    Show-Separator
    Write-Host ""

    # Reset all in-memory state and start a new transcript
    Reset-FabriqState

    # Re-initialize session (worker selection)
    $sessionResult = Initialize-Session
    if ($sessionResult -eq $false) {
        Show-Info "New session canceled - returning to main menu"
        Write-Host ""
        return
    }
    Write-Host ""

    # Re-load host list and re-select target device
    Show-Info "Loading hostlist.csv..."
    $hostListNew = Load-HostList
    if (-not $hostListNew) {
        Show-Error "Failed to load hostlist.csv - returning to main menu"
        Write-Host ""
        return
    }
    Write-Host ""

    $selectedHostNew = Select-Host -HostList $hostListNew
    if ($null -eq $selectedHostNew) {
        Show-Info "New session canceled - returning to main menu"
        Write-Host ""
        return
    }
    Write-Host ""
    Set-SelectedHostEnvironment -SelectedHost $selectedHostNew
    Write-Host ""

    Restore-ExecutionHistory
    Write-Host ""

    Show-Success "New session ready. Target: $env:SELECTED_NEW_PCNAME"
    Write-Host ""
}

# ========================================
# Function: Show Main Menu (Navigation Hub)
# ========================================
function Show-MainMenu {
    Write-Host ""
    Show-Separator
    Write-Host "Fabriq ver2.1 - Manifeste du Surkitinisme -" -ForegroundColor Green
    Show-Separator
    Write-Host "  Selected Host: $env:SELECTED_NEW_PCNAME" -ForegroundColor White
    Show-Separator
    Write-Host ""
    Write-Host "  [S] Script Menu" -ForegroundColor White
    Write-Host "  [A] FabriqApps" -ForegroundColor White
    Write-Host "  [C] Command" -ForegroundColor White
    Write-Host "  [P] Run Profile" -ForegroundColor White
    Write-Host "  [N] New Session" -ForegroundColor White
    Write-Host "  [R] History" -ForegroundColor White
    Write-Host ""
    Write-Host "----------------------------------------" -ForegroundColor DarkGray
    Write-Host "  [eh] History Export" -ForegroundColor DarkGray
    Write-Host "  [re] Windows Restart" -ForegroundColor DarkGray
    Write-Host "  [rf] Refabriq" -ForegroundColor DarkGray
    Write-Host "  [m]  Manifeste du Surkitinisme" -ForegroundColor DarkGray
    Write-Host "----------------------------------------" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  [Q]  Quit" -ForegroundColor DarkGray
}

# ========================================
# Function: Show Host Info
# ========================================
function Show-HostInfo {
    Write-Host ""
    Show-Separator
    Write-Host "Selected PC Info" -ForegroundColor Green
    Show-Separator
    Write-Host "[ID] $env:SELECTED_KANRI_NO"
    Write-Host "[Old Name] $env:SELECTED_OLD_PCNAME"
    Write-Host "[New Name] $env:SELECTED_NEW_PCNAME"
    Write-Host ""
    Write-Host "[Ethernet]" -ForegroundColor Yellow
    Write-Host "  IP: $env:SELECTED_ETH_IP"
    Write-Host "  Subnet: $env:SELECTED_ETH_SUBNET"
    Write-Host "  Gateway: $env:SELECTED_ETH_GATEWAY"
    Write-Host ""
    Write-Host "[Wi-Fi]" -ForegroundColor Yellow
    Write-Host "  IP: $env:SELECTED_WIFI_IP"
    Write-Host "  Subnet: $env:SELECTED_WIFI_SUBNET"
    Write-Host "  Gateway: $env:SELECTED_WIFI_GATEWAY"
    Write-Host ""
    Write-Host "[DNS]" -ForegroundColor Yellow
    Write-Host "  DNS1: $env:SELECTED_DNS1"
    Write-Host "  DNS2: $env:SELECTED_DNS2"
    Write-Host "  DNS3: $env:SELECTED_DNS3"
    Write-Host "  DNS4: $env:SELECTED_DNS4"

    $hasPrinter = $false
    for ($i = 1; $i -le 10; $i++) {
        $printerName = [Environment]::GetEnvironmentVariable("SELECTED_PRINTER_$($i)_NAME")
        if (-not [string]::IsNullOrEmpty($printerName)) {
            if (-not $hasPrinter) {
                Write-Host ""
                Write-Host "[Printer]" -ForegroundColor Yellow
                $hasPrinter = $true
            }
            $printerDriver = [Environment]::GetEnvironmentVariable("SELECTED_PRINTER_$($i)_DRIVER")
            $printerPort = [Environment]::GetEnvironmentVariable("SELECTED_PRINTER_$($i)_PORT")
            Write-Host "  Printer${i}:"
            Write-Host "    Name: $printerName"
            Write-Host "    Driver: $printerDriver"
            Write-Host "    Port: $printerPort"
        }
    }

    Show-Separator
    Wait-KeyPress
}

# ========================================
# Function: Script Menu (Module Selection)
# ========================================
function Enter-ScriptMenu {
    param(
        [array]$GroupedModules,
        [array]$AllModules
    )

    while ($true) {
        $menuMap = @{}

        Write-Host ""
        Show-Separator
        Write-Host "Script Menu" -ForegroundColor Magenta
        Show-Separator
        Write-Host "  Selected Host: $env:SELECTED_NEW_PCNAME" -ForegroundColor White
        Show-Separator

        $menuIndex = 1
        foreach ($category in $GroupedModules) {
            Show-CategorySeparator -Name $category.Name
            $items = $category.Group | Sort-Object Order
            foreach ($item in $items) {
                Write-Host "  [$menuIndex] $($item.MenuName)" -ForegroundColor White
                $menuMap[$menuIndex] = $item
                $menuIndex++
            }
        }

        Write-Host ""
        Write-Host "  * Batch Run: 1,3,5 or 1-5" -ForegroundColor DarkGray
        Write-Host "  [0] Back" -ForegroundColor Yellow
        Show-Separator

        Write-Host -NoNewline "Please select: "
        $choice = Read-Host

        # Back
        if ($choice -eq '0') { return }

        # Batch Input
        if (Test-BatchInput -InputString $choice) {
            $selectedNumbers = Parse-MenuSelection -InputString $choice
            $selectedModules = @()
            foreach ($num in $selectedNumbers) {
                if ($menuMap.ContainsKey($num)) {
                    $selectedModules += $menuMap[$num]
                }
            }
            if ($selectedModules.Count -gt 0) {
                Clear-Host
                Invoke-BatchExecution -SelectedModules $selectedModules
                Clear-Host
            }
            else {
                Write-Host ""
                Show-Error "No valid numbers specified"
                Write-Host ""
            }
            continue
        }

        # Single Number Selection
        $menuNum = 0
        if ([int]::TryParse($choice, [ref]$menuNum) -and $menuMap.ContainsKey($menuNum)) {
            $selectedModule = $menuMap[$menuNum]
            Clear-Host
            Show-Info "Executing [$($selectedModule.MenuName)]"
            Write-Host ""
            $null = Invoke-KittingScript -ScriptPath $selectedModule.Script `
                                         -ModuleName $selectedModule.MenuName `
                                         -Category $selectedModule.Category
            Wait-KeyPress
            Clear-Host
        }
        else {
            Write-Host ""
            Show-Error "Invalid selection"
            Write-Host ""
        }
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

            Add-ExecutionResult -Operation $ModuleName -Status $status -Message $message
            $null = Write-ExecutionHistory -ModuleName $ModuleName -Category $Category -Status $status -Message $message
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
        [switch]$StopOnError,
        [switch]$AutoPilot,
        [int]$AutoPilotWaitSec = 3,
        # Profile restart support (optional)
        [string]$ProfilePath = "",
        [string]$ProfileName = "",
        # Full profile module list for checklist (covers pre-restart modules in resume)
        # If omitted, $SelectedModules is used as-is
        [array]$FullProfileModules = $null
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
    $batchStartTime = Get-Date
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

            # Save resume state
            Save-ResumeState -ProfilePath $ProfilePath `
                             -ProfileName $ProfileName `
                             -StopOnError $StopOnError.IsPresent `
                             -ResumeAfterOrder $module.Order `
                             -CompletedModules $completedResults

            # Register RunOnce
            if (-not (Register-FabriqRunOnce)) {
                Remove-ResumeState
                Add-ExecutionResult -Operation "[RESTART]" -Status "Error" -Message "RunOnce registration failed"
                $null = Write-ExecutionHistory -ModuleName "[RESTART]" -Category "System" -Status "Error" -Message "RunOnce registration failed"
                if ($StopOnError) { break } else { continue }
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
                if ($StopOnError) { break }
            }
            $completedResults += @{ MenuName = "[REEXPLORER]"; Status = "Success" }
            continue
        }

        # __STOPLOG__ marker handling
        if ($module._IsStopLog) {
            Show-BatchProgress -Current $current -Total $total -ItemName "[STOPLOG]"
            Show-Info "Stopping transcript..."
            try {
                Stop-Transcript -ErrorAction SilentlyContinue | Out-Null
                Add-ExecutionResult -Operation "[STOPLOG]" -Status "Success" -Message "Transcript stopped"
                $null = Write-ExecutionHistory -ModuleName "[STOPLOG]" -Category "System" -Status "Success" -Message "Transcript stopped"
            }
            catch {
                Add-ExecutionResult -Operation "[STOPLOG]" -Status "Warning" -Message $_.Exception.Message
            }
            $completedResults += @{ MenuName = "[STOPLOG]"; Status = "Success" }
            continue
        }

        # __STARTLOG__ marker handling
        if ($module._IsStartLog) {
            Show-BatchProgress -Current $current -Total $total -ItemName "[STARTLOG]"
            Show-Info "Resuming transcript..."
            try {
                $transcriptPath = $global:FabriqTranscriptPath
                if ([string]::IsNullOrEmpty($transcriptPath)) {
                    $ts  = Get-Date -Format "yyyy_MM_dd_HHmmss"
                    $uid = Get-HardwareUniqueId
                    $hn  = $env:COMPUTERNAME
                    $transcriptPath = ".\logs\${ts}_${uid}_${hn}.log"
                    $global:FabriqTranscriptPath = $transcriptPath
                }
                Start-Transcript -Path $transcriptPath -Append | Out-Null
                Add-ExecutionResult -Operation "[STARTLOG]" -Status "Success" -Message "Transcript resumed: $transcriptPath"
                $null = Write-ExecutionHistory -ModuleName "[STARTLOG]" -Category "System" -Status "Success" -Message "Transcript resumed"
            }
            catch {
                Add-ExecutionResult -Operation "[STARTLOG]" -Status "Warning" -Message $_.Exception.Message
            }
            $completedResults += @{ MenuName = "[STARTLOG]"; Status = "Success" }
            continue
        }

        # __SHUTDOWN__ marker handling
        if ($module._IsShutdown) {
            Show-BatchProgress -Current $current -Total $total -ItemName "[SHUTDOWN]"
            Write-Host ""
            Write-Host "========================================" -ForegroundColor Red
            Write-Host "  Shutdown Phase" -ForegroundColor Red
            Write-Host "========================================" -ForegroundColor Red
            Write-Host ""
            Write-Host "  Progress: $($completedResults.Count) modules completed" -ForegroundColor White
            Write-Host ""

            Add-ExecutionResult -Operation "[SHUTDOWN]" -Status "Success" -Message "Shutting down..."
            $null = Write-ExecutionHistory -ModuleName "[SHUTDOWN]" -Category "System" -Status "Success" -Message "Shutdown initiated"

            Invoke-CountdownShutdown
            return
        }

        # __PAUSE__ marker handling
        if ($module._IsPause) {
            Show-BatchProgress -Current $current -Total $total -ItemName "[PAUSE]"
            Write-Host ""
            Write-Host "========================================" -ForegroundColor Yellow
            Write-Host "  Pause - Waiting for user input" -ForegroundColor Yellow
            Write-Host "========================================" -ForegroundColor Yellow
            Write-Host ""
            Write-Host "  Progress: $current/$total ($($completedResults.Count) completed)" -ForegroundColor White
            Write-Host ""
            Set-ConsoleForeground
            Wait-KeyPress -Message "Press Enter to continue..."
            Add-ExecutionResult -Operation "[PAUSE]" -Status "Success" -Message "User resumed"
            $null = Write-ExecutionHistory -ModuleName "[PAUSE]" -Category "System" -Status "Success" -Message "User resumed"
            $completedResults += @{ MenuName = "[PAUSE]"; Status = "Success" }
            continue
        }

        # Normal module execution
        Show-BatchProgress -Current $current -Total $total -ItemName $module.MenuName

        # AutoPilot: inter-module wait
        if ($global:AutoPilotMode -and $current -gt 1) {
            Write-Host "[AUTOPILOT] Next module in $($global:AutoPilotWaitSec)s..." -ForegroundColor Magenta
            Start-Sleep -Seconds $global:AutoPilotWaitSec
        }

        # __AUTO_to_xxx__ parameter passing via environment variable
        if ($module._AutoLogonNo) {
            $env:FABRIQ_AUTOLOGON_NO = $module._AutoLogonNo
        }

        $result = Invoke-SafeCommand -OperationName $module.MenuName -ScriptBlock {
            & $module.Script
        } -ContinueOnError

        # Clean up AutoLogon environment variable
        if ($module._AutoLogonNo) {
            $env:FABRIQ_AUTOLOGON_NO = $null
        }

        Add-ExecutionResult -Operation $module.MenuName -Status $result.Status -Message $result.Message
        $null = Write-ExecutionHistory -ModuleName $module.MenuName -Category $module.Category -Status $result.Status -Message $result.Message
        Capture-ScreenEvidence -ModuleName $module.MenuName -Status $result.Status

        # Track completed results for resume state
        $completedResults += @{ MenuName = $module.MenuName; Status = $result.Status }

        # StopOnError は Error ステータス時のみ発動（Cancelled/Skipped では停止しない）
        if (-not $result.Success -and $result.Status -eq "Error" -and $StopOnError) {
            Show-Error "Process aborted due to error"
            break
        }
    }

    # All modules completed (no restart, or all restarts done)
    Remove-ResumeState

    # Calculate elapsed time
    $batchElapsed = (Get-Date) - $batchStartTime

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

    Wait-KeyPress

    } # end try
    finally {
        # AutoPilot: always reset (Profile scope guarantee)
        $global:AutoPilotMode = $false
        $global:AutoPilotWaitSec = 3
    }
}

# ========================================
# Function: Profile Execution
# ========================================
function Invoke-ProfileExecution {
    param([array]$AllModules)

    $profiles = @(Load-Profiles -AllModules $AllModules)

    if ($profiles.Count -eq 0) {
        Show-Warning "No profile files found in profiles/ directory"
        Wait-KeyPress
        return
    }

    while ($true) {
        Show-ProfileMenu -Profiles $profiles

        Write-Host -NoNewline "Please select: "
        $choice = Read-Host

        if ($choice -eq '0') {
            return
        }

        $profileNum = 0
        if ([int]::TryParse($choice, [ref]$profileNum) -and $profileNum -ge 1 -and $profileNum -le $profiles.Count) {
            $selectedProfile = $profiles[$profileNum - 1]

            # Resolve modules from individual CSV
            $validation = Resolve-ProfileModules -ProfileCsvPath $selectedProfile.FilePath -AllModules $AllModules

            if ($validation.ValidModules.Count -eq 0) {
                Show-Error "No executable modules found"
                Wait-KeyPress
                continue
            }

            # Confirmation + execution mode selection
            Clear-Host
            $confirmation = Show-ProfileConfirmation `
                -SelectedProfile $selectedProfile `
                -Modules $validation.ValidModules `
                -InvalidPaths $validation.InvalidPaths `
                -AutoPilotFromCsv $validation.AutoPilot `
                -AutoPilotWaitSec $validation.AutoPilotWaitSec

            if ($null -ne $confirmation -and $confirmation.Confirmed) {
                Invoke-BatchExecution -SelectedModules $validation.ValidModules `
                    -StopOnError:$confirmation.StopOnError `
                    -AutoPilot:$confirmation.AutoPilot `
                    -AutoPilotWaitSec $confirmation.AutoPilotWaitSec `
                    -ProfilePath $selectedProfile.FilePath `
                    -ProfileName $selectedProfile.ProfileName
            }
            else {
                Show-Info "Profile execution canceled"
            }

            return
        }
        else {
            Write-Host ""
            Show-Error "Invalid selection"
            Write-Host ""
        }
    }
}

# ========================================
# Function: History Menu
# ========================================
function Enter-HistoryMode {
    while ($true) {
        Write-Host ""
        Show-Separator
        Write-Host "Execution History Menu" -ForegroundColor Magenta
        Show-Separator
        Write-Host ""
        Write-Host "  [1] Show last 20 entries (All hosts)" -ForegroundColor White
        Write-Host "  [2] Current host history only" -ForegroundColor White
        Write-Host "  [3] Export history" -ForegroundColor White
        Write-Host ""
        Write-Host "  [4] Clear Runtime Logs" -ForegroundColor Yellow
        Write-Host "  [5] Clear Evidence" -ForegroundColor Red
        Write-Host ""
        Write-Host "  [0] Back" -ForegroundColor Yellow
        Show-Separator

        Write-Host -NoNewline "Please select: "
        $choice = Read-Host

        switch ($choice) {
            '0' { return }
            '1' {
                Clear-Host
                Show-ExecutionHistory -Limit 20
                Wait-KeyPress
                Clear-Host
            }
            '2' {
                Clear-Host
                Show-ExecutionHistory -CurrentHostOnly -Limit 20
                Wait-KeyPress
                Clear-Host
            }
            '3' {
                Export-ExecutionHistory
                Wait-KeyPress
            }
            '4' {
                Clear-AllLogs
                Wait-KeyPress
                Clear-Host
            }
            '5' {
                Clear-Evidence
                Wait-KeyPress
                Clear-Host
            }
            default {
                Write-Host ""
                Show-Error "Invalid selection"
            }
        }
    }
}

# ========================================
# Function: Command Mode
# ========================================
function Enter-CommandMode {
    if (-not (Test-Path $COMMANDS_DIR)) {
        Show-Error "commands folder not found: $COMMANDS_DIR"
        Wait-KeyPress
        return
    }

    while ($true) {
        $scripts = @(Get-ChildItem -Path $COMMANDS_DIR -Filter "*.ps1" -File | Sort-Object Name)

        if ($scripts.Count -eq 0) {
            Show-Info "No executable commands found"
            Wait-KeyPress
            return
        }

        Write-Host ""
        Show-Separator
        Write-Host "Command Menu" -ForegroundColor Magenta
        Show-Separator
        Write-Host ""

        for ($i = 0; $i -lt $scripts.Count; $i++) {
            Write-Host "[$($i + 1)] $($scripts[$i].BaseName)"
        }

        Write-Host ""
        Write-Host "[0] Back" -ForegroundColor Yellow
        Show-Separator

        Write-Host -NoNewline "Please select: "
        $cmdChoice = Read-Host

        if ($cmdChoice -eq '0') {
            return
        }

        $cmdNum = 0
        if (-not [int]::TryParse($cmdChoice, [ref]$cmdNum) -or $cmdNum -lt 1 -or $cmdNum -gt $scripts.Count) {
            Write-Host ""
            Show-Error "Invalid selection"
            Write-Host ""
            continue
        }

        $targetScript = $scripts[$cmdNum - 1]
        Clear-Host
        Show-Info "Executing [$($targetScript.BaseName)]"
        Write-Host ""

        try {
            & $targetScript.FullName
            Write-Host ""
            Show-Success "Command execution completed"
        }
        catch {
            Write-Host ""
            Show-Error "Error executing command: $_"
        }

        Wait-KeyPress
        Clear-Host
    }
}

# ========================================
# Function: App Mode
# ========================================
function Enter-AppMode {
    if (-not (Test-Path $APPS_DIR)) {
        Show-Error "apps folder not found: $APPS_DIR"
        Wait-KeyPress
        return
    }

    while ($true) {
        # Discover apps: each subdirectory containing a .ps1 with the same name
        $appDirs = @(Get-ChildItem -Path $APPS_DIR -Directory | Sort-Object Name)
        $apps = @()
        foreach ($dir in $appDirs) {
            $entryScript = Join-Path $dir.FullName "$($dir.Name).ps1"
            if (Test-Path $entryScript) {
                $apps += [PSCustomObject]@{
                    Name = $dir.Name
                    Path = $entryScript
                }
            }
        }

        if ($apps.Count -eq 0) {
            Show-Info "No apps found"
            Wait-KeyPress
            return
        }

        Write-Host ""
        Show-Separator
        Write-Host "Apps Menu" -ForegroundColor Magenta
        Show-Separator
        Write-Host ""

        for ($i = 0; $i -lt $apps.Count; $i++) {
            Write-Host "[$($i + 1)] $($apps[$i].Name)"
        }

        Write-Host ""
        Write-Host "[0] Back" -ForegroundColor Yellow
        Show-Separator

        Write-Host -NoNewline "Please select: "
        $appChoice = Read-Host

        if ($appChoice -eq '0') {
            return
        }

        $appNum = 0
        if (-not [int]::TryParse($appChoice, [ref]$appNum) -or $appNum -lt 1 -or $appNum -gt $apps.Count) {
            Write-Host ""
            Show-Error "Invalid selection"
            Write-Host ""
            continue
        }

        $targetApp = $apps[$appNum - 1]
        Clear-Host
        Show-Info "Launching [$($targetApp.Name)]"
        Write-Host ""

        try {
            & $targetApp.Path
            Write-Host ""
            Show-Success "App closed"
        }
        catch {
            Write-Host ""
            Show-Error "Error launching app: $_"
        }

        Wait-KeyPress
        Clear-Host
    }
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

# Initialize history (Create backup)
Initialize-ExecutionHistory

# Initialize session (worker, media serial)
$sessionResult = Initialize-Session
if ($sessionResult -eq $false) {
    Exit-Fabriq
    exit 0
}

# Load hostlist.csv
Show-Info "Loading hostlist.csv..."
$hostList = Load-HostList
if (-not $hostList) {
    Write-Host ""
    Show-Error "Aborting process"
    Exit-Fabriq
    exit 1
}
Write-Host ""

# ========================================
# Resume Detection
# ========================================
$isResuming = $false
$resumeState = Load-ResumeState

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
    }
    else {
        Remove-ResumeState
        Show-Info "Resume state cleared. Starting normally."
    }
    Write-Host ""
}

# ========================================
# Host Selection (skip if resuming)
# ========================================
if (-not $isResuming) {
    $selectedHost = Select-Host -HostList $hostList
    if ($null -eq $selectedHost) {
        Exit-Fabriq
        exit 0
    }
    Write-Host ""
    Set-SelectedHostEnvironment -SelectedHost $selectedHost
    Write-Host ""
}

# Restore execution history for the selected PC
Restore-ExecutionHistory
Write-Host ""

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

        Invoke-BatchExecution -SelectedModules $remainingModules `
            -StopOnError:$resumeState.StopOnError `
            -AutoPilot:$resumeAutoPilot `
            -AutoPilotWaitSec $resumeAutoPilotWaitSec `
            -ProfilePath $resumeState.ProfilePath `
            -ProfileName $resumeState.ProfileName `
            -FullProfileModules $validation.ValidModules

        Remove-ResumeState

        # Offer new session after resume completion (same as [P] handler)
        Write-Host ""
        Show-Separator
        Write-Host "Profile Execution Completed" -ForegroundColor Green
        Show-Separator
        Write-Host ""
        if (Confirm-Execution -Message "Start a new kitting session for another device?") {
            Invoke-NewKittingSession
        }
        Clear-Host
    }
}

# Main Loop
while ($true) {
    Write-StatusFile -Phase "idle"
    Show-MainMenu

    Write-Host ""
    Write-Host -NoNewline "Please select: "
    $choice = Read-Host

    # Quit
    if ($choice -eq "Q" -or $choice -eq "q" -or $choice -eq "0") {
        Exit-Fabriq
        break
    }

    # Script Menu
    if ($choice -eq 'S' -or $choice -eq 's') {
        Clear-Host
        Enter-ScriptMenu -GroupedModules $groupedModules -AllModules $allModules
        Clear-Host
        continue
    }

    # Command Mode
    if ($choice -eq 'C' -or $choice -eq 'c') {
        Clear-Host
        Enter-CommandMode
        Clear-Host
        continue
    }

    # App Mode
    if ($choice -eq 'A' -or $choice -eq 'a') {
        Clear-Host
        Enter-AppMode
        Clear-Host
        continue
    }

    # Run Profile
    if ($choice -eq 'P' -or $choice -eq 'p') {
        Clear-Host
        Invoke-ProfileExecution -AllModules $allModules
        Write-Host ""
        Show-Separator
        Write-Host "Profile Execution Completed" -ForegroundColor Green
        Show-Separator
        Write-Host ""
        if (Confirm-Execution -Message "Start a new kitting session for another device?") {
            Invoke-NewKittingSession
        }
        Clear-Host
        continue
    }

    # New Kitting Session
    if ($choice -eq 'N' -or $choice -eq 'n') {
        Clear-Host
        Invoke-NewKittingSession
        Clear-Host
        continue
    }

    # History
    if ($choice -eq 'R' -or $choice -eq 'r') {
        Clear-Host
        Enter-HistoryMode
        Clear-Host
        continue
    }

    # Export History
    if ($choice -eq 'EH' -or $choice -eq 'eh') {
        Write-Host ""
        $null = Export-ExecutionHistory
        Write-Host ""
        Wait-KeyPress
        Clear-Host
        continue
    }

    # Restart with AutoRun
    if ($choice -eq 'RE' -or $choice -eq 're') {
        Write-Host ""
        Show-Separator
        Write-Host "Restart with AutoRun" -ForegroundColor Yellow
        Show-Separator
        Write-Host ""

        if (-not (Confirm-Execution -Message "Register RunOnce and restart the computer?")) {
            Show-Info "Canceled"
            Wait-KeyPress
            Clear-Host
            continue
        }

        Write-Host ""
        if (-not (Register-FabriqRunOnce)) {
            Wait-KeyPress
            Clear-Host
            continue
        }

        Invoke-CountdownRestart
        continue
    }

    # Refabriq (Restart Fabriq)
    if ($choice -eq 'RF' -or $choice -eq 'rf') {
        Write-Host ""
        Show-Info "Restarting Fabriq..."

        Stop-StatusMonitor -MonitorProcess $global:FabriqStatusMonitorProcess

        $fabriqRoot = (Resolve-Path ".").Path
        $fabriqBat = Join-Path $fabriqRoot "Fabriq.bat"

        if (Test-Path $fabriqBat) {
            Start-Process cmd.exe -ArgumentList "/c `"$fabriqBat`"" -WorkingDirectory $fabriqRoot
        }
        else {
            Show-Error "Fabriq.bat not found: $fabriqBat"
            Wait-KeyPress
            Clear-Host
            continue
        }

        try { Stop-Transcript | Out-Null } catch { }
        exit 0
    }

    # Manifeste du Surkitinisme
    if ($choice -eq 'M' -or $choice -eq 'm') {
        Show-Manifesto
        Clear-Host
        continue
    }

    # Invalid selection
    Write-Host ""
    Show-Error "Invalid selection"
    Write-Host ""
}

# Safety net: Exit-Fabriq is idempotent, safe to call even if already called
Exit-Fabriq