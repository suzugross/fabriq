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

# Enable sleep suppression while Fabriq is running
Enable-SleepSuppression

# Set compact console window size
Set-ConsoleSize -Columns 75 -Lines 35

# ========================================
# Constants
# ========================================
$HOSTLIST_CSV = ".\kernel\hostlist.csv"
$CATEGORIES_CSV = ".\kernel\categories.csv"
$MODULES_DIR = ".\modules"
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

# ========================================
# Function: Load Category Definitions
# ========================================
function Load-Categories {
    $categoryOrder = @{}

    if (Test-Path $CATEGORIES_CSV) {
        try {
            $categories = Import-Csv -Path $CATEGORIES_CSV -Encoding Default
            foreach ($cat in $categories) {
                $categoryOrder[$cat.Category] = [int]$cat.Order
            }
            Show-Success "Loaded categories.csv ($(($categories | Measure-Object).Count) items)"
        }
        catch {
            Show-Error "Failed to load categories.csv: $_"
        }
    }
    else {
        Show-Info "categories.csv not found. Using default order."
    }

    return $categoryOrder
}

# ========================================
# Function: Auto-detect Modules
# ========================================
function Load-ModulesFromDirectory {
    param(
        [string]$ModulesPath,
        [string]$ModuleType
    )

    $modules = @()

    if (-not (Test-Path $ModulesPath)) {
        return $modules
    }

    $dirs = Get-ChildItem $ModulesPath -Directory -ErrorAction SilentlyContinue

    foreach ($dir in $dirs) {
        $moduleCsv = Join-Path $dir.FullName "module.csv"
        if (Test-Path $moduleCsv) {
            try {
                $entries = Import-Csv $moduleCsv -Encoding Default
                foreach ($entry in $entries) {
                    # Check Enabled (Default is enabled if omitted)
                    $enabled = $entry.Enabled
                    if ($enabled -eq "0") {
                        continue
                    }

                    # Order (Default is 100)
                    $order = 100
                    if ($entry.Order -and $entry.Order -match '^\d+$') {
                        $order = [int]$entry.Order
                    }

                    $modules += [PSCustomObject]@{
                        MenuName     = $entry.MenuName
                        Category     = $entry.Category
                        Script       = Join-Path $dir.FullName $entry.Script
                        Order        = $order
                        ModuleType   = $ModuleType
                        ModuleDir    = $dir.Name
                        RelativePath = "$ModuleType\$($dir.Name)\$($entry.Script)"
                    }
                }
            }
            catch {
                Show-Error "Error loading module.csv: $($dir.Name) - $_"
            }
        }
    }

    return $modules
}

# ========================================
# Function: Load All Modules
# ========================================
function Load-AllModules {
    $allModules = @()

    # Standard modules
    $standardPath = Join-Path $MODULES_DIR "standard"
    $standardModules = Load-ModulesFromDirectory -ModulesPath $standardPath -ModuleType "standard"
    $allModules += $standardModules

    # Extended modules
    $extendedPath = Join-Path $MODULES_DIR "extended"
    $extendedModules = Load-ModulesFromDirectory -ModulesPath $extendedPath -ModuleType "extended"
    $allModules += $extendedModules

    $count = ($allModules | Measure-Object).Count
    Show-Success "Modules loaded ($count items)"

    return $allModules
}

# ========================================
# Function: Build Menu by Category
# ========================================
function Build-CategoryMenu {
    param(
        [array]$Modules,
        [hashtable]$CategoryOrder
    )

    # Group by category
    $grouped = $Modules | Group-Object -Property Category

    # Sort by category order
    $sorted = $grouped | Sort-Object {
        $order = $CategoryOrder[$_.Name]
        if ($null -eq $order) { 999 } else { $order }
    }

    return $sorted
}

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

    Show-Separator
    Write-Host ""

    while ($true) {
        Write-Host -NoNewline "Please enter the ID: "
        $userInput = Read-Host

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
# Function: Show Main Menu (By Category)
# ========================================
function Show-MainMenu {
    param(
        [array]$GroupedModules,
        [hashtable]$MenuMap
    )

    Write-Host ""
    Show-Separator
    Write-Host "Fabriq ver2.1 - Manifeste du Surkitinisme -" -ForegroundColor Green
    Show-Separator
    Write-Host "  Selected Host: $env:SELECTED_NEW_PCNAME" -ForegroundColor White
    Show-Separator

    $menuIndex = 1

    foreach ($category in $GroupedModules) {
        Show-CategorySeparator -Name $category.Name

        $items = $category.Group | Sort-Object Order

        foreach ($item in $items) {
            Write-Host "  [$menuIndex] $($item.MenuName)" -ForegroundColor White
            $MenuMap[$menuIndex] = $item
            $menuIndex++
        }
    }

    Write-Host ""
    Write-Host "----------------------------------------" -ForegroundColor DarkGray
    Write-Host "  [C] Command  [A] Apps  [H] Host Info  [Q] Quit" -ForegroundColor DarkGray
    Write-Host "  [P] Run Profile  [R] History" -ForegroundColor DarkGray
    Write-Host "  [re] Restart  [eh] Export  [rf] Refabriq" -ForegroundColor DarkGray
    Write-Host "  * Batch Run: 1,3,5 or 1-5 allowed" -ForegroundColor DarkGray
    Write-Host "----------------------------------------" -ForegroundColor DarkGray
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
            return ($status -eq "Success")
        }
        else {
            # レガシーパス: ModuleResult 未返却（全モジュール移行済み）
            Write-Host ""
            Write-Verbose "ModuleResult not returned from: $ScriptPath"
            Show-Warning "Script execution completed (status unverified)"
            Add-ExecutionResult -Operation $ModuleName -Status "Success" -Message "(legacy - unverified)"
            $null = Write-ExecutionHistory -ModuleName $ModuleName -Category $Category -Status "Success" -Message "(legacy - unverified)"
            return $true
        }
    }
    catch {
        Write-Host ""
        Show-Error "Error occurred during script execution: $_"
        Add-ExecutionResult -Operation $ModuleName -Status "Error" -Message $_.Exception.Message
        $null = Write-ExecutionHistory -ModuleName $ModuleName -Category $Category -Status "Error" -Message $_.Exception.Message
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
        # Profile restart support (optional)
        [string]$ProfilePath = "",
        [string]$ProfileName = ""
    )

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
            $fabriqRoot = (Resolve-Path ".").Path
            $fabriqBat = Join-Path $fabriqRoot "Fabriq.bat"
            $runOncePath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"

            try {
                if (-not (Test-Path $runOncePath)) {
                    New-Item -Path $runOncePath -Force | Out-Null
                }
                $runOnceValue = "cmd /c `"$fabriqBat`""
                New-ItemProperty -Path $runOncePath -Name "FabriqAutoStart" `
                    -Value $runOnceValue -PropertyType String -Force -ErrorAction Stop | Out-Null
                Write-Host "[SUCCESS] RunOnce registered" -ForegroundColor Green
            }
            catch {
                Write-Host "[ERROR] Failed to register RunOnce: $_" -ForegroundColor Red
                Remove-ResumeState
                Add-ExecutionResult -Operation "[RESTART]" -Status "Error" -Message "RunOnce registration failed"
                $null = Write-ExecutionHistory -ModuleName "[RESTART]" -Category "System" -Status "Error" -Message "RunOnce failed: $_"
                if ($StopOnError) { break } else { continue }
            }

            # Record in execution history
            Add-ExecutionResult -Operation "[RESTART]" -Status "Success" -Message "Restarting..."
            $null = Write-ExecutionHistory -ModuleName "[RESTART]" -Category "System" -Status "Success" -Message "Profile restart (ResumeAfter: $($module.Order))"

            # Countdown
            Write-Host ""
            Write-Host "The computer will restart in 5 seconds..." -ForegroundColor Yellow
            Write-Host "Press Ctrl+C to abort" -ForegroundColor Yellow
            Write-Host ""
            for ($i = 5; $i -ge 1; $i--) {
                Write-Host "`r  Restarting in $i seconds... " -NoNewline -ForegroundColor Yellow
                Start-Sleep -Seconds 1
            }
            Write-Host ""

            Restart-Computer -Force
            Start-Sleep -Seconds 30  # Fallback wait after Restart-Computer
            return
        }

        # __REEXPLORER__ marker handling
        if ($module._IsReexplorer) {
            Show-BatchProgress -Current $current -Total $total -ItemName "[REEXPLORER]"
            Show-Info "Restarting Explorer..."
            try {
                Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
                Start-Sleep -Seconds 2
                # Explorer usually auto-restarts; ensure it does
                if (-not (Get-Process -Name explorer -ErrorAction SilentlyContinue)) {
                    Start-Process explorer.exe
                }
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
                    $transcriptPath = ".\logs\log_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
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

            Write-Host "The computer will shut down in 5 seconds..." -ForegroundColor Yellow
            Write-Host "Press Ctrl+C to abort" -ForegroundColor Yellow
            Write-Host ""
            for ($i = 5; $i -ge 1; $i--) {
                Write-Host "`r  Shutting down in $i seconds... " -NoNewline -ForegroundColor Yellow
                Start-Sleep -Seconds 1
            }
            Write-Host ""

            Stop-Computer -Force
            Start-Sleep -Seconds 30
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

        $result = Invoke-SafeCommand -OperationName $module.MenuName -ScriptBlock {
            & $module.Script
        } -ContinueOnError

        Add-ExecutionResult -Operation $module.MenuName -Status $result.Status -Message $result.Message
        $null = Write-ExecutionHistory -ModuleName $module.MenuName -Category $module.Category -Status $result.Status -Message $result.Message

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

        # Auto-run log upload
        $logUploaderScript = ".\modules\extended\log_uploader\log_uploader.ps1"
        if (Test-Path $logUploaderScript) {
            $destConfig = ".\kernel\log_destinations.csv"
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
                    & $logUploaderScript
                }
                catch {
                    Show-Warning "Log upload failed: $($_.Exception.Message)"
                }
            }
        }
    }

    Wait-KeyPress
}

# ========================================
# Function: Profile Execution
# ========================================
function Invoke-ProfileExecution {
    param([array]$AllModules)

    $profiles = Load-Profiles -AllModules $AllModules

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

            # Confirmation + StopOnError selection
            Clear-Host
            $confirmation = Show-ProfileConfirmation -SelectedProfile $selectedProfile -Modules $validation.ValidModules -InvalidPaths $validation.InvalidPaths

            if ($null -ne $confirmation -and $confirmation.Confirmed) {
                Invoke-BatchExecution -SelectedModules $validation.ValidModules `
                    -StopOnError:$confirmation.StopOnError `
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
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile = Join-Path $logDir "log_$timestamp.txt"
$global:FabriqTranscriptPath = $logFile
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
Initialize-Session

# Load hostlist.csv
Show-Info "Loading hostlist.csv..."
$hostList = Load-HostList
if (-not $hostList) {
    Write-Host ""
    Show-Error "Aborting process"
    Stop-Transcript | Out-Null
    exit 1
}
Write-Host ""

# Check for resume state (profile restart)
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

    $resumeChoice = $null
    while ($true) {
        Write-Host -NoNewline "Resume profile execution? (Y/N): "
        $resumeChoice = Read-Host
        if ($resumeChoice -eq 'Y' -or $resumeChoice -eq 'y') { break }
        if ($resumeChoice -eq 'N' -or $resumeChoice -eq 'n') { break }
        Write-Host "[INFO] Please enter Y or N" -ForegroundColor Yellow
    }

    if ($resumeChoice -eq 'Y' -or $resumeChoice -eq 'y') {
        # Restore environment variables (skip host selection)
        Restore-HostEnvironment -HostEnv $resumeState.HostEnvironment
        Show-Success "Environment restored for: $($resumeState.HostEnvironment.SELECTED_NEW_PCNAME)"

        # Inherit SessionID for history continuity
        $script:SessionID = $resumeState.SessionID

        # Restore execution history for the selected PC
        Restore-ExecutionHistory
        Write-Host ""

        # Load modules
        Clear-Host
        Show-Info "Loading categories.csv..."
        $categoryOrder = Load-Categories
        Write-Host ""

        Show-Info "Detecting modules..."
        $allModules = Load-AllModules
        if (($allModules | Measure-Object).Count -eq 0) {
            Show-Error "No valid modules found"
            Remove-ResumeState
            Stop-Transcript | Out-Null
            exit 1
        }
        Write-Host ""

        $groupedModules = Build-CategoryMenu -Modules $allModules -CategoryOrder $categoryOrder

        # Launch Status Monitor early for resume flow
        Write-StatusFile -Phase "idle"
        $script:StatusMonitorProcess = $null
        try {
            $monitorScript = ".\kernel\status_monitor.ps1"
            if (Test-Path $monitorScript) {
                $statusFileFullPath = (Resolve-Path $script:StatusFilePath).Path
                $script:StatusMonitorProcess = Start-Process powershell.exe -ArgumentList @(
                    "-NoProfile", "-ExecutionPolicy", "Unrestricted",
                    "-File", $monitorScript,
                    "-StatusFilePath", $statusFileFullPath
                ) -WindowStyle Hidden -PassThru
                Show-Info "Status Monitor started (PID: $($script:StatusMonitorProcess.Id))"
                Write-Host ""

                Start-Sleep -Milliseconds 800
                try {
                    $wsh = New-Object -ComObject WScript.Shell
                    $wsh.AppActivate($PID) | Out-Null
                    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wsh) | Out-Null
                }
                catch { }
            }
        }
        catch {
            Show-Warning "Failed to start Status Monitor: $_"
            Write-Host ""
        }

        # Resolve remaining modules from profile
        $validation = Resolve-ProfileModules -ProfileCsvPath $resumeState.ProfilePath -AllModules $allModules

        $remainingModules = @($validation.ValidModules | Where-Object {
            $_.Order -gt $resumeState.ResumeAfterOrder
        })

        if ($remainingModules.Count -eq 0) {
            Show-Info "No remaining modules to execute"
            Remove-ResumeState
        }
        else {
            Write-Host ""
            Show-Info "Resuming profile: $($resumeState.ProfileName)"
            Show-Info "Remaining modules: $($remainingModules.Count)"
            Write-Host ""

            # Restore completed module results to summary
            foreach ($cm in $resumeState.CompletedModules) {
                Add-ExecutionResult -Operation $cm.MenuName -Status $cm.Status -Message "(completed before restart)"
            }
            Add-ExecutionResult -Operation "[RESTART]" -Status "Success" -Message "Resumed after restart"

            # Execute remaining modules
            Invoke-BatchExecution -SelectedModules $remainingModules `
                -StopOnError:$resumeState.StopOnError `
                -ProfilePath $resumeState.ProfilePath `
                -ProfileName $resumeState.ProfileName

            Remove-ResumeState
        }
    }
    else {
        # Resume declined: clear state, proceed with normal startup
        Remove-ResumeState
        Show-Info "Resume state cleared. Starting normally."
        Write-Host ""

        $selectedHost = Select-Host -HostList $hostList
        Write-Host ""
        Set-SelectedHostEnvironment -SelectedHost $selectedHost
        Write-Host ""
        Restore-ExecutionHistory
        Write-Host ""

        Clear-Host
        Show-Info "Loading categories.csv..."
        $categoryOrder = Load-Categories
        Write-Host ""

        Show-Info "Detecting modules..."
        $allModules = Load-AllModules
        if (($allModules | Measure-Object).Count -eq 0) {
            Show-Error "No valid modules found"
            Stop-Transcript | Out-Null
            exit 1
        }
        Write-Host ""

        $groupedModules = Build-CategoryMenu -Modules $allModules -CategoryOrder $categoryOrder
    }
}
else {
    # Normal startup (no resume state)
    $selectedHost = Select-Host -HostList $hostList
    Write-Host ""
    Set-SelectedHostEnvironment -SelectedHost $selectedHost
    Write-Host ""
    Restore-ExecutionHistory
    Write-Host ""

    Clear-Host
    Show-Info "Loading categories.csv..."
    $categoryOrder = Load-Categories
    Write-Host ""

    Show-Info "Detecting modules..."
    $allModules = Load-AllModules
    if (($allModules | Measure-Object).Count -eq 0) {
        Show-Error "No valid modules found"
        Stop-Transcript | Out-Null
        exit 1
    }
    Write-Host ""

    $groupedModules = Build-CategoryMenu -Modules $allModules -CategoryOrder $categoryOrder
}

# Initialize Menu Map
$menuMap = @{}

# Launch Status Monitor Window (skip if already started during resume flow)
if ($null -eq $script:StatusMonitorProcess) {
    Write-StatusFile -Phase "idle"
    try {
        $monitorScript = ".\kernel\status_monitor.ps1"
        if (Test-Path $monitorScript) {
            $statusFileFullPath = (Resolve-Path $script:StatusFilePath).Path
            $script:StatusMonitorProcess = Start-Process powershell.exe -ArgumentList @(
                "-NoProfile", "-ExecutionPolicy", "Unrestricted",
                "-File", $monitorScript,
                "-StatusFilePath", $statusFileFullPath
            ) -WindowStyle Hidden -PassThru
            Show-Info "Status Monitor started (PID: $($script:StatusMonitorProcess.Id))"
            Write-Host ""

            # Return focus to Fabriq console after monitor window appears
            Start-Sleep -Milliseconds 800
            Set-ConsoleForeground
        }
    }
    catch {
        Show-Warning "Failed to start Status Monitor: $_"
        Write-Host ""
    }
}

# Main Loop
while ($true) {
    Write-StatusFile -Phase "idle"
    $menuMap.Clear()
    Show-MainMenu -GroupedModules $groupedModules -MenuMap $menuMap

    Write-Host ""
    Write-Host -NoNewline "Please enter the number: "
    $choice = Read-Host

    # Quit
    if ($choice -eq "Q" -or $choice -eq "q" -or $choice -eq "0") {
        Write-Host ""
        Show-Info "Exiting"

        # Close Status Monitor Window
        if ($script:StatusMonitorProcess -and -not $script:StatusMonitorProcess.HasExited) {
            try {
                $script:StatusMonitorProcess.CloseMainWindow() | Out-Null
                if (-not $script:StatusMonitorProcess.WaitForExit(2000)) {
                    $script:StatusMonitorProcess.Kill()
                }
            }
            catch { }
        }
        Remove-StatusFile
        Disable-SleepSuppression

        break
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

    # Show Host Info
    if ($choice -eq 'H' -or $choice -eq 'h') {
        Clear-Host
        Show-HostInfo
        Clear-Host
        continue
    }

    # Run Profile
    if ($choice -eq 'P' -or $choice -eq 'p') {
        Clear-Host
        Invoke-ProfileExecution -AllModules $allModules
        Clear-Host
        continue
    }

    # Show History
    if ($choice -eq 'R' -or $choice -eq 'r') {
        Clear-Host
        Enter-HistoryMode
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

        $fabriqRoot = (Resolve-Path ".").Path
        $fabriqBat = Join-Path $fabriqRoot "Fabriq.bat"

        if (-not (Test-Path $fabriqBat)) {
            Show-Error "Fabriq.bat not found: $fabriqBat"
            Wait-KeyPress
            Clear-Host
            continue
        }

        $runOncePath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"
        $runOnceName = "FabriqAutoStart"
        $runOnceValue = "cmd /c `"$fabriqBat`""

        Write-Host "  RunOnce: $runOnceName" -ForegroundColor White
        Write-Host "  Value:   $runOnceValue" -ForegroundColor Gray
        Write-Host ""

        if (-not (Confirm-Execution -Message "Register RunOnce and restart the computer?")) {
            Show-Info "Canceled"
            Wait-KeyPress
            Clear-Host
            continue
        }

        Write-Host ""
        try {
            if (-not (Test-Path $runOncePath)) {
                New-Item -Path $runOncePath -Force | Out-Null
            }
            New-ItemProperty -Path $runOncePath -Name $runOnceName -Value $runOnceValue -PropertyType String -Force -ErrorAction Stop | Out-Null
            Write-Host "[SUCCESS] RunOnce registered" -ForegroundColor Green
        }
        catch {
            Show-Error "Failed to register RunOnce: $_"
            Wait-KeyPress
            Clear-Host
            continue
        }

        Write-Host ""
        Write-Host "The computer will restart in 5 seconds..." -ForegroundColor Yellow
        Write-Host "Press Ctrl+C to abort" -ForegroundColor Yellow
        Write-Host ""
        for ($i = 5; $i -ge 1; $i--) {
            Write-Host "`r  Restarting in $i seconds... " -NoNewline -ForegroundColor Yellow
            Start-Sleep -Seconds 1
        }
        Write-Host ""

        Restart-Computer -Force
        Start-Sleep -Seconds 30
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

    # Refabriq (Restart Fabriq)
    if ($choice -eq 'RF' -or $choice -eq 'rf') {
        Write-Host ""
        Show-Info "Restarting Fabriq..."

        # Close Status Monitor
        if ($script:StatusMonitorProcess -and -not $script:StatusMonitorProcess.HasExited) {
            try {
                $script:StatusMonitorProcess.CloseMainWindow() | Out-Null
                if (-not $script:StatusMonitorProcess.WaitForExit(2000)) {
                    $script:StatusMonitorProcess.Kill()
                }
            }
            catch { }
        }
        Remove-StatusFile

        # Launch new Fabriq instance
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

        # Stop transcript and exit current instance
        try { Stop-Transcript | Out-Null } catch { }
        exit 0
    }

    # Check Batch Input
    if (Test-BatchInput -InputString $choice) {
        $selectedNumbers = Parse-MenuSelection -InputString $choice

        # Filter valid numbers only
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
        Write-Host ""
        Clear-Host
        Show-Info "Executing [$($selectedModule.MenuName)]"
        Write-Host ""

        $null = Invoke-KittingScript -ScriptPath $selectedModule.Script -ModuleName $selectedModule.MenuName -Category $selectedModule.Category

        Wait-KeyPress
        Clear-Host
    }
    else {
        Write-Host ""
        Show-Error "Invalid selection"
        Write-Host ""
    }
}

Write-Host ""
Show-Separator

# Ensure Status Monitor is closed (safety net)
if ($script:StatusMonitorProcess -and -not $script:StatusMonitorProcess.HasExited) {
    try {
        $script:StatusMonitorProcess.CloseMainWindow() | Out-Null
        if (-not $script:StatusMonitorProcess.WaitForExit(2000)) {
            $script:StatusMonitorProcess.Kill()
        }
    }
    catch { }
}
Remove-StatusFile

# Stop Logging
Stop-Transcript | Out-Null