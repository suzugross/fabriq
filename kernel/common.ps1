# ========================================
# Easy Kitting Batch - Common Function Library v2.2
# ========================================

# ========================================
# Global Variables
# ========================================
$script:ExecutionResults = @()
$script:SessionID = Get-Date -Format "yyyyMMdd_HHmmss"
$script:HistoryPath = ".\logs\history\execution_history.csv"
$script:ProfilesDir = ".\profiles"
$script:StatusFilePath = ".\kernel\json\status.json"
$script:ResumeStatePath = ".\kernel\json\resume_state.json"
$script:SessionFilePath = ".\kernel\json\session.json"
$script:SourceMediaIdPath = ".\kernel\source_media.id"
$script:WorkersCsvPath = ".\kernel\csv\workers.csv"

# Session info (populated by Initialize-Session)
$script:SessionInfo = $null

# AutoPilot Mode (Profile execution only)
$global:AutoPilotMode = $false
$global:AutoPilotWaitSec = 3

# ========================================
# Sleep Suppression (SetThreadExecutionState)
# ========================================
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public class SleepSuppressor {
    [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    public static extern uint SetThreadExecutionState(uint esFlags);
    public const uint ES_CONTINUOUS       = 0x80000000;
    public const uint ES_SYSTEM_REQUIRED  = 0x00000001;
    public const uint ES_DISPLAY_REQUIRED = 0x00000002;
}
'@ -ErrorAction SilentlyContinue

# ========================================
# Console Focus (GetConsoleWindow + SetForegroundWindow)
# ========================================
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public class ConsoleFocus {
    [DllImport("kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int Left, Top, Right, Bottom; }
}
'@ -ErrorAction SilentlyContinue

function Set-ConsoleForeground {
    try {
        $hwnd = [ConsoleFocus]::GetConsoleWindow()
        if ($hwnd -ne [IntPtr]::Zero) {
            [ConsoleFocus]::SetForegroundWindow($hwnd) | Out-Null
        }
    }
    catch { }
}

function Set-ConsoleSize {
    param(
        [int]$Columns = 80,
        [int]$Lines   = 35
    )
    try {
        $rawUI = $Host.UI.RawUI
        # Window = compact visible area, Buffer = large scrollback
        $bufSize = $rawUI.BufferSize
        if ($bufSize.Width -gt $Columns) {
            $rawUI.WindowSize = New-Object System.Management.Automation.Host.Size($Columns, $Lines)
            $rawUI.BufferSize = New-Object System.Management.Automation.Host.Size($Columns, 9999)
        }
        else {
            $rawUI.BufferSize = New-Object System.Management.Automation.Host.Size($Columns, 9999)
            $rawUI.WindowSize = New-Object System.Management.Automation.Host.Size($Columns, $Lines)
        }
    }
    catch { }
}

function Enable-SleepSuppression {
    [SleepSuppressor]::SetThreadExecutionState(
        [SleepSuppressor]::ES_CONTINUOUS -bor
        [SleepSuppressor]::ES_SYSTEM_REQUIRED -bor
        [SleepSuppressor]::ES_DISPLAY_REQUIRED
    ) | Out-Null
}

function Disable-SleepSuppression {
    [SleepSuppressor]::SetThreadExecutionState(
        [SleepSuppressor]::ES_CONTINUOUS
    ) | Out-Null
}

# ========================================
# Display Functions
# ========================================

function Show-Separator {
    Write-Host "========================================" -ForegroundColor Cyan
}

function Show-CategorySeparator {
    param([string]$Name)
    Write-Host ""
    # Changed from Japanese symbols to standard equals signs for better compatibility
    Write-Host "=== $Name ===" -ForegroundColor Cyan
}

function Show-Info {
    param([string]$Message)
    Write-Host "[INFO] $Message" -ForegroundColor Cyan
}

function Show-Success {
    param([string]$Message)
    Write-Host "[SUCCESS] $Message" -ForegroundColor Green
}

function Show-Warning {
    param([string]$Message)
    Write-Host "[WARNING] $Message" -ForegroundColor Yellow
}

function Show-Error {
    param([string]$Message)
    Write-Host "[ERROR] $Message" -ForegroundColor Red
}

function Show-Skip {
    param([string]$Message)
    Write-Host "[SKIP] $Message" -ForegroundColor DarkGray
}

# ========================================
# Module Result Functions
# ========================================

function New-ModuleResult {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("Success", "Error", "Cancelled", "Skipped", "Partial")]
        [string]$Status,
        [string]$Message = "",
        [array]$Details = @()
    )
    $resultObj = [PSCustomObject]@{
        _IsModuleResult = $true
        Status          = $Status
        Message         = $Message
        Details         = $Details
        Timestamp       = Get-Date
    }
    # グローバル変数にも保存（パイプライン出力キャプチャ失敗時のフォールバック）
    $global:_LastModuleResult = $resultObj
    return $resultObj
}

# ========================================
# Pattern Layer Functions
# ========================================
# モジュール共通の定型パターンを関数化
# New-BatchResult: 結果集計表示 + ModuleResult 返却
# Confirm-ModuleExecution: 確認 + キャンセル処理
# Import-ModuleCsv: CSV 読み込み + フィルタ + カラム検証

function New-BatchResult {
    param(
        [int]$Success = 0,
        [int]$Skip = 0,
        [int]$Fail = 0,
        [string]$Title = "Execution Results",
        [string]$MessageSuffix = ""
    )

    Show-Separator
    Write-Host $Title -ForegroundColor Cyan
    Show-Separator

    if ($Success -gt 0) {
        Write-Host "  Success: $Success items" -ForegroundColor Green
    }
    if ($Skip -gt 0) {
        Write-Host "  Skipped: $Skip items" -ForegroundColor Gray
    }
    if ($Fail -gt 0) {
        Write-Host "  Failed:  $Fail items" -ForegroundColor Red
    }

    Show-Separator
    Write-Host ""

    $status = if ($Fail -eq 0 -and $Success -gt 0) { "Success" }
        elseif ($Success -gt 0 -and $Fail -gt 0) { "Partial" }
        elseif ($Fail -eq 0 -and $Skip -gt 0 -and $Success -eq 0) { "Skipped" }
        elseif ($Fail -gt 0 -and $Success -eq 0) { "Error" }
        else { "Success" }

    $msg = "Success: $Success, Skip: $Skip, Fail: $Fail"
    if ($MessageSuffix) { $msg += " $MessageSuffix" }

    return (New-ModuleResult -Status $status -Message $msg)
}

function Confirm-ModuleExecution {
    param(
        [string]$Message = "Are you sure you want to execute?"
    )

    if (-not (Confirm-Execution -Message $Message)) {
        Write-Host ""
        Show-Info "Canceled"
        Write-Host ""
        return (New-ModuleResult -Status "Cancelled" -Message "User canceled")
    }

    return $null
}

function Import-ModuleCsv {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [switch]$FilterEnabled,
        [string[]]$RequiredColumns
    )

    $allItems = Import-CsvSafe -Path $Path -Description ([System.IO.Path]::GetFileName($Path))
    if ($null -eq $allItems) { return $null }
    if ($allItems.Count -eq 0) { return $null }

    if ($RequiredColumns) {
        if (-not (Test-CsvColumns -CsvData $allItems -RequiredColumns $RequiredColumns -CsvName ([System.IO.Path]::GetFileName($Path)))) {
            return $null
        }
    }

    if ($FilterEnabled) {
        $filtered = @($allItems | Where-Object { $_.Enabled -eq "1" })
        if ($filtered.Count -eq 0) {
            Show-Skip "No enabled entries in $([System.IO.Path]::GetFileName($Path))"
            return @()
        }
        Show-Info "Loaded $($filtered.Count) enabled entries (total: $($allItems.Count))"
        return $filtered
    }

    return $allItems
}

function Show-Progress {
    param(
        [string]$Activity,
        [int]$Current,
        [int]$Total
    )
    $percent = [math]::Round(($Current / $Total) * 100)
    Write-Host "[$Current/$Total] $Activity ($percent%)" -ForegroundColor White
}

function Show-BatchProgress {
    param(
        [int]$Current,
        [int]$Total,
        [string]$ItemName
    )
    Write-Host ""
    Show-CategorySeparator "$Current/$Total : $ItemName"
}

# ========================================
# Confirmation Functions
# ========================================

function Confirm-Execution {
    param(
        [string]$Message = "Are you sure you want to execute?"
    )

    # AutoPilot: auto-confirm
    if ($global:AutoPilotMode) {
        Write-Host "[AUTOPILOT] $Message -> Y (auto)" -ForegroundColor Magenta
        return $true
    }

    while ($true) {
        Write-Host -NoNewline "$Message (Y/N): "
        $response = Read-Host

        if ($response -eq 'Y' -or $response -eq 'y') {
            return $true
        }
        if ($response -eq 'N' -or $response -eq 'n') {
            return $false
        }

        Write-Host "[INFO] Please enter Y or N" -ForegroundColor Yellow
    }
}

function Wait-KeyPress {
    param([string]$Message = "Press Enter to continue...")

    # AutoPilot: skip wait
    if ($global:AutoPilotMode) {
        return
    }

    Write-Host ""
    Write-Host $Message
    Read-Host
}

function Wait-NetworkReady {
    param(
        [string]$Target        = "8.8.8.8",
        [int]$RetryIntervalSec = 10,
        [int]$PingCount        = 1
    )
    while ($true) {
        Write-Host "Checking network connectivity ($Target)..." -ForegroundColor White
        $reachable = Test-Connection -ComputerName $Target -Count $PingCount `
                        -Quiet -ErrorAction SilentlyContinue
        if ($reachable) {
            Show-Success "Network connectivity OK ($Target)"
            return
        }
        Show-Warning "Network unreachable. Retrying in ${RetryIntervalSec}s... (Ctrl+C to abort)"
        Start-Sleep -Seconds $RetryIntervalSec
    }
}

function Get-HardwareUniqueId {
    # ========================================
    # Returns a hardware-unique ID for use in log file names.
    # Priority 1: BIOS Serial Number (Win32_BIOS)
    # Priority 2: Physical NIC MAC Address (first Up adapter)
    # Fallback:   "UNKNOWN"
    # ========================================

    $invalidSerials = @(
        "", "None", "N/A", "INVALID",
        "To be filled by O.E.M.", "To Be Filled By O.E.M.",
        "Default string", "System Serial Number", "00000000"
    )

    # --- Priority 1: BIOS Serial Number ---
    try {
        $bios   = Get-WmiObject -Class Win32_BIOS -ErrorAction SilentlyContinue
        $serial = if ($bios) { $bios.SerialNumber } else { "" }
        if (-not [string]::IsNullOrWhiteSpace($serial) -and
            $serial.Trim() -notin $invalidSerials) {
            # Replace characters invalid in file names with hyphens; trim leading/trailing hyphens
            $sanitized = ($serial.Trim() -replace '[^a-zA-Z0-9\-]', '-').Trim('-')
            if ($sanitized.Length -gt 0) { return $sanitized }
        }
    }
    catch { }

    # --- Priority 2: Physical NIC MAC Address (first Up adapter) ---
    try {
        $nic = Get-NetAdapter -Physical -ErrorAction SilentlyContinue |
               Where-Object { $_.Status -eq "Up" } |
               Select-Object -First 1
        if ($nic -and $nic.MacAddress) {
            # Normalize AA-BB-CC-DD-EE-FF and AA:BB:CC:DD:EE:FF to AABBCCDDEEFF
            return ($nic.MacAddress -replace '[-:]', '')
        }
    }
    catch { }

    return "UNKNOWN"
}

# ========================================
# CSV Operations
# ========================================

function Import-CsvSafe {
    param(
        [string]$Path,
        [string]$Description = "CSV"
    )

    if (-not (Test-Path $Path)) {
        Show-Error "${Description} not found: $Path"
        return $null
    }

    try {
        $data = @(Import-Csv -Path $Path -Encoding Default)
        if ($data.Count -eq 0) {
            Show-Warning "${Description} has no data: $Path"
            return @()
        }
        return $data
    }
    catch {
        Show-Error "Failed to load ${Description}: $_"
        return $null
    }
}

function Test-CsvColumns {
    param(
        [array]$CsvData,
        [string[]]$RequiredColumns,
        [string]$CsvName = "CSV"
    )

    if ($null -eq $CsvData -or $CsvData.Count -eq 0) {
        return $false
    }

    $firstRow = $CsvData[0]
    $existingColumns = $firstRow.PSObject.Properties.Name

    $missingColumns = @()
    foreach ($col in $RequiredColumns) {
        if ($col -notin $existingColumns) {
            $missingColumns += $col
        }
    }

    if ($missingColumns.Count -gt 0) {
        Show-Error "${CsvName} is missing required columns: $($missingColumns -join ', ')"
        return $false
    }

    return $true
}

# ========================================
# Error Handling Functions
# ========================================

function Invoke-SafeCommand {
    param(
        [scriptblock]$ScriptBlock,
        [string]$OperationName,
        [switch]$ContinueOnError
    )

    $startTime = Get-Date

    $result = [PSCustomObject]@{
        Operation = $OperationName
        Success   = $false
        Status    = "Error"
        Message   = ""
        Duration  = [TimeSpan]::Zero
        Error     = $null
    }

    try {
        # グローバルフォールバック変数をクリア
        $global:_LastModuleResult = $null

        $output = & $ScriptBlock

        # ModuleResult を出力から検出（パイプライン出力から）
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
            # モジュールが自己申告したステータスを使用
            $result.Status = $moduleResult.Status
            $result.Message = $moduleResult.Message
            $result.Success = ($moduleResult.Status -eq "Success")
        }
        else {
            # レガシーパス: ModuleResult 未返却（全モジュール移行済み）
            Write-Verbose "[$OperationName] ModuleResult not returned (legacy module)"
            $result.Success = $true
            $result.Status = "Success"
            $result.Message = "(legacy - unverified)"
        }
    }
    catch {
        $result.Success = $false
        $result.Status = "Error"
        $result.Message = $_.Exception.Message
        $result.Error = $_

        if (-not $ContinueOnError) {
            Show-Error "$OperationName : $($_.Exception.Message)"
        }
    }
    finally {
        $result.Duration = (Get-Date) - $startTime
    }

    return $result
}

function Add-ExecutionResult {
    param(
        [string]$Operation,
        [string]$Status,
        [string]$Message = ""
    )

    $script:ExecutionResults += [PSCustomObject]@{
        Operation = $Operation
        Status    = $Status
        Message   = $Message
        Timestamp = Get-Date
    }

    # ステータスモニター更新
    Write-StatusFile -Phase "executing"
}

function Clear-ExecutionResults {
    # 復元エントリとセパレーターは保持
    $restored = @($script:ExecutionResults | Where-Object {
        $_.IsRestored -eq $true -or $_.Status -eq "Separator"
    })

    if ($restored.Count -gt 0) {
        $script:ExecutionResults = $restored
    }
    else {
        $script:ExecutionResults = @()
    }

    # ステータスモニター更新
    Write-StatusFile -Phase "executing"
}

function Show-ExecutionSummary {
    param(
        [array]$Results = $null,
        [System.TimeSpan]$ElapsedTime = [System.TimeSpan]::Zero
    )

    # ステータスモニター更新（完了状態）
    Write-StatusFile -Phase "complete"

    if ($null -eq $Results) {
        $Results = $script:ExecutionResults
    }

    if ($Results.Count -eq 0) {
        return
    }

    $successCount  = ($Results | Where-Object { $_.Status -eq "Success" }).Count
    $skipCount     = ($Results | Where-Object { $_.Status -eq "Skip" -or $_.Status -eq "Skipped" }).Count
    $cancelCount   = ($Results | Where-Object { $_.Status -eq "Cancelled" }).Count
    $partialCount  = ($Results | Where-Object { $_.Status -eq "Partial" }).Count
    $warnCount     = ($Results | Where-Object { $_.Status -eq "Warning" }).Count
    $errorCount    = ($Results | Where-Object { $_.Status -eq "Error" }).Count

    Write-Host ""
    Show-Separator
    Write-Host "Execution Results" -ForegroundColor Cyan
    Show-Separator

    $successColor = if ($successCount -gt 0) { "Green" } else { "Gray" }
    $skipColor    = if ($skipCount -gt 0) { "DarkGray" } else { "Gray" }
    $cancelColor  = if ($cancelCount -gt 0) { "Yellow" } else { "Gray" }
    $partialColor = if ($partialCount -gt 0) { "Yellow" } else { "Gray" }
    $warnColor    = if ($warnCount -gt 0) { "Yellow" } else { "Gray" }
    $errorColor   = if ($errorCount -gt 0) { "Red" } else { "Gray" }

    Write-Host "  Success:   $successCount items" -ForegroundColor $successColor
    Write-Host "  Skipped:   $skipCount items" -ForegroundColor $skipColor
    Write-Host "  Cancelled: $cancelCount items" -ForegroundColor $cancelColor
    Write-Host "  Partial:   $partialCount items" -ForegroundColor $partialColor
    Write-Host "  Warnings:  $warnCount items" -ForegroundColor $warnColor
    Write-Host "  Errors:    $errorCount items" -ForegroundColor $errorColor

    # Elapsed time
    if ($ElapsedTime.TotalSeconds -gt 0) {
        Write-Host ""
        if ($ElapsedTime.TotalHours -ge 1) {
            $elapsedStr = "{0:0}h {1:0}m {2:0}s" -f [math]::Floor($ElapsedTime.TotalHours), $ElapsedTime.Minutes, $ElapsedTime.Seconds
        }
        elseif ($ElapsedTime.TotalMinutes -ge 1) {
            $elapsedStr = "{0:0}m {1:0}s" -f [math]::Floor($ElapsedTime.TotalMinutes), $ElapsedTime.Seconds
        }
        else {
            $elapsedStr = "{0:0}s" -f [math]::Floor($ElapsedTime.TotalSeconds)
        }
        Write-Host "  Elapsed:   $elapsedStr" -ForegroundColor Cyan
    }

    Show-Separator

    # Show Details
    if ($Results.Count -le 20) {
        Write-Host "Details:" -ForegroundColor White
        foreach ($r in $Results) {
            $icon = switch ($r.Status) {
                "Success"   { "[OK]";      $color = "Green" }
                "Skip"      { "[SKIP]";    $color = "DarkGray" }
                "Skipped"   { "[SKIP]";    $color = "DarkGray" }
                "Cancelled" { "[CANCEL]";  $color = "Yellow" }
                "Partial"   { "[PARTIAL]"; $color = "Yellow" }
                "Warning"   { "[WARN]";    $color = "Yellow" }
                "Error"     { "[ERROR]";   $color = "Red" }
                default     { "[?]";       $color = "Gray" }
            }

            $detail = if ($r.Message) { " ($($r.Message))" } else { "" }
            Write-Host "  $icon $($r.Operation)$detail" -ForegroundColor $color
        }
        Show-Separator
    }
}

# ========================================
# Batch Execution Functions
# ========================================

function Parse-MenuSelection {
    param([string]$InputString)

    $numbers = @()
    # Split by semicolon or comma
    $parts = $InputString -split '[,;]'

    foreach ($part in $parts) {
        $part = $part.Trim()

        # Range (e.g., 1-5)
        if ($part -match '^(\d+)-(\d+)$') {
            $start = [int]$Matches[1]
            $end = [int]$Matches[2]
            if ($start -le $end) {
                for ($i = $start; $i -le $end; $i++) {
                    $numbers += $i
                }
            }
        }
        # Single number
        elseif ($part -match '^\d+$') {
            $numbers += [int]$part
        }
    }

    return $numbers | Select-Object -Unique | Sort-Object
}

function Test-BatchInput {
    param([string]$InputString)

    # Check if input is batch (contains comma, semicolon, or hyphen)
    return ($InputString -match '[,;\-]' -and $InputString -match '\d')
}

function Show-BatchConfirmation {
    param(
        [array]$SelectedModules
    )

    Write-Host ""
    Show-Separator
    Write-Host "The following functions will be executed in batch:" -ForegroundColor Cyan
    Show-Separator

    $index = 1
    foreach ($module in $SelectedModules) {
        Write-Host "  [$index] $($module.MenuName)" -ForegroundColor White
        $index++
    }

    Show-Separator
    Write-Host ""

    return Confirm-Execution -Message "Are you sure you want to execute?"
}

# ========================================
# Execution History Functions
# ========================================

function Initialize-ExecutionHistory {
    # Ensure directory exists
    $historyDir = Split-Path $script:HistoryPath -Parent
    if (-not (Test-Path $historyDir)) {
        $null = New-Item -ItemType Directory -Path $historyDir -Force
    }

    # Migrate from old location (kernel/) if exists
    $oldPath = ".\kernel\execution_history.csv"
    if ((Test-Path $oldPath) -and -not (Test-Path $script:HistoryPath)) {
        Move-Item $oldPath $script:HistoryPath -Force
        $oldBak = "$oldPath.bak"
        if (Test-Path $oldBak) {
            Move-Item $oldBak "$($script:HistoryPath).bak" -Force
        }
        Show-Info "Migrated execution_history.csv to logs/history/"
    }

    # Create backup on startup
    if (Test-Path $script:HistoryPath) {
        try {
            Copy-Item $script:HistoryPath "$($script:HistoryPath).bak" -Force -ErrorAction SilentlyContinue
        }
        catch {
            # Warning only if backup fails
        }
    }

    # Migrate evidence from old locations (logs/) to evidence/
    # gyotaku: logs/gyotaku/ -> evidence/gyotaku/
    $oldGyotakuDir = ".\logs\gyotaku"
    $newGyotakuDir = ".\evidence\gyotaku"
    if ((Test-Path $oldGyotakuDir) -and @(Get-ChildItem $oldGyotakuDir -Recurse -File -ErrorAction SilentlyContinue).Count -gt 0) {
        if (-not (Test-Path $newGyotakuDir)) {
            $null = New-Item -ItemType Directory -Path $newGyotakuDir -Force
        }
        Get-ChildItem $oldGyotakuDir -Directory -ErrorAction SilentlyContinue | ForEach-Object {
            $dest = Join-Path $newGyotakuDir $_.Name
            if (-not (Test-Path $dest)) {
                Move-Item $_.FullName $dest -Force
            }
        }
        Show-Info "Migrated gyotaku to evidence/gyotaku/"
    }

    # pc_information: logs/pc_information_log/ -> evidence/pc_information/
    $oldPcInfoDir = ".\logs\pc_information_log"
    $newPcInfoDir = ".\evidence\pc_information"
    if ((Test-Path $oldPcInfoDir) -and @(Get-ChildItem $oldPcInfoDir -Recurse -File -ErrorAction SilentlyContinue).Count -gt 0) {
        if (-not (Test-Path $newPcInfoDir)) {
            $null = New-Item -ItemType Directory -Path $newPcInfoDir -Force
        }
        Get-ChildItem $oldPcInfoDir -Directory -ErrorAction SilentlyContinue | ForEach-Object {
            $dest = Join-Path $newPcInfoDir $_.Name
            if (-not (Test-Path $dest)) {
                Move-Item $_.FullName $dest -Force
            }
        }
        Show-Info "Migrated pc_information to evidence/pc_information/"
    }
}

function Write-ExecutionHistory {
    param(
        [string]$ModuleName,
        [string]$Category,
        [string]$Status,
        [string]$Message = ""
    )

    $maxRetry = 3
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $windowsUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name

    # Session info (worker name, media serial)
    $workerName = ""
    $mediaSerial = ""
    if ($null -ne $script:SessionInfo) {
        $workerName = $script:SessionInfo.WorkerName
        $mediaSerial = $script:SessionInfo.MediaSerial
    }

    # CSV Escape (if containing comma or newlines)
    $escapedMessage = $Message -replace '"', '""'
    if ($escapedMessage -match '[,\r\n]') {
        $escapedMessage = "`"$escapedMessage`""
    }

    $line = "$timestamp,$env:SELECTED_KANRI_NO,$env:SELECTED_NEW_PCNAME,$ModuleName,$Category,$Status,$escapedMessage,$windowsUser,$workerName,$mediaSerial,$($script:SessionID)"

    # Create with header if file does not exist
    $needHeader = -not (Test-Path $script:HistoryPath)

    for ($i = 0; $i -lt $maxRetry; $i++) {
        try {
            if ($needHeader) {
                $header = "Timestamp,KanriNo,PCName,ModuleName,Category,Status,Message,WindowsUser,Worker,MediaSerial,SessionID"
                $header | Out-File -FilePath $script:HistoryPath -Encoding Default -Force
                $needHeader = $false
            }
            $line | Out-File -FilePath $script:HistoryPath -Append -Encoding Default
            return $true
        }
        catch {
            Start-Sleep -Milliseconds 100
        }
    }

    # Warning only if write fails (continue process)
    return $false
}

function Import-ExecutionHistory {
    param(
        [string]$FilterKanriNo = $null,
        [int]$Limit = 0
    )

    if (-not (Test-Path $script:HistoryPath)) {
        return @()
    }

    try {
        # ロックフリー読み取り（リトライ付き）
        $data = $null
        for ($retry = 0; $retry -lt 3; $retry++) {
            try {
                $data = @(Import-Csv -Path $script:HistoryPath -Encoding Default)
                break
            }
            catch {
                if ($retry -lt 2) { Start-Sleep -Milliseconds 100 }
                else { throw }
            }
        }

        # Filtering
        if ($FilterKanriNo) {
            $data = @($data | Where-Object { $_.KanriNo -eq $FilterKanriNo })
        }

        # Sort descending
        $data = @($data | Sort-Object Timestamp -Descending)

        # Limit count
        if ($Limit -gt 0 -and $data.Count -gt $Limit) {
            $data = @($data | Select-Object -First $Limit)
        }

        return $data
    }
    catch {
        # Try restore from backup if corrupted
        $backupPath = "$($script:HistoryPath).bak"
        if (Test-Path $backupPath) {
            try {
                Copy-Item $backupPath $script:HistoryPath -Force
                Show-Warning "Restored history file from backup"
                return @(Import-Csv -Path $script:HistoryPath -Encoding Default)
            }
            catch {
                Show-Warning "Failed to restore history file"
            }
        }
        return @()
    }
}

function Restore-ExecutionHistory {
    if (-not $env:SELECTED_KANRI_NO) { return }

    try {
        [array]$history = @(Import-ExecutionHistory -FilterKanriNo $env:SELECTED_KANRI_NO -Limit 50)

        if ($history.Count -eq 0) {
            Show-Info "No previous execution history for AdminID $env:SELECTED_KANRI_NO"
            return
        }

        # Import-ExecutionHistory は降順 → 昇順に反転
        [array]::Reverse($history)

        $restoredResults = @()
        foreach ($entry in $history) {
            if ([string]::IsNullOrEmpty($entry.ModuleName)) { continue }

            $ts = $null
            try { $ts = [datetime]::ParseExact($entry.Timestamp, "yyyy-MM-dd HH:mm:ss", $null) }
            catch { $ts = Get-Date }

            $restoredResults += [PSCustomObject]@{
                Operation  = $entry.ModuleName
                Status     = $entry.Status
                Message    = $entry.Message
                Timestamp  = $ts
                IsRestored = $true
                SessionID  = $entry.SessionID
            }
        }

        if ($restoredResults.Count -gt 0) {
            # 復元データと現在セッションの境界セパレーター
            $restoredResults += [PSCustomObject]@{
                Operation  = "--- Current Session ---"
                Status     = "Separator"
                Message    = ""
                Timestamp  = Get-Date
                IsRestored = $false
                SessionID  = $script:SessionID
            }

            $script:ExecutionResults = $restoredResults
            Show-Success "Restored $($restoredResults.Count - 1) history entries"
        }
    }
    catch {
        Show-Warning "Failed to restore execution history: $($_.Exception.Message)"
    }
}

function Show-ExecutionHistory {
    param(
        [switch]$CurrentHostOnly,
        [int]$Limit = 20
    )

    Write-Host ""
    Show-Separator
    Write-Host "Execution History" -ForegroundColor Cyan
    Show-Separator

    $filterKanriNo = $null
    if ($CurrentHostOnly -and $env:SELECTED_KANRI_NO) {
        $filterKanriNo = $env:SELECTED_KANRI_NO
        Write-Host "Target: $env:SELECTED_NEW_PCNAME (AdminID: $filterKanriNo)" -ForegroundColor White
    }
    else {
        Write-Host "Target: All Hosts" -ForegroundColor White
    }
    Show-Separator

    $history = Import-ExecutionHistory -FilterKanriNo $filterKanriNo -Limit $Limit

    if ($history.Count -eq 0) {
        Write-Host "  No history found" -ForegroundColor DarkGray
    }
    else {
        foreach ($entry in $history) {
            $statusColor = switch ($entry.Status) {
                "Success"   { "Green" }
                "Error"     { "Red" }
                "Skip"      { "DarkGray" }
                "Skipped"   { "DarkGray" }
                "Cancelled" { "Yellow" }
                "Partial"   { "Yellow" }
                default     { "White" }
            }

            $pcInfo = if ($CurrentHostOnly) { "" } else { "[$($entry.PCName)] " }
            $msgInfo = if ($entry.Message) { " - $($entry.Message)" } else { "" }

            Write-Host "  $($entry.Timestamp) ${pcInfo}$($entry.ModuleName) [$($entry.Status)]$msgInfo" -ForegroundColor $statusColor
        }
    }

    Show-Separator
    Write-Host "  Displayed: $($history.Count) items (Max $Limit items)" -ForegroundColor DarkGray
    Show-Separator
}

function Export-ExecutionHistory {
    $exportDir = Split-Path $script:HistoryPath -Parent
    if (-not (Test-Path $exportDir)) {
        $null = New-Item -ItemType Directory -Path $exportDir -Force
    }

    $dateStr = Get-Date -Format 'yyyyMMdd_HHmmss'
    $exportPath = Join-Path $exportDir "history_export_$dateStr.csv"

    if (-not (Test-Path $script:HistoryPath)) {
        Show-Warning "No history to export"
        return $null
    }

    Copy-Item $script:HistoryPath $exportPath -Force
    Show-Success "History exported: $exportPath"

    # Copy to evidence/export_history/ with PC name in filename
    $pcName = if (-not [string]::IsNullOrEmpty($env:SELECTED_NEW_PCNAME)) {
        $env:SELECTED_NEW_PCNAME
    } else {
        $env:COMPUTERNAME
    }

    $evidenceExportDir = ".\evidence\export_history"
    if (-not (Test-Path $evidenceExportDir)) {
        $null = New-Item -ItemType Directory -Path $evidenceExportDir -Force
    }

    $evidenceDateStr = Get-Date -Format "yyyy_MM_dd_HHmmss"
    $uid = if ($global:FabriqUniqueId) { $global:FabriqUniqueId } else { Get-HardwareUniqueId }
    $evidenceExportPath = Join-Path $evidenceExportDir "history_export_${evidenceDateStr}_${uid}_${pcName}.csv"
    try {
        Copy-Item $script:HistoryPath $evidenceExportPath -Force
        Show-Success "Evidence copy:    $evidenceExportPath"
    }
    catch {
        Show-Warning "Failed to copy to evidence: $_"
    }

    return $exportPath
}

function Export-HtmlChecklist {
    param(
        [string]$ProfileName,
        [string]$ProfilePath,
        [array] $DefinedModules,
        [array] $ExecutionResults,
        [System.TimeSpan]$ElapsedTime = [System.TimeSpan]::Zero
    )

    # Load System.Web for HtmlEncode (not loaded by default in PS5.1)
    Add-Type -AssemblyName System.Web -ErrorAction SilentlyContinue

    # ----------------------------------------
    # Output path (same convention as history export)
    # ----------------------------------------
    $outputDir = ".\evidence\checklist"
    if (-not (Test-Path $outputDir)) {
        $null = New-Item -ItemType Directory -Path $outputDir -Force
    }

    $dateStr = Get-Date -Format "yyyy_MM_dd_HHmmss"
    $uid     = if ($global:FabriqUniqueId) { $global:FabriqUniqueId } else { Get-HardwareUniqueId }
    $pcName  = if (-not [string]::IsNullOrEmpty($env:SELECTED_NEW_PCNAME)) { $env:SELECTED_NEW_PCNAME } else { $env:COMPUTERNAME }
    $outPath = Join-Path $outputDir "checklist_${dateStr}_${uid}_${pcName}.html"

    # ----------------------------------------
    # Session metadata
    # ----------------------------------------
    $workerName  = if ($script:SessionInfo) { $script:SessionInfo.WorkerName }  else { "-" }
    $mediaSerial = if ($script:SessionInfo) { $script:SessionInfo.MediaSerial } else { "-" }
    $kanriNo     = if ($env:SELECTED_KANRI_NO)    { $env:SELECTED_KANRI_NO }    else { "-" }
    $oldPcName   = if ($env:SELECTED_OLD_PCNAME)  { $env:SELECTED_OLD_PCNAME }  else { "-" }
    $ethernetIp  = if ($env:SELECTED_ETH_IP)      { $env:SELECTED_ETH_IP }      else { "-" }
    $generatedAt = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $elapsedStr  = "{0:D2}:{1:D2}:{2:D2}" -f [int]$ElapsedTime.TotalHours, $ElapsedTime.Minutes, $ElapsedTime.Seconds

    # ----------------------------------------
    # System info: Printers (from env vars)
    # ----------------------------------------
    $printerList = @()
    for ($i = 1; $i -le 10; $i++) {
        $pName = [Environment]::GetEnvironmentVariable("SELECTED_PRINTER_$($i)_NAME")
        if (-not [string]::IsNullOrEmpty($pName)) {
            $pDriver = [Environment]::GetEnvironmentVariable("SELECTED_PRINTER_$($i)_DRIVER")
            $pPort   = [Environment]::GetEnvironmentVariable("SELECTED_PRINTER_$($i)_PORT")
            $printerList += [PSCustomObject]@{
                Name   = $pName
                Driver = if ($pDriver) { $pDriver } else { "-" }
                Port   = if ($pPort)   { $pPort }   else { "-" }
            }
        }
    }

    # ----------------------------------------
    # System info: Windows License status (WMI)
    # ----------------------------------------
    $licenseStatus  = "N/A"
    $licenseClass   = "notrun"
    $licenseProduct = ""
    try {
        $slp = @(Get-WmiObject SoftwareLicensingProduct -ErrorAction SilentlyContinue |
                 Where-Object { $_.PartialProductKey -and $_.Name -match "Windows" }) |
               Select-Object -First 1
        if ($slp) {
            $licMap = @{ 0="Unlicensed"; 1="Licensed"; 2="OOB Grace"; 3="OOT Grace"; 4="Non-Genuine Grace"; 5="Notification"; 6="Extended Grace" }
            $licenseStatus  = if ($licMap.ContainsKey([int]$slp.LicenseStatus)) { $licMap[[int]$slp.LicenseStatus] } else { "Unknown ($($slp.LicenseStatus))" }
            $licenseClass   = switch ([int]$slp.LicenseStatus) { 1 { "ok" } 0 { "ng" } default { "partial" } }
            $licenseProduct = if ($slp.Name) { $slp.Name } else { "" }
        }
    }
    catch { }

    # ----------------------------------------
    # System info: BitLocker status (C:)
    # ----------------------------------------
    $blProtection = "N/A"
    $blVolume     = "N/A"
    $blClass      = "notrun"
    try {
        $blv = Get-BitLockerVolume -MountPoint "C:" -ErrorAction SilentlyContinue
        if ($blv) {
            $blProtection = "$($blv.ProtectionStatus)"
            $blVolume     = "$($blv.VolumeStatus)"
            $blClass      = switch ("$($blv.ProtectionStatus)") {
                "On"    { "ok" }
                "Off"   { "partial" }
                default { "notrun" }
            }
        }
    }
    catch { }

    # ----------------------------------------
    # Read profile CSV Description column (supplemental)
    # ----------------------------------------
    $descriptionMap = @{}
    if (-not [string]::IsNullOrEmpty($ProfilePath) -and (Test-Path $ProfilePath)) {
        try {
            $profileRows = @(Import-Csv $ProfilePath -Encoding Default)
            foreach ($row in $profileRows) {
                if ($row.ScriptPath -and $row.Description) {
                    $descriptionMap[$row.ScriptPath.Trim()] = $row.Description
                }
            }
        }
        catch { }
    }

    # ----------------------------------------
    # Build checklist rows (profile definition vs actual result)
    # ----------------------------------------
    # Include IsRestored entries: pre-restart results are restored from CSV with IsRestored=true.
    # Select-Object -Last 1 ensures the most recent result wins when a module name appears
    # multiple times (across restarts or repeated sessions for the same KanriNo).
    $currentResults = @($ExecutionResults | Where-Object { $_.Status -ne "Separator" })

    $successTotal   = 0
    $skipTotal      = 0
    $errorTotal     = 0
    $notRunTotal    = 0

    $rowsHtml = ""
    foreach ($module in $DefinedModules) {
        # Match by MenuName (last occurrence wins for duplicated names)
        $result = $currentResults | Where-Object { $_.Operation -eq $module.MenuName } | Select-Object -Last 1

        $statusLabel = "Not Run"
        $statusClass = "notrun"
        $message     = "-"

        if ($null -ne $result) {
            switch ($result.Status) {
                "Success"   { $statusLabel = "OK";      $statusClass = "ok";      $successTotal++ }
                "Partial"   { $statusLabel = "Partial"; $statusClass = "partial"; $successTotal++ }
                "Skipped"   { $statusLabel = "Skip";    $statusClass = "skip";    $skipTotal++ }
                "Skip"      { $statusLabel = "Skip";    $statusClass = "skip";    $skipTotal++ }
                "Cancelled" { $statusLabel = "Cancel";  $statusClass = "skip";    $skipTotal++ }
                "Warning"   { $statusLabel = "Warn";    $statusClass = "partial"; $successTotal++ }
                "Error"     { $statusLabel = "NG";      $statusClass = "ng";      $errorTotal++ }
                default     { $statusLabel = $result.Status; $statusClass = "notrun" }
            }
            $message = if ($result.Message) { [System.Web.HttpUtility]::HtmlEncode($result.Message) } else { "-" }
            $ts      = if ($result.Timestamp) { $result.Timestamp.ToString("HH:mm:ss") } else { "-" }
        }
        else {
            $notRunTotal++
            $ts = "-"
        }

        # Description from profile CSV, fallback to MenuName
        $relPath = if ($module.RelativePath) { $module.RelativePath } else { "" }
        $desc    = if ($descriptionMap.ContainsKey($relPath)) { [System.Web.HttpUtility]::HtmlEncode($descriptionMap[$relPath]) } else { "" }

        # Marker row (RESTART / REEXPLORER etc.) - lighter styling
        $isMarker   = $module.MenuName -match '^\[.+\]$'
        $rowClass   = if ($isMarker) { ' class="marker-row"' } else { "" }

        $rowsHtml += @"
        <tr$rowClass>
            <td class="col-order">$($module.Order)</td>
            <td class="col-name">$([System.Web.HttpUtility]::HtmlEncode($module.MenuName))$(if($desc){"<br><span class='desc'>$desc</span>"})</td>
            <td class="col-cat">$([System.Web.HttpUtility]::HtmlEncode($module.Category))</td>
            <td class="col-status"><span class="badge $statusClass">$statusLabel</span></td>
            <td class="col-time">$ts</td>
            <td class="col-msg">$message</td>
        </tr>
"@
    }

    $overallClass = if ($errorTotal -gt 0) { "overall-ng" } elseif ($notRunTotal -gt 0) { "overall-partial" } else { "overall-ok" }
    $overallLabel = if ($errorTotal -gt 0) { "NG" } elseif ($notRunTotal -gt 0) { "Incomplete" } else { "OK" }

    # ----------------------------------------
    # Build supplemental section HTML
    # ----------------------------------------
    $licProductRow = if ($licenseProduct) {
        "<div class='sysinfo-row'><span class='sysinfo-label'>Product</span><span style='font-size:11px;color:#555;word-break:break-all;'>$([System.Web.HttpUtility]::HtmlEncode($licenseProduct))</span></div>"
    } else { "" }

    $printerRowsHtml = ""
    if ($printerList.Count -gt 0) {
        foreach ($p in $printerList) {
            $pn = [System.Web.HttpUtility]::HtmlEncode($p.Name)
            $pd = [System.Web.HttpUtility]::HtmlEncode($p.Driver)
            $pp = [System.Web.HttpUtility]::HtmlEncode($p.Port)
            $printerRowsHtml += "        <tr><td>$pn</td><td>$pd</td><td>$pp</td></tr>`n"
        }
        $printerSectionHtml = @"
  <div class="section">
    <div class="section-hd">Configured Printers ($($printerList.Count))</div>
    <table class="printer-table">
      <thead><tr><th>Name</th><th>Driver</th><th>Port</th></tr></thead>
      <tbody>
$printerRowsHtml      </tbody>
    </table>
  </div>
"@
    } else {
        $printerSectionHtml = ""
    }

    # ----------------------------------------
    # HTML document
    # ----------------------------------------
    $html = @"
<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<title>Fabriq Checklist - $([System.Web.HttpUtility]::HtmlEncode($pcName))</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Segoe UI', Meiryo, sans-serif; font-size: 13px; background: #f5f5f5; color: #222; }
  .page { max-width: 1100px; margin: 24px auto; padding: 0 16px 40px; }

  /* Header */
  .header { background: #1a1a2e; color: #fff; padding: 20px 24px; border-radius: 6px 6px 0 0; display: flex; justify-content: space-between; align-items: flex-start; }
  .header h1 { font-size: 20px; font-weight: 600; letter-spacing: 0.05em; }
  .header .subtitle { font-size: 11px; color: #aaa; margin-top: 4px; }
  .overall-badge { font-size: 22px; font-weight: 700; padding: 4px 18px; border-radius: 4px; }
  .overall-ok      { background: #27ae60; color: #fff; }
  .overall-ng      { background: #c0392b; color: #fff; }
  .overall-partial { background: #e67e22; color: #fff; }

  /* Meta cards */
  .meta-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(200px, 1fr)); gap: 1px; background: #ddd; border: 1px solid #ddd; }
  .meta-card { background: #fff; padding: 10px 14px; }
  .meta-card .label { font-size: 10px; color: #888; text-transform: uppercase; letter-spacing: 0.06em; margin-bottom: 3px; }
  .meta-card .value { font-size: 13px; font-weight: 600; word-break: break-all; }

  /* Summary bar */
  .summary-bar { display: flex; gap: 12px; padding: 10px 14px; background: #fff; border: 1px solid #ddd; border-top: none; align-items: center; }
  .summary-bar .label { font-size: 11px; color: #666; margin-right: 4px; }
  .chip { display: inline-block; padding: 2px 10px; border-radius: 12px; font-size: 12px; font-weight: 600; }
  .chip-ok      { background: #d4edda; color: #155724; }
  .chip-skip    { background: #e2e3e5; color: #383d41; }
  .chip-ng      { background: #f8d7da; color: #721c24; }
  .chip-notrun  { background: #fff3cd; color: #856404; }

  /* Table */
  .table-wrap { background: #fff; border: 1px solid #ddd; border-top: none; border-radius: 0 0 6px 6px; overflow: hidden; }
  table { width: 100%; border-collapse: collapse; }
  thead tr { background: #2c3e50; color: #fff; }
  thead th { padding: 9px 12px; text-align: left; font-size: 11px; font-weight: 600; letter-spacing: 0.05em; white-space: nowrap; }
  tbody tr { border-bottom: 1px solid #eee; }
  tbody tr:hover { background: #fafafa; }
  tbody tr.marker-row { background: #f8f8f8; }
  tbody tr.marker-row td { color: #888; font-style: italic; }
  td { padding: 8px 12px; vertical-align: middle; }

  .col-order  { width: 52px; text-align: center; color: #999; font-size: 12px; }
  .col-name   { min-width: 200px; }
  .col-cat    { width: 120px; color: #555; font-size: 12px; }
  .col-status { width: 72px; text-align: center; }
  .col-time   { width: 72px; text-align: center; color: #666; font-size: 12px; font-variant-numeric: tabular-nums; }
  .col-msg    { color: #555; font-size: 12px; word-break: break-all; }

  .desc { font-size: 11px; color: #999; font-weight: 400; }

  /* Badges */
  .badge { display: inline-block; padding: 2px 9px; border-radius: 3px; font-size: 11px; font-weight: 700; letter-spacing: 0.04em; }
  .ok      { background: #d4edda; color: #155724; }
  .partial { background: #fff3cd; color: #856404; }
  .skip    { background: #e2e3e5; color: #383d41; }
  .ng      { background: #f8d7da; color: #721c24; }
  .notrun  { background: #fff3cd; color: #856404; }

  /* System Info sections */
  .section { border: 1px solid #ddd; border-radius: 6px; overflow: hidden; margin-top: 20px; }
  .section-hd { background: #2c3e50; color: #fff; padding: 9px 14px; font-size: 11px; font-weight: 600; letter-spacing: 0.05em; text-transform: uppercase; }
  .sysinfo-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(280px, 1fr)); gap: 1px; background: #ddd; }
  .sysinfo-card { background: #fff; padding: 12px 16px; }
  .sysinfo-card-title { font-size: 10px; color: #888; text-transform: uppercase; letter-spacing: 0.06em; margin-bottom: 8px; padding-bottom: 5px; border-bottom: 1px solid #f0f0f0; }
  .sysinfo-row { display: flex; align-items: baseline; gap: 8px; margin-bottom: 5px; }
  .sysinfo-label { font-size: 11px; color: #888; min-width: 80px; flex-shrink: 0; }
  .printer-table { width: 100%; border-collapse: collapse; background: #fff; }
  .printer-table thead tr { background: #f5f5f5; }
  .printer-table th { padding: 7px 14px; text-align: left; font-size: 11px; color: #555; font-weight: 600; border-bottom: 1px solid #e0e0e0; }
  .printer-table td { padding: 7px 14px; font-size: 12px; border-bottom: 1px solid #f0f0f0; }
  .printer-table tr:last-child td { border-bottom: none; }

  .footer { text-align: center; font-size: 11px; color: #aaa; margin-top: 16px; }
</style>
</head>
<body>
<div class="page">

  <!-- Header -->
  <div class="header">
    <div>
      <div class="h1">Fabriq Kitting Checklist</div>
      <div class="subtitle">$([System.Web.HttpUtility]::HtmlEncode($ProfileName))</div>
    </div>
    <div class="overall-badge $overallClass">$overallLabel</div>
  </div>

  <!-- Meta -->
  <div class="meta-grid">
    <div class="meta-card"><div class="label">Target PC (New)</div><div class="value">$([System.Web.HttpUtility]::HtmlEncode($pcName))</div></div>
    <div class="meta-card"><div class="label">Target PC (Old)</div><div class="value">$([System.Web.HttpUtility]::HtmlEncode($oldPcName))</div></div>
    <div class="meta-card"><div class="label">Admin ID (KanriNo)</div><div class="value">$([System.Web.HttpUtility]::HtmlEncode($kanriNo))</div></div>
    <div class="meta-card"><div class="label">Ethernet IP</div><div class="value">$([System.Web.HttpUtility]::HtmlEncode($ethernetIp))</div></div>
    <div class="meta-card"><div class="label">Worker</div><div class="value">$([System.Web.HttpUtility]::HtmlEncode($workerName))</div></div>
    <div class="meta-card"><div class="label">Media Serial</div><div class="value">$([System.Web.HttpUtility]::HtmlEncode($mediaSerial))</div></div>
    <div class="meta-card"><div class="label">Hardware ID</div><div class="value">$([System.Web.HttpUtility]::HtmlEncode($uid))</div></div>
    <div class="meta-card"><div class="label">Generated At</div><div class="value">$generatedAt</div></div>
    <div class="meta-card"><div class="label">Elapsed Time</div><div class="value">$elapsedStr</div></div>
  </div>

  <!-- Summary bar -->
  <div class="summary-bar">
    <span class="label">Summary:</span>
    <span class="chip chip-ok">OK $successTotal</span>
    <span class="chip chip-skip">Skip $skipTotal</span>
    <span class="chip chip-ng">NG $errorTotal</span>
    <span class="chip chip-notrun">Not Run $notRunTotal</span>
    <span style="margin-left:auto; font-size:11px; color:#888;">Total: $($DefinedModules.Count) items</span>
  </div>

  <!-- Checklist table -->
  <div class="table-wrap">
    <table>
      <thead>
        <tr>
          <th class="col-order">#</th>
          <th class="col-name">Module</th>
          <th class="col-cat">Category</th>
          <th class="col-status">Result</th>
          <th class="col-time">Time</th>
          <th class="col-msg">Message</th>
        </tr>
      </thead>
      <tbody>
$rowsHtml      </tbody>
    </table>
  </div>

  <!-- System Status -->
  <div class="section">
    <div class="section-hd">System Status</div>
    <div class="sysinfo-grid">
      <div class="sysinfo-card">
        <div class="sysinfo-card-title">Windows License</div>
        <div class="sysinfo-row">
          <span class="sysinfo-label">Activation</span>
          <span class="badge $licenseClass">$licenseStatus</span>
        </div>
        $licProductRow
      </div>
      <div class="sysinfo-card">
        <div class="sysinfo-card-title">BitLocker (C:)</div>
        <div class="sysinfo-row">
          <span class="sysinfo-label">Protection</span>
          <span class="badge $blClass">$blProtection</span>
        </div>
        <div class="sysinfo-row">
          <span class="sysinfo-label">Volume</span>
          <span style="font-size:12px;">$blVolume</span>
        </div>
      </div>
    </div>
  </div>
$printerSectionHtml
  <div class="footer">Generated by Fabriq ver2.1 &mdash; $generatedAt</div>
</div>
</body>
</html>
"@

    try {
        $html | Out-File -FilePath $outPath -Encoding UTF8 -Force
        Show-Success "Checklist HTML: $outPath"
        return $outPath
    }
    catch {
        Show-Warning "Failed to generate HTML checklist: $_"
        return $null
    }
}

function Clear-AllLogs {
    Write-Host ""
    Show-Separator
    Write-Host "Clear Runtime Logs" -ForegroundColor Red
    Show-Separator
    Write-Host ""

    # Enumerate targets (runtime logs only, NOT evidence)
    $targets = @()

    # 1. Execution History
    if (Test-Path $script:HistoryPath) {
        $targets += [PSCustomObject]@{ Type = "Execution History"; Path = $script:HistoryPath; IsDir = $false }
    }
    $bakPath = "$($script:HistoryPath).bak"
    if (Test-Path $bakPath) {
        $targets += [PSCustomObject]@{ Type = "History Backup"; Path = $bakPath; IsDir = $false }
    }

    # 2. Exported History (logs/history/history_export_*.csv)
    $historyDir = Split-Path $script:HistoryPath -Parent
    if (Test-Path $historyDir) {
        $exports = @(Get-ChildItem $historyDir -Filter "history_export_*.csv" -File -ErrorAction SilentlyContinue)
        foreach ($f in $exports) {
            $targets += [PSCustomObject]@{ Type = "Exported History"; Path = $f.FullName; IsDir = $false }
        }
    }

    # 3. Transcript Logs (logs/*.log)
    if (Test-Path ".\logs") {
        $transcripts = @(Get-ChildItem ".\logs" -Filter "*.log" -File -ErrorAction SilentlyContinue)
        foreach ($f in $transcripts) {
            $targets += [PSCustomObject]@{ Type = "Transcript Log"; Path = $f.FullName; IsDir = $false }
        }
    }

    # 4. Status JSON
    if (Test-Path $script:StatusFilePath) {
        $targets += [PSCustomObject]@{ Type = "Status File"; Path = $script:StatusFilePath; IsDir = $false }
    }
    $statusTmp = "$($script:StatusFilePath).tmp"
    if (Test-Path $statusTmp) {
        $targets += [PSCustomObject]@{ Type = "Status Temp"; Path = $statusTmp; IsDir = $false }
    }

    # 5. Session JSON
    if (Test-Path $script:SessionFilePath) {
        $targets += [PSCustomObject]@{ Type = "Session File"; Path = $script:SessionFilePath; IsDir = $false }
    }

    if ($targets.Count -eq 0) {
        Show-Info "No log files to clear"
        return
    }

    # Show targets
    Write-Host "The following runtime logs will be deleted:" -ForegroundColor Yellow
    Write-Host "(Evidence data in evidence/ is NOT affected)" -ForegroundColor Gray
    Write-Host ""
    foreach ($t in $targets) {
        if ($t.IsDir) {
            $count = @(Get-ChildItem $t.Path -Recurse -File -ErrorAction SilentlyContinue).Count
            Write-Host "  [$($t.Type)] $($t.Path) ($count files)" -ForegroundColor White
        }
        else {
            Write-Host "  [$($t.Type)] $($t.Path)" -ForegroundColor White
        }
    }
    Write-Host ""

    # Confirm
    if (-not (Confirm-Execution -Message "Are you sure you want to clear runtime logs?")) {
        Show-Info "Canceled"
        return
    }

    # Delete
    $deleted = 0
    $skipped = 0
    $failed = 0
    foreach ($t in $targets) {
        try {
            if ($t.IsDir) {
                $dirFiles = @(Get-ChildItem $t.Path -Recurse -File)
                foreach ($df in $dirFiles) {
                    try {
                        Remove-Item $df.FullName -Force -ErrorAction Stop
                    }
                    catch {
                        Write-Host "  [SKIP] In use: $($df.Name)" -ForegroundColor DarkGray
                        $skipped++
                    }
                }
                Get-ChildItem $t.Path -Directory -Recurse | Sort-Object FullName -Descending | Remove-Item -Force -ErrorAction SilentlyContinue
                $deleted++
            }
            else {
                Remove-Item $t.Path -Force -ErrorAction Stop
                $deleted++
            }
        }
        catch {
            if ($_.Exception.Message -match 'used by another process|別のプロセス') {
                Write-Host "  [SKIP] In use: $($t.Path)" -ForegroundColor DarkGray
                $skipped++
            }
            else {
                Show-Warning "Failed to delete: $($t.Path)"
                $failed++
            }
        }
    }

    Write-Host ""
    Show-Success "Cleared: $deleted items"
    if ($skipped -gt 0) {
        Show-Info "Skipped: $skipped items (in use by current session)"
    }
    if ($failed -gt 0) {
        Show-Warning "Failed: $failed items"
    }
}

function Clear-Evidence {
    Write-Host ""
    Show-Separator
    Write-Host "Clear Evidence Data" -ForegroundColor Red
    Show-Separator
    Write-Host ""

    $targets = @()

    # 1. Gyotaku Screenshots (evidence/gyotaku/)
    $gyotakuDir = ".\evidence\gyotaku"
    if ((Test-Path $gyotakuDir) -and @(Get-ChildItem $gyotakuDir -Recurse -File -ErrorAction SilentlyContinue).Count -gt 0) {
        $targets += [PSCustomObject]@{ Type = "Gyotaku Screenshots"; Path = $gyotakuDir; IsDir = $true }
    }

    # 2. PC Information (evidence/pc_information/)
    $pcInfoDir = ".\evidence\pc_information"
    if ((Test-Path $pcInfoDir) -and @(Get-ChildItem $pcInfoDir -Recurse -File -ErrorAction SilentlyContinue).Count -gt 0) {
        $targets += [PSCustomObject]@{ Type = "PC Information"; Path = $pcInfoDir; IsDir = $true }
    }

    # 3. Exported History (evidence/export_history/)
    $exportDir = ".\evidence\export_history"
    if ((Test-Path $exportDir) -and @(Get-ChildItem $exportDir -Recurse -File -ErrorAction SilentlyContinue).Count -gt 0) {
        $targets += [PSCustomObject]@{ Type = "Exported History"; Path = $exportDir; IsDir = $true }
    }

    if ($targets.Count -eq 0) {
        Show-Info "No evidence files to clear"
        return
    }

    # Show targets
    Write-Host "The following evidence data will be deleted:" -ForegroundColor Yellow
    Write-Host ""
    foreach ($t in $targets) {
        $count = @(Get-ChildItem $t.Path -Recurse -File -ErrorAction SilentlyContinue).Count
        Write-Host "  [$($t.Type)] $($t.Path) ($count files)" -ForegroundColor White
    }
    Write-Host ""

    # Confirm
    if (-not (Confirm-Execution -Message "Are you sure you want to clear ALL evidence?")) {
        Show-Info "Canceled"
        return
    }

    # Delete
    $deleted = 0
    $skipped = 0
    $failed = 0
    foreach ($t in $targets) {
        try {
            $dirFiles = @(Get-ChildItem $t.Path -Recurse -File)
            foreach ($df in $dirFiles) {
                try {
                    Remove-Item $df.FullName -Force -ErrorAction Stop
                }
                catch {
                    Write-Host "  [SKIP] In use: $($df.Name)" -ForegroundColor DarkGray
                    $skipped++
                }
            }
            Get-ChildItem $t.Path -Directory -Recurse | Sort-Object FullName -Descending | Remove-Item -Force -ErrorAction SilentlyContinue
            $deleted++
        }
        catch {
            if ($_.Exception.Message -match 'used by another process|別のプロセス') {
                Write-Host "  [SKIP] In use: $($t.Path)" -ForegroundColor DarkGray
                $skipped++
            }
            else {
                Show-Warning "Failed to delete: $($t.Path)"
                $failed++
            }
        }
    }

    Write-Host ""
    Show-Success "Cleared: $deleted items"
    if ($skipped -gt 0) {
        Show-Info "Skipped: $skipped items (in use)"
    }
    if ($failed -gt 0) {
        Show-Warning "Failed: $failed items"
    }
}

# ========================================
# Resume State Functions (Profile Restart)
# ========================================

function Save-ResumeState {
    param(
        [string]$ProfilePath,
        [string]$ProfileName,
        [bool]$StopOnError,
        [int]$ResumeAfterOrder,
        [array]$CompletedModules
    )

    # Snapshot all host environment variables
    $hostEnv = @{}
    $envNames = @(
        "SELECTED_KANRI_NO", "SELECTED_OLD_PCNAME", "SELECTED_NEW_PCNAME",
        "SELECTED_ETH_IP", "SELECTED_ETH_SUBNET", "SELECTED_ETH_GATEWAY",
        "SELECTED_WIFI_IP", "SELECTED_WIFI_SUBNET", "SELECTED_WIFI_GATEWAY",
        "SELECTED_DNS1", "SELECTED_DNS2", "SELECTED_DNS3", "SELECTED_DNS4"
    )
    foreach ($name in $envNames) {
        $hostEnv[$name] = [Environment]::GetEnvironmentVariable($name)
    }
    for ($i = 1; $i -le 10; $i++) {
        foreach ($suffix in @("NAME", "DRIVER", "PORT")) {
            $key = "SELECTED_PRINTER_$($i)_$suffix"
            $hostEnv[$key] = [Environment]::GetEnvironmentVariable($key)
        }
    }

    $state = @{
        ProfilePath      = $ProfilePath
        ProfileName      = $ProfileName
        StopOnError      = $StopOnError
        AutoPilot        = $global:AutoPilotMode
        AutoPilotWaitSec = $global:AutoPilotWaitSec
        SessionID        = $script:SessionID
        ResumeAfterOrder = $ResumeAfterOrder
        CompletedModules = @($CompletedModules | ForEach-Object {
            @{ MenuName = $_.MenuName; Status = $_.Status }
        })
        HostEnvironment  = $hostEnv
    }

    $state | ConvertTo-Json -Depth 5 | Out-File -FilePath $script:ResumeStatePath -Encoding UTF8 -Force
}

function Load-ResumeState {
    if (-not (Test-Path $script:ResumeStatePath)) { return $null }
    try {
        $json = Get-Content $script:ResumeStatePath -Raw -Encoding UTF8
        return ($json | ConvertFrom-Json)
    }
    catch { return $null }
}

function Remove-ResumeState {
    if (Test-Path $script:ResumeStatePath) {
        Remove-Item $script:ResumeStatePath -Force -ErrorAction SilentlyContinue
    }
}

function Reset-FabriqState {
    # ========================================
    # Resets all in-memory session state so that a new kitting
    # session can begin on the same Fabriq process instance.
    # Evidence files on disk are NOT deleted.
    # ========================================

    Show-Info "Resetting Fabriq session state..."
    Write-Host ""

    # ----------------------------------------
    # 1. Transcript: stop current → start new
    # ----------------------------------------
    try { Stop-Transcript -ErrorAction SilentlyContinue | Out-Null } catch { }

    $logDir = ".\logs"
    if (-not (Test-Path $logDir)) { $null = New-Item -ItemType Directory -Path $logDir -Force }
    $newTs  = Get-Date -Format "yyyy_MM_dd_HHmmss"
    $uid    = $global:FabriqUniqueId    # hardware ID unchanged between sessions
    $hn     = $env:COMPUTERNAME
    $newLog = Join-Path $logDir "${newTs}_${uid}_${hn}.log"
    $global:FabriqTranscriptPath   = $newLog
    $global:FabriqSessionTimestamp = $newTs
    Start-Transcript -Path $newLog -Append | Out-Null

    # ----------------------------------------
    # 2. Execution Results & Session ID
    # ----------------------------------------
    $script:ExecutionResults = @()
    $script:SessionID        = Get-Date -Format "yyyyMMdd_HHmmss"

    # ----------------------------------------
    # 2b. Execution History CSV
    # Evidence export already ran at profile completion, so the CSV is no longer
    # needed for audit purposes. Delete it so Restore-ExecutionHistory finds nothing
    # and the status monitor shows a clean state on the next session.
    # ----------------------------------------
    if (Test-Path $script:HistoryPath) {
        Remove-Item $script:HistoryPath -Force -ErrorAction SilentlyContinue
    }
    $historyBak = "$($script:HistoryPath).bak"
    if (Test-Path $historyBak) {
        Remove-Item $historyBak -Force -ErrorAction SilentlyContinue
    }

    # ----------------------------------------
    # 3. Session Info + session.json (force worker re-selection)
    # ----------------------------------------
    $script:SessionInfo = $null
    if (Test-Path $script:SessionFilePath) {
        Remove-Item $script:SessionFilePath -Force -ErrorAction SilentlyContinue
    }

    # ----------------------------------------
    # 4. Global Flags
    # ----------------------------------------
    $global:AutoPilotMode     = $false
    $global:AutoPilotWaitSec  = 3
    $global:_LastModuleResult = $null

    # ----------------------------------------
    # 5. Environment Variables (selected host)
    # ----------------------------------------
    $envKeys = @(
        "SELECTED_KANRI_NO", "SELECTED_OLD_PCNAME", "SELECTED_NEW_PCNAME",
        "SELECTED_ETH_IP", "SELECTED_ETH_SUBNET", "SELECTED_ETH_GATEWAY",
        "SELECTED_WIFI_IP", "SELECTED_WIFI_SUBNET", "SELECTED_WIFI_GATEWAY",
        "SELECTED_DNS1", "SELECTED_DNS2", "SELECTED_DNS3", "SELECTED_DNS4",
        "FABRIQ_AUTOLOGON_NO"
    )
    foreach ($key in $envKeys) {
        [Environment]::SetEnvironmentVariable($key, $null, "Process")
    }
    for ($i = 1; $i -le 10; $i++) {
        foreach ($suffix in @("NAME", "DRIVER", "PORT")) {
            [Environment]::SetEnvironmentVariable("SELECTED_PRINTER_${i}_${suffix}", $null, "Process")
        }
    }

    # ----------------------------------------
    # 6. Resume State + Status File
    # ----------------------------------------
    Remove-ResumeState
    Write-StatusFile -Phase "idle"

    Show-Success "Session state reset. New log: $newLog"
    Write-Host ""
}

function Restore-HostEnvironment {
    param([object]$HostEnv)
    $HostEnv.PSObject.Properties | ForEach-Object {
        Set-Item -Path "env:$($_.Name)" -Value $_.Value -ErrorAction SilentlyContinue
    }
}

# ========================================
# Session Management Functions
# ========================================

function Get-VolumeSerial {
    param([string]$DriveLetter)
    try {
        $drive = $DriveLetter.TrimEnd(":\") + ":"
        $vol = Get-WmiObject -Class Win32_LogicalDisk -Filter "DeviceID='$drive'" -ErrorAction Stop
        if ($vol.VolumeSerialNumber) {
            return $vol.VolumeSerialNumber
        }
    }
    catch { }
    return "UNKNOWN"
}

function Initialize-Session {
    # Priority 1: Existing session.json (survives restart)
    if (Test-Path $script:SessionFilePath) {
        try {
            $json = Get-Content $script:SessionFilePath -Raw -Encoding UTF8
            $script:SessionInfo = $json | ConvertFrom-Json
            Show-Success "Session loaded: Worker=$($script:SessionInfo.WorkerName), MediaSerial=$($script:SessionInfo.MediaSerial)"
            return
        }
        catch {
            Show-Warning "Failed to load session.json, re-initializing..."
        }
    }

    # --- Determine Media Serial ---
    $mediaSerial = ""

    # Priority 2: source_media.id (created by Deploy.bat)
    if (Test-Path $script:SourceMediaIdPath) {
        try {
            $mediaSerial = (Get-Content $script:SourceMediaIdPath -Raw -ErrorAction Stop).Trim()
        }
        catch { }
    }

    # Priority 3: Current drive volume serial
    if ([string]::IsNullOrWhiteSpace($mediaSerial)) {
        $currentDrive = (Resolve-Path ".").Drive.Name + ":"
        $mediaSerial = Get-VolumeSerial -DriveLetter $currentDrive
    }

    # --- Determine Worker ---
    $workerID = ""
    $workerName = ""

    # Try loading workers.csv for selection
    if (Test-Path $script:WorkersCsvPath) {
        try {
            $workers = @(Import-Csv -Path $script:WorkersCsvPath -Encoding Default)
            if ($workers.Count -gt 0) {
                Write-Host ""
                Show-Separator
                Write-Host "Worker Selection" -ForegroundColor Magenta
                Show-Separator
                Write-Host ""

                for ($i = 0; $i -lt $workers.Count; $i++) {
                    Write-Host "  [$($i + 1)] $($workers[$i].ID) - $($workers[$i].Name)" -ForegroundColor White
                }
                Write-Host ""
                Write-Host "  [0] Manual input" -ForegroundColor Yellow
                Show-Separator
                Write-Host ""

                while ($true) {
                    Write-Host -NoNewline "Select worker: "
                    $wChoice = Read-Host

                    if ($wChoice -eq '0') {
                        Write-Host -NoNewline "Worker name: "
                        $workerName = Read-Host
                        $workerID = "MANUAL"
                        break
                    }

                    $wNum = 0
                    if ([int]::TryParse($wChoice, [ref]$wNum) -and $wNum -ge 1 -and $wNum -le $workers.Count) {
                        $selected = $workers[$wNum - 1]
                        $workerID = $selected.ID
                        $workerName = $selected.Name
                        break
                    }

                    Show-Error "Invalid selection"
                }
            }
        }
        catch {
            Show-Warning "Failed to load workers.csv: $_"
        }
    }

    # Fallback: manual input if no worker selected
    if ([string]::IsNullOrWhiteSpace($workerName)) {
        Write-Host ""
        Write-Host -NoNewline "Worker name: "
        $workerName = Read-Host
        if ([string]::IsNullOrWhiteSpace($workerName)) {
            $workerName = $env:USERNAME
        }
        $workerID = "MANUAL"
    }

    # Build session object
    $script:SessionInfo = [PSCustomObject]@{
        WorkerID     = $workerID
        WorkerName   = $workerName
        MediaSerial  = $mediaSerial
        StartTime    = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        WindowsUser  = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
        ComputerName = $env:COMPUTERNAME
    }

    # Save to session.json
    try {
        $script:SessionInfo | ConvertTo-Json -Depth 3 | Out-File -FilePath $script:SessionFilePath -Encoding UTF8 -Force
    }
    catch {
        Show-Warning "Failed to save session.json: $_"
    }

    Show-Success "Session initialized: Worker=$workerName, MediaSerial=$mediaSerial"
}

# ========================================
# Profile Functions
# ========================================

function Create-DefaultProfiles {
    param([array]$AllModules)

    if (-not (Test-Path $script:ProfilesDir)) {
        $null = New-Item -Path $script:ProfilesDir -ItemType Directory -Force
    }

    # Basic Setup
    $basicPath = Join-Path $script:ProfilesDir "Basic Setup.csv"
    if (-not (Test-Path $basicPath)) {
        $content = @(
            "Order,ScriptPath,Enabled,Description"
            "10,standard\hostname_config\hostname_config.ps1,1,Change Hostname"
            "20,standard\ipaddress_config\ipaddress_config.ps1,1,IP Address Settings"
            "30,standard\domain_join\domain_join.ps1,1,Domain Join"
        )
        $content -join "`r`n" | Out-File $basicPath -Encoding Default -Force
    }

    # Full Setup (all modules)
    $fullPath = Join-Path $script:ProfilesDir "Full Setup.csv"
    if (-not (Test-Path $fullPath) -and $AllModules.Count -gt 0) {
        $lines = @("Order,ScriptPath,Enabled,Description")
        $order = 10
        foreach ($m in $AllModules) {
            $lines += "$order,$($m.RelativePath),1,$($m.MenuName)"
            $order += 10
        }
        $lines -join "`r`n" | Out-File $fullPath -Encoding Default -Force
    }
}

function Load-Profiles {
    param([array]$AllModules)

    if (-not (Test-Path $script:ProfilesDir)) {
        Show-Info "Creating profiles directory"
        Create-DefaultProfiles -AllModules $AllModules
    }

    $profileFiles = @(Get-ChildItem $script:ProfilesDir -Filter "*.csv" -File -ErrorAction SilentlyContinue | Sort-Object Name)

    if ($profileFiles.Count -eq 0) {
        Show-Info "No profile files found, creating defaults"
        Create-DefaultProfiles -AllModules $AllModules
        $profileFiles = @(Get-ChildItem $script:ProfilesDir -Filter "*.csv" -File -ErrorAction SilentlyContinue | Sort-Object Name)
    }

    $profiles = @()
    foreach ($file in $profileFiles) {
        $profileName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
        $enabledCount = 0
        $totalCount = 0

        try {
            $entries = @(Import-Csv $file.FullName -Encoding Default)
            $totalCount = $entries.Count
            $enabledCount = @($entries | Where-Object { $_.Enabled -eq "1" }).Count
        }
        catch {
            # CSV read error — still show profile with 0 count
        }

        $profiles += [PSCustomObject]@{
            ProfileName  = $profileName
            FilePath     = $file.FullName
            ModuleCount  = $enabledCount
            TotalCount   = $totalCount
        }
    }

    return $profiles
}

function Resolve-ProfileModules {
    param(
        [string]$ProfileCsvPath,
        [array]$AllModules
    )

    $validModules = @()
    $invalidPaths = @()
    $autoPilot = $false
    $autoPilotWaitSec = 3

    try {
        $entries = @(Import-Csv $ProfileCsvPath -Encoding Default)
    }
    catch {
        return [PSCustomObject]@{
            ValidModules     = @()
            InvalidPaths     = @()
            AutoPilot        = $false
            AutoPilotWaitSec = 3
        }
    }

    # Filter enabled, sort by Order
    $enabledEntries = @($entries | Where-Object { $_.Enabled -eq "1" })
    $sortedEntries = @($enabledEntries | Sort-Object { [int]$_.Order })

    foreach ($entry in $sortedEntries) {
        $path = $entry.ScriptPath.Trim().Replace("/", "\")
        if ([string]::IsNullOrEmpty($path)) { continue }

        # AutoPilot metadata (extracted, not added to module list)
        if ($path -eq '__AUTOPILOT__') {
            $autoPilot = $true
            if ($entry.Description -match 'WaitSec=(\d+)') {
                $autoPilotWaitSec = [int]$Matches[1]
            }
            continue
        }

        # __AUTO_to_(No)__ pattern: resolve to autologon_config module with parameter
        if ($path -match '^__AUTO_to_(.+)__$') {
            $autoLogonNo = $Matches[1]
            $autoLogonModule = $AllModules | Where-Object { $_.ModuleDir -eq 'autologon_config' } | Select-Object -First 1
            if ($autoLogonModule) {
                $moduleWithOrder = $autoLogonModule.PSObject.Copy()
                $moduleWithOrder | Add-Member -NotePropertyName "Order" -NotePropertyValue ([int]$entry.Order) -Force
                $moduleWithOrder | Add-Member -NotePropertyName "_AutoLogonNo" -NotePropertyValue $autoLogonNo
                $moduleWithOrder.MenuName = "[AUTO:$autoLogonNo] $($autoLogonModule.MenuName)"
                $validModules += $moduleWithOrder
            }
            else {
                $invalidPaths += $path
            }
            continue
        }

        # Special markers
        $specialMarkers = @{
            '__RESTART__'    = @{ MenuName = "[RESTART]";    Flag = "_IsRestart" }
            '__REEXPLORER__' = @{ MenuName = "[REEXPLORER]"; Flag = "_IsReexplorer" }
            '__STOPLOG__'    = @{ MenuName = "[STOPLOG]";    Flag = "_IsStopLog" }
            '__STARTLOG__'   = @{ MenuName = "[STARTLOG]";   Flag = "_IsStartLog" }
            '__SHUTDOWN__'   = @{ MenuName = "[SHUTDOWN]";   Flag = "_IsShutdown" }
            '__PAUSE__'      = @{ MenuName = "[PAUSE]";      Flag = "_IsPause" }
        }

        if ($specialMarkers.ContainsKey($path)) {
            $marker = $specialMarkers[$path]
            $obj = [PSCustomObject]@{
                MenuName     = $marker.MenuName
                Category     = "System"
                Script       = $null
                RelativePath = $path
                Order        = [int]$entry.Order
            }
            $obj | Add-Member -NotePropertyName $marker.Flag -NotePropertyValue $true
            $validModules += $obj
            continue
        }

        $found = $AllModules | Where-Object { $_.RelativePath -eq $path } | Select-Object -First 1
        if ($found) {
            # Attach Order from profile CSV (used for resume filtering)
            $moduleWithOrder = $found.PSObject.Copy()
            $moduleWithOrder | Add-Member -NotePropertyName "Order" -NotePropertyValue ([int]$entry.Order) -Force
            $validModules += $moduleWithOrder
        }
        else {
            $invalidPaths += $path
        }
    }

    return [PSCustomObject]@{
        ValidModules     = $validModules
        InvalidPaths     = $invalidPaths
        AutoPilot        = $autoPilot
        AutoPilotWaitSec = $autoPilotWaitSec
    }
}

function Show-ProfileMenu {
    param([array]$Profiles)

    Write-Host ""
    Show-Separator
    Write-Host "Profile List" -ForegroundColor Magenta
    Show-Separator
    Write-Host ""

    for ($i = 0; $i -lt $Profiles.Count; $i++) {
        $p = $Profiles[$i]
        Write-Host "  [$($i + 1)] $($p.ProfileName)" -ForegroundColor White
        Write-Host "      Modules: $($p.ModuleCount) enabled / $($p.TotalCount) total" -ForegroundColor DarkGray
    }

    Write-Host ""
    Write-Host "  [0] Back" -ForegroundColor Yellow
    Show-Separator
}

function Show-ProfileConfirmation {
    param(
        [object]$SelectedProfile,
        [array]$Modules,
        [array]$InvalidPaths,
        [bool]$AutoPilotFromCsv = $false,
        [int]$AutoPilotWaitSec = 3
    )

    Write-Host ""
    Show-Separator
    Write-Host "Profile: $($SelectedProfile.ProfileName)" -ForegroundColor Magenta
    Show-Separator

    # AutoPilot banner (CSV-specified)
    if ($AutoPilotFromCsv) {
        Write-Host ""
        Write-Host "========================================" -ForegroundColor Magenta
        Write-Host "  AUTOPILOT MODE (from Profile CSV)" -ForegroundColor Magenta
        Write-Host "========================================" -ForegroundColor Magenta
        Write-Host "  All confirmations will be auto-approved" -ForegroundColor White
        Write-Host "  Wait between modules: ${AutoPilotWaitSec}s" -ForegroundColor White
        Write-Host "========================================" -ForegroundColor Magenta
        Write-Host ""
    }

    Write-Host "Modules to be executed:" -ForegroundColor Cyan
    $index = 1
    foreach ($m in $Modules) {
        # Check for any special marker
        $isSpecial = $m._IsRestart -or $m._IsReexplorer -or $m._IsStopLog -or
                     $m._IsStartLog -or $m._IsShutdown -or $m._IsPause
        if ($isSpecial) {
            Write-Host "  [$index] --- $($m.MenuName) ---" -ForegroundColor Yellow
        }
        else {
            Write-Host "  [$index] $($m.MenuName) ($($m.Category))" -ForegroundColor White
        }
        $index++
    }

    if ($InvalidPaths.Count -gt 0) {
        Write-Host ""
        Write-Host "Warning: Module paths not found:" -ForegroundColor Yellow
        foreach ($p in $InvalidPaths) {
            Write-Host "    $p" -ForegroundColor Yellow
        }
    }

    Show-Separator
    Write-Host ""

    # AutoPilot from CSV: require one final confirmation (safety valve)
    if ($AutoPilotFromCsv) {
        if (-not (Confirm-Execution -Message "Start AutoPilot execution?")) {
            return $null
        }

        return [PSCustomObject]@{
            Confirmed        = $true
            StopOnError      = $false
            AutoPilot        = $true
            AutoPilotWaitSec = $AutoPilotWaitSec
        }
    }

    # Normal flow: confirmation + mode selection
    if (-not (Confirm-Execution -Message "Are you sure you want to execute?")) {
        return $null
    }

    # Execution mode prompt
    Write-Host ""
    Write-Host "  [1] Continue on Error (Default)" -ForegroundColor White
    Write-Host "  [2] Stop on Error" -ForegroundColor White
    Write-Host "  [3] AutoPilot (Auto-confirm all)" -ForegroundColor Magenta
    Write-Host ""
    Write-Host -NoNewline "Execution mode [1]: "
    $modeChoice = Read-Host

    $stopOnError = ($modeChoice -eq "2")
    $autoPilot = ($modeChoice -eq "3")

    return [PSCustomObject]@{
        Confirmed        = $true
        StopOnError      = $stopOnError
        AutoPilot        = $autoPilot
        AutoPilotWaitSec = $AutoPilotWaitSec
    }
}

# ========================================
# Log Functions
# ========================================

function Write-KitLog {
    param(
        [ValidateSet("INFO", "WARN", "ERROR", "SUCCESS")]
        [string]$Level = "INFO",
        [string]$Message
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logLine = "[$timestamp] [$Level] $Message"

    # Output to console
    switch ($Level) {
        "INFO"    { Write-Host $logLine -ForegroundColor Cyan }
        "WARN"    { Write-Host $logLine -ForegroundColor Yellow }
        "ERROR"   { Write-Host $logLine -ForegroundColor Red }
        "SUCCESS" { Write-Host $logLine -ForegroundColor Green }
    }
}

function Save-RollbackInfo {
    param(
        [string]$Category,
        [string]$Key,
        [string]$OldValue,
        [string]$NewValue
    )

    $info = [PSCustomObject]@{
        Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Category  = $Category
        Key       = $Key
        OldValue  = $OldValue
        NewValue  = $NewValue
    }

    # Logging
    Write-KitLog -Level "INFO" -Message "Change Log: [$Category] $Key : '$OldValue' -> '$NewValue'"

    return $info
}

# ========================================
# Utility Functions
# ========================================

function Get-ModuleBasePath {
    # Get base path if called from module script
    if ($PSScriptRoot) {
        return $PSScriptRoot
    }
    return (Get-Location).Path
}

function Test-AdminPrivilege {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# ========================================
# Status Monitor Functions
# ========================================

function Get-CurrentPCInfo {
    $result = @{
        ComputerName    = $env:COMPUTERNAME
        EthernetIP      = ""
        EthernetSubnet  = ""
        EthernetGateway = ""
        WifiIP          = ""
        WifiSubnet      = ""
        WifiGateway     = ""
        DNS             = @()
        Printers        = @()
    }

    try {
        $adapters = @(Get-NetAdapter -Physical -ErrorAction SilentlyContinue |
                      Where-Object { $_.Status -ne "Disabled" })

        foreach ($adapter in $adapters) {
            $ipEntry = Get-NetIPAddress -InterfaceIndex $adapter.ifIndex `
                        -AddressFamily IPv4 -ErrorAction SilentlyContinue |
                       Where-Object { $_.PrefixOrigin -ne "WellKnown" } |
                       Select-Object -First 1

            if ($null -eq $ipEntry) { continue }

            # PrefixLength -> SubnetMask conversion
            $prefixLen = $ipEntry.PrefixLength
            $maskInt = if ($prefixLen -gt 0) {
                [uint32]([math]::Pow(2, 32) - [math]::Pow(2, 32 - $prefixLen))
            } else { [uint32]0 }
            $subnet = "{0}.{1}.{2}.{3}" -f `
                (($maskInt -shr 24) -band 0xFF),
                (($maskInt -shr 16) -band 0xFF),
                (($maskInt -shr 8) -band 0xFF),
                ($maskInt -band 0xFF)

            # Gateway
            $gwConfig = Get-NetIPConfiguration -InterfaceIndex $adapter.ifIndex `
                        -ErrorAction SilentlyContinue
            $gateway = if ($gwConfig.IPv4DefaultGateway) {
                $gwConfig.IPv4DefaultGateway.NextHop
            } else { "" }

            # Wi-Fi or Ethernet (InterfaceDescription is always English regardless of OS locale)
            $isWifi = $adapter.InterfaceDescription -match "Wi-Fi|Wireless|WLAN|802\.11"

            if ($isWifi) {
                $result.WifiIP      = $ipEntry.IPAddress
                $result.WifiSubnet  = $subnet
                $result.WifiGateway = $gateway
            }
            elseif ([string]::IsNullOrEmpty($result.EthernetIP)) {
                $result.EthernetIP      = $ipEntry.IPAddress
                $result.EthernetSubnet  = $subnet
                $result.EthernetGateway = $gateway
            }
        }

        # DNS from all active adapters, deduplicated
        $dnsAll = @()
        foreach ($adapter in $adapters) {
            $dnsEntry = Get-DnsClientServerAddress -InterfaceIndex $adapter.ifIndex `
                        -AddressFamily IPv4 -ErrorAction SilentlyContinue
            if ($dnsEntry.ServerAddresses) {
                $dnsAll += $dnsEntry.ServerAddresses
            }
        }
        $result.DNS = @($dnsAll | Select-Object -Unique | Select-Object -First 4)

        # Printers (network printers + IP port printers only)
        # PortName examples: "IP_192.168.0.1", "TCPIP_192.168.0.1", "192.168.0.1"
        $printers = @(Get-Printer -ErrorAction SilentlyContinue |
                      Where-Object {
                          $_.Type -eq "Connection" -or
                          $_.PortName -match "\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}"
                      })
        foreach ($p in $printers) {
            $result.Printers += @{
                Name = $p.Name
                Port = $p.PortName
            }
        }
    }
    catch { }

    return $result
}

function Write-StatusFile {
    param(
        [string]$Phase = "idle"
    )

    try {
        # PC情報を環境変数から収集
        $pcInfo = @{
            AdminID         = $env:SELECTED_KANRI_NO
            OldPCName       = $env:SELECTED_OLD_PCNAME
            NewPCName       = $env:SELECTED_NEW_PCNAME
            EthernetIP      = $env:SELECTED_ETH_IP
            EthernetSubnet  = $env:SELECTED_ETH_SUBNET
            EthernetGateway = $env:SELECTED_ETH_GATEWAY
            WifiIP          = $env:SELECTED_WIFI_IP
            WifiSubnet      = $env:SELECTED_WIFI_SUBNET
            WifiGateway     = $env:SELECTED_WIFI_GATEWAY
            DNS             = @($env:SELECTED_DNS1, $env:SELECTED_DNS2, $env:SELECTED_DNS3, $env:SELECTED_DNS4) | Where-Object { -not [string]::IsNullOrEmpty($_) }
            Printers        = @()
        }

        # プリンタ情報を収集 (1-10)
        for ($i = 1; $i -le 10; $i++) {
            $pName = [Environment]::GetEnvironmentVariable("SELECTED_PRINTER_$($i)_NAME")
            if (-not [string]::IsNullOrEmpty($pName)) {
                $pcInfo.Printers += @{
                    Name   = $pName
                    Driver = [Environment]::GetEnvironmentVariable("SELECTED_PRINTER_$($i)_DRIVER")
                    Port   = [Environment]::GetEnvironmentVariable("SELECTED_PRINTER_$($i)_PORT")
                }
            }
        }

        # 実行結果サマリーを集計
        $results = @($script:ExecutionResults)
        $executionInfo = @{
            Phase          = $Phase
            TotalCount     = $results.Count
            SuccessCount   = @($results | Where-Object { $_.Status -eq "Success" }).Count
            ErrorCount     = @($results | Where-Object { $_.Status -eq "Error" }).Count
            SkippedCount   = @($results | Where-Object { $_.Status -eq "Skip" -or $_.Status -eq "Skipped" }).Count
            CancelledCount = @($results | Where-Object { $_.Status -eq "Cancelled" }).Count
            PartialCount   = @($results | Where-Object { $_.Status -eq "Partial" }).Count
            WarningCount   = @($results | Where-Object { $_.Status -eq "Warning" }).Count
            Details        = @()
        }

        foreach ($r in $results) {
            $executionInfo.Details += @{
                Operation  = $r.Operation
                Status     = $r.Status
                Message    = $r.Message
                Timestamp  = if ($r.Timestamp) { $r.Timestamp.ToString("yyyy-MM-dd HH:mm:ss") } else { "" }
                IsRestored = if ($r.IsRestored) { $true } else { $false }
            }
        }

        $currentPC = Get-CurrentPCInfo

        $statusData = @{
            UpdatedAt     = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            PCInfo        = $pcInfo
            CurrentPCInfo = $currentPC
            Execution     = $executionInfo
        }

        # アトミック書き込み: tmpファイルに書いてからリネーム
        $tempPath = "$($script:StatusFilePath).tmp"
        $statusData | ConvertTo-Json -Depth 5 | Out-File -FilePath $tempPath -Encoding UTF8 -Force
        Move-Item -Path $tempPath -Destination $script:StatusFilePath -Force
    }
    catch {
        # ステータスモニターは補助機能のため、エラーは無視
        try {
            # Move-Itemが失敗した場合のフォールバック: 直接書き込み
            if ($statusData) {
                $statusData | ConvertTo-Json -Depth 5 | Out-File -FilePath $script:StatusFilePath -Encoding UTF8 -Force
            }
        }
        catch { }
    }
}

function Remove-StatusFile {
    try {
        if (Test-Path $script:StatusFilePath) {
            Remove-Item $script:StatusFilePath -Force -ErrorAction SilentlyContinue
        }
        $tempPath = "$($script:StatusFilePath).tmp"
        if (Test-Path $tempPath) {
            Remove-Item $tempPath -Force -ErrorAction SilentlyContinue
        }
    }
    catch { }
}

# ========================================
# Status Monitor Lifecycle
# ========================================
function Start-StatusMonitor {
    Write-StatusFile -Phase "idle"
    $monitorProcess = $null
    try {
        $monitorScript = ".\kernel\ps1\status_monitor.ps1"
        if (Test-Path $monitorScript) {
            $statusFileFullPath = (Resolve-Path $script:StatusFilePath).Path
            $monitorProcess = Start-Process powershell.exe -ArgumentList @(
                "-NoProfile", "-ExecutionPolicy", "Unrestricted",
                "-File", $monitorScript,
                "-StatusFilePath", $statusFileFullPath
            ) -WindowStyle Hidden -PassThru
            Show-Info "Status Monitor started (PID: $($monitorProcess.Id))"
            Write-Host ""

            Start-Sleep -Milliseconds 1200
            Set-ConsoleForeground
        }
    }
    catch {
        Show-Warning "Failed to start Status Monitor: $_"
        Write-Host ""
    }
    return $monitorProcess
}

function Stop-StatusMonitor {
    param([System.Diagnostics.Process]$MonitorProcess)

    if ($MonitorProcess -and -not $MonitorProcess.HasExited) {
        try {
            $MonitorProcess.CloseMainWindow() | Out-Null
            if (-not $MonitorProcess.WaitForExit(2000)) {
                $MonitorProcess.Kill()
            }
        }
        catch { }
    }
    Remove-StatusFile
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
# Module System Initialization
# ========================================
function Initialize-ModuleSystem {
    param(
        [string]$CategoriesCsv = ".\kernel\csv\categories.csv",
        [string]$ModulesDir = ".\modules"
    )

    Show-Info "Loading categories.csv..."
    $categoryOrder = @{}
    if (Test-Path $CategoriesCsv) {
        try {
            $categories = Import-Csv -Path $CategoriesCsv -Encoding Default
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
    Write-Host ""

    Show-Info "Detecting modules..."
    $allModules = @()
    $standardPath = Join-Path $ModulesDir "standard"
    $extendedPath = Join-Path $ModulesDir "extended"

    foreach ($type in @(@{Path=$standardPath;Type="standard"}, @{Path=$extendedPath;Type="extended"})) {
        if (Test-Path $type.Path) {
            $dirs = Get-ChildItem $type.Path -Directory -ErrorAction SilentlyContinue
            foreach ($dir in $dirs) {
                $moduleCsv = Join-Path $dir.FullName "module.csv"
                if (Test-Path $moduleCsv) {
                    try {
                        $entries = Import-Csv $moduleCsv -Encoding Default
                        foreach ($entry in $entries) {
                            if ($entry.Enabled -eq "0") { continue }
                            $order = 100
                            if ($entry.Order -and $entry.Order -match '^\d+$') {
                                $order = [int]$entry.Order
                            }
                            $allModules += [PSCustomObject]@{
                                MenuName     = $entry.MenuName
                                Category     = $entry.Category
                                Script       = Join-Path $dir.FullName $entry.Script
                                Order        = $order
                                ModuleType   = $type.Type
                                ModuleDir    = $dir.Name
                                RelativePath = "$($type.Type)\$($dir.Name)\$($entry.Script)"
                            }
                        }
                    }
                    catch {
                        Show-Error "Error loading module.csv: $($dir.Name) - $_"
                    }
                }
            }
        }
    }

    $count = ($allModules | Measure-Object).Count
    if ($count -eq 0) {
        Show-Error "No valid modules found"
        return $null
    }
    Show-Success "Modules loaded ($count items)"
    Write-Host ""

    $groupedModules = Build-CategoryMenu -Modules $allModules -CategoryOrder $categoryOrder

    return [PSCustomObject]@{
        AllModules     = $allModules
        GroupedModules = $groupedModules
        CategoryOrder  = $categoryOrder
    }
}

# ========================================
# RunOnce Registration & Countdown
# ========================================
function Register-FabriqRunOnce {
    $fabriqRoot = (Resolve-Path ".").Path
    $fabriqBat = Join-Path $fabriqRoot "Fabriq.bat"

    if (-not (Test-Path $fabriqBat)) {
        Show-Error "Fabriq.bat not found: $fabriqBat"
        return $false
    }

    $runOncePath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"
    try {
        if (-not (Test-Path $runOncePath)) {
            New-Item -Path $runOncePath -Force | Out-Null
        }
        $runOnceValue = "cmd /c `"$fabriqBat`""
        New-ItemProperty -Path $runOncePath -Name "FabriqAutoStart" `
            -Value $runOnceValue -PropertyType String -Force -ErrorAction Stop | Out-Null
        Show-Success "RunOnce registered"
        return $true
    }
    catch {
        Show-Error "Failed to register RunOnce: $_"
        return $false
    }
}

function Invoke-CountdownRestart {
    param([int]$Seconds = 5)

    Write-Host ""
    Write-Host "The computer will restart in $Seconds seconds..." -ForegroundColor Yellow
    Write-Host "Press Ctrl+C to abort" -ForegroundColor Yellow
    Write-Host ""
    for ($i = $Seconds; $i -ge 1; $i--) {
        Write-Host "`r  Restarting in $i seconds... " -NoNewline -ForegroundColor Yellow
        Start-Sleep -Seconds 1
    }
    Write-Host ""
    Restart-Computer -Force
    Start-Sleep -Seconds 30
}

# ========================================
# Function: Capture Screen Evidence
# ========================================
# Captures a screenshot of the primary screen
# and saves it as PNG for quality assurance.
# Silently fails on error (never stops execution).
# ========================================
function Capture-ScreenEvidence {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ModuleName,
        [string]$Status = ""
    )

    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
        Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue

        # Enable DPI awareness for accurate full-screen capture on scaled displays
        Add-Type -TypeDefinition @"
            using System.Runtime.InteropServices;
            public class DPIUtil {
                [DllImport("user32.dll")]
                public static extern bool SetProcessDPIAware();
            }
"@ -ErrorAction SilentlyContinue
        $null = [DPIUtil]::SetProcessDPIAware()

        # Build save directory: yyyy_MM_dd_{SN}_{PCname}
        # Date-only (no time) so all captures within the same day share one directory
        $pcName  = if ($env:SELECTED_NEW_PCNAME) { $env:SELECTED_NEW_PCNAME } else { $env:COMPUTERNAME }
        $dateOnly = Get-Date -Format "yyyy_MM_dd"
        $uid      = if ($global:FabriqUniqueId) { $global:FabriqUniqueId } else { Get-HardwareUniqueId }
        $saveDir = Join-Path $PSScriptRoot "..\evidence\auto_capture\${dateOnly}_${uid}_${pcName}"
        if (-not (Test-Path $saveDir)) {
            New-Item -Path $saveDir -ItemType Directory -Force | Out-Null
        }

        # Build filename: yyyy_mm_dd_HHmmss_ModuleName_Status_PCName.png
        $timestamp = Get-Date -Format "yyyy_MM_dd_HHmmss"
        $safeName = ($ModuleName -replace '[\\/:*?"<>|\s]', '_')
        $statusSuffix = if ($Status) { "_$Status" } else { "" }
        $fileName = "${timestamp}_${safeName}${statusSuffix}_${pcName}.png"
        $filePath = Join-Path $saveDir $fileName

        # Capture primary screen
        $screen = [System.Windows.Forms.Screen]::PrimaryScreen
        $bounds = $screen.Bounds
        $bitmap = New-Object System.Drawing.Bitmap($bounds.Width, $bounds.Height)
        $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
        $graphics.CopyFromScreen($bounds.Location, [System.Drawing.Point]::Empty, $bounds.Size)
        $graphics.Dispose()

        # Save as PNG
        $bitmap.Save($filePath, [System.Drawing.Imaging.ImageFormat]::Png)
        $bitmap.Dispose()
    }
    catch {
        Write-Warning "Screen capture failed: $($_.Exception.Message)"
    }
}

function Invoke-CountdownShutdown {
    param([int]$Seconds = 5)

    Write-Host ""
    Write-Host "The computer will shut down in $Seconds seconds..." -ForegroundColor Yellow
    Write-Host "Press Ctrl+C to abort" -ForegroundColor Yellow
    Write-Host ""
    for ($i = $Seconds; $i -ge 1; $i--) {
        Write-Host "`r  Shutting down in $i seconds... " -NoNewline -ForegroundColor Yellow
        Start-Sleep -Seconds 1
    }
    Write-Host ""
    Stop-Computer -Force
    Start-Sleep -Seconds 30
}

function Invoke-CountdownSignout {
    param([int]$Seconds = 7)

    Write-Host ""
    Write-Host "Signing out in $Seconds seconds..." -ForegroundColor Yellow
    Write-Host "Press Ctrl+C to abort" -ForegroundColor Yellow
    Write-Host ""
    for ($i = $Seconds; $i -ge 1; $i--) {
        Write-Host "`r  Signing out in $i seconds... " -NoNewline -ForegroundColor Yellow
        Start-Sleep -Seconds 1
    }
    Write-Host ""
    try { Stop-Transcript | Out-Null } catch { }
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public class FabriqSignOut {
    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool ExitWindowsEx(uint uFlags, uint dwReason);
}
'@ -ErrorAction SilentlyContinue
    [FabriqSignOut]::ExitWindowsEx(4, 0)
    Start-Sleep -Seconds 30
}