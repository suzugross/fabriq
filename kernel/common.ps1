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
$script:StatusFilePath = ".\kernel\status.json"
$script:ResumeStatePath = ".\kernel\resume_state.json"

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
    Write-Host ""
    Write-Host $Message
    Read-Host
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
    $operator = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name

    # CSV Escape (if containing comma or newlines)
    $escapedMessage = $Message -replace '"', '""'
    if ($escapedMessage -match '[,\r\n]') {
        $escapedMessage = "`"$escapedMessage`""
    }

    $line = "$timestamp,$env:SELECTED_KANRI_NO,$env:SELECTED_NEW_PCNAME,$ModuleName,$Category,$Status,$escapedMessage,$operator,$($script:SessionID)"

    # Create with header if file does not exist
    $needHeader = -not (Test-Path $script:HistoryPath)

    for ($i = 0; $i -lt $maxRetry; $i++) {
        try {
            if ($needHeader) {
                $header = "Timestamp,KanriNo,PCName,ModuleName,Category,Status,Message,Operator,SessionID"
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

    $evidenceExportPath = Join-Path $evidenceExportDir "history_export_${pcName}_$dateStr.csv"
    try {
        Copy-Item $script:HistoryPath $evidenceExportPath -Force
        Show-Success "Evidence copy:    $evidenceExportPath"
    }
    catch {
        Show-Warning "Failed to copy to evidence: $_"
    }

    return $exportPath
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

    # 3. Transcript Logs (logs/log_*.txt)
    if (Test-Path ".\logs") {
        $transcripts = @(Get-ChildItem ".\logs" -Filter "log_*.txt" -File -ErrorAction SilentlyContinue)
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

function Restore-HostEnvironment {
    param([object]$HostEnv)
    $HostEnv.PSObject.Properties | ForEach-Object {
        Set-Item -Path "env:$($_.Name)" -Value $_.Value -ErrorAction SilentlyContinue
    }
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

    try {
        $entries = @(Import-Csv $ProfileCsvPath -Encoding Default)
    }
    catch {
        return [PSCustomObject]@{
            ValidModules = @()
            InvalidPaths = @()
        }
    }

    # Filter enabled, sort by Order
    $enabledEntries = @($entries | Where-Object { $_.Enabled -eq "1" })
    $sortedEntries = @($enabledEntries | Sort-Object { [int]$_.Order })

    foreach ($entry in $sortedEntries) {
        $path = $entry.ScriptPath.Trim().Replace("/", "\")
        if ([string]::IsNullOrEmpty($path)) { continue }

        # __RESTART__ marker
        if ($path -eq '__RESTART__') {
            $validModules += [PSCustomObject]@{
                MenuName     = "[RESTART]"
                Category     = "System"
                Script       = $null
                RelativePath = "__RESTART__"
                Order        = [int]$entry.Order
                _IsRestart   = $true
            }
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
        ValidModules = $validModules
        InvalidPaths = $invalidPaths
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
        [array]$InvalidPaths
    )

    Write-Host ""
    Show-Separator
    Write-Host "Profile: $($SelectedProfile.ProfileName)" -ForegroundColor Magenta
    Show-Separator

    Write-Host "Modules to be executed:" -ForegroundColor Cyan
    $index = 1
    foreach ($m in $Modules) {
        if ($m._IsRestart) {
            Write-Host "  [$index] --- RESTART ---" -ForegroundColor Yellow
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

    if (-not (Confirm-Execution -Message "Are you sure you want to execute?")) {
        return $null
    }

    # StopOnError runtime prompt
    Write-Host ""
    Write-Host "  [1] Continue on Error (Default)" -ForegroundColor White
    Write-Host "  [2] Stop on Error" -ForegroundColor White
    Write-Host ""
    Write-Host -NoNewline "Error handling [1]: "
    $errorChoice = Read-Host
    $stopOnError = ($errorChoice -eq "2")

    return [PSCustomObject]@{
        Confirmed   = $true
        StopOnError = $stopOnError
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