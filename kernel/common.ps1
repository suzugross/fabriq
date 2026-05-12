# ========================================
# Easy Kitting Batch - Common Function Library v3.4.0
# ========================================

# ========================================
# Global Variables
# ========================================
$script:ExecutionResults = @()

# Last batch execution results (per-module).
# Populated by Invoke-BatchExecution at natural completion (and from
# finally on cancel / mid-throw paths). Each entry is a hashtable:
#   Order    : int     - profile CSV Order (markers included with their Order)
#   MenuName : string  - module display name (with [seg:..] suffix etc)
#   Status   : string  - Success / Error / Skipped / Cancelled / Partial
#   Verified : Nullable[bool] - Post-Apply Verification result, or $null
#   Message  : string  - module-reported message
# Consumed by FlexProfile dashboard to flash post-execution state without
# re-importing execution_history.csv. Reset by Reset-FabriqState.
$script:LastBatchResults = @()

$script:SessionID = Get-Date -Format "yyyyMMdd_HHmmss"
$script:HistoryPath = ".\logs\history\execution_history.csv"
$script:ProfilesDir = ".\profiles"
$script:StatusFilePath = ".\kernel\json\status.json"
$script:ResumeStatePath = ".\kernel\json\resume_state.json"
$script:SessionFilePath = ".\kernel\json\session.json"
$script:SourceMediaIdPath = ".\kernel\source_media.id"
$script:WorkersCsvPath = ".\kernel\csv\workers.csv"
$global:ArtPulseFilePath = ".\kernel\json\art_pulse.txt"
$global:ArtPulseCounter = 0

# Session info (populated by Initialize-Session)
$script:SessionInfo = $null

# AutoPilot Mode (Profile execution only)
$global:AutoPilotMode = $false
$global:AutoPilotWaitSec = 3

# AutoConfirm Mode (FlexProfile single-execution only).
# Suppresses Y/N prompts and Wait-KeyPress so a one-click Run-This run
# completes without blocking. Strictly a subset of AutoPilot: does NOT
# enable inter-module wait, ErrorMode dispatch, or the
# Show-AutoPilotErrorDialog retry loop. Set $true only for the duration
# of one Flex single-module run; reset in finally.
$global:AutoConfirmMode = $false

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
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
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

# ========================================
# Console Window Visibility Control
# ========================================
function Hide-ConsoleWindow {
    $hwnd = [ConsoleFocus]::GetConsoleWindow()
    if ($hwnd -ne [IntPtr]::Zero) {
        [ConsoleFocus]::ShowWindow($hwnd, 0) | Out-Null  # SW_HIDE
    }
}

function Show-ConsoleWindow {
    $hwnd = [ConsoleFocus]::GetConsoleWindow()
    if ($hwnd -ne [IntPtr]::Zero) {
        [ConsoleFocus]::ShowWindow($hwnd, 5) | Out-Null  # SW_SHOW
        [ConsoleFocus]::SetForegroundWindow($hwnd) | Out-Null
    }
}

# ========================================
# QuickEdit Mode Disabler (SetConsoleMode)
# ========================================
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public class QuickEditDisabler {
    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern IntPtr GetStdHandle(int nStdHandle);
    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern bool GetConsoleMode(IntPtr hConsoleHandle, out uint lpMode);
    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern bool SetConsoleMode(IntPtr hConsoleHandle, uint dwMode);

    public const int  STD_INPUT_HANDLE       = -10;
    public const uint ENABLE_QUICK_EDIT_MODE  = 0x0040;
    public const uint ENABLE_EXTENDED_FLAGS   = 0x0080;
}
'@ -ErrorAction SilentlyContinue

function Disable-QuickEditMode {
    try {
        $handle = [QuickEditDisabler]::GetStdHandle(
                      [QuickEditDisabler]::STD_INPUT_HANDLE)
        if ($handle -eq [IntPtr]::Zero -or $handle -eq [IntPtr]::new(-1)) {
            return
        }
        $mode = [uint32]0
        if (-not [QuickEditDisabler]::GetConsoleMode($handle, [ref]$mode)) {
            return
        }
        $mode = $mode -band (-bnot [QuickEditDisabler]::ENABLE_QUICK_EDIT_MODE)
        $mode = $mode -bor  [QuickEditDisabler]::ENABLE_EXTENDED_FLAGS
        [QuickEditDisabler]::SetConsoleMode($handle, $mode) | Out-Null
    }
    catch { }
}

function Set-ConsoleForeground {
    try {
        $hwnd = [ConsoleFocus]::GetConsoleWindow()
        if ($hwnd -ne [IntPtr]::Zero) {
            [ConsoleFocus]::SetForegroundWindow($hwnd) | Out-Null
        }
    }
    catch { }
}

# ========================================
# Error Notification (Beep + Foreground)
# ========================================
function Invoke-ErrorNotification {
    param(
        [string]$ModuleName = "",
        [ValidateSet("Error", "Partial")]
        [string]$Status = "Error"
    )

    # Bring console to foreground so the operator notices
    Set-ConsoleForeground

    try {
        if ($Status -eq "Error") {
            # Three ascending beeps for error
            [console]::Beep(600, 200); Start-Sleep -Milliseconds 100
            [console]::Beep(800, 200); Start-Sleep -Milliseconds 100
            [console]::Beep(1000, 400)
        }
        else {
            # Two beeps for partial
            [console]::Beep(600, 300); Start-Sleep -Milliseconds 150
            [console]::Beep(600, 300)
        }
    }
    catch {
        # Non-fatal: sound failure should not affect execution
    }
}

# ========================================
# AutoPilot Error Dialog (Retry / Skip)
# ========================================
function Show-AutoPilotErrorDialog {
    param(
        [Parameter(Mandatory)][string]$ModuleName,
        [Parameter(Mandatory)][string]$Status,
        [string]$Message = ""
    )

    Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue

    $text = "Module: $ModuleName`nStatus: $Status"
    if (-not [string]::IsNullOrEmpty($Message)) {
        $text += "`nDetail: $Message"
    }
    $text += "`n`nRetry = Re-execute this module`nCancel = Skip and continue"

    $dialogResult = [System.Windows.Forms.MessageBox]::Show(
        $text,
        "fabriq - AutoPilot Error",
        [System.Windows.Forms.MessageBoxButtons]::RetryCancel,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    if ($dialogResult -eq [System.Windows.Forms.DialogResult]::Retry) {
        return "Retry"
    }
    return "Skip"
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
# Telemetry Layer (internal — see dev/TELEMETRY_INTERNAL.md)
# ========================================
# AI development corpus: capture per-module envelope + Show-* events
# + full ErrorRecord into JSONL under logs/telemetry/{SessionID}/.
# Module side stays untouched — instrumentation lives entirely in the
# Show-* family + Invoke-SafeCommand[Async] envelope hooks.
#
# CRITICAL invariants:
#   1. Telemetry never affects kitting outcomes. Every write path is
#      wrapped in try/catch with no fallback surface to operator.
#   2. Reentrancy: Show-* called from inside a telemetry write must NOT
#      recurse back into telemetry. Guarded by $global:_TelemetryWriting.
#   3. Privacy: every emitted string passes through the redact map built
#      at envelope start from $env:SELECTED_*, $env:FABRIQ_WORKER_NAME,
#      $env:COMPUTERNAME, $global:FabriqUniqueId.
# ========================================

# Reentrancy guard. Set true while Write-TelemetryEvent is appending.
$global:_TelemetryWriting = $false

# Cached salt (loaded lazily from kernel/json/telemetry_salt.txt).
$global:_TelemetrySalt = $null
$global:_TelemetrySaltDigest = $null

# Chronological sequence counter within current session (1-based).
$global:_TelemetryModuleSeq = 0

# Active per-module envelope. Show-* reads this; Start/Complete write it.
# Schema: @{ Path; RedactMap; Module; Order; Sequence; StartTime;
#            ShowCounts = @{info,success,warning,error,skip} }
$global:_CurrentModuleTelemetry = $null

# UTF-8 without BOM (PSv5 [System.Text.Encoding]::UTF8 includes BOM).
$global:_TelemetryUtf8NoBom = New-Object System.Text.UTF8Encoding($false)

function Get-TelemetrySalt {
    if ($null -ne $global:_TelemetrySalt) { return $global:_TelemetrySalt }

    $saltPath = ".\kernel\json\telemetry_salt.txt"
    try {
        $saltDir = Split-Path $saltPath -Parent
        if (-not (Test-Path $saltDir)) {
            New-Item -ItemType Directory -Path $saltDir -Force | Out-Null
        }

        if (Test-Path $saltPath) {
            $salt = (Get-Content $saltPath -Raw -Encoding UTF8 -ErrorAction Stop).Trim()
        }
        else {
            # Generate fresh 32-byte (256-bit) salt
            $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()
            $bytes = New-Object byte[] 32
            $rng.GetBytes($bytes)
            $rng.Dispose()
            $salt = [Convert]::ToBase64String($bytes)
            [System.IO.File]::WriteAllText($saltPath, $salt, $global:_TelemetryUtf8NoBom)
        }

        if ([string]::IsNullOrWhiteSpace($salt)) { return $null }

        # Salt digest (non-secret correlation indicator across sessions).
        $sha = [System.Security.Cryptography.SHA256]::Create()
        $digestBytes = $sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($salt))
        $sha.Dispose()
        $digestHex = -join ($digestBytes[0..3] | ForEach-Object { $_.ToString('x2') })

        $global:_TelemetrySalt = $salt
        $global:_TelemetrySaltDigest = "sha256:$digestHex"
        return $salt
    }
    catch {
        # Salt machinery broken — disable hashing (every value will become
        # [REDACTED] in the map; safe-by-default).
        $global:_TelemetrySalt = $null
        return $null
    }
}

function _HashTelemetryValue {
    param([string]$Value, [string]$Prefix)
    if ([string]::IsNullOrEmpty($Value)) { return $Value }
    $salt = Get-TelemetrySalt
    if ($null -eq $salt) { return "[REDACTED]" }

    try {
        $sha = [System.Security.Cryptography.SHA256]::Create()
        $bytes = [System.Text.Encoding]::UTF8.GetBytes("$salt|$Value")
        $hash = $sha.ComputeHash($bytes)
        $sha.Dispose()
        # 6 bytes = 12 hex chars = ~10^14 collision space, plenty for
        # cross-session correlation within one site.
        $hex = -join ($hash[0..5] | ForEach-Object { $_.ToString('x2') })
        return "${Prefix}:$hex"
    }
    catch { return "[REDACTED]" }
}

function New-TelemetryRedactMap {
    # Snapshot all known sensitive values from current env. Map key =
    # literal value, map value = redacted token. Applied longest-first
    # at redact time to avoid partial overlap.
    $map = @{}

    $hashSpec = @(
        @{ Var = $env:SELECTED_KANRI_NO     ; Pre = "KANRI"  }
        @{ Var = $env:SELECTED_OLD_PCNAME   ; Pre = "PC"     }
        @{ Var = $env:SELECTED_NEW_PCNAME   ; Pre = "PC"     }
        @{ Var = $env:SELECTED_ETH_IP       ; Pre = "IP"     }
        @{ Var = $env:SELECTED_ETH_SUBNET   ; Pre = "IP"     }
        @{ Var = $env:SELECTED_ETH_GATEWAY  ; Pre = "IP"     }
        @{ Var = $env:SELECTED_WIFI_IP      ; Pre = "IP"     }
        @{ Var = $env:SELECTED_WIFI_SUBNET  ; Pre = "IP"     }
        @{ Var = $env:SELECTED_WIFI_GATEWAY ; Pre = "IP"     }
        @{ Var = $env:SELECTED_DNS1         ; Pre = "IP"     }
        @{ Var = $env:SELECTED_DNS2         ; Pre = "IP"     }
        @{ Var = $env:SELECTED_DNS3         ; Pre = "IP"     }
        @{ Var = $env:SELECTED_DNS4         ; Pre = "IP"     }
        @{ Var = $env:FABRIQ_WORKER_NAME    ; Pre = "WORKER" }
        @{ Var = $env:COMPUTERNAME          ; Pre = "HOST"   }
        @{ Var = $global:FabriqUniqueId     ; Pre = "HW"     }
    )

    for ($i = 1; $i -le 10; $i++) {
        foreach ($suffix in @('NAME','DRIVER','PORT')) {
            $envName = "SELECTED_PRINTER_${i}_${suffix}"
            $val = [Environment]::GetEnvironmentVariable($envName)
            if (-not [string]::IsNullOrWhiteSpace($val)) {
                $hashSpec += @{ Var = $val; Pre = "PRINTER" }
            }
        }
    }

    foreach ($s in $hashSpec) {
        $v = $s.Var
        # Skip empty / very short values (1-2 chars cause false positives
        # on common substrings like "01", "OK", etc.).
        if ([string]::IsNullOrWhiteSpace($v)) { continue }
        if ($v.Length -lt 3) { continue }
        if (-not $map.ContainsKey($v)) {
            $map[$v] = (_HashTelemetryValue -Value $v -Prefix $s.Pre)
        }
    }

    # Hard-redacted (PIN — never stored even as hash).
    if (-not [string]::IsNullOrWhiteSpace($env:SELECTED_PIN)) {
        $map[$env:SELECTED_PIN] = "[REDACTED]"
    }

    return $map
}

function Invoke-TelemetryRedact {
    param([string]$Text, [hashtable]$Map)
    if ([string]::IsNullOrEmpty($Text)) { return $Text }
    if ($null -eq $Map -or $Map.Count -eq 0) { return $Text }

    # Longest-first replacement avoids partial-match bleed (e.g. a short
    # value being a substring of a longer sensitive value).
    $keys = @($Map.Keys | Sort-Object { $_.Length } -Descending)
    $out = $Text
    foreach ($k in $keys) {
        if ([string]::IsNullOrEmpty($k)) { continue }
        $out = $out.Replace($k, [string]$Map[$k])
    }
    return $out
}

function Write-TelemetryEvent {
    param(
        [Parameter(Mandatory)][string]$Type,
        [hashtable]$Data = @{}
    )

    if ($global:_TelemetryWriting) { return }

    $env_ = $global:_CurrentModuleTelemetry
    if ($null -eq $env_) { return }
    if ([string]::IsNullOrWhiteSpace($env_.Path)) { return }

    $global:_TelemetryWriting = $true
    try {
        $line = [ordered]@{
            ts   = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss.fffzzz")
            type = $Type
        }
        foreach ($k in $Data.Keys) { $line[$k] = $Data[$k] }

        $json = ([PSCustomObject]$line | ConvertTo-Json -Depth 6 -Compress)
        # AppendAllText auto-creates the file on first call; subsequent
        # calls append. Encoding without BOM (already created by salt /
        # _meta.json or here on first write).
        [System.IO.File]::AppendAllText($env_.Path, $json + "`n", $global:_TelemetryUtf8NoBom)
    }
    catch {
        # Telemetry must never bubble. Swallow.
    }
    finally {
        $global:_TelemetryWriting = $false
    }
}

function _GetShowTag {
    param([string]$Function, [string]$Message)
    if ([string]::IsNullOrEmpty($Message)) { return $Function }
    # Trim leading whitespace: many modules indent prefix tags for visual
    # alignment ("  [APPLY] foo") and we want them tagged consistently.
    $m = $Message.TrimStart()
    if ($m.StartsWith('[APPLY]'))         { return 'apply' }
    if ($m.StartsWith('[SKIP]'))          { return 'skip' }
    if ($m.StartsWith('[NOT FOUND]'))     { return 'notFound' }
    if ($m.StartsWith('[VERIFIED]'))      { return 'verifyPass' }
    if ($m.StartsWith('[VERIFY FAILED]')) { return 'verifyFail' }
    if ($m.StartsWith('[AUTOPILOT]'))     { return 'autopilot' }
    if ($m.StartsWith('[ASYNC]'))         { return 'async' }
    if ($m.StartsWith('[RESTART]'))       { return 'restart' }
    return $Function
}

function _TrackShowEvent {
    param([string]$Function, [string]$Message)
    $env_ = $global:_CurrentModuleTelemetry
    if ($null -eq $env_) { return }

    try {
        if ($env_.ShowCounts.ContainsKey($Function)) {
            $env_.ShowCounts[$Function]++
        }

        $tag = _GetShowTag -Function $Function -Message $Message
        $redacted = Invoke-TelemetryRedact -Text $Message -Map $env_.RedactMap

        Write-TelemetryEvent -Type "show.$Function" -Data @{
            tag = $tag
            msg = $redacted
        }
    }
    catch { }
}

function Get-TelemetryHostInfo {
    # Cached host info for _meta.json. WMI calls take ~50-100ms each so we
    # only compute once per session. Manufacturer/model are fleet-level
    # (e.g. "ThinkCentre M75q") not customer-identifying, OK to record raw.
    if ($null -ne $global:_TelemetryHostInfo) { return $global:_TelemetryHostInfo }

    $info = [ordered]@{
        os         = [ordered]@{ caption = ""; version = ""; build = "" }
        hardware   = [ordered]@{ manufacturer = ""; model = ""; ram_gb = 0 }
        powershell = ""
    }
    try {
        $os = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($os) {
            $info.os.caption = "$($os.Caption)".Trim()
            $info.os.version = "$($os.Version)".Trim()
            $info.os.build   = "$($os.BuildNumber)".Trim()
        }
    } catch { }
    try {
        $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($cs) {
            $info.hardware.manufacturer = "$($cs.Manufacturer)".Trim()
            $info.hardware.model        = "$($cs.Model)".Trim()
            if ($cs.TotalPhysicalMemory) {
                $info.hardware.ram_gb = [int]([math]::Round([double]$cs.TotalPhysicalMemory / 1GB))
            }
        }
    } catch { }
    try { $info.powershell = "$($PSVersionTable.PSVersion)" } catch { }

    $global:_TelemetryHostInfo = $info
    return $info
}

function _WriteTelemetryMeta {
    # Idempotent: writes _meta.json for current SessionID if not yet present.
    # Called by Start-ModuleTelemetry and Write-KernelTelemetryEvent so the
    # session-level metadata is written by whichever telemetry write fires
    # first.
    $sessionId = $script:SessionID
    if ([string]::IsNullOrWhiteSpace($sessionId)) { return }

    $sessionDir = Join-Path ".\logs\telemetry" $sessionId
    try {
        if (-not (Test-Path $sessionDir)) {
            New-Item -ItemType Directory -Path $sessionDir -Force | Out-Null
        }
        $metaPath = Join-Path $sessionDir "_meta.json"
        if (Test-Path $metaPath) { return }

        $null = Get-TelemetrySalt
        $kernelVer = ""
        try { $kernelVer = (Get-Content ".\kernel\KERNEL_VERSION" -Raw -ErrorAction Stop).Trim() } catch { }

        $meta = [ordered]@{
            telemetrySchemaVersion = 1
            sessionId       = $sessionId
            kernelVersion   = $kernelVer
            startedAt       = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss.fffzzz")
            redactionPolicy = "hash-and-redact-v1"
            saltDigest      = $global:_TelemetrySaltDigest
            host            = (Get-TelemetryHostInfo)
        }
        $metaJson = ([PSCustomObject]$meta | ConvertTo-Json -Depth 6)
        [System.IO.File]::WriteAllText($metaPath, $metaJson, $global:_TelemetryUtf8NoBom)
    } catch { }
}

function Write-KernelTelemetryEvent {
    # Session-level event channel (parallel to per-module envelopes).
    # Emits to logs/telemetry/{SessionID}/_kernel.jsonl. Used for
    # session-lifecycle events (profile.start/end, restart.invoked,
    # resume.consumed, finalize.start/end). Best-effort, never affects
    # kernel behavior.
    param(
        [Parameter(Mandatory)][string]$Type,
        [hashtable]$Data = @{}
    )

    if ($global:_TelemetryWriting) { return }

    $sessionId = $script:SessionID
    if ([string]::IsNullOrWhiteSpace($sessionId)) { return }

    _WriteTelemetryMeta

    $sessionDir = Join-Path ".\logs\telemetry" $sessionId
    $kernelPath = Join-Path $sessionDir "_kernel.jsonl"

    $global:_TelemetryWriting = $true
    try {
        $line = [ordered]@{
            ts   = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss.fffzzz")
            type = $Type
        }
        foreach ($k in $Data.Keys) { $line[$k] = $Data[$k] }
        $json = ([PSCustomObject]$line | ConvertTo-Json -Depth 6 -Compress)
        [System.IO.File]::AppendAllText($kernelPath, $json + "`n", $global:_TelemetryUtf8NoBom)
    } catch { }
    finally {
        $global:_TelemetryWriting = $false
    }
}

function Start-ModuleTelemetry {
    param(
        [Parameter(Mandatory)][string]$ModuleName,
        [int]$Order = 0,
        [string]$Segment = "",
        [string]$ErrorMode = "",
        [string]$Group = "",
        [bool]$IsAsync = $false
    )

    # Defensive: if a prior envelope leaked (caller forgot finally), close it.
    if ($null -ne $global:_CurrentModuleTelemetry) {
        try { Complete-ModuleTelemetry -Status 'Cancelled' -Message 'envelope superseded' } catch { }
    }

    $sessionId = $script:SessionID
    if ([string]::IsNullOrWhiteSpace($sessionId)) { return }

    $sessionDir = Join-Path ".\logs\telemetry" $sessionId
    $modulesDir = Join-Path $sessionDir "modules"

    try {
        if (-not (Test-Path $modulesDir)) {
            New-Item -ItemType Directory -Path $modulesDir -Force | Out-Null
        }
        _WriteTelemetryMeta
    }
    catch {
        # Cannot prepare directory: telemetry disabled for this envelope.
        $global:_CurrentModuleTelemetry = $null
        return
    }

    $global:_TelemetryModuleSeq++
    $seq = $global:_TelemetryModuleSeq
    $safeName = ($ModuleName -replace '[^A-Za-z0-9_\-\.]', '_')
    if ([string]::IsNullOrWhiteSpace($safeName)) { $safeName = 'unknown' }
    $seqStr = "{0:D4}" -f $seq
    $jsonlPath = Join-Path $modulesDir "${seqStr}_${safeName}.jsonl"

    $envelope = @{
        Path       = $jsonlPath
        RedactMap  = (New-TelemetryRedactMap)
        Module     = $ModuleName
        Order      = $Order
        Sequence   = $seq
        StartTime  = (Get-Date)
        ShowCounts = @{ info=0; success=0; warning=0; error=0; skip=0 }
    }
    $global:_CurrentModuleTelemetry = $envelope

    # Scope $Error to this envelope so envelope.end captures only its own
    # entries (catch-and-swallowed exceptions included).
    try { $global:Error.Clear() } catch { }

    # Profile execution context (set by Invoke-BatchExecution per module).
    # Cross-module dependency analysis: AI can see "module X failed when
    # prevModule was Skipped" patterns.
    $envFields = [ordered]@{
        module    = $ModuleName
        sequence  = $seq
        order     = $Order
        segment   = $Segment
        errorMode = $ErrorMode
        group     = $Group
        isAsync   = $IsAsync
    }
    $ctx = $global:_FabriqCurrentProfileContext
    if ($null -ne $ctx) {
        if (-not [string]::IsNullOrWhiteSpace($ctx.ProfileName))   { $envFields.profileName    = $ctx.ProfileName }
        if ($null -ne $ctx.ProfileOrder -and $ctx.ProfileOrder -gt 0) { $envFields.profileOrder = [int]$ctx.ProfileOrder }
        if (-not [string]::IsNullOrWhiteSpace($ctx.ExecutionMode)) { $envFields.executionMode  = $ctx.ExecutionMode }
        if (-not [string]::IsNullOrWhiteSpace($ctx.PrevModuleName)) {
            $envFields.prevModuleName   = $ctx.PrevModuleName
            $envFields.prevModuleStatus = $ctx.PrevModuleStatus
        }
    }

    Write-TelemetryEvent -Type "envelope.start" -Data $envFields
}

function Complete-ModuleTelemetry {
    param(
        [string]$Status = "",
        $Verified = $null,
        [string]$Message = ""
    )

    $env_ = $global:_CurrentModuleTelemetry
    if ($null -eq $env_) { return }

    try {
        # Capture every $Error entry that accumulated during this envelope.
        $errorRecords = @()
        try { $errorRecords = @($global:Error) } catch { }
        foreach ($er in $errorRecords) {
            if ($null -eq $er) { continue }
            try {
                $msg = ""
                $stack = ""
                $cat = ""
                $tgt = ""
                $errType = ""
                $hr = $null
                try { $msg     = Invoke-TelemetryRedact -Text $er.Exception.Message  -Map $env_.RedactMap } catch { }
                try { $stack   = Invoke-TelemetryRedact -Text $er.ScriptStackTrace   -Map $env_.RedactMap } catch { }
                try { $cat     = Invoke-TelemetryRedact -Text "$($er.CategoryInfo)"  -Map $env_.RedactMap } catch { }
                try { $tgt     = Invoke-TelemetryRedact -Text "$($er.TargetObject)"  -Map $env_.RedactMap } catch { }
                try { $errType = $er.Exception.GetType().FullName } catch { }
                try { $hr      = $er.Exception.HResult } catch { }

                Write-TelemetryEvent -Type "error" -Data ([ordered]@{
                    errorType    = $errType
                    hresult      = $hr
                    msg          = $msg
                    scriptStack  = $stack
                    categoryInfo = $cat
                    targetObject = $tgt
                })
            } catch { }
        }

        $duration = (Get-Date) - $env_.StartTime
        Write-TelemetryEvent -Type "envelope.end" -Data ([ordered]@{
            status      = $Status
            verified    = $Verified
            message     = (Invoke-TelemetryRedact -Text $Message -Map $env_.RedactMap)
            duration_ms = [int]$duration.TotalMilliseconds
            errorCount  = $errorRecords.Count
            showCounts  = $env_.ShowCounts
        })
    }
    catch { }
    finally {
        $global:_CurrentModuleTelemetry = $null
    }
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

function Write-ArtPulse {
    $global:ArtPulseCounter++
    try {
        [System.IO.File]::WriteAllText(
            (Join-Path (Get-Location) $global:ArtPulseFilePath),
            $global:ArtPulseCounter.ToString()
        )
    }
    catch { }
}

function Show-Info {
    param([string]$Message)
    Write-Host "[INFO] $Message" -ForegroundColor Cyan
    Write-ArtPulse
    _TrackShowEvent -Function 'info' -Message $Message
}

function Show-Success {
    param([string]$Message)
    Write-Host "[SUCCESS] $Message" -ForegroundColor Green
    Write-ArtPulse
    _TrackShowEvent -Function 'success' -Message $Message
}

function Show-Warning {
    param([string]$Message)
    Write-Host "[WARNING] $Message" -ForegroundColor Yellow
    Write-ArtPulse
    _TrackShowEvent -Function 'warning' -Message $Message
}

function Show-Error {
    param([string]$Message)
    Write-Host "[ERROR] $Message" -ForegroundColor Red
    Write-ArtPulse
    _TrackShowEvent -Function 'error' -Message $Message
}

function Show-Skip {
    param([string]$Message)
    Write-Host "[SKIP] $Message" -ForegroundColor DarkGray
    Write-ArtPulse
    _TrackShowEvent -Function 'skip' -Message $Message
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
        [array]$Details = @(),
        [Nullable[bool]]$Verified = $null
    )
    $resultObj = [PSCustomObject]@{
        _IsModuleResult = $true
        Status          = $Status
        Message         = $Message
        Details         = $Details
        Verified        = $Verified
        Timestamp       = Get-Date
    }
    # Also stash on a global as a fallback for pipeline capture failures
    $global:_LastModuleResult = $resultObj
    return $resultObj
}

# ========================================
# Pattern Layer Functions
# ========================================
# Reusable building blocks shared by every module
# New-BatchResult       : aggregate counts -> Show result -> return ModuleResult
# Confirm-ModuleExecution: prompt for confirmation, handle cancel
# Import-ModuleCsv      : load CSV, filter Enabled, validate required columns

function New-BatchResult {
    param(
        [int]$Success = 0,
        [int]$Skip = 0,
        [int]$Fail = 0,
        [string]$Title = "Execution Results",
        [string]$MessageSuffix = "",
        [Nullable[bool]]$Verified = $null
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
    if ($null -ne $Verified) {
        if ($Verified) {
            Write-Host "  Verified: PASS" -ForegroundColor Green
        } else {
            Write-Host "  Verified: FAIL" -ForegroundColor Red
        }
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

    return (New-ModuleResult -Status $status -Message $msg -Verified $Verified)
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

# ========================================
# DPAPI Passphrase Protection (Resume)
# ========================================

function Protect-PassphraseForResume {
    <#
    .SYNOPSIS
        Encrypts a passphrase via DPAPI (LocalMachine) and returns
        a Base64 string. Used for safe storage in resume_state.json.
    #>
    param([Parameter(Mandatory)][string]$Passphrase)

    Add-Type -AssemblyName System.Security -ErrorAction Stop
    $plainBytes = [System.Text.Encoding]::UTF8.GetBytes($Passphrase)
    $encryptedBytes = [System.Security.Cryptography.ProtectedData]::Protect(
        $plainBytes, $null,
        [System.Security.Cryptography.DataProtectionScope]::LocalMachine
    )
    return [Convert]::ToBase64String($encryptedBytes)
}

function Unprotect-PassphraseFromResume {
    <#
    .SYNOPSIS
        Decrypts the Base64 string produced by Protect-PassphraseForResume
        via DPAPI (LocalMachine) and returns the plaintext passphrase.
    #>
    param([Parameter(Mandatory)][string]$ProtectedBase64)

    Add-Type -AssemblyName System.Security -ErrorAction Stop
    $encryptedBytes = [Convert]::FromBase64String($ProtectedBase64)
    $plainBytes = [System.Security.Cryptography.ProtectedData]::Unprotect(
        $encryptedBytes, $null,
        [System.Security.Cryptography.DataProtectionScope]::LocalMachine
    )
    return [System.Text.Encoding]::UTF8.GetString($plainBytes)
}

# ========================================
# Encryption / Decryption (AES-256-CBC)
# ========================================

function Unprotect-FabriqValue {
    <#
    .SYNOPSIS
        Decrypts an "ENC:"-prefixed AES-256-CBC ciphertext value.
    .DESCRIPTION
        Algorithm spec (must match the C# CryptoPoC exactly):
          Key derivation : PBKDF2-HMAC-SHA256, 100,000 iterations, fixed salt
          Cipher         : AES-256-CBC with PKCS7 padding
          Encoding       : UTF-8 (plaintext), Base64 (ciphertext)
    .PARAMETER EncryptedValue
        Ciphertext value of the form "ENC:<Base64-string>".
        Values without the ENC: prefix are returned unchanged.
    .PARAMETER Passphrase
        Master passphrase used for PBKDF2 key derivation.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$EncryptedValue,
        [Parameter(Mandatory)]
        [string]$Passphrase
    )

    # No ENC: prefix => treat the value as plaintext and return as-is
    if (-not $EncryptedValue.StartsWith('ENC:')) {
        return $EncryptedValue
    }
    $base64 = $EncryptedValue.Substring(4)

    # -- Shared parameters (must match the C# CryptoPoC exactly) --
    $salt       = [System.Text.Encoding]::UTF8.GetBytes("fabriq-fixed-salt-2024")
    $iterations = 100000
    $keySize    = 32   # AES-256
    $ivSize     = 16   # AES block size

    # PBKDF2-HMAC-SHA256 key derivation
    $kdf = New-Object System.Security.Cryptography.Rfc2898DeriveBytes(
        $Passphrase, $salt, $iterations,
        [System.Security.Cryptography.HashAlgorithmName]::SHA256
    )
    $key = $kdf.GetBytes($keySize)
    $iv  = $kdf.GetBytes($ivSize)
    $kdf.Dispose()

    # AES-256-CBC decryption
    $aes         = [System.Security.Cryptography.Aes]::Create()
    $aes.Mode    = [System.Security.Cryptography.CipherMode]::CBC
    $aes.Padding = [System.Security.Cryptography.PaddingMode]::PKCS7
    $aes.Key     = $key
    $aes.IV      = $iv

    $cipherBytes = [Convert]::FromBase64String($base64)
    $decryptor   = $aes.CreateDecryptor()
    $ms          = New-Object System.IO.MemoryStream(, $cipherBytes)
    $cs          = New-Object System.Security.Cryptography.CryptoStream(
                       $ms, $decryptor,
                       [System.Security.Cryptography.CryptoStreamMode]::Read)
    $sr          = New-Object System.IO.StreamReader($cs, [System.Text.Encoding]::UTF8)
    $plainText   = $sr.ReadToEnd()

    $sr.Dispose(); $cs.Dispose(); $ms.Dispose(); $decryptor.Dispose(); $aes.Dispose()
    return $plainText
}

function Test-MasterPassphrase {
    param(
        [Parameter(Mandatory)][string]$Passphrase,
        [Parameter(Mandatory)][string]$VerifyTokenPath
    )
    $VERIFY_PLAINTEXT = "surkitinisme"

    $token = (Get-Content -Path $VerifyTokenPath -Raw -ErrorAction Stop).Trim()
    if ([string]::IsNullOrWhiteSpace($token) -or -not $token.StartsWith('ENC:')) {
        Show-Warning "Verification token file is invalid."
        return $false
    }
    try {
        $decrypted = Unprotect-FabriqValue -EncryptedValue $token -Passphrase $Passphrase
        return ($decrypted -eq $VERIFY_PLAINTEXT)
    }
    catch {
        return $false
    }
}

function Import-ModuleCsv {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [switch]$FilterEnabled,
        [string[]]$RequiredColumns,
        [string]$Segment = $env:FABRIQ_SEGMENT
    )

    $allItems = Import-CsvSafe -Path $Path -Description ([System.IO.Path]::GetFileName($Path))
    if ($null -eq $allItems) { return $null }
    if ($allItems.Count -eq 0) { return $null }

    # Transparent decryption: decrypt ENC: prefixed values if master passphrase is available
    if (-not [string]::IsNullOrWhiteSpace($global:FabriqMasterPassphrase)) {
        foreach ($item in $allItems) {
            foreach ($prop in $item.PSObject.Properties) {
                if ($prop.Value -is [string] -and $prop.Value.StartsWith('ENC:')) {
                    try {
                        $prop.Value = Unprotect-FabriqValue -EncryptedValue $prop.Value -Passphrase $global:FabriqMasterPassphrase
                    }
                    catch {
                        Show-Warning "Failed to decrypt field '$($prop.Name)' in $([System.IO.Path]::GetFileName($Path)): $_"
                    }
                }
            }
        }
    }

    if ($RequiredColumns) {
        if (-not (Test-CsvColumns -CsvData $allItems -RequiredColumns $RequiredColumns -CsvName ([System.IO.Path]::GetFileName($Path)))) {
            return $null
        }
    }

    $totalCount = @($allItems).Count

    if ($FilterEnabled) {
        $filtered = @($allItems | Where-Object { $_.Enabled -eq "1" })
        if ($filtered.Count -eq 0) {
            Show-Skip "No enabled entries in $([System.IO.Path]::GetFileName($Path))"
            return @()
        }
        $allItems = $filtered
    }

    # Segment filtering: strict match (empty matches empty, value matches value)
    $csvColumns = $allItems[0].PSObject.Properties.Name
    if ('Segment' -in $csvColumns) {
        $effectiveSegment = if ([string]::IsNullOrWhiteSpace($Segment)) { "" } else { $Segment.Trim() }
        $beforeCount = $allItems.Count
        $allItems = @($allItems | Where-Object {
            $rowSegment = if ([string]::IsNullOrWhiteSpace($_.Segment)) { "" } else { $_.Segment.Trim() }
            $rowSegment -eq $effectiveSegment
        })
        if (-not [string]::IsNullOrWhiteSpace($effectiveSegment)) {
            Show-Info "Segment filter [$effectiveSegment]: $($allItems.Count) of $beforeCount entries matched"
        }
        if ($allItems.Count -eq 0) {
            $segLabel = if ($effectiveSegment -eq "") { "(default)" } else { "'$effectiveSegment'" }
            Show-Skip "No entries matched Segment $segLabel in $([System.IO.Path]::GetFileName($Path))"
            return @()
        }
    }

    if ($FilterEnabled) {
        Show-Info "Loaded $($allItems.Count) enabled entries (total: $totalCount)"
    }

    # Telemetry: structured csv.load event for AI corpus (decision-trace).
    # Captures file structural metadata only (no row values). Best-effort,
    # never affects load outcome.
    try {
        if ($null -ne $global:_CurrentModuleTelemetry) {
            Write-TelemetryEvent -Type "csv.load" -Data ([ordered]@{
                fileName      = [System.IO.Path]::GetFileName($Path)
                path          = $Path
                totalRows     = $totalCount
                returnedRows  = @($allItems).Count
                filterEnabled = [bool]$FilterEnabled
                segment       = $Segment
                columns       = @($csvColumns)
            })
        }
    } catch { }

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

    # AutoConfirm: Flex single-execution short-circuit. Mirrors AutoPilot's
    # Y/N suppression but stays out of AutoPilot's other paths (no
    # inter-module wait, no ErrorMode retry dialog).
    if ($global:AutoConfirmMode) {
        Write-Host "[AUTOCONFIRM] $Message -> Y (auto)" -ForegroundColor DarkCyan
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

    # AutoPilot or AutoConfirm: skip wait so unattended / one-click
    # flows do not block on Press-Enter.
    if ($global:AutoPilotMode -or $global:AutoConfirmMode) {
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

function Wait-SystemReady {
    # ========================================
    # Waits until required Windows services are running
    # and (optionally) network is reachable.
    # Designed for post-reboot AutoPilot resume.
    # Always exits after MaxWaitSec to prevent hangs.
    # ========================================
    param(
        [int]$MaxWaitSec        = 120,
        [string[]]$RequiredServices = @("LanmanWorkstation", "Dnscache"),
        [string]$NetworkTarget  = "",
        [int]$CheckIntervalSec  = 5
    )

    Show-Info "Waiting for system readiness (max ${MaxWaitSec}s)..."
    $deadline = (Get-Date).AddSeconds($MaxWaitSec)

    while ((Get-Date) -lt $deadline) {
        $notReady = @()

        foreach ($svcName in $RequiredServices) {
            $svc = Get-Service -Name $svcName -ErrorAction SilentlyContinue
            if ($null -eq $svc -or $svc.Status -ne "Running") {
                $notReady += $svcName
            }
        }

        if ($notReady.Count -eq 0 -and $NetworkTarget -ne "") {
            $reachable = Test-Connection -ComputerName $NetworkTarget -Count 1 `
                            -Quiet -ErrorAction SilentlyContinue
            if (-not $reachable) {
                $notReady += "Network:$NetworkTarget"
            }
        }

        if ($notReady.Count -eq 0) {
            Show-Success "System ready"
            return
        }

        Show-Info "Not ready yet: $($notReady -join ', ') — retrying in ${CheckIntervalSec}s..."
        Start-Sleep -Seconds $CheckIntervalSec
    }

    Show-Warning "System readiness timeout (${MaxWaitSec}s). Proceeding anyway."
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
# Evidence Base Path
# ========================================

function Initialize-EvidenceBasePath {
    # ========================================
    # Generates the unified evidence parent directory path and sets
    # $global:FabriqEvidenceBasePath and $env:FABRIQ_EVIDENCE_BASE.
    #
    # Format: {SessionTimestamp}_{PCName}_{SerialNumber}_evidence
    # Example: 2026_03_12_143025_NEW-PC-01_ABC123XYZ_evidence
    #
    # Must be called AFTER host selection completes, since it
    # depends on $env:SELECTED_NEW_PCNAME.
    # ========================================

    # ---- Component values with fallbacks ----
    $ts     = if (-not [string]::IsNullOrWhiteSpace($global:FabriqSessionTimestamp)) {
                  $global:FabriqSessionTimestamp
              } else {
                  Get-Date -Format "yyyy_MM_dd_HHmmss"
              }

    $pcName = if (-not [string]::IsNullOrWhiteSpace($env:SELECTED_NEW_PCNAME)) {
                  $env:SELECTED_NEW_PCNAME
              } else {
                  "Unknown_PC"
              }

    $sn     = if (-not [string]::IsNullOrWhiteSpace($global:FabriqUniqueId)) {
                  $global:FabriqUniqueId
              } else {
                  "Unknown_SN"
              }

    # ---- Sanitize: replace Windows path-invalid characters with hyphen ----
    $pcName = ($pcName -replace '[\\/:*?"<>|]', '-').Trim('-')
    $sn     = ($sn     -replace '[\\/:*?"<>|]', '-').Trim('-')

    # Final safety: if sanitized to empty, use fallback
    if ([string]::IsNullOrWhiteSpace($pcName)) { $pcName = "Unknown_PC" }
    if ([string]::IsNullOrWhiteSpace($sn))     { $sn     = "Unknown_SN" }

    # ---- Build path ----
    # Structure: .\evidence\{name}_evidence\evidence\{category}\
    # Matches log_uploader output so evidence manager can read both
    $dirName  = "${ts}_${pcName}_${sn}_evidence"
    $rootPath = Join-Path ".\evidence" $dirName
    $basePath = Join-Path $rootPath "evidence"

    $global:FabriqEvidenceBasePath = $basePath
    $global:FabriqEvidenceRootPath = $rootPath
    $env:FABRIQ_EVIDENCE_BASE     = $basePath

    Show-Info "Evidence base path: $basePath"
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

# ========================================
# Telemetry: Verbose Stream Capture (standard deployment, default ON)
# ========================================
# Standard deployment ships with kernel/json/verbose_capture.flag present
# (tracked in git). Presence = enable; deletion = opt-out escape hatch.
#
# When active, Invoke-SafeCommand wraps module execution with
# $VerbosePreference='Continue' + $PSDefaultParameterValues['*:Verbose']=$true
# + 4>&1 stream redirect so cmdlets emit verbose records into the pipeline,
# where they are filtered out and converted to cmdlet.verbose telemetry
# events. ModuleResult / Show-* host output are unaffected (Show-* uses
# Information stream / Write-Host, not affected by 4>&1).
#
# Trade-offs:
# - Verbose records contain cmdlet args including potential customer values;
#   the existing redact map is applied before writing to JSONL.
# - No process-restart needed — VerbosePreference and PSDefaultParameterValues
#   are evaluated at runtime per cmdlet call.
# - Output is local-only: log_uploader excludes logs/telemetry/, so verbose
#   data never leaves the kitting PC. Customers wanting zero on-disk telemetry
#   can delete the flag file.
# ========================================
$global:FabriqVerboseCaptureActive = $false

function Enable-FabriqVerboseCapture {
    # Standard deployment: kernel/json/verbose_capture.flag is shipped present.
    # Operator/customer can delete the flag to opt out of cmdlet.verbose
    # capture (escape hatch — fabriq runs in pre-P5 telemetry mode).
    $flagPath = ".\kernel\json\verbose_capture.flag"
    if (-not (Test-Path $flagPath)) { return $false }
    $global:FabriqVerboseCaptureActive = $true
    Show-Info "Verbose stream capture active (cmdlet.verbose telemetry events)"
    return $true
}

function Disable-FabriqVerboseCapture {
    if (-not $global:FabriqVerboseCaptureActive) { return }
    $global:FabriqVerboseCaptureActive = $false
}

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
        Verified  = $null
    }

    # Telemetry envelope (best-effort; never affects outcome)
    try { Start-ModuleTelemetry -ModuleName $OperationName -IsAsync $false } catch { }

    # Verbose stream capture context (dev trial). When active, save/restore
    # both $VerbosePreference (controls Write-Verbose host display) and
    # $PSDefaultParameterValues['*:Verbose'] (forces built-in cmdlets'
    # ShouldProcess machinery to emit "Performing the operation..." verbose
    # — necessary because $VerbosePreference='Continue' alone does NOT
    # trigger this; only an explicit -Verbose param does, which we synthesize
    # via PSDefaultParameterValues).
    $verboseCaptureOn      = [bool]$global:FabriqVerboseCaptureActive
    $verbosePrefSaved      = $null
    $psdpvHadVerbose       = $false
    $psdpvVerboseSaved     = $null
    if ($verboseCaptureOn) {
        $verbosePrefSaved = $VerbosePreference
        $VerbosePreference = 'Continue'
        if ($null -eq $global:PSDefaultParameterValues) {
            $global:PSDefaultParameterValues = @{}
        }
        if ($global:PSDefaultParameterValues.ContainsKey('*:Verbose')) {
            $psdpvHadVerbose = $true
            $psdpvVerboseSaved = $global:PSDefaultParameterValues['*:Verbose']
        }
        $global:PSDefaultParameterValues['*:Verbose'] = $true
    }

    try {
        # Clear the global fallback before invocation
        $global:_LastModuleResult = $null

        if ($verboseCaptureOn) {
            # 4>&1 redirects the verbose stream into stream 1; ForEach-Object
            # filters VerboseRecord objects out and routes them to telemetry,
            # passing other output (incl. ModuleResult) through unchanged.
            $output = & $ScriptBlock 4>&1 | ForEach-Object {
                if ($_ -is [System.Management.Automation.VerboseRecord]) {
                    try {
                        $tEnv = $global:_CurrentModuleTelemetry
                        if ($null -ne $tEnv) {
                            $vMsg = Invoke-TelemetryRedact -Text "$($_.Message)" -Map $tEnv.RedactMap
                            Write-TelemetryEvent -Type 'cmdlet.verbose' -Data @{ msg = $vMsg }
                        }
                    } catch { }
                    # Filter out: don't emit downstream
                } else {
                    $_
                }
            }
        } else {
            $output = & $ScriptBlock
        }

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
            # Trust the status the module reported about itself
            $result.Status = $moduleResult.Status
            $result.Message = $moduleResult.Message
            $result.Success = ($moduleResult.Status -eq "Success")
            $result.Verified = $moduleResult.Verified
        }
        else {
            # Legacy path: module did not return a ModuleResult (all modules migrated)
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
        # Restore $VerbosePreference + $PSDefaultParameterValues before any
        # other cleanup so subsequent telemetry writes don't accidentally
        # emit verbose records that leak past this scope.
        if ($verboseCaptureOn) {
            if ($null -ne $verbosePrefSaved) {
                $VerbosePreference = $verbosePrefSaved
            }
            if ($psdpvHadVerbose) {
                $global:PSDefaultParameterValues['*:Verbose'] = $psdpvVerboseSaved
            } else {
                $global:PSDefaultParameterValues.Remove('*:Verbose') | Out-Null
            }
        }
        $result.Duration = (Get-Date) - $startTime
        # Close telemetry envelope (best-effort)
        try { Complete-ModuleTelemetry -Status $result.Status -Verified $result.Verified -Message $result.Message } catch { }
    }

    return $result
}

# ========================================
# Async Execution (Parallel Path, opt-in via __ASYNC__ marker)
# ========================================
# Invoke-SafeCommandAsync runs a module script in a separate PowerShell
# runspace so that the main thread can monitor for a user-triggered Skip
# flag or a timeout, and forcibly stop the runspace when needed. The
# existing Invoke-SafeCommand remains the default (synchronous) path.
# Config lives in kernel/json/async_config.json. If Enabled=false, callers
# should fall back to Invoke-SafeCommand.
# ========================================

function Get-FabriqAsyncConfig {
    $configPath = ".\kernel\json\async_config.json"
    # Fallback defaults (used when async_config.json is missing or corrupt).
    # DefaultAsync defaults to $false here so that a missing config file
    # cannot silently switch the whole framework into async mode — the
    # shipped config opts in explicitly with "DefaultAsync": true.
    $default = [PSCustomObject]@{
        Enabled           = $true
        DefaultAsync      = $false
        DefaultTimeoutSec = 0
        PollIntervalMs    = 500
        SkipFlagPath      = ".\kernel\json\skip_request.flag"
    }
    if (-not (Test-Path $configPath)) { return $default }
    try {
        $json = Get-Content $configPath -Raw -Encoding UTF8 | ConvertFrom-Json
        return [PSCustomObject]@{
            Enabled           = if ($null -ne $json.Enabled) { [bool]$json.Enabled } else { $default.Enabled }
            DefaultAsync      = if ($null -ne $json.DefaultAsync) { [bool]$json.DefaultAsync } else { $default.DefaultAsync }
            DefaultTimeoutSec = if ($null -ne $json.DefaultTimeoutSec) { [int]$json.DefaultTimeoutSec } else { $default.DefaultTimeoutSec }
            PollIntervalMs    = if ($null -ne $json.PollIntervalMs) { [int]$json.PollIntervalMs } else { $default.PollIntervalMs }
            SkipFlagPath      = if (-not [string]::IsNullOrWhiteSpace($json.SkipFlagPath)) { $json.SkipFlagPath } else { $default.SkipFlagPath }
        }
    }
    catch {
        Show-Warning "Failed to parse async_config.json, using defaults: $_"
        return $default
    }
}

function Invoke-SafeCommandAsync {
    param(
        [Parameter(Mandatory)][string]$ScriptPath,
        [Parameter(Mandatory)][string]$OperationName,
        [int]$TimeoutSec = 0,
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
        Verified  = $null
    }

    # Telemetry envelope (best-effort; never affects outcome).
    # Envelope hashtable is referenced (not copied) when injected into
    # the child runspace, so ShowCounts increments and writes from the
    # child are visible to Complete-ModuleTelemetry in the parent.
    try { Start-ModuleTelemetry -ModuleName $OperationName -IsAsync $true } catch { }

    # Resolve config
    $cfg = Get-FabriqAsyncConfig
    $skipFlagPath = $cfg.SkipFlagPath
    $pollMs = [math]::Max(50, $cfg.PollIntervalMs)
    if ($TimeoutSec -le 0) { $TimeoutSec = $cfg.DefaultTimeoutSec }

    # Normalize skip flag path to absolute so directory changes during a
    # long polling loop cannot make it miss a Skip request from the
    # separate Status Monitor process (which writes an absolute path).
    if (-not [System.IO.Path]::IsPathRooted($skipFlagPath)) {
        $skipFlagPath = Join-Path (Get-Location).Path $skipFlagPath
    }
    $skipFlagPath = [System.IO.Path]::GetFullPath($skipFlagPath)

    # Clear any stale skip flag before starting
    if (Test-Path $skipFlagPath) {
        Remove-Item $skipFlagPath -Force -ErrorAction SilentlyContinue
    }

    # Resolve absolute paths (runspace may not share PWD reliably)
    $commonPath = (Resolve-Path ".\kernel\common.ps1").Path
    $fabriqRoot = (Resolve-Path ".").Path
    $absScript  = (Resolve-Path $ScriptPath -ErrorAction SilentlyContinue)
    if ($null -eq $absScript) {
        $result.Status  = "Error"
        $result.Message = "Script not found: $ScriptPath"
        $result.Duration = (Get-Date) - $startTime
        return $result
    }
    $absScriptPath = $absScript.Path

    # Globals that modules depend on. env: vars are Process-scoped so they
    # are automatically visible in the child runspace; only $global:* needs
    # to be injected explicitly.
    $inject = @{
        FabriqMasterPassphrase = $global:FabriqMasterPassphrase
        AutoPilotMode          = $global:AutoPilotMode
        AutoConfirmMode        = $global:AutoConfirmMode
        AutoPilotWaitSec       = $global:AutoPilotWaitSec
        FabriqTranscriptPath   = $global:FabriqTranscriptPath
        FabriqUniqueId         = $global:FabriqUniqueId
        FabriqSessionTimestamp = $global:FabriqSessionTimestamp
        FabriqEvidenceBasePath = $global:FabriqEvidenceBasePath
        FabriqEvidenceRootPath = $global:FabriqEvidenceRootPath
        # Telemetry envelope: child runspace's Show-* writes to the same
        # JSONL file via this reference. Hashtable is reference-typed so
        # ShowCounts increments are visible to the parent on EndInvoke.
        _CurrentModuleTelemetry = $global:_CurrentModuleTelemetry
    }

    $runspace = $null
    $ps = $null
    $asyncHandle = $null
    $interrupted = $false
    $interruptReason = $null

    try {
        $runspace = [runspacefactory]::CreateRunspace($Host)
        $runspace.ApartmentState = "STA"
        $runspace.ThreadOptions = "ReuseThread"
        $runspace.Open()

        # Align working directory so relative paths inside the module behave
        # the same as when called synchronously from main.ps1.
        try {
            $runspace.SessionStateProxy.Path.SetLocation($fabriqRoot) | Out-Null
        } catch { }

        $ps = [PowerShell]::Create()
        $ps.Runspace = $runspace

        [void]$ps.AddScript({
            param($CommonPath, $ModuleScript, $Inject, $FabriqRoot)

            Set-Location -Path $FabriqRoot

            # Reload common.ps1 into this runspace so Show-Info /
            # Import-ModuleCsv / New-ModuleResult etc. are available.
            . $CommonPath

            # Overwrite globals initialized by common.ps1 with values from
            # the parent session (passphrase, AutoPilot flags, evidence path).
            foreach ($key in $Inject.Keys) {
                Set-Variable -Name $key -Value $Inject[$key] -Scope Global -Force
            }

            # Fallback buffer for modules that return ModuleResult via
            # $global:_LastModuleResult instead of pipeline.
            $global:_LastModuleResult = $null

            $output = & $ModuleScript

            [PSCustomObject]@{
                _AsyncWrapper = $true
                Output        = $output
                LastResult    = $global:_LastModuleResult
            }
        })
        [void]$ps.AddArgument($commonPath)
        [void]$ps.AddArgument($absScriptPath)
        [void]$ps.AddArgument($inject)
        [void]$ps.AddArgument($fabriqRoot)

        $asyncHandle = $ps.BeginInvoke()

        # Monitor loop: wait for completion, Skip flag, or timeout.
        # Inner sub-sleep pumps WinForms messages at ~16ms so the
        # execution toolbar / FlexProfile dashboard stay smooth during
        # drag while an async module runs. Skip flag is checked inside
        # the inner loop so Skip is honored within ~16ms instead of up
        # to $pollMs. DoEvents is try/catch'd for headless contexts
        # where System.Windows.Forms may not be loaded.
        $uiPumpMs = 16
        while (-not $asyncHandle.IsCompleted) {
            $waitUntil = (Get-Date).AddMilliseconds($pollMs)
            $skipDetected = $false
            while ((Get-Date) -lt $waitUntil) {
                Start-Sleep -Milliseconds $uiPumpMs
                try { [System.Windows.Forms.Application]::DoEvents() } catch { }
                if ($asyncHandle.IsCompleted) { break }
                if (Test-Path $skipFlagPath) { $skipDetected = $true; break }
            }

            if ($skipDetected -or (Test-Path $skipFlagPath)) {
                Remove-Item $skipFlagPath -Force -ErrorAction SilentlyContinue
                $interrupted = $true
                $interruptReason = "Skip"
                try { $ps.Stop() } catch { }
                break
            }

            if ($TimeoutSec -gt 0) {
                $elapsed = ((Get-Date) - $startTime).TotalSeconds
                if ($elapsed -ge $TimeoutSec) {
                    $interrupted = $true
                    $interruptReason = "Timeout"
                    try { $ps.Stop() } catch { }
                    break
                }
            }
        }

        if (-not $interrupted) {
            $wrappedOutput = $null
            try {
                $wrappedOutput = $ps.EndInvoke($asyncHandle)
            }
            catch {
                $result.Status  = "Error"
                $result.Message = "Async runspace error: $($_.Exception.Message)"
                $result.Error   = $_
                return $result
            }

            $wrapper = $null
            foreach ($item in @($wrappedOutput)) {
                if ($item -is [PSCustomObject] -and $item._AsyncWrapper -eq $true) {
                    $wrapper = $item
                    break
                }
            }

            $moduleResult = $null
            if ($wrapper) {
                if ($null -ne $wrapper.Output) {
                    foreach ($out in @($wrapper.Output)) {
                        if ($out -is [PSCustomObject] -and $out._IsModuleResult -eq $true) {
                            $moduleResult = $out
                        }
                    }
                }
                if (-not $moduleResult -and $null -ne $wrapper.LastResult) {
                    $moduleResult = $wrapper.LastResult
                }
            }

            if ($moduleResult) {
                $result.Status   = $moduleResult.Status
                $result.Message  = $moduleResult.Message
                $result.Success  = ($moduleResult.Status -eq "Success")
                $result.Verified = $moduleResult.Verified
            }
            else {
                # Legacy path mirrors Invoke-SafeCommand: no ModuleResult returned
                $result.Status  = "Success"
                $result.Success = $true
                $result.Message = "(async legacy - unverified)"
            }
        }
        else {
            if ($interruptReason -eq "Skip") {
                $result.Status  = "Error"
                $result.Message = "Module skipped by operator (async runspace stopped; system state may be incomplete)"
            }
            elseif ($interruptReason -eq "Timeout") {
                $result.Status  = "Error"
                $result.Message = "Module exceeded timeout of ${TimeoutSec}s (async runspace stopped; system state may be incomplete)"
            }
        }
    }
    catch {
        $result.Status  = "Error"
        $result.Message = "Async execution error: $($_.Exception.Message)"
        $result.Error   = $_
        if (-not $ContinueOnError) {
            Show-Error "$OperationName : $($_.Exception.Message)"
        }
    }
    finally {
        $result.Duration = (Get-Date) - $startTime
        if ($null -ne $ps) {
            try { $ps.Dispose() } catch { }
        }
        if ($null -ne $runspace) {
            try { $runspace.Close() } catch { }
            try { $runspace.Dispose() } catch { }
        }
        # Close telemetry envelope. The child runspace's Show-* events
        # have already been appended to the JSONL by reference; we only
        # write envelope.end + accumulated $Error here.
        try { Complete-ModuleTelemetry -Status $result.Status -Verified $result.Verified -Message $result.Message } catch { }
    }

    return $result
}

function Add-ExecutionResult {
    param(
        [string]$Operation,
        [string]$Status,
        [string]$Message = "",
        [Nullable[bool]]$Verified = $null,
        # Profile CSV row Order this result corresponds to. 0 means
        # "no Profile row" (e.g., [RESTART NOW], log uploader, ad-hoc
        # module runs). Used by FlexProfile dashboard / HTML checklist
        # for per-row state tracking when multiple Profile rows share
        # the same MenuName.
        [int]$Order = 0
    )

    $script:ExecutionResults += [PSCustomObject]@{
        Operation = $Operation
        Status    = $Status
        Message   = $Message
        Timestamp = Get-Date
        Verified  = $Verified
        Order     = $Order
    }

    # Refresh the status monitor
    Write-StatusFile -Phase "executing"
}

function Clear-ExecutionResults {
    # Preserve restored entries and separators across the clear
    $restored = @($script:ExecutionResults | Where-Object {
        $_.IsRestored -eq $true -or $_.Status -eq "Separator"
    })

    if ($restored.Count -gt 0) {
        $script:ExecutionResults = $restored
    }
    else {
        $script:ExecutionResults = @()
    }

    # Refresh the status monitor
    Write-StatusFile -Phase "executing"
}

function Show-ExecutionSummary {
    param(
        [array]$Results = $null,
        [System.TimeSpan]$ElapsedTime = [System.TimeSpan]::Zero
    )

    # Refresh the status monitor (final/complete state)
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
        [string]$Message = "",
        [string]$Verified = "",
        # Profile CSV row Order this entry corresponds to. 0 means
        # "no Profile row" (markers / ad-hoc / log uploader); the CSV
        # cell is left empty in that case. Positive integer = exact
        # Profile row identity, used for per-row state matching.
        # New column added in kernel 3.1.3 — backward compatible
        # (legacy rows without Order column read as Order=0 via
        # Restore-ExecutionHistory).
        [int]$Order = 0
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

    # Order column: emit empty for "no Profile row" (markers / ad-hoc),
    # otherwise the integer value.
    $orderStr = if ($Order -gt 0) { "$Order" } else { "" }

    $line = "$timestamp,$env:SELECTED_KANRI_NO,$env:SELECTED_NEW_PCNAME,$ModuleName,$Category,$Status,$escapedMessage,$windowsUser,$workerName,$mediaSerial,$($script:SessionID),$Verified,$orderStr"

    # Create with header if file does not exist
    $needHeader = -not (Test-Path $script:HistoryPath)

    for ($i = 0; $i -lt $maxRetry; $i++) {
        try {
            if ($needHeader) {
                $header = "Timestamp,KanriNo,PCName,ModuleName,Category,Status,Message,WindowsUser,Worker,MediaSerial,SessionID,Verified,Order"
                $header | Out-File -FilePath $script:HistoryPath -Encoding UTF8 -Force
                $needHeader = $false
            }
            $line | Out-File -FilePath $script:HistoryPath -Append -Encoding UTF8
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
        # Lock-free read with retry
        $data = $null
        for ($retry = 0; $retry -lt 3; $retry++) {
            try {
                $data = @(Import-Csv -Path $script:HistoryPath -Encoding UTF8)
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
                return @(Import-Csv -Path $script:HistoryPath -Encoding UTF8)
            }
            catch {
                Show-Warning "Failed to restore history file"
            }
        }
        return @()
    }
}

function Restore-ExecutionHistory {
    param(
        # When set, restrict the load to entries whose SessionID equals
        # this value, with no row-count limit. Used by Invoke-BatchExecution
        # at the start of each batch (in-batch reload) to refresh the
        # IsRestored set with the current session's history exactly,
        # surviving Clear-ExecutionResults's wipe of non-IsRestored
        # entries between batches. Empty string = legacy behavior
        # (cross-session pull, top 50 most recent, with separator row
        # appended). See KERNEL_API.md S6 internal implementation.
        [string]$SessionIDFilter = ""
    )

    if (-not $env:SELECTED_KANRI_NO) { return }

    $isFilterMode = -not [string]::IsNullOrEmpty($SessionIDFilter)

    try {
        # Filter mode: no limit (current session is naturally bounded).
        # Legacy mode: top 50 most recent for display continuity.
        $limit = if ($isFilterMode) { 0 } else { 50 }
        [array]$history = @(Import-ExecutionHistory -FilterKanriNo $env:SELECTED_KANRI_NO -Limit $limit)

        if ($isFilterMode) {
            $history = @($history | Where-Object { $_.SessionID -eq $SessionIDFilter })
        }

        if ($history.Count -eq 0) {
            if ($isFilterMode) {
                # Filter mode + 0 hits = first batch of a new session.
                # Explicitly REPLACE ExecutionResults with empty so any
                # cross-session IsRestored entries from session-start
                # Restore are evicted (otherwise stale prev-session
                # entries with matching MenuName / Order could pollute
                # current Profile's HTML rendering).
                $script:ExecutionResults = @()
            }
            else {
                Show-Info "No previous execution history for AdminID $env:SELECTED_KANRI_NO"
            }
            return
        }

        # Import-ExecutionHistory returns descending; flip to ascending
        [array]::Reverse($history)

        $restoredResults = @()
        foreach ($entry in $history) {
            if ([string]::IsNullOrEmpty($entry.ModuleName)) { continue }

            $ts = $null
            try { $ts = [datetime]::ParseExact($entry.Timestamp, "yyyy-MM-dd HH:mm:ss", $null) }
            catch { $ts = Get-Date }

            # Restore Verified field from CSV ("True"/"False"/"" -> bool/$null)
            $restoredVerified = $null
            if (-not [string]::IsNullOrWhiteSpace($entry.Verified)) {
                $restoredVerified = ($entry.Verified -eq "True")
            }

            # Order column may be missing (legacy CSV pre-3.1.3) or
            # empty (markers / ad-hoc rows). Default to 0 = "no Profile
            # row association"; non-zero = exact Profile row identity
            # for per-Order matching.
            $restoredOrder = 0
            $hasOrderCol = $entry.PSObject.Properties.Name -contains 'Order'
            if ($hasOrderCol -and -not [string]::IsNullOrWhiteSpace($entry.Order)) {
                try { $restoredOrder = [int]$entry.Order } catch { $restoredOrder = 0 }
            }

            $restoredResults += [PSCustomObject]@{
                Operation  = $entry.ModuleName
                Status     = $entry.Status
                Message    = $entry.Message
                Timestamp  = $ts
                IsRestored = $true
                SessionID  = $entry.SessionID
                Verified   = $restoredVerified
                Order      = $restoredOrder
            }
        }

        if ($restoredResults.Count -gt 0) {
            if (-not $isFilterMode) {
                # Boundary separator between restored history and the current session
                $restoredResults += [PSCustomObject]@{
                    Operation  = "--- Current Session ---"
                    Status     = "Separator"
                    Message    = ""
                    Timestamp  = Get-Date
                    IsRestored = $false
                    SessionID  = $script:SessionID
                    Order      = 0
                }

                $script:ExecutionResults = $restoredResults
                Show-Success "Restored $($restoredResults.Count - 1) history entries"
            }
            else {
                # Filter mode: silent replace — the in-batch reload is a
                # mechanical state refresh, no UX message warranted.
                $script:ExecutionResults = $restoredResults
            }
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

    # Copy to evidence directory
    if (-not [string]::IsNullOrWhiteSpace($global:FabriqEvidenceBasePath)) {
        $evidenceExportDir = Join-Path $global:FabriqEvidenceBasePath "export_history"
    }
    else {
        $evidenceExportDir = ".\evidence\export_history"
    }
    if (-not (Test-Path $evidenceExportDir)) {
        $null = New-Item -ItemType Directory -Path $evidenceExportDir -Force
    }

    $evidenceDateStr = Get-Date -Format "yyyy_MM_dd_HHmmss"
    if (-not [string]::IsNullOrWhiteSpace($global:FabriqEvidenceBasePath)) {
        # The evidence directory name already encodes UID+PCName,
        # so the export filename only needs the timestamp.
        $evidenceExportPath = Join-Path $evidenceExportDir "history_export_${evidenceDateStr}.csv"
    }
    else {
        # Fallback path: encode the host identity in the filename instead
        $pcName = if (-not [string]::IsNullOrEmpty($env:SELECTED_NEW_PCNAME)) {
            $env:SELECTED_NEW_PCNAME
        } else {
            $env:COMPUTERNAME
        }
        $uid = if ($global:FabriqUniqueId) { $global:FabriqUniqueId } else { Get-HardwareUniqueId }
        $evidenceExportPath = Join-Path $evidenceExportDir "history_export_${evidenceDateStr}_${uid}_${pcName}.csv"
    }
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
    # Output path
    # ----------------------------------------
    if (-not [string]::IsNullOrWhiteSpace($global:FabriqEvidenceBasePath)) {
        $outputDir = Join-Path $global:FabriqEvidenceBasePath "checklist"
    }
    else {
        $outputDir = ".\evidence\checklist"
    }
    if (-not (Test-Path $outputDir)) {
        $null = New-Item -ItemType Directory -Path $outputDir -Force
    }

    $dateStr = Get-Date -Format "yyyy_MM_dd_HHmmss"
    # $uid / $pcName are also rendered into the HTML body (Meta grid, <title>),
    # so they must be resolved in both branches — not only in the fallback filename path.
    $uid    = if ($global:FabriqUniqueId) { $global:FabriqUniqueId } else { Get-HardwareUniqueId }
    $pcName = if (-not [string]::IsNullOrEmpty($env:SELECTED_NEW_PCNAME)) { $env:SELECTED_NEW_PCNAME } else { $env:COMPUTERNAME }

    if (-not [string]::IsNullOrWhiteSpace($global:FabriqEvidenceBasePath)) {
        $outPath = Join-Path $outputDir "checklist_${dateStr}.html"
    }
    else {
        $outPath = Join-Path $outputDir "checklist_${dateStr}_${uid}_${pcName}.html"
    }

    # ----------------------------------------
    # Session metadata
    # ----------------------------------------
    $workerName  = if ($script:SessionInfo) { $script:SessionInfo.WorkerName }  else { "-" }
    $mediaSerial = if ($script:SessionInfo) { $script:SessionInfo.MediaSerial } else { "-" }
    $kanriNo     = if ($env:SELECTED_KANRI_NO)    { $env:SELECTED_KANRI_NO }    else { "-" }
    $oldPcName   = if ($env:SELECTED_OLD_PCNAME)  { $env:SELECTED_OLD_PCNAME }  else { "-" }
    $generatedAt = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $elapsedStr  = "{0:D2}:{1:D2}:{2:D2}" -f [int][math]::Floor($ElapsedTime.TotalHours), $ElapsedTime.Minutes, $ElapsedTime.Seconds

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
    # Actual PC info (from OS) for verification
    # ----------------------------------------
    $actualPC = Get-CurrentPCInfo

    # --- Network settings comparison ---
    $verifyRows = @()
    $verifyItems = @(
        @{ Label = "PC Name";      Expected = $env:SELECTED_NEW_PCNAME;   Actual = $actualPC.ComputerName }
        @{ Label = "Ethernet IP";   Expected = $env:SELECTED_ETH_IP;       Actual = $actualPC.EthernetIP }
        @{ Label = "Eth Subnet";    Expected = $env:SELECTED_ETH_SUBNET;   Actual = $actualPC.EthernetSubnet }
        @{ Label = "Eth Gateway";   Expected = $env:SELECTED_ETH_GATEWAY;  Actual = $actualPC.EthernetGateway }
        @{ Label = "Wi-Fi IP";      Expected = $env:SELECTED_WIFI_IP;      Actual = $actualPC.WifiIP }
        @{ Label = "Wi-Fi Subnet";  Expected = $env:SELECTED_WIFI_SUBNET;  Actual = $actualPC.WifiSubnet }
        @{ Label = "Wi-Fi Gateway"; Expected = $env:SELECTED_WIFI_GATEWAY; Actual = $actualPC.WifiGateway }
    )
    foreach ($item in $verifyItems) {
        if ([string]::IsNullOrEmpty($item.Expected)) { continue }
        $exp = $item.Expected
        $act = if ([string]::IsNullOrEmpty($item.Actual)) { "(none)" } else { $item.Actual }
        $match = ($exp -eq $act)
        $verifyRows += @{ Label = $item.Label; Expected = $exp; Actual = $act; Match = $match }
    }

    # DNS comparison (set-based, order-independent)
    $expectedDns = @($env:SELECTED_DNS1, $env:SELECTED_DNS2, $env:SELECTED_DNS3, $env:SELECTED_DNS4) |
                   Where-Object { -not [string]::IsNullOrEmpty($_) } | Sort-Object
    $actualDns   = @($actualPC.DNS) | Where-Object { -not [string]::IsNullOrEmpty($_) } | Sort-Object
    if ($expectedDns.Count -gt 0) {
        $expStr = $expectedDns -join ", "
        $actStr = if ($actualDns.Count -gt 0) { $actualDns -join ", " } else { "(none)" }
        $dnsMatch = ($expStr -eq $actStr)
        $verifyRows += @{ Label = "DNS"; Expected = $expStr; Actual = $actStr; Match = $dnsMatch }
    }

    # --- Printer cross-check (3-way) ---
    $expectedPrinterNames = @($printerList | ForEach-Object { $_.Name })
    $actualPrinterNames   = @($actualPC.Printers | ForEach-Object { $_.Name })

    $printerVerifyRows = @()
    # Expected printers: check if installed
    foreach ($ep in $expectedPrinterNames) {
        $installed = $actualPrinterNames -contains $ep
        $printerVerifyRows += @{ Name = $ep; Source = "Expected"; Installed = $installed }
    }
    # Actual printers not in expected list: mark as Extra
    foreach ($ap in $actualPrinterNames) {
        if ($expectedPrinterNames -notcontains $ap) {
            $printerVerifyRows += @{ Name = $ap; Source = "Extra"; Installed = $true }
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
        # Match prefers Order (per-row precision; supports multiple
        # Profile rows sharing the same MenuName). Falls back to
        # MenuName for legacy entries that lack Order (pre-3.1.3
        # history.csv rows). The fallback is STRICT: only entries
        # with Order=0 (legacy / non-Profile) or matching Order are
        # accepted, so an executed sibling row sharing the same
        # MenuName does not leak its status into this row's display.
        $result = $null
        $moduleOrder = if ($null -ne $module.Order) { [int]$module.Order } else { 0 }
        if ($moduleOrder -gt 0) {
            $result = $currentResults | Where-Object {
                ($null -ne $_.Order) -and ([int]$_.Order -eq $moduleOrder)
            } | Select-Object -Last 1
        }
        if ($null -eq $result) {
            $result = $currentResults | Where-Object {
                if ($_.Operation -ne $module.MenuName) { return $false }
                # Accept legacy entry (Order=0/null) or matching-Order entry only.
                $candOrder = if ($null -ne $_.Order) { [int]$_.Order } else { 0 }
                return ($candOrder -eq 0 -or $candOrder -eq $moduleOrder)
            } | Select-Object -Last 1
        }

        $statusLabel = "Not Run"
        $statusClass = "notrun"
        $message     = "-"

        $verifiedLabel = "-"
        $verifiedClass = "notrun"

        if ($null -ne $result) {
            switch ($result.Status) {
                "Success"   { $statusLabel = "OK";      $statusClass = "ok";      $successTotal++ }
                "Partial"   { $statusLabel = "Partial"; $statusClass = "partial"; $successTotal++ }
                "Skipped"   { $statusLabel = "Skip";    $statusClass = "skip";    $skipTotal++ }
                "Skip"      { $statusLabel = "Skip";    $statusClass = "skip";    $skipTotal++ }
                "Cancelled" { $statusLabel = "Cancel";  $statusClass = "skip";    $skipTotal++ }
                "Warning"   { $statusLabel = "Warn";    $statusClass = "partial"; $successTotal++ }
                "Error"     { $statusLabel = "NG";      $statusClass = "ng";      $errorTotal++ }
                # Pending: explicit status set by FlexProfile [Mark as Pending]
                # action. Counts toward notRunTotal so HTML totals reconcile
                # with $DefinedModules.Count (otherwise Pending fell into the
                # default branch with no counter increment, leaving a gap).
                "Pending"   { $statusLabel = "Pending"; $statusClass = "notrun"; $notRunTotal++ }
                default     { $statusLabel = $result.Status; $statusClass = "notrun" }
            }
            $message = if ($result.Message) { [System.Web.HttpUtility]::HtmlEncode($result.Message) } else { "-" }
            $ts      = if ($result.Timestamp) { $result.Timestamp.ToString("HH:mm:ss") } else { "-" }

            # Post-Apply Verification status
            if ($null -ne $result.Verified) {
                if ($result.Verified -eq $true) {
                    $verifiedLabel = "PASS"
                    $verifiedClass = "ok"
                } else {
                    $verifiedLabel = "FAIL"
                    $verifiedClass = "ng"
                }
            }
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
            <td class="col-status"><span class="badge $verifiedClass">$verifiedLabel</span></td>
            <td class="col-time">$ts</td>
            <td class="col-msg">$message</td>
        </tr>
"@
    }

    # Verification failure detection
    $verifyHasFailure = ($verifyRows | Where-Object { -not $_.Match }).Count -gt 0
    $printerHasNotFound = ($printerVerifyRows | Where-Object {
        $_.Source -eq "Expected" -and -not $_.Installed
    }).Count -gt 0
    $printerHasExtra = ($printerVerifyRows | Where-Object { $_.Source -eq "Extra" }).Count -gt 0
    $hasVerifyNG = $verifyHasFailure -or $printerHasNotFound

    $overallClass = if ($errorTotal -gt 0 -or $hasVerifyNG) { "overall-ng" } elseif ($notRunTotal -gt 0) { "overall-partial" } elseif ($printerHasExtra) { "overall-ca" } else { "overall-ok" }
    $overallLabel = if ($errorTotal -gt 0 -or $hasVerifyNG) { "NG" } elseif ($notRunTotal -gt 0) { "Incomplete" } elseif ($printerHasExtra) { "CA" } else { "OK" }

    # ----------------------------------------
    # Build supplemental section HTML
    # ----------------------------------------
    $licProductRow = if ($licenseProduct) {
        "<div class='sysinfo-row'><span class='sysinfo-label'>Product</span><span style='font-size:11px;color:#555;word-break:break-all;'>$([System.Web.HttpUtility]::HtmlEncode($licenseProduct))</span></div>"
    } else { "" }

    # ----------------------------------------
    # Build verification section HTML
    # ----------------------------------------
    $verifyNetSectionHtml = ""
    if ($verifyRows.Count -gt 0) {
        $vNetRows = ""
        foreach ($vr in $verifyRows) {
            $statusBadge = if ($vr.Match) { '<span class="badge ok">Match</span>' } else { '<span class="badge ng">Mismatch</span>' }
            $rowCls = if (-not $vr.Match) { ' class="verify-mismatch"' } else { "" }
            $vNetRows += "        <tr$rowCls><td>$([System.Web.HttpUtility]::HtmlEncode($vr.Label))</td><td>$([System.Web.HttpUtility]::HtmlEncode($vr.Expected))</td><td>$([System.Web.HttpUtility]::HtmlEncode($vr.Actual))</td><td class=`"col-status`">$statusBadge</td></tr>`n"
        }
        $verifyNetSectionHtml = @"
  <div class="section">
    <div class="section-hd">PC Settings Verification</div>
    <table class="verify-table">
      <thead><tr><th>Item</th><th>Expected (hostlist)</th><th>Actual (OS)</th><th>Status</th></tr></thead>
      <tbody>
$vNetRows      </tbody>
    </table>
  </div>
"@
    }

    $verifyPrinterSectionHtml = ""
    if ($printerVerifyRows.Count -gt 0) {
        $vPrtRows = ""
        foreach ($pr in $printerVerifyRows) {
            if ($pr.Source -eq "Expected" -and $pr.Installed) {
                $badge = '<span class="badge ok">Match</span>'
                $rowCls = ""
            } elseif ($pr.Source -eq "Expected" -and -not $pr.Installed) {
                $badge = '<span class="badge ng">Not Found</span>'
                $rowCls = ' class="verify-mismatch"'
            } else {
                $badge = '<span class="badge extra">Extra</span>'
                $rowCls = ' class="verify-extra"'
            }
            $vPrtRows += "        <tr$rowCls><td>$([System.Web.HttpUtility]::HtmlEncode($pr.Name))</td><td class=`"col-status`">$badge</td></tr>`n"
        }
        $verifyPrinterSectionHtml = @"
  <div class="section">
    <div class="section-hd">Printer Verification</div>
    <table class="verify-table">
      <thead><tr><th>Printer Name</th><th>Status</th></tr></thead>
      <tbody>
$vPrtRows      </tbody>
    </table>
  </div>
"@
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
  .overall-ca      { background: #f1c40f; color: #333; }

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

  /* Verification tables */
  .verify-table { width: 100%; border-collapse: collapse; background: #fff; }
  .verify-table thead tr { background: #f5f5f5; }
  .verify-table th { padding: 7px 14px; text-align: left; font-size: 11px; color: #555; font-weight: 600; border-bottom: 1px solid #e0e0e0; }
  .verify-table td { padding: 7px 14px; font-size: 12px; border-bottom: 1px solid #f0f0f0; }
  .verify-table tr:last-child td { border-bottom: none; }
  .verify-mismatch { background: #fff5f5; }
  .verify-extra { background: #fffbf0; }
  .extra { background: #ffeaa7; color: #856404; }

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

$verifyNetSectionHtml
$verifyPrinterSectionHtml

  <!-- Checklist table -->
  <div class="table-wrap">
    <table>
      <thead>
        <tr>
          <th class="col-order">#</th>
          <th class="col-name">Module</th>
          <th class="col-cat">Category</th>
          <th class="col-status">Result</th>
          <th class="col-status">Verified</th>
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

# ========================================
# Profile Completion Pipeline (Linear auto-finalize / [cl] regen / Flex Complete)
# ========================================
# Single source of truth for the post-profile pipeline:
#   1. Export execution history CSV into the evidence directory
#   2. Generate the HTML checklist
#   3. Update $global:FabriqLastProfile{Name,Path,Modules}
#   4. Run log_uploader if log_destinations.csv has enabled rows
#   5. Launch view_report.ps1
#
# Mode preserves two pre-existing behaviors verbatim:
#   'Auto'   : Linear AutoPilot finalize (Invoke-BatchExecution end).
#              Silent log upload, viewer launched AFTER upload, "Auto-..."
#              wording.
#   'Manual' : main.ps1 [cl] / RegenerateChecklist action.
#              Log upload recorded as ExecutionResult / history entry
#              ("Log Upload (cl)"), viewer launched BEFORE upload,
#              "Regenerating..." wording.
# Flex [Complete] (P5/P7) will use 'Manual'.
# ========================================
function Complete-ProfileExecution {
    param(
        [Parameter(Mandatory)][string]$ProfileName,
        [Parameter(Mandatory)][string]$ProfilePath,
        [Parameter(Mandatory)][array]$DefinedModules,
        [TimeSpan]$ElapsedTime = [TimeSpan]::Zero,
        [ValidateSet('Auto','Manual')][string]$Mode = 'Auto'
    )

    $finalizeStart = Get-Date
    try {
        Write-KernelTelemetryEvent -Type "finalize.start" -Data ([ordered]@{
            profileName    = $ProfileName
            mode           = $Mode
            elapsedMs      = [int]$ElapsedTime.TotalMilliseconds
            definedModules = @($DefinedModules).Count
        })
    } catch { }

    Write-Host ""

    # Step 1: Export execution history CSV → evidence directory
    if ($Mode -eq 'Auto') {
        Write-Host "[INFO] Auto-exporting execution history as evidence..." -ForegroundColor Cyan
    } else {
        Show-Info "Regenerating checklist..."
    }
    $null = Export-ExecutionHistory

    # Step 2: Generate HTML checklist
    if ($Mode -eq 'Auto') {
        Write-Host "[INFO] Generating HTML checklist..." -ForegroundColor Cyan
    }
    $checklistPath = Export-HtmlChecklist `
        -ProfileName      $ProfileName `
        -ProfilePath      $ProfilePath `
        -DefinedModules   $DefinedModules `
        -ExecutionResults $script:ExecutionResults `
        -ElapsedTime      $ElapsedTime

    # Step 3: Retain profile info for [cl] regeneration on the same session
    $global:FabriqLastProfileName    = $ProfileName
    $global:FabriqLastProfilePath    = $ProfilePath
    $global:FabriqLastProfileModules = $DefinedModules

    # Step 4 (Manual only): viewer BEFORE upload, preserving [cl] behavior
    if ($Mode -eq 'Manual' -and -not [string]::IsNullOrEmpty($checklistPath) -and (Test-Path $checklistPath)) {
        $viewerScript = ".\kernel\ps1\view_report.ps1"
        if (Test-Path $viewerScript) {
            try { & $viewerScript -HtmlPath $checklistPath } catch { }
        }
    }

    # Step 5: Log uploader (only when log_destinations.csv has enabled rows)
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
            if ($Mode -eq 'Auto') {
                Write-Host ""
                Write-Host "[INFO] Auto-uploading logs and evidence..." -ForegroundColor Cyan
                try {
                    $null = & $logUploaderScript
                }
                catch {
                    Show-Warning "Log upload failed: $($_.Exception.Message)"
                }
            }
            else {
                Show-Info "Uploading updated evidence..."
                $uploadResult = & $logUploaderScript
                if ($null -ne $uploadResult -and $uploadResult._IsModuleResult) {
                    # Log upload is profile-external — Order=0 means
                    # "no Profile row association" (CSV cell empty).
                    Add-ExecutionResult -Operation "Log Upload (cl)" -Status $uploadResult.Status -Message $uploadResult.Message -Order 0
                    $null = Write-ExecutionHistory -ModuleName "Log Upload (cl)" -Category "System" -Status $uploadResult.Status -Message $uploadResult.Message -Order 0
                }
            }
        }
    }

    # Step 6 (Auto only): viewer AFTER upload, preserving Linear finalize behavior
    if ($Mode -eq 'Auto' -and -not [string]::IsNullOrEmpty($checklistPath) -and (Test-Path $checklistPath)) {
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

    try {
        Write-KernelTelemetryEvent -Type "finalize.end" -Data ([ordered]@{
            profileName        = $ProfileName
            mode               = $Mode
            durationMs         = [int]((Get-Date) - $finalizeStart).TotalMilliseconds
            checklistGenerated = (-not [string]::IsNullOrEmpty($checklistPath)) -and (Test-Path $checklistPath)
        })
    } catch { }

    return $checklistPath
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
        [int]$ResumeAfterOrder,
        [array]$CompletedModules,
        [datetime]$ProfileStartTime = (Get-Date),
        # ----- v2 schema fields (FlexProfile) -----
        # Set ExecutionMode='Flex' (and optionally pass SelectedOrders /
        # ModuleStates) to emit schemaVersion=2 with Flex-specific fields.
        # Linear callers omit these and the output is byte-for-byte
        # compatible with pre-P4 v1 format (no schemaVersion field).
        # The reading path (Load-ResumeState) is unchanged in P4 — Flex
        # resume detection lands in P6.
        [ValidateSet('Linear','Flex')][string]$ExecutionMode = 'Linear',
        [int[]]$SelectedOrders = @(),
        [hashtable]$ModuleStates = @{}
    )

    # Snapshot all host environment variables
    $hostEnv = @{}
    $envNames = @(
        "SELECTED_KANRI_NO", "SELECTED_OLD_PCNAME", "SELECTED_NEW_PCNAME",
        "SELECTED_ETH_IP", "SELECTED_ETH_SUBNET", "SELECTED_ETH_GATEWAY",
        "SELECTED_WIFI_IP", "SELECTED_WIFI_SUBNET", "SELECTED_WIFI_GATEWAY",
        "SELECTED_DNS1", "SELECTED_DNS2", "SELECTED_DNS3", "SELECTED_DNS4",
        "SELECTED_PIN"
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
        AutoPilot        = $global:AutoPilotMode
        AutoPilotWaitSec = $global:AutoPilotWaitSec
        SessionID        = $script:SessionID
        ResumeAfterOrder = $ResumeAfterOrder
        CompletedModules = @($CompletedModules | ForEach-Object {
            # Order included so post-restart resume can re-Add-ExecutionResult
            # with the correct Profile row Order, preserving per-row state
            # tracking across the restart boundary.
            $cmOrder = if ($null -ne $_.Order) { [int]$_.Order } else { 0 }
            @{ MenuName = $_.MenuName; Status = $_.Status; Order = $cmOrder }
        })
        HostEnvironment  = $hostEnv
        EvidenceBasePath = $global:FabriqEvidenceBasePath
        # Absolute start timestamp of the profile (in ISO 8601 / round-trip
        # format). Final elapsed time is computed at completion as a simple
        # subtraction (Get-Date) - ProfileStartTime, naturally including
        # reboot/login/startup gaps across __RESTART__ cycles.
        ProfileStartTime = $ProfileStartTime.ToString("o")
    }

    # Persist master passphrase (DPAPI LocalMachine encrypted) for post-reboot resume
    if (-not [string]::IsNullOrWhiteSpace($global:FabriqMasterPassphrase)) {
        try {
            $state["ProtectedPassphrase"] = Protect-PassphraseForResume -Passphrase $global:FabriqMasterPassphrase
        }
        catch {
            Show-Warning "Failed to protect passphrase for resume: $_"
        }
    }

    # v2 schema additions: only emit when caller declared Flex mode.
    # When ExecutionMode='Linear' (default) the output is byte-for-byte
    # compatible with pre-P4 v1 format (no schemaVersion / ExecutionMode
    # / SelectedOrders / ModuleStates fields are written).
    if ($ExecutionMode -eq 'Flex') {
        $state['schemaVersion']  = 2
        $state['ExecutionMode']  = 'Flex'
        $state['SelectedOrders'] = @($SelectedOrders)
        $state['ModuleStates']   = $ModuleStates
    }

    $state | ConvertTo-Json -Depth 5 | Out-File -FilePath $script:ResumeStatePath -Encoding UTF8 -Force

    # Telemetry kernel event: restart-invoked record. Save-ResumeState
    # is the canonical "we are about to reboot" boundary for both
    # __RESTART__ markers and FlexProfile [Restart Now]; ResumeAfterOrder=-1
    # signals the [Restart Now] sentinel path.
    try {
        Write-KernelTelemetryEvent -Type "restart.invoked" -Data ([ordered]@{
            profileName       = $ProfileName
            executionMode     = $ExecutionMode
            resumeAfterOrder  = [int]$ResumeAfterOrder
            completedCount    = @($CompletedModules).Count
            schemaVersion     = if ($ExecutionMode -eq 'Flex') { 2 } else { 1 }
        })
    } catch { }
}

function Load-ResumeState {
    if (-not (Test-Path $script:ResumeStatePath)) { return $null }
    try {
        $json = Get-Content $script:ResumeStatePath -Raw -Encoding UTF8
        $loaded = $json | ConvertFrom-Json
        # Telemetry kernel event: resume-consumed. Note that Load-ResumeState
        # may be called multiple times; we accept potential duplicates rather
        # than tracking "first call" state. SessionID at load time may be
        # pre-resume; downstream callers reset it after Restore-HostEnvironment.
        try {
            $sv = if ($null -ne $loaded.schemaVersion) { [int]$loaded.schemaVersion } else { 1 }
            Write-KernelTelemetryEvent -Type "resume.consumed" -Data ([ordered]@{
                profileName      = "$($loaded.ProfileName)"
                schemaVersion    = $sv
                executionMode    = if ($null -ne $loaded.ExecutionMode) { "$($loaded.ExecutionMode)" } else { 'Linear' }
                resumeAfterOrder = if ($null -ne $loaded.ResumeAfterOrder) { [int]$loaded.ResumeAfterOrder } else { 0 }
                completedCount   = @($loaded.CompletedModules).Count
            })
        } catch { }
        return $loaded
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
    $script:LastBatchResults = @()
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
    # 5. Evidence Base Path & Profile Info
    # ----------------------------------------
    $global:FabriqEvidenceBasePath = $null
    $global:FabriqEvidenceRootPath = $null
    [Environment]::SetEnvironmentVariable("FABRIQ_EVIDENCE_BASE", $null, "Process")
    $global:FabriqLastProfileName    = $null
    $global:FabriqLastProfilePath    = $null
    $global:FabriqLastProfileModules = $null

    # ----------------------------------------
    # 6. Environment Variables (selected host)
    # ----------------------------------------
    $envKeys = @(
        "SELECTED_KANRI_NO", "SELECTED_OLD_PCNAME", "SELECTED_NEW_PCNAME",
        "SELECTED_ETH_IP", "SELECTED_ETH_SUBNET", "SELECTED_ETH_GATEWAY",
        "SELECTED_WIFI_IP", "SELECTED_WIFI_SUBNET", "SELECTED_WIFI_GATEWAY",
        "SELECTED_DNS1", "SELECTED_DNS2", "SELECTED_DNS3", "SELECTED_DNS4",
        "FABRIQ_AUTOLOGON_USER"
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
    # 7. Resume State + Status File
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
            $env:FABRIQ_WORKER_NAME = $script:SessionInfo.WorkerName
            Show-Success "Session loaded: Worker=$($script:SessionInfo.WorkerName), MediaSerial=$($script:SessionInfo.MediaSerial)"
            return $true
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
                Write-Host "  [Q] Quit" -ForegroundColor DarkGray
                Show-Separator
                Write-Host ""

                while ($true) {
                    Write-Host -NoNewline "Select worker: "
                    $wChoice = Read-Host

                    if ($wChoice -eq 'q' -or $wChoice -eq 'Q') {
                        return $false
                    }

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

    $env:FABRIQ_WORKER_NAME = $workerName
    Show-Success "Session initialized: Worker=$workerName, MediaSerial=$mediaSerial"
    return $true
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
        [array]$AllModules,
        # When set, returns rows with Enabled=0 too, each tagged with
        # _IsCheckedDefault reflecting the original CSV Enabled value.
        # __AUTOPILOT__ / __ASYNC__ markers still require Enabled=1 to
        # take effect, so disabled marker rows do not flip global state.
        # Consumed by FlexProfile to populate the dashboard with all
        # rows while preserving CSV-driven default checkbox state.
        [switch]$IncludeDisabled
    )

    $validModules = @()
    $invalidPaths = @()
    $autoPilot = $false
    $autoPilotWaitSec = 3

    # Kill switch: if async is disabled in config, __ASYNC__ markers are
    # silently ignored and all modules fall back to the synchronous path.
    # DefaultAsync: when true (and kill switch is also on), every module
    # is treated as async even without an __ASYNC__ marker — the marker
    # then becomes an idempotent no-op for backward compatibility with
    # existing profiles. The shipped async_config.json opts into this
    # default so exhausted operators don't have to remember the marker.
    $asyncCfg              = Get-FabriqAsyncConfig
    $asyncEnabledGlobally  = [bool]$asyncCfg.Enabled
    $asyncMode             = $asyncEnabledGlobally -and [bool]$asyncCfg.DefaultAsync

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

    # Filter enabled (unless -IncludeDisabled), sort by Order
    if ($IncludeDisabled) {
        $sortedEntries = @($entries | Sort-Object { [int]$_.Order })
    } else {
        $sortedEntries = @($entries | Where-Object { $_.Enabled -eq "1" } | Sort-Object { [int]$_.Order })
    }

    foreach ($entry in $sortedEntries) {
        $path = $entry.ScriptPath.Trim().Replace("/", "\")
        if ([string]::IsNullOrEmpty($path)) { continue }

        # AutoPilot metadata (extracted, not added to module list).
        # Effect requires Enabled=1 even under -IncludeDisabled, so a
        # disabled marker row never flips global AutoPilot state.
        if ($path -eq '__AUTOPILOT__') {
            if ($entry.Enabled -eq "1") {
                $autoPilot = $true
                if ($entry.Description -match 'WaitSec=(\d+)') {
                    $autoPilotWaitSec = [int]$Matches[1]
                }
            }
            continue
        }

        # Async mode: subsequent modules run in a child runspace so they
        # can be skipped via flag file or cut off on timeout. The marker
        # itself is not added to the module list.
        # Effect requires Enabled=1 even under -IncludeDisabled.
        if ($path -eq '__ASYNC__') {
            if ($entry.Enabled -eq "1") {
                if ($asyncEnabledGlobally) {
                    $asyncMode = $true
                }
                else {
                    Show-Info "__ASYNC__ marker ignored (async disabled in async_config.json)"
                }
            }
            continue
        }

        # __AUTO_to_<User>__ pattern: resolve to autologon_config module with User parameter
        if ($path -match '^__AUTO_to_(.+)__$') {
            $autoLogonUser = $Matches[1]
            $autoLogonModule = $AllModules | Where-Object { $_.ModuleDir -eq 'autologon_config' } | Select-Object -First 1
            if ($autoLogonModule) {
                $moduleWithOrder = $autoLogonModule.PSObject.Copy()
                $moduleWithOrder | Add-Member -NotePropertyName "Order" -NotePropertyValue ([int]$entry.Order) -Force
                $moduleWithOrder | Add-Member -NotePropertyName "_AutoLogonUser" -NotePropertyValue $autoLogonUser
                $moduleWithOrder | Add-Member -NotePropertyName "_IsAsync" -NotePropertyValue $asyncMode
                $groupValue = if ($entry.PSObject.Properties.Name -contains 'Group') { "$($entry.Group)".Trim() } else { "" }
                $moduleWithOrder | Add-Member -NotePropertyName "_Group" -NotePropertyValue $groupValue
                if ($IncludeDisabled) {
                    $moduleWithOrder | Add-Member -NotePropertyName "_IsCheckedDefault" -NotePropertyValue ($entry.Enabled -eq "1")
                }
                $moduleWithOrder.MenuName = "[AUTO:$autoLogonUser] $($autoLogonModule.MenuName)"
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
            $groupValue = if ($entry.PSObject.Properties.Name -contains 'Group') { "$($entry.Group)".Trim() } else { "" }
            $obj | Add-Member -NotePropertyName "_Group" -NotePropertyValue $groupValue
            if ($IncludeDisabled) {
                $obj | Add-Member -NotePropertyName "_IsCheckedDefault" -NotePropertyValue ($entry.Enabled -eq "1")
            }
            $validModules += $obj
            continue
        }

        $found = $AllModules | Where-Object { $_.RelativePath -eq $path } | Select-Object -First 1
        if ($found) {
            # Attach Order from profile CSV (used for resume filtering)
            $moduleWithOrder = $found.PSObject.Copy()
            $moduleWithOrder | Add-Member -NotePropertyName "Order" -NotePropertyValue ([int]$entry.Order) -Force

            # Segment parameter passing (same pattern as _AutoLogonUser)
            $segmentValue = if ($entry.PSObject.Properties.Name -contains 'Segment') { $entry.Segment } else { "" }
            $moduleWithOrder | Add-Member -NotePropertyName "_Segment" -NotePropertyValue $segmentValue
            if (-not [string]::IsNullOrWhiteSpace($segmentValue)) {
                $moduleWithOrder.MenuName = "$($found.MenuName) [seg:$segmentValue]"
            }

            # ErrorMode (AutoPilot per-module error handling): "" / Ask / Skip / Retry
            $errorModeValue = if ($entry.PSObject.Properties.Name -contains 'ErrorMode') { "$($entry.ErrorMode)".Trim() } else { "" }
            $moduleWithOrder | Add-Member -NotePropertyName "_ErrorMode" -NotePropertyValue $errorModeValue

            # Async dispatch flag (sticky after __ASYNC__ marker until end of profile)
            $moduleWithOrder | Add-Member -NotePropertyName "_IsAsync" -NotePropertyValue $asyncMode

            # Group association (FlexProfile Groups bar). Empty / missing
            # column = "no group", row not surfaced as a [Run: <Group>]
            # button. Linear path ignores this attribute entirely.
            $groupValue = if ($entry.PSObject.Properties.Name -contains 'Group') { "$($entry.Group)".Trim() } else { "" }
            $moduleWithOrder | Add-Member -NotePropertyName "_Group" -NotePropertyValue $groupValue

            # Default checkbox state for FlexProfile (only when -IncludeDisabled).
            # Reflects the CSV's original Enabled value at load time.
            if ($IncludeDisabled) {
                $moduleWithOrder | Add-Member -NotePropertyName "_IsCheckedDefault" -NotePropertyValue ($entry.Enabled -eq "1")
            }

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
        $isSpecial = $m._IsRestart -or $m._IsReexplorer
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
    Write-Host "  [2] AutoPilot (Auto-confirm all)" -ForegroundColor Magenta
    Write-Host ""
    Write-Host -NoNewline "Execution mode [1]: "
    $modeChoice = Read-Host

    $autoPilot = ($modeChoice -eq "2")

    return [PSCustomObject]@{
        Confirmed        = $true
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
# User Environment Variable Expansion
# ========================================
# When running elevated (Run As Administrator) from a different user account,
# %USERPROFILE%, %LOCALAPPDATA%, %APPDATA% expand to the admin's profile.
# This function resolves them to the logged-on user's paths instead.

# Cache for logged-on user info (populated on first call)
$script:_LoggedOnUserProfile = $null
$script:_LoggedOnUserResolved = $false
$script:_LoggedOnUserSid = $null
$script:_LoggedOnUserName = $null

# Internal helper: detect logged-on user and populate cache
function _Resolve-LoggedOnUser {
    if ($script:_LoggedOnUserResolved) { return }
    $script:_LoggedOnUserResolved = $true
    try {
        if (-not (Test-AdminPrivilege)) { return }
        $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
        $loggedOnUser = $cs.UserName
        if ([string]::IsNullOrWhiteSpace($loggedOnUser)) { return }
        $username = $loggedOnUser.Split('\')[-1]
        $currentUser = [System.Environment]::UserName
        # Only apply correction when elevated user differs from logged-on user
        if ($username -eq $currentUser) { return }
        $sid = (New-Object System.Security.Principal.NTAccount($loggedOnUser)).Translate(
            [System.Security.Principal.SecurityIdentifier]
        ).Value
        $script:_LoggedOnUserSid = $sid
        $script:_LoggedOnUserName = $username
        $profilePath = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$sid" -ErrorAction Stop).ProfileImagePath
        if (Test-Path $profilePath) {
            $script:_LoggedOnUserProfile = @{
                UserProfile  = $profilePath
                LocalAppData = Join-Path $profilePath "AppData\Local"
                AppData      = Join-Path $profilePath "AppData\Roaming"
            }
        }
    }
    catch {
        # Detection failed - cache remains null (fallback behavior)
    }
}

function Expand-UserEnvironmentVariables {
    param(
        [Parameter(Mandatory)]
        [string]$Value
    )

    # Fast path: no environment variables to expand
    if ($Value -notmatch '%') {
        return $Value
    }

    # Ensure logged-on user info is resolved (cached after first call)
    _Resolve-LoggedOnUser

    # If logged-on user differs, replace user-specific variables before standard expansion
    if ($null -ne $script:_LoggedOnUserProfile) {
        $Value = $Value -ireplace '%USERPROFILE%',  $script:_LoggedOnUserProfile.UserProfile
        $Value = $Value -ireplace '%LOCALAPPDATA%', $script:_LoggedOnUserProfile.LocalAppData
        $Value = $Value -ireplace '%APPDATA%',      $script:_LoggedOnUserProfile.AppData
    }

    # Expand any remaining standard variables (%TEMP%, %SystemRoot%, etc.)
    return [System.Environment]::ExpandEnvironmentVariables($Value)
}

# ========================================
# Zone.Identifier Removal
# ========================================
# Removes the Zone.Identifier alternate data stream (Mark of the Web)
# from a file or all files in a directory tree.
# Safe to call on files without Zone.Identifier (no-op).
# Uses best-effort approach: errors are silently ignored.
# ========================================

function Remove-ZoneIdentifier {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    if (-not (Test-Path $Path)) { return }

    if (Test-Path -Path $Path -PathType Container) {
        Get-ChildItem -Path $Path -Recurse -File -ErrorAction SilentlyContinue |
            ForEach-Object { Unblock-File -Path $_.FullName -ErrorAction SilentlyContinue }
    }
    else {
        Unblock-File -Path $Path -ErrorAction SilentlyContinue
    }
}

# ========================================
# HKCU Root Resolution for Elevated Sessions
# ========================================
# When running elevated as a different admin account, HKCU: points to
# the admin's registry hive. This function resolves the correct root
# for the logged-on user, using HKU:\<SID> when redirection is needed.
#
# Returns hashtable:
#   PsDrivePath  - "HKCU:" or "HKU:\<SID>" (for PowerShell cmdlets)
#   RegExePath   - "HKEY_CURRENT_USER" or "HKEY_USERS\<SID>" (for reg.exe)
#   Label        - Display label (e.g., "Current User" or "username (via HKU)")
#   Redirected   - $true if targeting a different user's hive
#   SID          - User SID string (or $null if not redirected)

function Resolve-HkcuRoot {
    _Resolve-LoggedOnUser

    if ($null -ne $script:_LoggedOnUserSid) {
        $sid = $script:_LoggedOnUserSid
        # Ensure HKU PSDrive exists
        if (-not (Get-PSDrive -Name HKU -ErrorAction SilentlyContinue)) {
            New-PSDrive -Name HKU -PSProvider Registry -Root HKEY_USERS | Out-Null
        }
        if (Test-Path "HKU:\$sid") {
            return @{
                PsDrivePath = "HKU:\$sid"
                RegExePath  = "HKEY_USERS\$sid"
                Label       = "$($script:_LoggedOnUserName) (via HKU)"
                Redirected  = $true
                SID         = $sid
            }
        }
        else {
            Show-Warning "Logged-on user hive not found in HKU. HKCU will target the elevated admin user."
        }
    }

    return @{
        PsDrivePath = "HKCU:"
        RegExePath  = "HKEY_CURRENT_USER"
        Label       = "Current User"
        Redirected  = $false
        SID         = $null
    }
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
        # Collect PC info from environment variables
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

        # Collect printer info (slots 1-10)
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

        # Aggregate the execution result summary
        $results = @($script:ExecutionResults)
        $executionInfo = @{
            Phase          = $Phase
            TotalCount     = @($results | Where-Object { $_.Status -ne "Separator" }).Count
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
                Verified   = if ($null -ne $r.Verified) { $r.Verified } else { $null }
            }
        }

        $currentPC = Get-CurrentPCInfo

        $statusData = @{
            UpdatedAt     = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            WorkerName    = $env:FABRIQ_WORKER_NAME
            PCInfo        = $pcInfo
            CurrentPCInfo = $currentPC
            Execution     = $executionInfo
        }

        # Atomic write: write to a tmp file first, then rename
        $tempPath = "$($script:StatusFilePath).tmp"
        $statusData | ConvertTo-Json -Depth 5 | Out-File -FilePath $tempPath -Encoding UTF8 -Force
        Move-Item -Path $tempPath -Destination $script:StatusFilePath -Force
    }
    catch {
        # Status monitor is best-effort; swallow errors silently
        try {
            # Fallback when Move-Item fails: write the file directly
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
        if (Test-Path $global:ArtPulseFilePath) {
            Remove-Item $global:ArtPulseFilePath -Force -ErrorAction SilentlyContinue
        }
    }
    catch { }
}

# ========================================
# Status Monitor Lifecycle
# ========================================

# Show a modal warning so the operator gets a clear acknowledgment that
# Status Monitor is unavailable for this session, even if conhost has
# already been hidden by the dashboard. Best-effort — if MessageBox
# itself is blocked by the same host policy, we silently fall back to
# the console warning that will already have been written.
function Show-MonitorFailureDialog {
    param(
        [string]$Title = "Fabriq - Status Monitor disabled",
        [Parameter(Mandatory)][string]$Body
    )
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
        [System.Windows.Forms.MessageBox]::Show(
            $Body,
            $Title,
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
    } catch {
        # Dialog blocked too — console warning is the only fallback.
    }
}

function Start-StatusMonitor {
    Write-StatusFile -Phase "idle"
    $monitorProcess = $null
    try {
        $monitorScript = ".\kernel\ps1\status_monitor.ps1"
        if (-not (Test-Path $monitorScript)) {
            Show-Warning "Status Monitor script not found: $monitorScript"
            Write-Host ""
            return $null
        }

        $statusFileFullPath = (Resolve-Path $script:StatusFilePath).Path
        $pulseFileFullPath = (Join-Path (Get-Location) $global:ArtPulseFilePath)
        $sentenceFile = ".\kernel\txt\art_sentences.txt"
        $sentenceFileFullPath = if (Test-Path $sentenceFile) {
            (Resolve-Path $sentenceFile).Path
        } else { "" }
        $silenceFlagFullPath = (Join-Path (Get-Location) ".\kernel\txt\silence.flag")

        # Per-session diagnostic log so child startup failures can be triaged
        # post-mortem. -WindowStyle Hidden eats every uncaught exception in
        # the child, so without a file-backed log the only symptom is "no
        # taskbar entry, no monitor window".
        $logsDir = Join-Path (Get-Location) "logs"
        if (-not (Test-Path $logsDir)) {
            New-Item -Path $logsDir -ItemType Directory -Force | Out-Null
        }
        $diagLogPath = Join-Path $logsDir ("status_monitor_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".log")

        $argList = @(
            "-NoProfile", "-ExecutionPolicy", "Unrestricted",
            "-File", $monitorScript,
            "-StatusFilePath", $statusFileFullPath,
            "-PulseFilePath", $pulseFileFullPath,
            "-SilenceFlagPath", $silenceFlagFullPath,
            "-DiagLogPath", $diagLogPath
        )
        if (-not [string]::IsNullOrWhiteSpace($sentenceFileFullPath)) {
            $argList += @("-SentenceFilePath", $sentenceFileFullPath)
        }

        $monitorProcess = Start-Process powershell.exe -ArgumentList $argList -WindowStyle Hidden -PassThru

        Show-Info "Status Monitor spawn: PID=$($monitorProcess.Id)"
        Show-Info "  Diagnostic log: $diagLogPath"
        Write-Host ""

        # ----------------------------------------------------------------
        # Outcome detection: poll for window existence rather than trust
        # HasExited timing alone.
        #
        # HasExited timing is unreliable on hardened hosts: WDAC / AppLocker
        # / Defender ASR can kill the child anywhere between 500ms and
        # several seconds, *or* leave it alive but silently neutered so the
        # form never shows. The only ground truth is "did a window titled
        # 'Fabriq' actually appear under our PID".
        #
        # We poll up to 4 seconds (every 200ms). Three exit conditions:
        #   1. MainWindowTitle matches "Fabriq" → success, return process
        #   2. Process HasExited           → failure (early death)
        #   3. Timeout, still alive, no window → failure (hung in startup)
        # ----------------------------------------------------------------
        $pollDeadline = (Get-Date).AddMilliseconds(4000)
        $windowFound = $false
        $earlyExit = $false
        while ((Get-Date) -lt $pollDeadline) {
            if ($null -eq $monitorProcess) { break }
            if ($monitorProcess.HasExited) {
                $earlyExit = $true
                break
            }
            try {
                $monitorProcess.Refresh()
                $title = $monitorProcess.MainWindowTitle
                if (-not [string]::IsNullOrWhiteSpace($title) -and $title -match "Fabriq") {
                    $windowFound = $true
                    break
                }
            } catch { }
            Start-Sleep -Milliseconds 200
        }

        if ($windowFound) {
            Show-Success "Status Monitor window confirmed."
            Write-Host ""
            Set-ConsoleForeground
            return $monitorProcess
        }

        # ----- Failure path -----

        $logExists = Test-Path $diagLogPath
        $logSize = if ($logExists) { (Get-Item $diagLogPath -ErrorAction SilentlyContinue).Length } else { 0 }
        $exitCode = if ($monitorProcess.HasExited) {
            try { $monitorProcess.ExitCode } catch { "n/a" }
        } else { "(still alive, will be terminated)" }

        # Shared message header so console + modal stay in sync.
        $headerLines = @()
        $causeLine = ""

        if ($earlyExit -and ($logSize -eq 0)) {
            # Process died without writing a single log line — strong signal
            # of host security policy denying script execution / file I/O.
            $causeLine = "Likely cause: host security policy is blocking the child PowerShell process (WDAC / AppLocker / Defender ASR / equivalent)."
            $headerLines += "Status Monitor exited without writing any diagnostic log (ExitCode=$exitCode)."
        }
        elseif ($earlyExit) {
            # Child wrote something then died. Show last log entry to point
            # operator to the failure point.
            $lastLine = ""
            try {
                $lastLine = (Get-Content $diagLogPath -Tail 1 -ErrorAction SilentlyContinue) -join ""
            } catch { }
            $causeLine = "Cause: child PowerShell process exited mid-startup."
            $headerLines += "Status Monitor exited mid-startup (ExitCode=$exitCode)."
            if (-not [string]::IsNullOrWhiteSpace($lastLine)) {
                $headerLines += "Last diagnostic log entry: $lastLine"
            }
        }
        else {
            # Process is alive but window never appeared. Either hung in
            # Add-Type compilation, or stuck before reaching Application.Run.
            # Kill it so we do not leave an orphan eating CPU.
            $causeLine = "Cause: child process is alive but no window appeared within 4 seconds (likely hung during startup, possibly under host security inspection)."
            $headerLines += "Status Monitor window did not appear within 4 seconds."
            try {
                if (-not $monitorProcess.HasExited) {
                    $monitorProcess.Kill()
                    $headerLines += "Hung child process terminated (PID=$($monitorProcess.Id))."
                }
            } catch {
                $headerLines += "Failed to terminate hung child: $($_.Exception.Message)"
            }
        }

        # Console output (visible if conhost is still on screen)
        foreach ($line in $headerLines) {
            Show-Warning $line
        }
        Show-Warning $causeLine
        Show-Warning "  Diagnostic log: $diagLogPath"
        Show-Warning "  Kitting will continue normally without the Status Monitor."
        Write-Host ""

        # Modal dialog so the operator gets the message even if conhost
        # has been hidden by a transient dashboard switch.
        $body = ($headerLines -join "`r`n") + "`r`n`r`n" + $causeLine + "`r`n`r`n" +
                "Kitting will continue normally without the Status Monitor. " +
                "All other fabriq features (profile execution, evidence capture, " +
                "HTML checklist) operate independently.`r`n`r`n" +
                "Diagnostic log:`r`n$diagLogPath"
        Show-MonitorFailureDialog -Body $body

        return $null
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
# Function: Exit Fabriq (Centralized Cleanup)
# ========================================
function Exit-Fabriq {
    # Idempotency guard: safe to call multiple times
    if ($global:_FabriqExitCalled) { return }
    $global:_FabriqExitCalled = $true

    Write-Host ""
    Show-Separator
    Show-Info "Exiting Fabriq..."
    Show-Separator

    # Stop Status Monitor (if running, also stops integrated Art Display)
    if ($null -ne $global:FabriqStatusMonitorProcess) {
        Stop-StatusMonitor -MonitorProcess $global:FabriqStatusMonitorProcess
        $global:FabriqStatusMonitorProcess = $null
    }

    # Hide in-process execution toolbar (defined in fabriq_operator;
    # may not exist in headless contexts where the GUI failed to load).
    if (Get-Command Hide-ExecutionToolbar -ErrorAction SilentlyContinue) {
        try { Hide-ExecutionToolbar } catch { }
    }

    # Verbose capture cleanup (no-op if not active; just resets the flag)
    try { Disable-FabriqVerboseCapture } catch { }

    # Disable sleep suppression
    Disable-SleepSuppression

    # Stop transcript
    try { Stop-Transcript -ErrorAction SilentlyContinue | Out-Null } catch { }
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
    $fabriqExe = Join-Path $fabriqRoot "Fabriq.exe"

    if (-not (Test-Path $fabriqExe)) {
        Show-Error "Fabriq.exe not found: $fabriqExe"
        return $false
    }

    $runOnceValue = "`"$fabriqExe`""
    $entryPoint = "Fabriq.exe"

    $runOncePath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"
    try {
        if (-not (Test-Path $runOncePath)) {
            New-Item -Path $runOncePath -Force | Out-Null
        }
        New-ItemProperty -Path $runOncePath -Name "FabriqAutoStart" `
            -Value $runOnceValue -PropertyType String -Force -ErrorAction Stop | Out-Null
        Show-Success "RunOnce registered ($entryPoint)"
        return $true
    }
    catch {
        Show-Error "Failed to register RunOnce: $_"
        return $false
    }
}

function Register-FabriqActiveSetup {
    param(
        [Parameter(Mandatory)] [string]$GUID,
        [Parameter(Mandatory)] [string]$Description,
        [Parameter(Mandatory)] [string]$ScriptName,
        [Parameter(Mandatory)] [string[]]$ScriptLines
    )

    if (-not (Test-AdminPrivilege)) {
        Show-Warning "Admin privileges required for Active Setup registration"
        return $false
    }

    $scriptDir  = "C:\ProgramData\fabriq"
    $scriptPath = Join-Path $scriptDir $ScriptName

    try {
        if (-not (Test-Path $scriptDir)) {
            New-Item -Path $scriptDir -ItemType Directory -Force | Out-Null
        }
        $ScriptLines | Set-Content -Path $scriptPath -Force -Encoding UTF8
        Show-Success "Script deployed: $scriptPath"

        $asPath = "HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\$GUID"
        if (-not (Test-Path $asPath)) {
            New-Item -Path $asPath -Force | Out-Null
        }
        $stubPath = "cmd.exe /c powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""
        Set-ItemProperty -Path $asPath -Name "(Default)"  -Value $Description -Force
        Set-ItemProperty -Path $asPath -Name "StubPath"   -Value $stubPath -Force
        Set-ItemProperty -Path $asPath -Name "Version"    -Value "1,0,0,0" -Force
        Show-Success "Active Setup registered: $GUID"
        return $true
    }
    catch {
        Show-Error "Active Setup registration failed: $_"
        return $false
    }
}

function Deploy-FabriqUserSetupLauncher {
    # Deploy the launcher script that runs apply_*.ps1 scripts and restarts Explorer.
    # Called by modules that need deferred user-level setup (reg_hkcu_config, spi_config).
    # Idempotent: safe to call multiple times (overwrites same file).

    if (-not (Test-AdminPrivilege)) {
        Show-Warning "Admin privileges required for Startup Batch deployment"
        return $false
    }

    $scriptDir  = "C:\ProgramData\fabriq"
    $scriptPath = Join-Path $scriptDir "fabriq_user_setup.ps1"

    try {
        if (-not (Test-Path $scriptDir)) {
            New-Item -Path $scriptDir -ItemType Directory -Force | Out-Null
        }

        $launcherContent = @'
# Fabriq User Setup Launcher
# Runs once at first logon via Startup folder trigger.
# Applies deferred HKCU/SPI settings and restarts Explorer.

$flagDir  = Join-Path $env:LOCALAPPDATA "fabriq"
$flagFile = Join-Path $flagDir "user_setup_done.flag"

# Flag check: skip if already executed
if (Test-Path $flagFile) { exit 0 }

# Run all apply_*.ps1 scripts
$scriptDir = "C:\ProgramData\fabriq"
$scripts = @(Get-ChildItem -Path $scriptDir -Filter "apply_*.ps1" -File -ErrorAction SilentlyContinue | Sort-Object Name)
foreach ($s in $scripts) {
    try { & $s.FullName } catch { }
}

# Restart Explorer to apply visual settings
try {
    Stop-Process -Name explorer -Force -ErrorAction Stop
} catch { }

$maxWait = 15
$elapsed = 0
while ($elapsed -lt $maxWait) {
    Start-Sleep -Seconds 1
    $elapsed++
    if (@(Get-Process -Name explorer -ErrorAction SilentlyContinue).Count -gt 0) { break }
}
if ($elapsed -ge $maxWait) {
    Start-Process explorer.exe
}

# Create flag so this does not run again
if (-not (Test-Path $flagDir)) {
    New-Item -Path $flagDir -ItemType Directory -Force | Out-Null
}
Get-Date -Format "yyyy-MM-dd HH:mm:ss" | Set-Content -Path $flagFile -Force -Encoding UTF8

# Clean up: remove Startup trigger since setup is complete
$userStartup = [Environment]::GetFolderPath("Startup")
$triggerCmd  = Join-Path $userStartup "FabriqUserSetup.cmd"
if (Test-Path $triggerCmd) {
    Remove-Item -Path $triggerCmd -Force -ErrorAction SilentlyContinue
}
'@

        $launcherContent | Set-Content -Path $scriptPath -Force -Encoding UTF8
        Show-Success "Launcher deployed: $scriptPath"
        return $true
    }
    catch {
        Show-Error "Failed to deploy launcher: $_"
        return $false
    }
}

function Deploy-FabriqStartupTrigger {
    # Deploy a .cmd trigger to Default Profile Startup folder.
    # When a new user logs in, this .cmd launches fabriq_user_setup.ps1 via Startup.
    # Idempotent: safe to call multiple times (overwrites same file).

    if (-not (Test-AdminPrivilege)) {
        Show-Warning "Admin privileges required for Startup trigger deployment"
        return $false
    }

    $startupDir = "C:\Users\Default\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"
    $cmdPath    = Join-Path $startupDir "FabriqUserSetup.cmd"

    try {
        if (-not (Test-Path $startupDir)) {
            New-Item -Path $startupDir -ItemType Directory -Force | Out-Null
        }

        $cmdContent = @"
@echo off
if exist "%LOCALAPPDATA%\fabriq\user_setup_done.flag" exit /b 0
powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "C:\ProgramData\fabriq\fabriq_user_setup.ps1"
"@

        $cmdContent | Set-Content -Path $cmdPath -Force -Encoding ASCII
        Show-Success "Startup trigger deployed: $cmdPath"
        return $true
    }
    catch {
        Show-Error "Failed to deploy Startup trigger: $_"
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

function Invoke-AutoResumeCountdown {
    # ========================================
    # Countdown for AutoPilot post-reboot resume.
    # Returns $true  → resume execution
    # Returns $false → abort (clear resume state)
    # Keys during countdown:
    #   Enter / Y → immediate resume
    #   Esc   / N → abort
    #   Other     → ignored (countdown continues)
    # ========================================
    param([int]$Seconds = 60)

    Write-Host ""
    Write-Host "[AUTOPILOT] Auto-resume countdown" -ForegroundColor Magenta
    Write-Host "  [Enter] = Resume now   [Esc] = Abort" -ForegroundColor DarkGray
    Write-Host ""

    for ($i = $Seconds; $i -ge 1; $i--) {
        Write-Host "`r  Resuming in $i seconds...   " -NoNewline -ForegroundColor Magenta
        Start-Sleep -Milliseconds 900

        if ($Host.UI.RawUI.KeyAvailable) {
            $key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            # Enter (13) or Y
            if ($key.VirtualKeyCode -eq 13 -or
                $key.Character -eq 'Y' -or $key.Character -eq 'y') {
                Write-Host ""
                Show-Success "Resuming now"
                return $true
            }
            # Escape (27) or N / Q
            if ($key.VirtualKeyCode -eq 27 -or
                $key.Character -eq 'N' -or $key.Character -eq 'n' -or
                $key.Character -eq 'Q' -or $key.Character -eq 'q') {
                Write-Host ""
                Show-Warning "Auto-resume aborted by operator"
                return $false
            }
            # Other key: drain buffer and continue countdown
        }
    }

    Write-Host ""
    Show-Info "Countdown complete. Auto-resuming..."
    return $true
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

        # Build save directory
        $pcName  = if ($env:SELECTED_NEW_PCNAME) { $env:SELECTED_NEW_PCNAME } else { $env:COMPUTERNAME }

        if (-not [string]::IsNullOrWhiteSpace($global:FabriqEvidenceBasePath)) {
            # Unified path: flat (no date/uid/pc subfolder)
            $saveDir = Join-Path $global:FabriqEvidenceBasePath "auto_capture"
        }
        else {
            # Fallback: legacy path with date/uid/pc subfolder
            $dateOnly = Get-Date -Format "yyyy_MM_dd"
            $uid      = if ($global:FabriqUniqueId) { $global:FabriqUniqueId } else { Get-HardwareUniqueId }
            $saveDir  = Join-Path $PSScriptRoot "..\evidence\auto_capture\${dateOnly}_${uid}_${pcName}"
        }

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

# ========================================
# Function: Save Screenshot (Manual)
# ========================================
# On-demand screenshot for manual evidence capture.
# Unlike Capture-ScreenEvidence (auto-capture during
# module execution), this is triggered by the user
# (e.g. via a button in Status Monitor).
#
# NOTE: Does NOT call SetProcessDPIAware(). That call
# irreversibly changes the process DPI mode, which shrinks
# WinForms windows in scaled displays. The caller (e.g.
# Status Monitor) is a GUI process, so DPI changes are
# destructive. Screenshots are captured at logical resolution.
#
# Returns: file path on success, $null on failure.
# ========================================
function Save-Screenshot {
    param(
        [Parameter(Mandatory=$true)]
        [string]$BaseDir
    )

    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
        Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue

        # PC name resolution (same convention as Capture-ScreenEvidence)
        $pcName = if ($env:SELECTED_NEW_PCNAME) { $env:SELECTED_NEW_PCNAME } else { $env:COMPUTERNAME }

        # Build save directory
        if (-not [string]::IsNullOrWhiteSpace($global:FabriqEvidenceBasePath)) {
            # Unified path: flat (no date/pc subfolder)
            $saveDir = Join-Path $global:FabriqEvidenceBasePath "gyotaku"
        }
        else {
            # Fallback: legacy path with date/pc subfolder
            $dateOnly = Get-Date -Format "yyyy_MM_dd"
            $saveDir = Join-Path $BaseDir "${dateOnly}_${pcName}"
        }

        if (-not (Test-Path $saveDir)) {
            New-Item -Path $saveDir -ItemType Directory -Force | Out-Null
        }

        # Build filename: YYYY_MM_DD_HHmmss_{PCname}.png
        $timestamp = Get-Date -Format "yyyy_MM_dd_HHmmss"
        $fileName = "${timestamp}_${pcName}.png"
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

        return $filePath
    }
    catch {
        return $null
    }
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