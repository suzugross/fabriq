# ========================================
# Pianist - Multi-phase GUI Configuration Maestro
# ========================================
# Extended fabriq module for GUI-based configuration of business
# applications that cannot be automated through registry/CLI alone.
#
# v1.0.0 first release as a module (migrated from the v0.2.x app
# experiment in apps/pianist/). Profile authoring/editing belongs
# to Fabriq Studio. Pianist concentrates on letting an operator
# step through phases of an existing profile during a kitting
# session, then returns a ModuleResult so the kernel records the
# run in the standard execution history.
#
# UI policy: mouse-only. No keyboard accelerators - keyboard
# input is reserved for SendKeys-to-target.
#
# NOTES:
# - Requires module to run inside fabriq (kernel/common.ps1 must
#   be dot-sourced so Show-Info / New-ModuleResult / Confirm-Execution
#   etc. are available)
# - Master passphrase is shared with fabriq for ENC: decryption
#   of values.csv entries
# - Layout uses absolute positioning + Anchor only (no nested Dock),
#   matching modules/extended/manual_kitting_assistant for predictable
#   resizing
# ========================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ========================================
# Win32 Window Finder
# Enumerates all top-level windows including modal dialogs owned
# by other processes (e.g. Run dialog hosted by explorer.exe,
# Save As dialog hosted by notepad). Plain Get-Process |
# Where MainWindowTitle misses these.
# ========================================
if (-not ([System.Management.Automation.PSTypeName]'PianistWin32').Type) {
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Text;

public class PianistWin32 {
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern int GetWindowTextW(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    public const int SW_RESTORE = 9;

    public static IntPtr FindByTitleContains(string substr) {
        IntPtr result = IntPtr.Zero;
        if (string.IsNullOrEmpty(substr)) return result;
        EnumWindows((hWnd, lParam) => {
            if (!IsWindowVisible(hWnd)) return true;
            var sb = new StringBuilder(512);
            int len = GetWindowTextW(hWnd, sb, sb.Capacity);
            if (len > 0) {
                string title = sb.ToString();
                if (title.IndexOf(substr, StringComparison.OrdinalIgnoreCase) >= 0) {
                    result = hWnd;
                    return false;
                }
            }
            return true;
        }, IntPtr.Zero);
        return result;
    }

    public static bool Focus(IntPtr hWnd) {
        if (hWnd == IntPtr.Zero) return false;
        ShowWindow(hWnd, SW_RESTORE);
        return SetForegroundWindow(hWnd);
    }
}
"@
}

# ========================================
# Color Scheme (Fabriq Standard)
# ========================================
$bgDark       = [System.Drawing.Color]::FromArgb(30, 30, 30)
$bgPanel      = [System.Drawing.Color]::FromArgb(40, 40, 40)
$bgGrid       = [System.Drawing.Color]::FromArgb(35, 35, 35)
$bgCell       = [System.Drawing.Color]::FromArgb(45, 45, 45)
$bgHeader     = [System.Drawing.Color]::FromArgb(55, 55, 55)
$bgButton     = [System.Drawing.Color]::FromArgb(60, 60, 60)
$bgButtonHov  = [System.Drawing.Color]::FromArgb(80, 80, 80)
$bgAccent     = [System.Drawing.Color]::FromArgb(0, 120, 215)
$bgWarn       = [System.Drawing.Color]::FromArgb(200, 140, 40)
$bgError      = [System.Drawing.Color]::FromArgb(180, 40, 40)
$bgRun        = [System.Drawing.Color]::FromArgb(40, 140, 60)
$fgText       = [System.Drawing.Color]::FromArgb(220, 220, 220)
$fgDim        = [System.Drawing.Color]::FromArgb(150, 150, 150)
$fgHeader     = [System.Drawing.Color]::FromArgb(100, 180, 255)
$gridLine     = [System.Drawing.Color]::FromArgb(60, 60, 60)
$bgInput      = [System.Drawing.Color]::FromArgb(50, 50, 50)

# ========================================
# Phase Color Palette (procedure.csv "Color" column to RGB)
# ========================================
$script:phaseColors = @{
    "Blue"   = [System.Drawing.Color]::FromArgb( 60, 120, 200)
    "Green"  = [System.Drawing.Color]::FromArgb( 50, 150,  80)
    "Yellow" = [System.Drawing.Color]::FromArgb(180, 150,  40)
    "Orange" = [System.Drawing.Color]::FromArgb(200, 110,  40)
    "Red"    = [System.Drawing.Color]::FromArgb(180,  60,  60)
    "Purple" = [System.Drawing.Color]::FromArgb(140,  80, 180)
    "Cyan"   = [System.Drawing.Color]::FromArgb( 60, 160, 180)
    "Pink"   = [System.Drawing.Color]::FromArgb(200, 100, 150)
    "Gray"   = [System.Drawing.Color]::FromArgb( 90,  90,  90)
}

# ========================================
# Module-level state
# ========================================
$script:currentProfile      = $null
$script:isRunning           = $false
$script:wsShell             = $null
$script:phasesOrdered       = @()
$script:currentPhaseIndex   = -1
$script:phaseStatus         = @{}
$script:UserAction          = $null   # "done" | "cancel" | $null
$script:profilesRoot        = Join-Path $PSScriptRoot "profiles"

# ========================================
# Logging helper - mirror to console (Show-Info) and the in-form log
# ========================================
function Write-PianistLog {
    param([string]$Message, [string]$Level = "INFO")
    if ($null -ne $script:logBox) {
        $stamp = Get-Date -Format "HH:mm:ss"
        $color = switch ($Level) {
            "ERROR" { $bgError }
            "WARN"  { $bgWarn }
            "OK"    { $bgRun }
            default { $fgText }
        }
        $script:logBox.SelectionStart  = $script:logBox.TextLength
        $script:logBox.SelectionLength = 0
        $script:logBox.SelectionColor  = $fgDim
        $script:logBox.AppendText("[$stamp] ")
        $script:logBox.SelectionColor  = $color
        $script:logBox.AppendText("[$Level] ")
        $script:logBox.SelectionColor  = $fgText
        $script:logBox.AppendText("$Message`r`n")
        $script:logBox.ScrollToCaret()
        [System.Windows.Forms.Application]::DoEvents()
    }
    # Also surface to fabriq console
    switch ($Level) {
        "ERROR" { Show-Error $Message }
        "WARN"  { Show-Warning $Message }
        "OK"    { Show-Success $Message }
        default { Show-Info $Message }
    }
}

# ========================================
# WinForms Helpers
# ========================================
function New-PianistButton {
    param(
        [string]$Text,
        [int]$X, [int]$Y = 0,
        [int]$Width = 100, [int]$Height = 28,
        $BgColor = $bgButton, $FgColor = $fgText,
        $Font = $null
    )
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = $Text
    $btn.Location = New-Object System.Drawing.Point($X, $Y)
    $btn.Size = New-Object System.Drawing.Size($Width, $Height)
    $btn.FlatStyle = "Flat"
    $btn.FlatAppearance.BorderColor = $gridLine
    $btn.FlatAppearance.MouseOverBackColor = $bgButtonHov
    $btn.BackColor = $BgColor
    $btn.ForeColor = $FgColor
    $btn.Cursor = [System.Windows.Forms.Cursors]::Hand
    if ($null -ne $Font) { $btn.Font = $Font }
    return $btn
}

# ========================================
# Profile data loading
# ========================================
function Read-CsvSafe {
    param([string]$Path)
    if (-not (Test-Path $Path)) { return @() }
    try {
        return @(Import-Csv -Path $Path -Encoding UTF8)
    } catch {
        Write-PianistLog "Failed to read CSV: $Path - $_" "ERROR"
        return @()
    }
}

# Decrypt a single cell. ENC:<Base64> prefix triggers decryption via the
# kernel's Unprotect-FabriqValue (matches the convention used by hostlist.csv
# and Import-ModuleCsv). Cells without the prefix pass through unchanged.
function Resolve-PianistEncryptedCell {
    param([string]$Cell)
    if ([string]::IsNullOrEmpty($Cell)) { return "" }
    if (-not $Cell.StartsWith('ENC:')) { return $Cell }
    if ([string]::IsNullOrWhiteSpace($global:FabriqMasterPassphrase)) { return $Cell }
    if (-not (Get-Command Unprotect-FabriqValue -ErrorAction SilentlyContinue)) { return $Cell }
    try {
        return (Unprotect-FabriqValue -EncryptedValue $Cell -Passphrase $global:FabriqMasterPassphrase)
    } catch {
        Write-PianistLog "Decrypt failed for an encrypted cell (using ciphertext as-is): $_" "WARN"
        return $Cell
    }
}

# Build the resolved $VarName -> value dictionary from values.csv.
# Two schemas are supported:
#   1. Wide format (current, since pianist 1.1.0): NewPCName + arbitrary
#      variable columns. Row with NewPCName='*' (or empty) provides the
#      defaults; row matching the current selected NewPCName overrides it
#      column-by-column. Empty cells fall through to the default row.
#   2. Legacy long format: Key,Value,Encrypted,Note (pianist 1.0.0 shape).
#      Detected when the first row has no NewPCName column. Encrypted=1
#      flag triggers decryption; behavior is preserved bit-for-bit.
# Variable column names must match the [A-Za-z_][A-Za-z0-9_]* pattern that
# Expand-Variables uses, otherwise procedure.csv cannot reference them.
function Build-PianistValuesDict {
    param([array]$Rows)
    $dict = @{}
    if ($null -eq $Rows -or $Rows.Count -eq 0) { return $dict }

    $columns = @($Rows[0].PSObject.Properties.Name)
    $isWideFormat = $columns -contains 'NewPCName'

    if ($isWideFormat) {
        $currentPC = [string]$env:SELECTED_NEW_PCNAME
        $valueColumns = @($columns | Where-Object { $_ -ne 'NewPCName' })

        $defaultRow = $Rows | Where-Object {
            $_.NewPCName -eq '*' -or [string]::IsNullOrWhiteSpace($_.NewPCName)
        } | Select-Object -First 1

        $pcRow = $null
        if (-not [string]::IsNullOrWhiteSpace($currentPC)) {
            $pcRow = $Rows | Where-Object { $_.NewPCName -eq $currentPC } | Select-Object -First 1
        }

        if ($null -ne $pcRow) {
            Write-PianistLog "values.csv: matched row for NewPCName='$currentPC' (defaults: $([bool]$defaultRow))"
        } elseif (-not [string]::IsNullOrWhiteSpace($currentPC)) {
            Write-PianistLog "values.csv: no row for NewPCName='$currentPC', using defaults only" "WARN"
        }

        foreach ($col in $valueColumns) {
            $cell = $null
            if ($null -ne $pcRow) {
                $candidate = [string]$pcRow.$col
                if (-not [string]::IsNullOrEmpty($candidate)) { $cell = $candidate }
            }
            if ($null -eq $cell -and $null -ne $defaultRow) {
                $candidate = [string]$defaultRow.$col
                if (-not [string]::IsNullOrEmpty($candidate)) { $cell = $candidate }
            }
            if ($null -ne $cell) {
                $dict[$col] = Resolve-PianistEncryptedCell -Cell $cell
            }
        }
        return $dict
    }

    # Legacy long format (pianist <= 1.0.0)
    foreach ($r in $Rows) {
        if ([string]::IsNullOrWhiteSpace($r.Key)) { continue }
        $val = [string]$r.Value
        if ($r.Encrypted -eq '1' -and -not [string]::IsNullOrEmpty($val)) {
            if (-not [string]::IsNullOrWhiteSpace($global:FabriqMasterPassphrase) -and (Get-Command Unprotect-FabriqValue -ErrorAction SilentlyContinue)) {
                try {
                    $val = Unprotect-FabriqValue -EncryptedValue $val -Passphrase $global:FabriqMasterPassphrase
                } catch {
                    Write-PianistLog "Decrypt failed (legacy values.csv row Key=$($r.Key)): $_" "WARN"
                }
            }
        }
        $dict[$r.Key] = $val
    }
    return $dict
}

function Load-PianistProfileData {
    param([string]$ProfileName)
    $profilePath = Join-Path $script:profilesRoot $ProfileName
    if (-not (Test-Path $profilePath)) {
        Show-Error "Profile folder not found: $profilePath"
        return $null
    }

    $metaPath = Join-Path $profilePath "pianist.json"
    $meta = $null
    if (Test-Path $metaPath) {
        try {
            $meta = Get-Content -Path $metaPath -Raw -Encoding UTF8 | ConvertFrom-Json
        } catch {
            Show-Warning "Invalid pianist.json: $_"
            $meta = [PSCustomObject]@{ label = $ProfileName; description = "(invalid pianist.json)"; default_phase = "" }
        }
    } else {
        $meta = [PSCustomObject]@{ label = $ProfileName; description = ""; default_phase = "" }
    }

    $procRaw = Read-CsvSafe (Join-Path $profilePath "procedure.csv")
    $valsRaw = Read-CsvSafe (Join-Path $profilePath "values.csv")
    $cutsRaw = Read-CsvSafe (Join-Path $profilePath "shortcuts.csv")

    if ($procRaw.Count -eq 0) {
        Show-Error "procedure.csv is empty or missing for profile [$ProfileName]"
        return $null
    }

    $valuesDict = Build-PianistValuesDict -Rows $valsRaw

    return [PSCustomObject]@{
        Name        = $ProfileName
        Path        = $profilePath
        Meta        = $meta
        Procedure   = $procRaw
        Values      = $valsRaw
        ValuesDict  = $valuesDict
        Shortcuts   = $cutsRaw
    }
}

# ========================================
# Variable Substitution: $VarName -> values.csv lookup
# ========================================
function Expand-Variables {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return $Text }
    if ($null -eq $script:currentProfile) { return $Text }

    $dict = $script:currentProfile.ValuesDict
    return [System.Text.RegularExpressions.Regex]::Replace(
        $Text,
        '\$([A-Za-z_][A-Za-z0-9_]*)',
        {
            param($m)
            $key = $m.Groups[1].Value
            if ($dict.ContainsKey($key)) { return [string]$dict[$key] }
            return $m.Value
        }
    )
}

# ========================================
# Action Executors (10 actions: Open / WaitWin / AppFocus / Type /
# Key / Wait / Copy / Paste / Screenshot / Prompt)
# ========================================
# Invoke an Open step. Three input shapes are supported:
#   1. URL / shell scheme (https://, ms-settings:, shell:)        - Start-Process direct
#   2. Quoted path: "C:\path\with space\app.exe" [args...]        - explicit split on closing quote
#   3. Unquoted path: works for both no-space (legacy) and space-in-path
#      with no args (resolved via Test-Path -LiteralPath). Unquoted path
#      with space PLUS args still requires quoting (Windows convention).
function Invoke-PianistOpen {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
    $v = $Value.Trim()

    if ($v -match '^(https?://|ms-settings:|shell:)') {
        try { $null = Start-Process $v -ErrorAction Stop; return $true }
        catch { Write-PianistLog "Open failed: $v - $_" "ERROR"; return $false }
    }

    $filePath  = $null
    $arguments = ''

    if ($v.StartsWith('"')) {
        $end = $v.IndexOf('"', 1)
        if ($end -gt 0) {
            $filePath  = $v.Substring(1, $end - 1)
            $arguments = $v.Substring($end + 1).TrimStart()
        }
    }

    if ($null -eq $filePath) {
        if (Test-Path -LiteralPath $v -PathType Leaf) {
            $filePath  = $v
            $arguments = ''
        }
        elseif ($v -match ' ') {
            $parts     = $v -split ' ', 2
            $filePath  = $parts[0]
            $arguments = $parts[1]
        }
        else {
            $filePath  = $v
            $arguments = ''
        }
    }

    try {
        if ([string]::IsNullOrEmpty($arguments)) {
            $null = Start-Process -FilePath $filePath -ErrorAction Stop
        } else {
            $null = Start-Process -FilePath $filePath -ArgumentList $arguments -ErrorAction Stop
        }
        return $true
    } catch {
        # Fallback via cmd /c start (rescues PATH-resolved targets, env vars, etc.)
        try {
            $cmdArgs = if ([string]::IsNullOrEmpty($arguments)) {
                "/c start `"`" `"$filePath`""
            } else {
                "/c start `"`" `"$filePath`" $arguments"
            }
            $null = Start-Process -FilePath "cmd" -ArgumentList $cmdArgs -WindowStyle Hidden
            return $true
        } catch {
            Write-PianistLog "Open failed: $v - $_" "ERROR"
            return $false
        }
    }
}

function Invoke-PianistWaitWin {
    param([string]$Title, [int]$TimeoutMs)
    if ($TimeoutMs -le 0) { $TimeoutMs = 10000 }
    if ([string]::IsNullOrWhiteSpace($Title)) { return $false }
    $elapsed = 0
    Write-PianistLog "Waiting for window matching '$Title' (max $($TimeoutMs/1000)s)..."
    while ($elapsed -lt $TimeoutMs) {
        $hwnd = [PianistWin32]::FindByTitleContains($Title)
        if ($hwnd -ne [IntPtr]::Zero) {
            [void][PianistWin32]::Focus($hwnd)
            Start-Sleep -Milliseconds 300
            return $true
        }
        Start-Sleep -Milliseconds 400
        $elapsed += 400
        [System.Windows.Forms.Application]::DoEvents()
    }
    return $false
}

function Invoke-PianistAppFocus {
    param([string]$Title)
    if ([string]::IsNullOrWhiteSpace($Title)) { return $false }
    $hwnd = [PianistWin32]::FindByTitleContains($Title)
    if ($hwnd -ne [IntPtr]::Zero) {
        [void][PianistWin32]::Focus($hwnd)
        Start-Sleep -Milliseconds 200
        return $true
    }
    try {
        $r = $script:wsShell.AppActivate($Title)
        Start-Sleep -Milliseconds 200
        return [bool]$r
    } catch { return $false }
}

function Invoke-PianistType {
    param([string]$Text)
    try { [System.Windows.Forms.SendKeys]::SendWait($Text); return $true }
    catch { Write-PianistLog "Type failed: $_" "ERROR"; return $false }
}

function Invoke-PianistKey {
    param([string]$KeySequence)
    try { [System.Windows.Forms.SendKeys]::SendWait($KeySequence); return $true }
    catch { Write-PianistLog "Key failed: $_" "ERROR"; return $false }
}

function Invoke-PianistCopy {
    param([string]$Value)
    try {
        if ([string]::IsNullOrEmpty($Value)) { [System.Windows.Forms.Clipboard]::Clear() }
        else                                  { [System.Windows.Forms.Clipboard]::SetText($Value) }
        return $true
    } catch { Write-PianistLog "Copy failed: $_" "ERROR"; return $false }
}

function Invoke-PianistPaste {
    param([string]$Value)
    if (-not (Invoke-PianistCopy -Value $Value)) { return $false }
    Start-Sleep -Milliseconds 80
    return (Invoke-PianistKey -KeySequence "^v")
}

function Invoke-PianistScreenshot {
    param([string]$Tag)
    try {
        if (Get-Command Capture-ScreenEvidence -ErrorAction SilentlyContinue) {
            $modName = "pianist"
            if ($script:currentProfile) { $modName = "pianist_$($script:currentProfile.Name)" }
            Capture-ScreenEvidence -ModuleName $modName -Status $Tag
            return $true
        }
        return $false
    } catch { Write-PianistLog "Screenshot failed: $_" "ERROR"; return $false }
}

function Invoke-PianistPrompt {
    param([string]$Message)
    $expanded = Expand-Variables $Message
    [void][System.Windows.Forms.MessageBox]::Show(
        $expanded, "Pianist - Manual Step",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )
    return $true
}

# ========================================
# Step / Phase runner
# ========================================
function Invoke-PianistStep {
    param([PSCustomObject]$Step)
    if (-not $Step) { return $false }

    $action = $Step.Action
    $value  = Expand-Variables $Step.Value
    $waitMs = 0
    [int]::TryParse($Step.Wait, [ref]$waitMs) | Out-Null

    $label = if ($Step.PhaseID) { "$($Step.PhaseID).$($Step.StepNo)" } else { "Step$($Step.StepNo)" }
    $note  = if ($Step.Note) { " - $($Step.Note)" } else { "" }
    Write-PianistLog "$label $action`: $value$note"

    $ok = $false
    switch ($action) {
        "Open"       { $ok = Invoke-PianistOpen -Value $value }
        "WaitWin"    { $ok = Invoke-PianistWaitWin -Title $value -TimeoutMs $waitMs }
        "AppFocus"   { $ok = Invoke-PianistAppFocus -Title $value }
        "Type"       { $ok = Invoke-PianistType -Text $value }
        "Key"        { $ok = Invoke-PianistKey -KeySequence $value }
        "Wait" {
            $ms = 0
            [int]::TryParse($value, [ref]$ms) | Out-Null
            if ($ms -le 0) { $ms = $waitMs }
            if ($ms -gt 0) { Start-Sleep -Milliseconds $ms }
            $ok = $true
        }
        "Copy"       { $ok = Invoke-PianistCopy -Value $value }
        "Paste"      { $ok = Invoke-PianistPaste -Value $value }
        "Screenshot" { $ok = Invoke-PianistScreenshot -Tag $value }
        "Prompt"     { $ok = Invoke-PianistPrompt -Message $Step.Value }
        default      { Write-PianistLog "Unknown action: $action" "WARN"; $ok = $false }
    }

    if ($ok) { Write-PianistLog "  -> OK" "OK" } else { Write-PianistLog "  -> FAILED" "ERROR" }

    if ($action -ne "Wait" -and $action -ne "WaitWin" -and $waitMs -gt 0) {
        Start-Sleep -Milliseconds $waitMs
    }
    [System.Windows.Forms.Application]::DoEvents()
    return $ok
}

function Invoke-PianistPhase {
    param([string]$PhaseID)
    if ($null -eq $script:currentProfile) { Write-PianistLog "No profile loaded" "WARN"; return }
    if ($script:isRunning) { Write-PianistLog "Already running another phase" "WARN"; return }

    $steps = @($script:currentProfile.Procedure | Where-Object { $_.PhaseID -eq $PhaseID })
    if ($steps.Count -eq 0) { Write-PianistLog "Phase '$PhaseID' has no steps" "WARN"; return }

    $phaseLabel = $steps[0].PhaseLabel
    $script:isRunning = $true
    Set-AllControlsEnabled $false

    if ($null -ne $script:phaseStatus -and $script:phaseStatus.ContainsKey($PhaseID)) {
        $script:phaseStatus[$PhaseID].Auto = "Running"
    }
    Update-StatusBadges

    Write-PianistLog "===== Phase $PhaseID '$phaseLabel' ($($steps.Count) steps) ====="

    $ok = 0; $fail = 0
    foreach ($s in $steps) {
        if (Invoke-PianistStep -Step $s) { $ok++ } else { $fail++ }
    }

    $autoStatus = if ($fail -eq 0) { "Success" }
                  elseif ($ok -eq 0) { "Error" }
                  else { "Partial" }
    if ($null -ne $script:phaseStatus -and $script:phaseStatus.ContainsKey($PhaseID)) {
        $script:phaseStatus[$PhaseID].Auto = $autoStatus
    }
    Update-StatusBadges

    Write-PianistLog "===== Phase $PhaseID done: $ok OK / $fail FAIL (Auto=$autoStatus) ====="
    $script:isRunning = $false
    Set-AllControlsEnabled $true
}

# ========================================
# Phase ordering / navigation state
# ========================================
function Build-PhasesOrdered {
    $script:phasesOrdered = @()
    $script:phaseStatus   = @{}
    if ($null -eq $script:currentProfile) { return }
    $seen = @{}
    foreach ($s in $script:currentProfile.Procedure) {
        if ([string]::IsNullOrWhiteSpace($s.PhaseID)) { continue }
        if ($seen.ContainsKey($s.PhaseID)) { continue }
        $seen[$s.PhaseID] = $true
        $stepsOfPhase = @($script:currentProfile.Procedure | Where-Object { $_.PhaseID -eq $s.PhaseID })
        $script:phasesOrdered += [PSCustomObject]@{
            ID    = $s.PhaseID
            Label = $s.PhaseLabel
            Color = $s.Color
            Steps = $stepsOfPhase
        }
        $script:phaseStatus[$s.PhaseID] = @{ Auto = "NotRun"; Manual = "Unset"; Note = "" }
    }
}

function Set-CurrentPhase {
    param([int]$Index)
    if ($script:phasesOrdered.Count -eq 0) {
        $script:currentPhaseIndex = -1
        Update-PhaseView
        return
    }
    $bounded = [Math]::Max(0, [Math]::Min($Index, $script:phasesOrdered.Count - 1))
    $script:currentPhaseIndex = $bounded
    Update-PhaseView
}

# ========================================
# UI: Update phase view, badges, nav
# ========================================
function Get-StatusColor {
    param([string]$Status)
    switch ($Status) {
        "Success" { return $bgRun }
        "OK"      { return $bgRun }
        "Partial" { return $bgWarn }
        "Warning" { return $bgWarn }
        "Running" { return $bgAccent }
        "Error"   { return $bgError }
        "Skip"    { return [System.Drawing.Color]::FromArgb(110, 110, 110) }
        default   { return [System.Drawing.Color]::FromArgb(70, 70, 70) }
    }
}

function Update-NavButtons {
    if ($null -eq $script:btnPrev -or $null -eq $script:btnNext) { return }
    if ($script:phasesOrdered.Count -eq 0) {
        $script:btnPrev.Enabled = $false
        $script:btnNext.Enabled = $false
        if ($null -ne $script:lblManualHint) { $script:lblManualHint.Visible = $false }
        return
    }
    $script:btnPrev.Enabled = ($script:currentPhaseIndex -gt 0)

    # Forward navigation requires the current phase's Manual status to be
    # explicitly set (anything other than "Unset"). This forces the operator
    # to record a judgment for every phase before advancing or finishing.
    $manualSet = $false
    if ($script:currentPhaseIndex -ge 0 -and $script:currentPhaseIndex -lt $script:phasesOrdered.Count) {
        $phase = $script:phasesOrdered[$script:currentPhaseIndex]
        $manualSet = ($script:phaseStatus[$phase.ID].Manual -ne "Unset")
    }

    $isLast = ($script:currentPhaseIndex -eq ($script:phasesOrdered.Count - 1))
    if ($isLast) {
        # Right arrow becomes the "Complete profile" button on the last phase.
        # Width stays at 60 to keep Anchor=Top,Right,Bottom layout stable;
        # font is 12pt bold so "Done" fits without wrapping.
        $script:btnNext.Text = "Done"
        $script:btnNext.BackColor = $bgRun
        $script:btnNext.ForeColor = [System.Drawing.Color]::White
        $script:btnNext.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
        $script:btnNext.Enabled = $manualSet
    } else {
        $script:btnNext.Text = ">"
        $script:btnNext.BackColor = $bgButton
        $script:btnNext.ForeColor = $fgText
        $script:btnNext.Font = New-Object System.Drawing.Font("Segoe UI", 28, [System.Drawing.FontStyle]::Bold)
        $script:btnNext.Enabled = $manualSet
    }

    # Hint label tells the operator why forward is blocked
    if ($null -ne $script:lblManualHint) {
        if ($manualSet) {
            $script:lblManualHint.Visible = $false
        } else {
            $script:lblManualHint.Text = "Set Phase Status to advance"
            $script:lblManualHint.Visible = $true
        }
    }
}

function Update-StatusBadges {
    if ($null -eq $script:lblAutoStatus) { return }
    if ($script:currentPhaseIndex -lt 0 -or $script:currentPhaseIndex -ge $script:phasesOrdered.Count) {
        $script:lblAutoStatus.Text = "  Auto: -"
        $script:lblAutoStatus.BackColor = $bgPanel
        $script:lblManualStatus.Text = "  Manual: -"
        $script:lblManualStatus.BackColor = $bgPanel
        return
    }
    $phase = $script:phasesOrdered[$script:currentPhaseIndex]
    $st = $script:phaseStatus[$phase.ID]
    $script:lblAutoStatus.Text = "  Auto:  $($st.Auto)  "
    $script:lblAutoStatus.BackColor = Get-StatusColor $st.Auto
    $script:lblManualStatus.Text = "  Manual:  $($st.Manual)  "
    $script:lblManualStatus.BackColor = Get-StatusColor $st.Manual
}

function Update-PhaseView {
    if ($null -eq $script:lblPhaseHeader) { return }
    if ($script:currentPhaseIndex -lt 0 -or $script:currentPhaseIndex -ge $script:phasesOrdered.Count) {
        $script:lblPhaseHeader.Text = "  (no phase loaded)"
        $script:phaseHeaderPanel.BackColor = $bgPanel
        $script:lblPhaseIndex.Text = "- / -"
        $script:txtInstruction.Text = ""
        $script:lstSteps.Items.Clear()
        Update-NavButtons
        Update-StatusBadges
        return
    }
    $phase = $script:phasesOrdered[$script:currentPhaseIndex]

    $script:lblPhaseHeader.Text = "  $($phase.ID)    $($phase.Label)"
    $clr = $script:phaseColors[$phase.Color]
    if (-not $clr) { $clr = $script:phaseColors["Gray"] }
    $script:phaseHeaderPanel.BackColor = $clr
    $script:lblPhaseIndex.Text = "$($script:currentPhaseIndex + 1) / $($script:phasesOrdered.Count)"

    # Instruction text
    $instrPath = Join-Path (Join-Path $script:currentProfile.Path "instructions") "$($phase.ID).txt"
    if (Test-Path $instrPath) {
        try {
            $txt = [System.IO.File]::ReadAllText($instrPath, [System.Text.Encoding]::UTF8)
            if ($null -eq $txt) { $txt = "" }
            # Normalize line endings to CRLF for WinForms TextBox
            $txt = $txt -replace "`r`n", "`n"
            $txt = $txt -replace "`r", "`n"
            $txt = $txt -replace "`n", "`r`n"
            $script:txtInstruction.Text = $txt
            $script:txtInstruction.SelectionStart = 0
            $script:txtInstruction.SelectionLength = 0
            $script:txtInstruction.ScrollToCaret()
        } catch {
            $script:txtInstruction.Text = "(failed to read instructions/$($phase.ID).txt: $_)"
        }
    } else {
        $script:txtInstruction.Text = "(no instruction file at instructions/$($phase.ID).txt)"
    }

    # Steps preview
    $script:lstSteps.Items.Clear()
    foreach ($s in $phase.Steps) {
        $line = ("  {0,2}. {1,-10} {2}" -f $s.StepNo, $s.Action, $s.Value)
        if ($s.Note) { $line += "    -- $($s.Note)" }
        $script:lstSteps.Items.Add($line) | Out-Null
    }

    # Copy Values button label reflects how many distinct $VarName
    # references this phase has. Stays clickable even when zero so the
    # operator can still open the dialog and toggle "Show all" to access
    # values not referenced from procedure.csv (since v1.3.0).
    if ($null -ne $script:btnCopyValues) {
        $phaseRefs = @(Get-PhaseReferencedVariables -PhaseID $phase.ID)
        $allVars   = @(Get-AllProfileVariables)
        $n = $phaseRefs.Count
        $totalN = $allVars.Count
        $script:btnCopyValues.Text = if ($n -eq 0) { "Copy Values..." } else { "Copy Values ($n)..." }
        # Disable only when the entire profile has no resolvable values
        # for this PC (nothing to show even with Show all on).
        $script:btnCopyValues.Enabled = ($totalN -gt 0)
    }

    Update-NavButtons
    Update-StatusBadges
}

function Set-AllControlsEnabled {
    param([bool]$Enabled)
    if ($null -ne $script:btnRunPhase)    { $script:btnRunPhase.Enabled    = $Enabled }
    if ($null -ne $script:btnScreenshot)  { $script:btnScreenshot.Enabled  = $Enabled }
    if ($null -ne $script:btnPhaseStatus) { $script:btnPhaseStatus.Enabled = $Enabled }
    if ($null -ne $script:btnCopyValues)  { $script:btnCopyValues.Enabled  = $Enabled }
    if ($null -ne $script:btnPrev)        { $script:btnPrev.Enabled        = $Enabled }
    if ($null -ne $script:btnNext)        { $script:btnNext.Enabled        = $Enabled }
    if ($Enabled) { Update-NavButtons }
}

# ========================================
# Phase Status dialog (operator judgment)
# ========================================
function Show-PhaseStatusDialog {
    param([string]$PhaseID, [string]$PhaseLabel)
    $st = $script:phaseStatus[$PhaseID]

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Pianist - Phase Status"
    $dlg.ClientSize = New-Object System.Drawing.Size(440, 320)
    $dlg.StartPosition = "CenterParent"
    $dlg.BackColor = $bgPanel
    $dlg.ForeColor = $fgText
    $dlg.FormBorderStyle = "FixedDialog"
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "$PhaseID    $PhaseLabel"
    $lbl.Location = New-Object System.Drawing.Point(16, 12)
    $lbl.Size = New-Object System.Drawing.Size(400, 24)
    $lbl.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $lbl.ForeColor = $fgHeader
    $dlg.Controls.Add($lbl)

    $lbl2 = New-Object System.Windows.Forms.Label
    $lbl2.Text = "Operator judgment - did this phase achieve its goal?"
    $lbl2.Location = New-Object System.Drawing.Point(16, 42)
    $lbl2.Size = New-Object System.Drawing.Size(400, 20)
    $lbl2.ForeColor = $fgDim
    $dlg.Controls.Add($lbl2)

    $rbList = @()
    $opts = @(
        @{ Tag = "OK";      Text = "OK             achieved";              Y =  72 }
        @{ Tag = "Warning"; Text = "Warning   achieved with concerns";     Y =  98 }
        @{ Tag = "Error";   Text = "Error        not achieved";            Y = 124 }
        @{ Tag = "Skip";    Text = "Skip          intentionally skipped";  Y = 150 }
    )
    foreach ($o in $opts) {
        $rb = New-Object System.Windows.Forms.RadioButton
        $rb.Text = $o.Text
        $rb.Location = New-Object System.Drawing.Point(24, $o.Y)
        $rb.Size = New-Object System.Drawing.Size(380, 22)
        $rb.ForeColor = $fgText
        $rb.Tag = $o.Tag
        if ($st.Manual -eq $o.Tag) { $rb.Checked = $true }
        $dlg.Controls.Add($rb)
        $rbList += $rb
    }

    $lblNote = New-Object System.Windows.Forms.Label
    $lblNote.Text = "Note (optional):"
    $lblNote.Location = New-Object System.Drawing.Point(16, 184)
    $lblNote.Size = New-Object System.Drawing.Size(400, 20)
    $lblNote.ForeColor = $fgDim
    $dlg.Controls.Add($lblNote)

    $txtNote = New-Object System.Windows.Forms.TextBox
    $txtNote.Location = New-Object System.Drawing.Point(16, 204)
    $txtNote.Size = New-Object System.Drawing.Size(400, 60)
    $txtNote.Multiline = $true
    $txtNote.BackColor = $bgInput
    $txtNote.ForeColor = $fgText
    $txtNote.BorderStyle = "FixedSingle"
    $txtNote.Text = $st.Note
    $dlg.Controls.Add($txtNote)

    $btnSave = New-PianistButton -Text "Save" -X 218 -Y 276 -Width 96 -Height 32 -BgColor $bgRun -FgColor ([System.Drawing.Color]::White)
    $btnSave.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $dlg.Controls.Add($btnSave)

    $btnCancel = New-PianistButton -Text "Cancel" -X 320 -Y 276 -Width 96 -Height 32
    $dlg.Controls.Add($btnCancel)

    $script:_psd_PhaseID = $PhaseID
    $script:_psd_radios  = $rbList
    $script:_psd_note    = $txtNote
    $script:_psd_dlg     = $dlg

    $btnSave.Add_Click({
        $picked = $null
        foreach ($c in $script:_psd_radios) {
            if ($c.Checked) { $picked = $c.Tag; break }
        }
        if ($picked) {
            $pid2 = $script:_psd_PhaseID
            $script:phaseStatus[$pid2].Manual = $picked
            $script:phaseStatus[$pid2].Note   = $script:_psd_note.Text
            $msgNote = if ($script:_psd_note.Text) { " ($($script:_psd_note.Text))" } else { "" }
            Write-PianistLog "Phase $pid2 Manual=$picked$msgNote" "OK"
        }
        $script:_psd_dlg.Close()
    })
    $btnCancel.Add_Click({ $script:_psd_dlg.Close() })

    [void]$dlg.ShowDialog()
    $dlg.Dispose()
    Update-StatusBadges
    # Manual change may unlock the > / Done button
    Update-NavButtons
}

# ========================================
# Copy Values: per-Phase variable picker (since v1.2.0)
# Auto-discovers $VarName references in this phase's procedure rows and
# offers a per-row [Copy] button that places the resolved value (decrypted
# if ENC:) onto the clipboard.
# ========================================
function Get-PhaseReferencedVariables {
    param([string]$PhaseID)
    if ($null -eq $script:currentProfile) { return @() }

    $names = New-Object System.Collections.Generic.HashSet[string]
    $regex = [regex]'\$([A-Za-z_][A-Za-z0-9_]*)'
    foreach ($step in $script:currentProfile.Procedure) {
        if ($step.PhaseID -ne $PhaseID) { continue }
        foreach ($field in @($step.Value, $step.Note)) {
            if ([string]::IsNullOrEmpty($field)) { continue }
            foreach ($m in $regex.Matches($field)) {
                $null = $names.Add($m.Groups[1].Value)
            }
        }
    }

    $dict = $script:currentProfile.ValuesDict
    $result = @()
    foreach ($n in ($names | Sort-Object)) {
        $resolved = $dict.ContainsKey($n)
        $val      = if ($resolved) { [string]$dict[$n] } else { '' }
        $result  += [PSCustomObject]@{
            Name     = $n
            Value    = $val
            Resolved = $resolved
        }
    }
    # Plain return - callers must wrap with @() if they need guaranteed
    # array semantics for 1-element results. The ,$result trick was
    # avoided because it interacts poorly with @() at the call site
    # (collapses N-element results to a single outer wrap).
    return $result
}

# Returns every variable in values.csv that resolves to a value for the
# current PC (including * fallback). Used by the "Show all" toggle in the
# Copy Values dialog so operators can grab values that aren't referenced
# from procedure.csv (e.g. read-aloud serials, side info to paste into
# tools outside Pianist's automation scope).
function Get-AllProfileVariables {
    if ($null -eq $script:currentProfile) { return @() }
    $dict = $script:currentProfile.ValuesDict
    if ($null -eq $dict -or $dict.Count -eq 0) { return @() }

    $result = @()
    foreach ($name in ($dict.Keys | Sort-Object)) {
        $result += [PSCustomObject]@{
            Name     = $name
            Value    = [string]$dict[$name]
            Resolved = $true
        }
    }
    return $result
}

# Builds the variable rows inside the dialog's scrolling panel.
# Called both at initial render and on Show-all toggle changes.
# Build a single variable row as a self-contained Panel. Using a container
# Panel per row (rather than absolute-positioning labels onto the outer
# scrollable Panel) avoids Label AutoSize / repaint quirks that caused
# 2nd-and-later name labels to disappear in absolute-positioned rendering.
function New-PianistVariableRow {
    param([PSCustomObject]$Var)

    $row = New-Object System.Windows.Forms.Panel
    $row.Size = New-Object System.Drawing.Size(548, 36)
    $row.Margin = New-Object System.Windows.Forms.Padding(0, 2, 0, 2)
    $row.BackColor = $bgGrid

    $nameLbl = New-Object System.Windows.Forms.Label
    $nameLbl.AutoSize = $false
    $nameLbl.Size = New-Object System.Drawing.Size(140, 28)
    $nameLbl.Location = New-Object System.Drawing.Point(4, 4)
    $nameLbl.Font = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Bold)
    $nameLbl.ForeColor = $fgHeader
    $nameLbl.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $nameLbl.Text = '$' + $Var.Name
    $null = $row.Controls.Add($nameLbl)

    $valLbl = New-Object System.Windows.Forms.Label
    $valLbl.AutoSize = $false
    $valLbl.Size = New-Object System.Drawing.Size(320, 28)
    $valLbl.Location = New-Object System.Drawing.Point(148, 4)
    $valLbl.Font = New-Object System.Drawing.Font("Consolas", 9)
    $valLbl.AutoEllipsis = $true
    $valLbl.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    if ($Var.Resolved) {
        $valLbl.ForeColor = $fgText
        $valLbl.Text = $Var.Value
    } else {
        $valLbl.ForeColor = $fgDim
        $valLbl.Text = "(undefined - not in values.csv)"
    }
    $null = $row.Controls.Add($valLbl)

    $copyBtn = New-PianistButton -Text "Copy" -X 472 -Y 4 -Width 70 -Height 28
    if (-not $Var.Resolved) { $copyBtn.Enabled = $false }
    # Bind the row's data via Tag so each button reads from itself
    # (PowerShell closures capture by ref otherwise -> all buttons would
    # copy the last loop iteration's value).
    $copyBtn.Tag = [PSCustomObject]@{ Name = $Var.Name; Value = $Var.Value }
    $copyBtn.Add_Click({
        $tag = $this.Tag
        try {
            if ([string]::IsNullOrEmpty($tag.Value)) {
                [System.Windows.Forms.Clipboard]::Clear()
            } else {
                [System.Windows.Forms.Clipboard]::SetText($tag.Value)
            }
            Write-PianistLog ("Copied `${0} to clipboard" -f $tag.Name) "OK"
        } catch {
            Write-PianistLog ("Copy failed for `${0}: {1}" -f $tag.Name, $_) "ERROR"
        }
    })
    $null = $row.Controls.Add($copyBtn)

    return $row
}

function Update-PianistVariablesPanel {
    if ($null -eq $script:_pvd_panel) { return }
    $script:_pvd_panel.SuspendLayout()
    $script:_pvd_panel.Controls.Clear()

    $vars = if ($script:_pvd_cb.Checked) {
        $script:_pvd_allVars
    } else {
        $script:_pvd_phaseVars
    }
    $emptyMsg = if ($script:_pvd_cb.Checked) {
        "(No values defined for this PC in values.csv)"
    } else {
        "(No `$VarName references in this phase's procedure rows. Toggle 'Show all' to see every value defined for this PC.)"
    }

    if (@($vars).Count -eq 0) {
        $emptyLbl = New-Object System.Windows.Forms.Label
        $emptyLbl.AutoSize = $false
        $emptyLbl.Text = $emptyMsg
        $emptyLbl.Margin = New-Object System.Windows.Forms.Padding(8, 8, 8, 8)
        $emptyLbl.Size = New-Object System.Drawing.Size(540, 60)
        $emptyLbl.ForeColor = $fgDim
        $null = $script:_pvd_panel.Controls.Add($emptyLbl)
    } else {
        foreach ($v in $vars) {
            $row = New-PianistVariableRow -Var $v
            $null = $script:_pvd_panel.Controls.Add($row)
        }
    }

    $script:_pvd_panel.ResumeLayout($true)
    $script:_pvd_panel.PerformLayout()
}

function Show-PianistVariablesDialog {
    param([string]$PhaseID, [string]$PhaseLabel)
    # @() wrap forces array semantics regardless of element count
    # (PowerShell auto-unwraps single-element arrays at pipeline boundary).
    $script:_pvd_phaseVars = @(Get-PhaseReferencedVariables -PhaseID $PhaseID)
    $script:_pvd_allVars   = @(Get-AllProfileVariables)

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Pianist - Copy Values"
    $dlg.ClientSize = New-Object System.Drawing.Size(600, 460)
    $dlg.StartPosition = "CenterParent"
    $dlg.BackColor = $bgPanel
    $dlg.ForeColor = $fgText
    $dlg.FormBorderStyle = "FixedDialog"
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.AutoSize = $false
    $lbl.Text = "$PhaseID    $PhaseLabel"
    $lbl.Location = New-Object System.Drawing.Point(16, 12)
    $lbl.Size = New-Object System.Drawing.Size(560, 24)
    $lbl.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $lbl.ForeColor = $fgHeader
    $null = $dlg.Controls.Add($lbl)

    $lbl2 = New-Object System.Windows.Forms.Label
    $lbl2.AutoSize = $false
    $lbl2.Text = "Click [Copy] to place the value on the clipboard."
    $lbl2.Location = New-Object System.Drawing.Point(16, 40)
    $lbl2.Size = New-Object System.Drawing.Size(560, 20)
    $lbl2.ForeColor = $fgDim
    $null = $dlg.Controls.Add($lbl2)

    # Show-all toggle: OFF (default) shows only $VarName references in the
    # current phase's procedure rows. ON shows every variable in values.csv
    # resolved for this PC (handy for read-aloud or paste into tools outside
    # Pianist's automation scope).
    $cb = New-Object System.Windows.Forms.CheckBox
    $cb.Text = "Show all values for this PC"
    $cb.Location = New-Object System.Drawing.Point(16, 66)
    $cb.Size = New-Object System.Drawing.Size(560, 22)
    $cb.ForeColor = $fgText
    $cb.BackColor = $bgPanel
    $cb.FlatStyle = "Flat"
    $cb.Checked = $false
    $null = $dlg.Controls.Add($cb)
    $script:_pvd_cb = $cb

    # FlowLayoutPanel TopDown so we don't have to manage row Y coordinates.
    # Each row is a self-contained Panel (built by New-PianistVariableRow)
    # which avoids Label AutoSize quirks that previously made later rows'
    # name labels invisible.
    $panel = New-Object System.Windows.Forms.FlowLayoutPanel
    $panel.Location = New-Object System.Drawing.Point(16, 92)
    $panel.Size = New-Object System.Drawing.Size(568, 320)
    $panel.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown
    $panel.WrapContents = $false
    $panel.AutoScroll = $true
    $panel.BackColor = $bgGrid
    $panel.BorderStyle = "FixedSingle"
    $null = $dlg.Controls.Add($panel)
    $script:_pvd_panel = $panel

    $cb.Add_CheckedChanged({ Update-PianistVariablesPanel })
    Update-PianistVariablesPanel  # initial render

    $btnClose = New-PianistButton -Text "Close" -X 504 -Y 420 -Width 80 -Height 30
    $script:_pvd_dlg = $dlg
    $btnClose.Add_Click({ $script:_pvd_dlg.Close() })
    $null = $dlg.Controls.Add($btnClose)

    [void]$dlg.ShowDialog()
    $dlg.Dispose()
    $script:_pvd_panel = $null
    $script:_pvd_cb    = $null
    $script:_pvd_dlg   = $null
}

# ========================================
# Profile selection dialog (used when multiple candidates)
# ========================================
function Show-PianistProfileSelector {
    param([array]$Items)
    if ($Items.Count -eq 0) { return $null }
    if ($Items.Count -eq 1) { return $Items[0] }

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Pianist - Select Profile"
    $dlg.ClientSize = New-Object System.Drawing.Size(540, 220)
    $dlg.StartPosition = "CenterScreen"
    $dlg.BackColor = $bgPanel
    $dlg.ForeColor = $fgText
    $dlg.FormBorderStyle = "FixedDialog"
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "Select the Pianist profile to execute:"
    $lbl.Location = New-Object System.Drawing.Point(16, 14)
    $lbl.Size = New-Object System.Drawing.Size(500, 20)
    $lbl.ForeColor = $fgHeader
    $dlg.Controls.Add($lbl)

    $cmb = New-Object System.Windows.Forms.ComboBox
    $cmb.Location = New-Object System.Drawing.Point(16, 42)
    $cmb.Size = New-Object System.Drawing.Size(508, 26)
    $cmb.DropDownStyle = "DropDownList"
    $cmb.BackColor = $bgInput
    $cmb.ForeColor = $fgText
    $cmb.FlatStyle = "Flat"
    foreach ($it in $Items) {
        $disp = "[$($it.Group)] $($it.ProfileName)"
        if ($it.Description) { $disp += " - $($it.Description)" }
        [void]$cmb.Items.Add($disp)
    }
    $cmb.SelectedIndex = 0
    $dlg.Controls.Add($cmb)

    $lblHint = New-Object System.Windows.Forms.Label
    $lblHint.Text = "Profile-driven runs (Profile CSV with Segment) auto-skip this picker."
    $lblHint.Location = New-Object System.Drawing.Point(16, 78)
    $lblHint.Size = New-Object System.Drawing.Size(508, 40)
    $lblHint.ForeColor = $fgDim
    $dlg.Controls.Add($lblHint)

    $btnOk = New-PianistButton -Text "Open Profile" -X 308 -Y 168 -Width 130 -Height 32 -BgColor $bgRun -FgColor ([System.Drawing.Color]::White)
    $btnOk.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $dlg.Controls.Add($btnOk)

    $btnCxl = New-PianistButton -Text "Cancel" -X 444 -Y 168 -Width 80 -Height 32
    $dlg.Controls.Add($btnCxl)

    $script:_psel_idx = -1
    $script:_psel_dlg = $dlg
    $script:_psel_cmb = $cmb

    $btnOk.Add_Click({
        $script:_psel_idx = $script:_psel_cmb.SelectedIndex
        $script:_psel_dlg.Close()
    })
    $btnCxl.Add_Click({
        $script:_psel_idx = -1
        $script:_psel_dlg.Close()
    })

    [void]$dlg.ShowDialog()
    $dlg.Dispose()

    if ($script:_psel_idx -lt 0 -or $script:_psel_idx -ge $Items.Count) { return $null }
    return $Items[$script:_psel_idx]
}


# ============================================================
# Module Entry: kernel-driven flow (template-style steps 1-6)
# ============================================================

Write-Host ""
Show-Separator
Write-Host "Pianist" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# ========================================
# Step 1: Load pianist_list.csv
# Import-ModuleCsv auto-filters by $env:FABRIQ_SEGMENT
# ========================================
$csvPath = Join-Path $PSScriptRoot "pianist_list.csv"
$enabledItems = Import-ModuleCsv -Path $csvPath -FilterEnabled `
    -RequiredColumns @("Enabled", "ProfileName", "Group", "Description", "Segment")

if ($null -eq $enabledItems) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load pianist_list.csv")
}
if ($enabledItems.Count -eq 0) {
    return (New-ModuleResult -Status "Skipped" -Message "No enabled profile entries match the current segment filter")
}


# ========================================
# Step 2: Validate referenced profile folders
# ========================================
$missing = @()
foreach ($p in $enabledItems) {
    $pPath = Join-Path $script:profilesRoot $p.ProfileName
    if (-not (Test-Path $pPath)) {
        $missing += $p.ProfileName
        continue
    }
    $procPath = Join-Path $pPath "procedure.csv"
    if (-not (Test-Path $procPath)) {
        $missing += "$($p.ProfileName) (missing procedure.csv)"
    }
}
if ($missing.Count -gt 0) {
    Show-Error "Missing profile resources:`n  $($missing -join "`n  ")"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Missing profiles: $($missing -join ', ')")
}


# ========================================
# Step 3: Dry-run summary
# ========================================
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "Candidate Profiles" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""
foreach ($p in $enabledItems) {
    Write-Host "  [$($p.Group)] $($p.ProfileName)" -ForegroundColor Yellow
    if ($p.Description) {
        Write-Host "    $($p.Description)" -ForegroundColor DarkGray
    }
    Write-Host ""
}
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""


# ========================================
# Step 4: Confirm execution
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Open Pianist for the candidate profile(s) above?"
if ($null -ne $cancelResult) { return $cancelResult }
Write-Host ""


# ========================================
# Step 5: Resolve which profile to run
# - Single candidate (segment-filtered) -> use directly
# - Multiple candidates                 -> show selector
# ========================================
$selected = Show-PianistProfileSelector -Items $enabledItems
if ($null -eq $selected) {
    Show-Info "Profile selection cancelled"
    Write-Host ""
    return (New-ModuleResult -Status "Cancelled" -Message "Profile selection cancelled")
}
Show-Info "Loading profile: $($selected.ProfileName)"

$script:currentProfile = Load-PianistProfileData -ProfileName $selected.ProfileName
if ($null -eq $script:currentProfile) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load profile [$($selected.ProfileName)]")
}

Build-PhasesOrdered
if ($script:phasesOrdered.Count -eq 0) {
    return (New-ModuleResult -Status "Error" -Message "Profile [$($selected.ProfileName)] has no phases")
}


# ========================================
# Step 6: Build the main GUI (absolute positioning + Anchor)
# ========================================
$form = New-Object System.Windows.Forms.Form
$form.Text = "Pianist - $($script:currentProfile.Meta.label)"
$form.ClientSize = New-Object System.Drawing.Size(1080, 720)
$form.MinimumSize = New-Object System.Drawing.Size(900, 600)
$form.StartPosition = "CenterScreen"
$form.BackColor = $bgDark
$form.ForeColor = $fgText
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

# ---- Top bar (Y=0, H=64) ----
$topBar = New-Object System.Windows.Forms.Panel
$topBar.Location = New-Object System.Drawing.Point(0, 0)
$topBar.Size = New-Object System.Drawing.Size(1080, 64)
$topBar.Anchor = "Top,Left,Right"
$topBar.BackColor = $bgPanel
$null = $form.Controls.Add($topBar)

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "Pianist"
$lblTitle.Location = New-Object System.Drawing.Point(16, 6)
$lblTitle.Size = New-Object System.Drawing.Size(110, 28)
$lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$lblTitle.ForeColor = $fgHeader
$null = $topBar.Controls.Add($lblTitle)

$lblProfile = New-Object System.Windows.Forms.Label
$lblProfile.Text = "Profile:  $($script:currentProfile.Name)"
$lblProfile.Location = New-Object System.Drawing.Point(140, 12)
$lblProfile.Size = New-Object System.Drawing.Size(640, 20)
$lblProfile.ForeColor = $fgText
$null = $topBar.Controls.Add($lblProfile)

$lblPhaseIndex = New-Object System.Windows.Forms.Label
$lblPhaseIndex.Text = "- / -"
$lblPhaseIndex.Location = New-Object System.Drawing.Point(940, 8)
$lblPhaseIndex.Size = New-Object System.Drawing.Size(120, 28)
$lblPhaseIndex.Anchor = "Top,Right"
$lblPhaseIndex.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$lblPhaseIndex.ForeColor = $fgHeader
$lblPhaseIndex.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$null = $topBar.Controls.Add($lblPhaseIndex)
$script:lblPhaseIndex = $lblPhaseIndex

$lblTarget = New-Object System.Windows.Forms.Label
$lblTarget.Text = "$($script:currentProfile.Meta.target_app)  -  $($script:currentProfile.Meta.description)"
$lblTarget.Location = New-Object System.Drawing.Point(140, 38)
$lblTarget.Size = New-Object System.Drawing.Size(900, 20)
$lblTarget.Anchor = "Top,Left,Right"
$lblTarget.ForeColor = $fgDim
$null = $topBar.Controls.Add($lblTarget)

# ---- Left navigation arrow ----
$btnPrev = New-Object System.Windows.Forms.Button
$btnPrev.Text = "<"
$btnPrev.Location = New-Object System.Drawing.Point(0, 64)
$btnPrev.Size = New-Object System.Drawing.Size(60, 516)
$btnPrev.Anchor = "Top,Left,Bottom"
$btnPrev.FlatStyle = "Flat"
$btnPrev.FlatAppearance.BorderColor = $gridLine
$btnPrev.FlatAppearance.MouseOverBackColor = $bgButtonHov
$btnPrev.BackColor = $bgButton
$btnPrev.ForeColor = $fgText
$btnPrev.Font = New-Object System.Drawing.Font("Segoe UI", 28, [System.Drawing.FontStyle]::Bold)
$btnPrev.Cursor = [System.Windows.Forms.Cursors]::Hand
$null = $form.Controls.Add($btnPrev)
$script:btnPrev = $btnPrev

# ---- Right navigation arrow ----
$btnNext = New-Object System.Windows.Forms.Button
$btnNext.Text = ">"
$btnNext.Location = New-Object System.Drawing.Point(1020, 64)
$btnNext.Size = New-Object System.Drawing.Size(60, 516)
$btnNext.Anchor = "Top,Right,Bottom"
$btnNext.FlatStyle = "Flat"
$btnNext.FlatAppearance.BorderColor = $gridLine
$btnNext.FlatAppearance.MouseOverBackColor = $bgButtonHov
$btnNext.BackColor = $bgButton
$btnNext.ForeColor = $fgText
$btnNext.Font = New-Object System.Drawing.Font("Segoe UI", 28, [System.Drawing.FontStyle]::Bold)
$btnNext.Cursor = [System.Windows.Forms.Cursors]::Hand
$null = $form.Controls.Add($btnNext)
$script:btnNext = $btnNext

# ---- Phase header bar (color-coded) ----
$phaseHeaderPanel = New-Object System.Windows.Forms.Panel
$phaseHeaderPanel.Location = New-Object System.Drawing.Point(60, 64)
$phaseHeaderPanel.Size = New-Object System.Drawing.Size(960, 56)
$phaseHeaderPanel.Anchor = "Top,Left,Right"
$phaseHeaderPanel.BackColor = $bgPanel
$null = $form.Controls.Add($phaseHeaderPanel)
$script:phaseHeaderPanel = $phaseHeaderPanel

$lblPhaseHeader = New-Object System.Windows.Forms.Label
$lblPhaseHeader.Location = New-Object System.Drawing.Point(0, 0)
$lblPhaseHeader.Size = New-Object System.Drawing.Size(960, 56)
$lblPhaseHeader.Anchor = "Top,Left,Right,Bottom"
$lblPhaseHeader.Text = "  (loading...)"
$lblPhaseHeader.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$lblPhaseHeader.ForeColor = [System.Drawing.Color]::White
$lblPhaseHeader.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$null = $phaseHeaderPanel.Controls.Add($lblPhaseHeader)
$script:lblPhaseHeader = $lblPhaseHeader

# ---- Instruction text (grows on resize) ----
$txtInstruction = New-Object System.Windows.Forms.TextBox
$txtInstruction.Location = New-Object System.Drawing.Point(60, 124)
$txtInstruction.Size = New-Object System.Drawing.Size(960, 220)
$txtInstruction.Anchor = "Top,Left,Right,Bottom"
$txtInstruction.Multiline = $true
$txtInstruction.ReadOnly = $true
$txtInstruction.ScrollBars = "Vertical"
$txtInstruction.WordWrap = $true
$txtInstruction.BackColor = $bgGrid
$txtInstruction.ForeColor = $fgText
$txtInstruction.Font = New-Object System.Drawing.Font("Segoe UI", 11)
$txtInstruction.BorderStyle = "FixedSingle"
$null = $form.Controls.Add($txtInstruction)
$script:txtInstruction = $txtInstruction

# ---- Steps preview list ----
$lblStepsHdr = New-Object System.Windows.Forms.Label
$lblStepsHdr.Text = "Steps in this phase (preview)"
$lblStepsHdr.Location = New-Object System.Drawing.Point(60, 350)
$lblStepsHdr.Size = New-Object System.Drawing.Size(400, 18)
$lblStepsHdr.Anchor = "Bottom,Left"
$lblStepsHdr.ForeColor = $fgDim
$null = $form.Controls.Add($lblStepsHdr)

$lstSteps = New-Object System.Windows.Forms.ListBox
$lstSteps.Location = New-Object System.Drawing.Point(60, 370)
$lstSteps.Size = New-Object System.Drawing.Size(960, 108)
$lstSteps.Anchor = "Bottom,Left,Right"
$lstSteps.BackColor = $bgGrid
$lstSteps.ForeColor = $fgText
$lstSteps.Font = New-Object System.Drawing.Font("Consolas", 9)
$lstSteps.BorderStyle = "FixedSingle"
$lstSteps.IntegralHeight = $false
$lstSteps.SelectionMode = "One"
$null = $form.Controls.Add($lstSteps)
$script:lstSteps = $lstSteps

# ---- Action buttons row ----
$btnRunPhase = New-PianistButton -Text "Run Phase" -X 72 -Y 492 -Width 180 -Height 36 -BgColor $bgRun -FgColor ([System.Drawing.Color]::White) `
    -Font (New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold))
$btnRunPhase.Anchor = "Bottom,Left"
$null = $form.Controls.Add($btnRunPhase)
$script:btnRunPhase = $btnRunPhase

$btnScreenshot = New-PianistButton -Text "Screenshot" -X 264 -Y 492 -Width 160 -Height 36 `
    -Font (New-Object System.Drawing.Font("Segoe UI", 10))
$btnScreenshot.Anchor = "Bottom,Left"
$null = $form.Controls.Add($btnScreenshot)
$script:btnScreenshot = $btnScreenshot

$btnPhaseStatus = New-PianistButton -Text "Phase Status..." -X 436 -Y 492 -Width 160 -Height 36 -BgColor $bgAccent -FgColor ([System.Drawing.Color]::White) `
    -Font (New-Object System.Drawing.Font("Segoe UI", 10))
$btnPhaseStatus.Anchor = "Bottom,Left"
$null = $form.Controls.Add($btnPhaseStatus)
$script:btnPhaseStatus = $btnPhaseStatus

# Copy Values: opens a per-Phase variable picker with [Copy] buttons.
# Label suffix "(N)" reflects how many $VarName references the current
# phase has - updated in Update-PhaseView.
$btnCopyValues = New-PianistButton -Text "Copy Values..." -X 608 -Y 492 -Width 160 -Height 36 `
    -Font (New-Object System.Drawing.Font("Segoe UI", 10))
$btnCopyValues.Anchor = "Bottom,Left"
$null = $form.Controls.Add($btnCopyValues)
$script:btnCopyValues = $btnCopyValues

# ---- Status badges ----
$lblAutoStatus = New-Object System.Windows.Forms.Label
$lblAutoStatus.Text = "  Auto: -"
$lblAutoStatus.Location = New-Object System.Drawing.Point(72, 540)
$lblAutoStatus.Size = New-Object System.Drawing.Size(280, 26)
$lblAutoStatus.Anchor = "Bottom,Left"
$lblAutoStatus.BackColor = $bgPanel
$lblAutoStatus.ForeColor = [System.Drawing.Color]::White
$lblAutoStatus.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$lblAutoStatus.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$null = $form.Controls.Add($lblAutoStatus)
$script:lblAutoStatus = $lblAutoStatus

$lblManualStatus = New-Object System.Windows.Forms.Label
$lblManualStatus.Text = "  Manual: -"
$lblManualStatus.Location = New-Object System.Drawing.Point(364, 540)
$lblManualStatus.Size = New-Object System.Drawing.Size(280, 26)
$lblManualStatus.Anchor = "Bottom,Left"
$lblManualStatus.BackColor = $bgPanel
$lblManualStatus.ForeColor = [System.Drawing.Color]::White
$lblManualStatus.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$lblManualStatus.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$null = $form.Controls.Add($lblManualStatus)
$script:lblManualStatus = $lblManualStatus

# Hint label - shown when current phase has no Manual judgment yet,
# explains why the > / Done button is greyed out.
$lblManualHint = New-Object System.Windows.Forms.Label
$lblManualHint.Text = ""
$lblManualHint.Location = New-Object System.Drawing.Point(656, 542)
$lblManualHint.Size = New-Object System.Drawing.Size(360, 22)
$lblManualHint.Anchor = "Bottom,Left,Right"
$lblManualHint.ForeColor = $bgWarn
$lblManualHint.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
$lblManualHint.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$lblManualHint.Visible = $false
$null = $form.Controls.Add($lblManualHint)
$script:lblManualHint = $lblManualHint

# ---- Bottom log ----
$logBox = New-Object System.Windows.Forms.RichTextBox
$logBox.Location = New-Object System.Drawing.Point(0, 580)
$logBox.Size = New-Object System.Drawing.Size(1080, 140)
$logBox.Anchor = "Bottom,Left,Right"
$logBox.BackColor = $bgGrid
$logBox.ForeColor = $fgText
$logBox.Font = New-Object System.Drawing.Font("Consolas", 9)
$logBox.ReadOnly = $true
$logBox.BorderStyle = "FixedSingle"
$null = $form.Controls.Add($logBox)
$script:logBox = $logBox


# ========================================
# Event handlers
# ========================================
$btnPrev.Add_Click({
    if ($script:currentPhaseIndex -gt 0) {
        Set-CurrentPhase ($script:currentPhaseIndex - 1)
    }
})

$btnNext.Add_Click({
    if ($script:phasesOrdered.Count -eq 0) { return }
    if ($script:currentPhaseIndex -eq ($script:phasesOrdered.Count - 1)) {
        # Last phase: ask to finish
        $ans = [System.Windows.Forms.MessageBox]::Show(
            "All phases reviewed. Finish and record to fabriq history?",
            "Pianist - Done",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        if ($ans -eq [System.Windows.Forms.DialogResult]::Yes) {
            $script:UserAction = "done"
        }
    } else {
        Set-CurrentPhase ($script:currentPhaseIndex + 1)
    }
})

$btnRunPhase.Add_Click({
    if ($script:currentPhaseIndex -lt 0 -or $script:currentPhaseIndex -ge $script:phasesOrdered.Count) { return }
    $phase = $script:phasesOrdered[$script:currentPhaseIndex]
    Invoke-PianistPhase -PhaseID $phase.ID
})

$btnScreenshot.Add_Click({
    $tag = if ($script:currentPhaseIndex -ge 0 -and $script:currentPhaseIndex -lt $script:phasesOrdered.Count) {
        "manual_$($script:phasesOrdered[$script:currentPhaseIndex].ID)"
    } else { "manual" }
    if (Invoke-PianistScreenshot -Tag $tag) {
        Write-PianistLog "Screenshot saved (tag=$tag)" "OK"
    }
})

$btnPhaseStatus.Add_Click({
    if ($script:currentPhaseIndex -lt 0 -or $script:currentPhaseIndex -ge $script:phasesOrdered.Count) { return }
    $phase = $script:phasesOrdered[$script:currentPhaseIndex]
    Show-PhaseStatusDialog -PhaseID $phase.ID -PhaseLabel $phase.Label
})

$btnCopyValues.Add_Click({
    if ($script:currentPhaseIndex -lt 0 -or $script:currentPhaseIndex -ge $script:phasesOrdered.Count) { return }
    $phase = $script:phasesOrdered[$script:currentPhaseIndex]
    Show-PianistVariablesDialog -PhaseID $phase.ID -PhaseLabel $phase.Label
})

$form.Add_FormClosing({
    param($sender, $e)
    if ($script:UserAction -ne "done" -and $script:UserAction -ne "cancel") {
        $ans = [System.Windows.Forms.MessageBox]::Show(
            "Cancel Pianist? Phase progress will be recorded as Cancelled.",
            "Pianist - Cancel",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        if ($ans -eq [System.Windows.Forms.DialogResult]::Yes) {
            $script:UserAction = "cancel"
        } else {
            $e.Cancel = $true
        }
    }
})


# ========================================
# Step 7: Show form (modeless) and run main loop
# ========================================
$script:wsShell = New-Object -ComObject WScript.Shell
$form.Show()

# Initial phase (default_phase if specified, else first)
$initialIdx = 0
$defaultPhase = $script:currentProfile.Meta.default_phase
if ($defaultPhase) {
    for ($i = 0; $i -lt $script:phasesOrdered.Count; $i++) {
        if ($script:phasesOrdered[$i].ID -eq $defaultPhase) { $initialIdx = $i; break }
    }
}
Set-CurrentPhase $initialIdx
Write-PianistLog "Pianist v1.0.0 ready. Profile [$($script:currentProfile.Name)] - $($script:phasesOrdered.Count) phases."

# Main wait loop - DoEvents polling until UserAction set
$script:UserAction = $null
while ($null -eq $script:UserAction) {
    [System.Windows.Forms.Application]::DoEvents()
    Start-Sleep -Milliseconds 50
}


# ========================================
# Step 8: Cleanup
# ========================================
try {
    if ($form.Visible) { $form.Close() }
    $form.Dispose()
} catch {}
if ($script:wsShell) {
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:wsShell) | Out-Null } catch {}
    $script:wsShell = $null
}


# ========================================
# Step 9: Aggregate phase status -> ModuleResult
# ========================================
if ($script:UserAction -eq "cancel") {
    Show-Info "Pianist cancelled by operator"
    Write-Host ""
    return (New-ModuleResult -Status "Cancelled" -Message "Operator cancelled Pianist for [$($script:currentProfile.Name)]")
}

$counts = [ordered]@{ OK = 0; Warning = 0; Error = 0; Skip = 0; Unset = 0 }
foreach ($p in $script:phasesOrdered) {
    $st = $script:phaseStatus[$p.ID].Manual
    switch ($st) {
        "OK"      { $counts.OK++ }
        "Warning" { $counts.Warning++ }
        "Error"   { $counts.Error++ }
        "Skip"    { $counts.Skip++ }
        default   { $counts.Unset++ }
    }
}

$aggStatus = if ($counts.Error -gt 0)                                     { "Error" }
             elseif ($counts.Warning -gt 0 -or $counts.Unset -gt 0)       { "Partial" }
             else                                                          { "Success" }

$verified = if ($counts.Error -eq 0 -and $counts.Warning -eq 0 -and $counts.Unset -eq 0) { $true }
            elseif ($counts.Error -gt 0)                                                  { $false }
            else                                                                           { $null }

$msg = "Profile [$($script:currentProfile.Name)] phases: OK=$($counts.OK) Warning=$($counts.Warning) Error=$($counts.Error) Skip=$($counts.Skip) Unset=$($counts.Unset)"

Show-Info $msg
Write-Host ""

if ($null -eq $verified) {
    return (New-ModuleResult -Status $aggStatus -Message $msg)
}
return (New-ModuleResult -Status $aggStatus -Message $msg -Verified $verified)
