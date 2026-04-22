# ========================================
# Fabriq Status Monitor Window
# ========================================
# Launched as a separate process by main.ps1
# Reads status.json and displays PC info + execution status
# Usage: powershell -NoProfile -ExecutionPolicy Unrestricted -File .\kernel\ps1\status_monitor.ps1 -StatusFilePath ".\kernel\json\status.json"

param(
    [string]$StatusFilePath = ".\kernel\json\status.json",
    [string]$PulseFilePath = ".\kernel\json\art_pulse.txt",
    [string]$SentenceFilePath = ".\kernel\txt\art_sentences.txt",
    [string]$SilenceFlagPath = ".\kernel\txt\silence.flag"
)

# ========================================
# DPI Awareness (must be set BEFORE any Forms/Drawing operations)
# ========================================
# SetProcessDPIAware() makes Screen.Bounds return physical pixels,
# preventing screenshot cropping on scaled displays.
# Form dimensions are then scaled by the DPI factor below.
Add-Type -AssemblyName System.Drawing
Add-Type -TypeDefinition @"
    using System.Runtime.InteropServices;
    public class DPIUtil {
        [DllImport("user32.dll")]
        public static extern bool SetProcessDPIAware();
    }
"@ -ErrorAction SilentlyContinue
$null = [DPIUtil]::SetProcessDPIAware()

# Get DPI scale factor (96 DPI = 100% = scale 1.0)
$tmpG = [System.Drawing.Graphics]::FromHwnd([IntPtr]::Zero)
$script:dpiScale = $tmpG.DpiX / 96.0
$tmpG.Dispose()

Add-Type -AssemblyName System.Windows.Forms

# Hide the PowerShell console window (keep only the Forms window visible)
Add-Type -Name Win32 -Namespace Native -MemberDefinition @'
    [DllImport("kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
'@

# NoActivateForm: Form subclass that does not steal focus on show
# WndProc override returns MA_NOACTIVATE for WM_MOUSEACTIVATE so that
# clicks on ToolStrip buttons work on the first click without requiring
# the form to be activated first.
Add-Type -ReferencedAssemblies System.Windows.Forms -TypeDefinition @'
using System;
using System.Windows.Forms;
public class NoActivateForm : Form {
    protected override bool ShowWithoutActivation { get { return true; } }
    private const int WS_EX_NOACTIVATE = 0x08000000;
    private const int WM_MOUSEACTIVATE = 0x0021;
    private const int MA_NOACTIVATE = 3;
    protected override CreateParams CreateParams {
        get {
            CreateParams cp = base.CreateParams;
            cp.ExStyle |= WS_EX_NOACTIVATE;
            return cp;
        }
    }
    protected override void WndProc(ref Message m) {
        if (m.Msg == WM_MOUSEACTIVATE) {
            m.Result = (IntPtr)MA_NOACTIVATE;
            return;
        }
        base.WndProc(ref m);
    }
}
public class ClickThroughStatusStrip : StatusStrip {
    private const int WM_MOUSEACTIVATE = 0x0021;
    private const int MA_NOACTIVATE = 3;
    protected override void WndProc(ref Message m) {
        if (m.Msg == WM_MOUSEACTIVATE) {
            m.Result = (IntPtr)MA_NOACTIVATE;
            return;
        }
        base.WndProc(ref m);
    }
}
'@
$consoleHwnd = [Native.Win32]::GetConsoleWindow()
if ($consoleHwnd -ne [IntPtr]::Zero) {
    [Native.Win32]::ShowWindow($consoleHwnd, 0) | Out-Null  # SW_HIDE = 0
}

# ========================================
# Derive evidence directory from StatusFilePath
# ========================================
# IMPORTANT: Must be done BEFORE dot-sourcing common.ps1, because
# common.ps1 overwrites $script:StatusFilePath with a relative path.
$script:fabriqRoot = [System.IO.Path]::GetFullPath((Join-Path (Split-Path $StatusFilePath) "..\.."))
$script:gyotakuDir = Join-Path $script:fabriqRoot "evidence\gyotaku"

# ========================================
# Load common.ps1 (for Save-Screenshot)
# ========================================
. (Join-Path $PSScriptRoot "..\common.ps1")

# ========================================
# Evidence base path (from parent process via env var)
# ========================================
if (-not [string]::IsNullOrWhiteSpace($env:FABRIQ_EVIDENCE_BASE)) {
    $global:FabriqEvidenceBasePath = $env:FABRIQ_EVIDENCE_BASE
    $global:FabriqEvidenceRootPath = Split-Path $env:FABRIQ_EVIDENCE_BASE -Parent
}

# ========================================
# Color Definitions
# ========================================
$darkBg       = [System.Drawing.Color]::FromArgb(30, 30, 30)
$panelBg      = [System.Drawing.Color]::FromArgb(40, 40, 40)
$accentCyan   = [System.Drawing.Color]::FromArgb(0, 200, 200)
$textWhite    = [System.Drawing.Color]::White
$textGray     = [System.Drawing.Color]::FromArgb(160, 160, 160)
$successGreen = [System.Drawing.Color]::FromArgb(80, 220, 80)
$errorRed     = [System.Drawing.Color]::FromArgb(255, 80, 80)
$warnYellow   = [System.Drawing.Color]::FromArgb(255, 200, 0)

# ========================================
# Font Definitions
# ========================================
$fontNormal = New-Object System.Drawing.Font("Consolas", 9)
$fontBold   = New-Object System.Drawing.Font("Consolas", 9, [System.Drawing.FontStyle]::Bold)
$fontTitle  = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Bold)

# ========================================
# Art Display Configuration
# ========================================
$script:ART_RENDER_INTERVAL = 40    # render timer interval (ms) ~25fps
$script:ART_PULSE_INTERVAL  = 200   # pulse file poll interval (ms)
$script:ART_BURST_SPEED     = 8     # ms per character (triggered)
$script:ART_IDLE_SPEED      = 35    # ms per character (idle)
$script:ART_MAX_LINES       = 50    # max completed lines before trimming
$script:ART_LINE_HEIGHT     = 16    # line height in pixels (base)
$script:GLITCH_CHARS = @('_','#','@','!','^','~','`','|','{','}','[',']','<','>','/','?','+','=','*','0','1')

# Art Display Fonts & Brushes
$script:artFont     = New-Object System.Drawing.Font("Consolas", 8)
$script:artFontBold = New-Object System.Drawing.Font("Consolas", 8, [System.Drawing.FontStyle]::Bold)
$script:artLineH    = $script:ART_LINE_HEIGHT
$script:artBgColor       = $darkBg
$script:artCyanColor     = [System.Drawing.Color]::FromArgb(0, 130, 130)
$script:artDimGreen      = [System.Drawing.Color]::FromArgb(0, 140, 100)
$script:artDimGray       = [System.Drawing.Color]::FromArgb(60, 70, 60)
$script:artGlitchWhite   = [System.Drawing.Color]::FromArgb(255, 255, 255)
$script:artAlphaColor        = [System.Drawing.Color]::FromArgb(130, 80, 180)
$script:artAlphaDimGreen     = [System.Drawing.Color]::FromArgb(80, 50, 120)
$script:artAlphaDimGray      = [System.Drawing.Color]::FromArgb(50, 40, 55)
$script:artBgBrush           = New-Object System.Drawing.SolidBrush($darkBg)
$script:artCyanBrush         = New-Object System.Drawing.SolidBrush($script:artCyanColor)
$script:artDimGreenBrush     = New-Object System.Drawing.SolidBrush($script:artDimGreen)
$script:artDimGrayBrush      = New-Object System.Drawing.SolidBrush($script:artDimGray)
$script:artAlphaBrush        = New-Object System.Drawing.SolidBrush($script:artAlphaColor)
$script:artAlphaDimGreenBrush = New-Object System.Drawing.SolidBrush($script:artAlphaDimGreen)
$script:artAlphaDimGrayBrush  = New-Object System.Drawing.SolidBrush($script:artAlphaDimGray)

# Art Display State
$script:artRng = New-Object System.Random
$script:artDisplayLines = [System.Collections.ArrayList]::new()
$script:artCurrentText = ""
$script:artCursorPos = 0
$script:artState = "waiting"
$script:artTypeSpeed = $script:ART_BURST_SPEED
$script:artLastTypeTime = [DateTime]::Now
$script:artCursorVisible = $true
$script:artCursorBlinkTime = [DateTime]::Now
$script:artGlitchFrames = 0
$script:artFlashFrames = 0
$script:artFlashColor = $null
$script:artContinueOnSameLine = $false
$script:artLastPulseValue = 0
$script:artTriggerQueue = 0
$script:artBufferBitmap = $null
$script:artBufferGraphics = $null

# Art Display - Status polling state (for flash effects)
$script:artLastStatusWriteTime = [DateTime]::MinValue
$script:artLastDetailCount = 0
$script:artCurrentPhase = "idle"

# Art Display - Silence mode (toggled by presence of silence.flag)
$script:artSilent = $false

# Load Sentences
# Load art text and split into 3 granularity pools
$script:artByParagraph = @()
$script:artBySentence  = @()
$script:artByClause    = @()
$resolvedSentencePath = $SentenceFilePath
if (-not [System.IO.Path]::IsPathRooted($SentenceFilePath)) {
    $resolvedSentencePath = Join-Path (Get-Location) $SentenceFilePath
}
if (Test-Path $resolvedSentencePath) {
    $rawText = [System.IO.File]::ReadAllText($resolvedSentencePath, [System.Text.Encoding]::UTF8)
    # Paragraphs: split by newline (each line = one paragraph entry)
    $script:artByParagraph = @($rawText -split '\r?\n' | ForEach-Object { $_.Trim() } | Where-Object { $_.Length -ge 8 })
    # Sentences: split by Japanese period (keep the delimiter at end)
    $script:artBySentence = @($rawText -split '(?<=\u3002)' | ForEach-Object { $_.Trim() } | Where-Object { $_.Length -ge 8 })
    # Clauses: split by Japanese comma (keep the delimiter at end)
    $script:artByClause = @($rawText -split '(?<=\u3001)' | ForEach-Object { $_.Trim() } | Where-Object { $_.Length -ge 8 })
}
if ($script:artBySentence.Count -eq 0) {
    $fallback = @(
        "SYSTEM INITIALIZED.",
        "DATA STREAM ACTIVE.",
        "PROCESSING SIGNAL.",
        "KERNEL READY.",
        "CONFIGURATION LOADED.",
        "VERIFY SEQUENCE COMPLETE."
    )
    $script:artByParagraph = $fallback
    $script:artBySentence  = $fallback
    $script:artByClause    = $fallback
}

# ========================================
# Form Setup
# ========================================
$form = New-Object NoActivateForm
$form.Text = "Fabriq - Status Monitor"
# Scale form dimensions by DPI factor (designed at 96 DPI / 100%)
$form.Size = New-Object System.Drawing.Size(
    [int](750 * $script:dpiScale),
    [int](800 * $script:dpiScale)
)
$form.StartPosition = "Manual"
$form.Location = New-Object System.Drawing.Point(
    ([System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Right - [int](770 * $script:dpiScale)),
    [int](50 * $script:dpiScale)
)
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox = $false
$form.BackColor = $darkBg
$form.ForeColor = $textWhite
$form.Font = $fontNormal

# ========================================
# Main Layout (TableLayoutPanel: 2 rows, 2 columns)
# ========================================
$mainLayout = New-Object System.Windows.Forms.TableLayoutPanel
$mainLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
$mainLayout.RowCount = 2
$mainLayout.ColumnCount = 2
$mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 75))) | Out-Null
$mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 25))) | Out-Null
$mainLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50))) | Out-Null
$mainLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50))) | Out-Null
$mainLayout.Padding = New-Object System.Windows.Forms.Padding(6, 6, 6, 0)
$form.Controls.Add($mainLayout)

# ========================================
# Execution Summary Panel (Left Column)
# ========================================
$execGroup = New-Object System.Windows.Forms.GroupBox
$execGroup.Text = " Execution Summary "
$execGroup.Dock = [System.Windows.Forms.DockStyle]::Fill
$execGroup.ForeColor = $accentCyan
$execGroup.Font = $fontBold
$mainLayout.Controls.Add($execGroup, 0, 0)

$execLabel = New-Object System.Windows.Forms.RichTextBox
$execLabel.Dock = [System.Windows.Forms.DockStyle]::Fill
$execLabel.ForeColor = $textWhite
$execLabel.BackColor = $darkBg
$execLabel.Font = $fontNormal
$execLabel.ReadOnly = $true
$execLabel.BorderStyle = "None"
$execLabel.TabStop = $false
$execLabel.Text = "No execution data yet."
$execGroup.Controls.Add($execLabel)

# ========================================
# PC Info Panel (Right Column)
# ========================================
$pcInfoGroup = New-Object System.Windows.Forms.GroupBox
$pcInfoGroup.Text = " PC Info Comparison "
$pcInfoGroup.Dock = [System.Windows.Forms.DockStyle]::Fill
$pcInfoGroup.ForeColor = $accentCyan
$pcInfoGroup.Font = $fontBold
$mainLayout.Controls.Add($pcInfoGroup, 1, 0)

$pcInfoRtb = New-Object System.Windows.Forms.RichTextBox
$pcInfoRtb.Dock = [System.Windows.Forms.DockStyle]::Fill
$pcInfoRtb.ForeColor = $textWhite
$pcInfoRtb.BackColor = $darkBg
$pcInfoRtb.Font = $fontNormal
$pcInfoRtb.ReadOnly = $true
$pcInfoRtb.BorderStyle = "None"
$pcInfoRtb.TabStop = $false
$pcInfoRtb.Text = "Waiting for status data..."
$pcInfoGroup.Controls.Add($pcInfoRtb)

# ========================================
# Art Display Panel (Row 1, ColSpan=2)
# ========================================
$artGroup = New-Object System.Windows.Forms.GroupBox
$artGroup.Text = " Surkitinisme "
$artGroup.Dock = [System.Windows.Forms.DockStyle]::Fill
$artGroup.ForeColor = $accentCyan
$artGroup.Font = $fontBold
$mainLayout.Controls.Add($artGroup, 0, 1)
$mainLayout.SetColumnSpan($artGroup, 2)

$artCanvas = New-Object System.Windows.Forms.PictureBox
$artCanvas.Dock = [System.Windows.Forms.DockStyle]::Fill
$artCanvas.BackColor = $darkBg
$artGroup.Controls.Add($artCanvas)

# ========================================
# Status Bar
# ========================================
$statusBar = New-Object ClickThroughStatusStrip
$statusBar.BackColor = [System.Drawing.Color]::FromArgb(25, 25, 25)
# Screenshot button (leftmost item)
$btnScreenshot = New-Object System.Windows.Forms.ToolStripButton
$btnScreenshot.Text = "Screenshot"
$btnScreenshot.ForeColor = $accentCyan
$btnScreenshot.Font = New-Object System.Drawing.Font("Consolas", 9, [System.Drawing.FontStyle]::Bold)
$btnScreenshot.Margin = New-Object System.Windows.Forms.Padding(4, 2, 8, 0)
$statusBar.Items.Add($btnScreenshot) | Out-Null
# Skip button (only effective when the running module is in __ASYNC__ block)
$btnSkipAsync = New-Object System.Windows.Forms.ToolStripButton
$btnSkipAsync.Text = "Skip"
$btnSkipAsync.ForeColor = [System.Drawing.Color]::FromArgb(255, 170, 60)
$btnSkipAsync.Font = New-Object System.Drawing.Font("Consolas", 9, [System.Drawing.FontStyle]::Bold)
$btnSkipAsync.Margin = New-Object System.Windows.Forms.Padding(0, 2, 8, 0)
$btnSkipAsync.ToolTipText = "Request async module skip. Only effective for modules running after __ASYNC__ marker."
$statusBar.Items.Add($btnSkipAsync) | Out-Null
# Separator between button and status text
$statusSep = New-Object System.Windows.Forms.ToolStripSeparator
$statusBar.Items.Add($statusSep) | Out-Null
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.ForeColor = $textGray
$statusLabel.Text = "Waiting for data..."
$statusLabel.Spring = $true
$statusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$statusBar.Items.Add($statusLabel) | Out-Null
$emailLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$emailLabel.ForeColor = [System.Drawing.Color]::FromArgb(80, 80, 80)
$emailLabel.Text = "yuki.suzuki@suzugross.com"
$statusBar.Items.Add($emailLabel) | Out-Null
$form.Controls.Add($statusBar)

# ========================================
# Update Function
# ========================================
$script:lastWriteTime = [datetime]::MinValue

function Set-ColorizedText {
    param(
        [System.Windows.Forms.RichTextBox]$RichTextBox,
        [string]$Text
    )
    $RichTextBox.Text = $Text
    $RichTextBox.SelectAll()
    $RichTextBox.SelectionFont = $fontNormal
    $RichTextBox.SelectionColor = $textWhite

    # [OK] -> green
    $pos = 0
    while (($idx = $RichTextBox.Text.IndexOf("[OK]", $pos)) -ge 0) {
        $RichTextBox.Select($idx, 4)
        $RichTextBox.SelectionColor = $successGreen
        $pos = $idx + 4
    }

    # [!!] -> red
    $pos = 0
    while (($idx = $RichTextBox.Text.IndexOf("[!!]", $pos)) -ge 0) {
        $RichTextBox.Select($idx, 4)
        $RichTextBox.SelectionColor = $errorRed
        $pos = $idx + 4
    }

    # [--] -> red
    $pos = 0
    while (($idx = $RichTextBox.Text.IndexOf("[--]", $pos)) -ge 0) {
        $RichTextBox.Select($idx, 4)
        $RichTextBox.SelectionColor = $errorRed
        $pos = $idx + 4
    }

    # [ER] -> red
    $pos = 0
    while (($idx = $RichTextBox.Text.IndexOf("[ER]", $pos)) -ge 0) {
        $RichTextBox.Select($idx, 4)
        $RichTextBox.SelectionColor = $errorRed
        $pos = $idx + 4
    }

    # [SK] -> gray
    $pos = 0
    while (($idx = $RichTextBox.Text.IndexOf("[SK]", $pos)) -ge 0) {
        $RichTextBox.Select($idx, 4)
        $RichTextBox.SelectionColor = $textGray
        $pos = $idx + 4
    }

    # [CA] -> gray
    $pos = 0
    while (($idx = $RichTextBox.Text.IndexOf("[CA]", $pos)) -ge 0) {
        $RichTextBox.Select($idx, 4)
        $RichTextBox.SelectionColor = $textGray
        $pos = $idx + 4
    }

    # [PT] -> yellow
    $pos = 0
    while (($idx = $RichTextBox.Text.IndexOf("[PT]", $pos)) -ge 0) {
        $RichTextBox.Select($idx, 4)
        $RichTextBox.SelectionColor = $warnYellow
        $pos = $idx + 4
    }

    # [WN] -> yellow
    $pos = 0
    while (($idx = $RichTextBox.Text.IndexOf("[WN]", $pos)) -ge 0) {
        $RichTextBox.Select($idx, 4)
        $RichTextBox.SelectionColor = $warnYellow
        $pos = $idx + 4
    }

    $RichTextBox.Select(0, 0)
}

# Right-align [OK]/[!!]/[--] markers
function Format-StatusLine {
    param([string]$Content, [string]$Marker, [int]$Width = 44)
    $padding = $Width - $Content.Length - $Marker.Length
    if ($padding -lt 1) { $padding = 1 }
    return "$Content$(" " * $padding)$Marker"
}

function Update-StatusDisplay {
    try {
        if (-not (Test-Path $StatusFilePath)) {
            $statusLabel.ForeColor = $textGray
            $statusLabel.Text = "Status file not found - waiting..."
            return
        }

        # File change check (skip reparse if unchanged)
        $fileInfo = Get-Item $StatusFilePath -ErrorAction SilentlyContinue
        if ($null -eq $fileInfo) { return }
        if ($fileInfo.LastWriteTime -eq $script:lastWriteTime) { return }
        $script:lastWriteTime = $fileInfo.LastWriteTime

        # ロックフリー読み取り（リトライ付き）
        $jsonText = $null
        for ($retry = 0; $retry -lt 3; $retry++) {
            try {
                $jsonText = [System.IO.File]::ReadAllText(
                    (Resolve-Path $StatusFilePath).Path,
                    [System.Text.Encoding]::UTF8
                )
                break
            }
            catch {
                Start-Sleep -Milliseconds 50
            }
        }
        if ([string]::IsNullOrEmpty($jsonText)) { return }

        $status = $jsonText | ConvertFrom-Json

        # --- PC Info 比較更新 ---
        $pc  = $status.PCInfo
        $cur = $status.CurrentPCInfo
        $pcText = ""

        # Worker name (if available)
        $workerName = $status.WorkerName
        if (-not [string]::IsNullOrEmpty($workerName)) {
            $pcText += "Worker:    $workerName`r`n`r`n"
        }

        # CurrentPCInfo が存在しない場合は従来表示にフォールバック
        if ($null -eq $cur) {
            $pcText += "ID:        $($pc.AdminID)`r`n"
            $pcText += "Old Name:  $($pc.OldPCName)`r`n"
            $pcText += "New Name:  $($pc.NewPCName)`r`n"
            $pcText += "`r`n"
            if (-not [string]::IsNullOrEmpty($pc.EthernetIP)) {
                $pcText += "[Ethernet]`r`n"
                $pcText += "  IP:      $($pc.EthernetIP)`r`n"
                $pcText += "  Subnet:  $($pc.EthernetSubnet)`r`n"
                $pcText += "  Gateway: $($pc.EthernetGateway)`r`n"
            }
            Set-ColorizedText -RichTextBox $pcInfoRtb -Text $pcText
        }
        else {
            $pcText += "ID:       $($pc.AdminID)`r`n"

            # --- PC Name 比較 ---
            $curName = $cur.ComputerName
            $tgtName = $pc.NewPCName
            if ([string]::IsNullOrEmpty($tgtName)) {
                $pcText += "PC Name:  $curName`r`n"
            }
            elseif ($curName -eq $tgtName) {
                $pcText += (Format-StatusLine "PC Name:  $curName" "[OK]") + "`r`n"
            }
            else {
                $pcText += (Format-StatusLine "PC Name:  $curName" "[!!]") + "`r`n"
                $pcText += "          -> $tgtName`r`n"
            }
            $pcText += "`r`n"

            # --- Ethernet 比較 ---
            if (-not [string]::IsNullOrEmpty($pc.EthernetIP)) {
                $pcText += "[Ethernet]`r`n"

                # IP
                $curVal = if ($cur.EthernetIP) { $cur.EthernetIP } else { "(none)" }
                if ($curVal -eq $pc.EthernetIP) {
                    $pcText += (Format-StatusLine "  IP:     $curVal" "[OK]") + "`r`n"
                } else {
                    $pcText += (Format-StatusLine "  IP:     $curVal" "[!!]") + "`r`n"
                    $pcText += "          -> $($pc.EthernetIP)`r`n"
                }

                # Subnet
                $curVal = if ($cur.EthernetSubnet) { $cur.EthernetSubnet } else { "(none)" }
                if ($curVal -eq $pc.EthernetSubnet) {
                    $pcText += (Format-StatusLine "  Subnet: $curVal" "[OK]") + "`r`n"
                } else {
                    $pcText += (Format-StatusLine "  Subnet: $curVal" "[!!]") + "`r`n"
                    $pcText += "          -> $($pc.EthernetSubnet)`r`n"
                }

                # Gateway
                $curVal = if ($cur.EthernetGateway) { $cur.EthernetGateway } else { "(none)" }
                if ($curVal -eq $pc.EthernetGateway) {
                    $pcText += (Format-StatusLine "  GW:     $curVal" "[OK]") + "`r`n"
                } else {
                    $pcText += (Format-StatusLine "  GW:     $curVal" "[!!]") + "`r`n"
                    $pcText += "          -> $($pc.EthernetGateway)`r`n"
                }
                $pcText += "`r`n"
            }

            # --- Wi-Fi 比較 ---
            if (-not [string]::IsNullOrEmpty($pc.WifiIP)) {
                $pcText += "[Wi-Fi]`r`n"

                $curVal = if ($cur.WifiIP) { $cur.WifiIP } else { "(none)" }
                if ($curVal -eq $pc.WifiIP) {
                    $pcText += (Format-StatusLine "  IP:     $curVal" "[OK]") + "`r`n"
                } else {
                    $pcText += (Format-StatusLine "  IP:     $curVal" "[!!]") + "`r`n"
                    $pcText += "          -> $($pc.WifiIP)`r`n"
                }

                $curVal = if ($cur.WifiSubnet) { $cur.WifiSubnet } else { "(none)" }
                if ($curVal -eq $pc.WifiSubnet) {
                    $pcText += (Format-StatusLine "  Subnet: $curVal" "[OK]") + "`r`n"
                } else {
                    $pcText += (Format-StatusLine "  Subnet: $curVal" "[!!]") + "`r`n"
                    $pcText += "          -> $($pc.WifiSubnet)`r`n"
                }

                $curVal = if ($cur.WifiGateway) { $cur.WifiGateway } else { "(none)" }
                if ($curVal -eq $pc.WifiGateway) {
                    $pcText += (Format-StatusLine "  GW:     $curVal" "[OK]") + "`r`n"
                } else {
                    $pcText += (Format-StatusLine "  GW:     $curVal" "[!!]") + "`r`n"
                    $pcText += "          -> $($pc.WifiGateway)`r`n"
                }
                $pcText += "`r`n"
            }

            # --- DNS 比較 ---
            $targetDns = @($pc.DNS) | Where-Object { -not [string]::IsNullOrEmpty($_) } | Sort-Object
            $currentDns = @($cur.DNS) | Where-Object { -not [string]::IsNullOrEmpty($_) } | Sort-Object
            if ($targetDns.Count -gt 0) {
                $tgtStr = $targetDns -join ", "
                $curStr = $currentDns -join ", "
                if ($curStr -eq $tgtStr) {
                    $pcText += (Format-StatusLine "[DNS]  $curStr" "[OK]") + "`r`n"
                } else {
                    $curDisplay = if ($curStr) { $curStr } else { "(none)" }
                    $pcText += (Format-StatusLine "[DNS]  $curDisplay" "[!!]") + "`r`n"
                    $pcText += "       -> $tgtStr`r`n"
                }
                $pcText += "`r`n"
            }

            # --- Printers 比較 ---
            $targetPrinters = @($pc.Printers)
            if ($targetPrinters.Count -gt 0) {
                $pcText += "[Printers]`r`n"
                $currentPrinterNames = @($cur.Printers | ForEach-Object { $_.Name })

                foreach ($tp in $targetPrinters) {
                    $installed = $currentPrinterNames -contains $tp.Name
                    $pName = $tp.Name
                    if ($pName.Length -gt 30) { $pName = $pName.Substring(0, 27) + "..." }
                    if ($installed) {
                        $pcText += (Format-StatusLine "  $pName" "[OK]") + "`r`n"
                    } else {
                        $pcText += (Format-StatusLine "  $pName" "[--]") + "`r`n"
                    }
                }
            }

            Set-ColorizedText -RichTextBox $pcInfoRtb -Text $pcText
        }

        # --- Execution Summary 更新 ---
        $exec = $status.Execution
        $execText = ""

        if ($exec.TotalCount -eq 0 -and $exec.Phase -eq "idle") {
            $execText = "No execution data yet."
        }
        else {
            $phaseLabel = switch ($exec.Phase) {
                "idle"      { "Idle" }
                "executing" { ">> Running..." }
                "complete"  { "Complete" }
                default     { $exec.Phase }
            }

            $execText += "Phase: $phaseLabel`r`n"
            $execText += "Total: $($exec.TotalCount)`r`n"
            $execText += "`r`n"
            $execText += "  Success:   $($exec.SuccessCount)`r`n"
            $execText += "  Error:     $($exec.ErrorCount)`r`n"
            $execText += "  Skipped:   $($exec.SkippedCount)`r`n"
            $execText += "  Cancelled: $($exec.CancelledCount)`r`n"
            $execText += "  Partial:   $($exec.PartialCount)`r`n"

            $details = @($exec.Details)
            if ($details.Count -gt 0) {
                $execText += "`r`n--- Details ---`r`n"
                $maxShow = $details.Count
                for ($i = 0; $i -lt $maxShow; $i++) {
                    $d = $details[$i]

                    # セッション境界セパレーター
                    if ($d.Status -eq "Separator") {
                        $execText += "------------------------------`r`n"
                        continue
                    }

                    $icon = switch ($d.Status) {
                        "Success"   { "[OK]" }
                        "Error"     { "[ER]" }
                        "Skipped"   { "[SK]" }
                        "Skip"      { "[SK]" }
                        "Cancelled" { "[CA]" }
                        "Partial"   { "[PT]" }
                        "Warning"   { "[WN]" }
                        default     { "[--]" }
                    }

                    # 復元エントリには ^ プレフィックス
                    $prefix = ""
                    if ($d.IsRestored -eq $true) {
                        $prefix = "^ "
                    }

                    $msg = if ($d.Message) { " $($d.Message)" } else { "" }
                    # メッセージが長い場合は切り詰め
                    $line = "$prefix$icon $($d.Operation)$msg"
                    if ($line.Length -gt 70) { $line = $line.Substring(0, 67) + "..." }
                    $execText += "$line`r`n"
                }
            }
        }

        Set-ColorizedText -RichTextBox $execLabel -Text $execText

        $statusLabel.ForeColor = $textGray
        $statusLabel.Text = "Last update: $($status.UpdatedAt)"
    }
    catch {
        $statusLabel.ForeColor = $warnYellow
        $statusLabel.Text = "Read error: $($_.Exception.Message)"
    }
}

# ========================================
# Art Display Functions
# ========================================
function Initialize-ArtBuffer {
    $w = $artCanvas.ClientSize.Width
    $h = $artCanvas.ClientSize.Height
    if ($w -le 0 -or $h -le 0) { return }

    if ($null -ne $script:artBufferGraphics) { $script:artBufferGraphics.Dispose() }
    if ($null -ne $script:artBufferBitmap) { $script:artBufferBitmap.Dispose() }
    $script:artBufferBitmap = New-Object System.Drawing.Bitmap($w, $h)
    $script:artBufferGraphics = [System.Drawing.Graphics]::FromImage($script:artBufferBitmap)
    $script:artBufferGraphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit
}

function Select-NextArtSentence {
    # Randomly pick granularity: sentence 60%, paragraph 20%, clause 20%
    $roll = $script:artRng.Next(100)
    $pool = if ($roll -lt 60) { $script:artBySentence }
            elseif ($roll -lt 80) { $script:artByParagraph }
            else { $script:artByClause }
    $script:artCurrentText = $pool[$script:artRng.Next($pool.Count)]
    $script:artCursorPos = 0
    $script:artLastTypeTime = [DateTime]::Now
}

function Invoke-ArtTrigger {
    $script:artTypeSpeed = $script:ART_BURST_SPEED
    $script:artGlitchFrames = [Math]::Max($script:artGlitchFrames, 4)

    if ($script:artState -eq "typing" -and $script:artCurrentText.Length -gt 0 -and $script:artCursorPos -lt $script:artCurrentText.Length) {
        # Force-complete current sentence, then start new one
        $script:artDisplayLines.Add(@{
            Text  = $script:artCurrentText
            Color = "dim"
            Age   = 0
        }) | Out-Null
        while ($script:artDisplayLines.Count -gt $script:ART_MAX_LINES) {
            $script:artDisplayLines.RemoveAt(0)
        }
    }

    # Random reset: 25% chance to clear all displayed lines
    if ($script:artRng.Next(100) -lt 25) {
        $script:artDisplayLines.Clear()
    }

    # Random line continuation: 40% chance to continue on same line
    $script:artContinueOnSameLine = ($script:artDisplayLines.Count -gt 0 -and $script:artRng.Next(100) -lt 40)

    # Select new sentence and begin typing
    Select-NextArtSentence
    $script:artState = "typing"
}

function Poll-ArtSilenceFlag {
    # Toggle silent mode by presence of silence.flag
    try {
        $resolvedFlagPath = $SilenceFlagPath
        if (-not [System.IO.Path]::IsPathRooted($SilenceFlagPath)) {
            $resolvedFlagPath = Join-Path (Get-Location) $SilenceFlagPath
        }
        $present = Test-Path $resolvedFlagPath
        if ($present -eq $script:artSilent) { return }

        if ($present) {
            # OFF -> ON: clear all art state so silent canvas is clean
            $script:artDisplayLines.Clear()
            $script:artCurrentText   = ""
            $script:artCursorPos     = 0
            $script:artState         = "waiting"
            $script:artTriggerQueue  = 0
            $script:artFlashFrames   = 0
            $script:artGlitchFrames  = 0
            $script:artContinueOnSameLine = $false
        }
        else {
            # ON -> OFF: catch up pulse counter so accumulated pulses during
            # silence do not flood the canvas on release
            try {
                $resolvedPulsePath = $PulseFilePath
                if (-not [System.IO.Path]::IsPathRooted($PulseFilePath)) {
                    $resolvedPulsePath = Join-Path (Get-Location) $PulseFilePath
                }
                if (Test-Path $resolvedPulsePath) {
                    $val = [int][System.IO.File]::ReadAllText($resolvedPulsePath).Trim()
                    $script:artLastPulseValue = $val
                }
            }
            catch { }
        }

        $script:artSilent = $present
    }
    catch { }
}

function Poll-ArtPulseFile {
    try {
        $resolvedPulsePath = $PulseFilePath
        if (-not [System.IO.Path]::IsPathRooted($PulseFilePath)) {
            $resolvedPulsePath = Join-Path (Get-Location) $PulseFilePath
        }
        if (Test-Path $resolvedPulsePath) {
            $val = [int][System.IO.File]::ReadAllText($resolvedPulsePath).Trim()
            if ($val -gt $script:artLastPulseValue) {
                $diff = $val - $script:artLastPulseValue
                # Only queue triggers during profile execution
                # Do not update lastPulseValue until executing, so early pulses accumulate
                if ($script:artCurrentPhase -eq "executing") {
                    $script:artTriggerQueue += $diff
                    $script:artLastPulseValue = $val
                }
            }
        }
    }
    catch { }
}

function Poll-ArtStatusFile {
    # Phase changes and error/success flash effects
    $resolvedPath = $StatusFilePath
    if (-not [System.IO.Path]::IsPathRooted($StatusFilePath)) {
        $resolvedPath = Join-Path (Get-Location) $StatusFilePath
    }
    if (-not (Test-Path $resolvedPath)) { return }

    try {
        $fileInfo = Get-Item $resolvedPath -ErrorAction Stop
        if ($fileInfo.LastWriteTime -eq $script:artLastStatusWriteTime) { return }
        $script:artLastStatusWriteTime = $fileInfo.LastWriteTime

        $json = $null
        for ($retry = 0; $retry -lt 3; $retry++) {
            try {
                $json = [System.IO.File]::ReadAllText($resolvedPath, [System.Text.Encoding]::UTF8)
                break
            }
            catch { Start-Sleep -Milliseconds 50 }
        }
        if ([string]::IsNullOrWhiteSpace($json)) { return }

        $artStatus = $json | ConvertFrom-Json

        $newPhase = $artStatus.Execution.Phase
        if ($newPhase -ne $script:artCurrentPhase) {
            $script:artCurrentPhase = $newPhase
            if ($newPhase -eq "complete") {
                $script:artFlashColor = [System.Drawing.Color]::White
                $script:artFlashFrames = 8
            }
        }

        $detailCount = 0
        if ($null -ne $artStatus.Execution.Details) {
            $detailCount = @($artStatus.Execution.Details).Count
        }

        if ($detailCount -gt $script:artLastDetailCount) {
            $latestDetail = $artStatus.Execution.Details[-1]
            if ($null -ne $latestDetail -and $latestDetail.Status -eq "Error") {
                $script:artFlashColor = [System.Drawing.Color]::FromArgb(200, 30, 30)
                $script:artFlashFrames = 6
                $script:artGlitchFrames = 8
            }
            elseif ($null -ne $latestDetail -and $latestDetail.Status -eq "Success") {
                $script:artFlashColor = [System.Drawing.Color]::FromArgb(0, 200, 180)
                $script:artFlashFrames = 3
            }
        }
        $script:artLastDetailCount = $detailCount
    }
    catch { }
}

# Build overlay string: alpha chars preserved, CJK → 2 spaces, other ASCII → 1 space
# Draw text with alphabetic chars in alphaBrush color, others in baseBrush color (single-pass)
function Draw-ColoredText {
    param(
        [System.Drawing.Graphics]$Graphics,
        [string]$Text,
        [System.Drawing.Font]$Font,
        [System.Drawing.Brush]$BaseBrush,
        [System.Drawing.Brush]$AlphaBrush,
        [System.Drawing.RectangleF]$Rect,
        [System.Drawing.StringFormat]$SF
    )

    # Fast path: no alpha chars → single draw
    if ($Text -notmatch '[a-zA-Z]') {
        $Graphics.DrawString($Text, $Font, $BaseBrush, $Rect, $SF)
        return
    }
    # Fast path: all alpha → single draw
    if ($Text -match '^[a-zA-Z ]+$') {
        $Graphics.DrawString($Text, $Font, $AlphaBrush, $Rect, $SF)
        return
    }

    # Check if text fits in one line (common case)
    $sfTypo = [System.Drawing.StringFormat]::GenericTypographic.Clone()
    $sfTypo.FormatFlags = [System.Drawing.StringFormatFlags]::MeasureTrailingSpaces -bor [System.Drawing.StringFormatFlags]::NoWrap
    $fullSize = $Graphics.MeasureString($Text, $Font, [System.Drawing.PointF]::new(0, 0), $sfTypo)

    if ($fullSize.Width -gt $Rect.Width) {
        # Multi-line: fall back to DrawString with wrapping for base, then overlay alpha
        # Draw full text in base color with word wrap
        $Graphics.DrawString($Text, $Font, $BaseBrush, $Rect, $SF)
        # Overlay alpha chars using MeasureCharacterRanges for exact positions
        $alphaRuns = @()
        $ci = 0
        while ($ci -lt $Text.Length -and $alphaRuns.Count -lt 32) {
            if ($Text[$ci] -match '[a-zA-Z]') {
                $start = $ci
                while ($ci -lt $Text.Length -and $Text[$ci] -match '[a-zA-Z]') { $ci++ }
                $alphaRuns += @{ First = $start; Length = $ci - $start }
            } else { $ci++ }
        }
        if ($alphaRuns.Count -gt 0) {
            $crArray = New-Object 'System.Drawing.CharacterRange[]' $alphaRuns.Count
            for ($j = 0; $j -lt $alphaRuns.Count; $j++) {
                $crArray[$j] = New-Object System.Drawing.CharacterRange($alphaRuns[$j].First, $alphaRuns[$j].Length)
            }
            $sfMeasure = $SF.Clone()
            try {
                $sfMeasure.SetMeasurableCharacterRanges($crArray)
                $regions = $Graphics.MeasureCharacterRanges($Text, $Font, $Rect, $sfMeasure)
                $padX = $Font.Size / 6.0
                for ($j = 0; $j -lt $regions.Length; $j++) {
                    $bounds = $regions[$j].GetBounds($Graphics)
                    $regions[$j].Dispose()
                    if ($bounds.Width -gt 0 -and $bounds.Height -gt 0) {
                        $substr = $Text.Substring($alphaRuns[$j].First, $alphaRuns[$j].Length)
                        $Graphics.DrawString($substr, $Font, $AlphaBrush, ($bounds.X - $padX), $bounds.Y, $SF)
                    }
                }
            } catch { }
            finally { $sfMeasure.Dispose() }
        }
        $sfTypo.Dispose()
        return
    }

    # Single-line: draw run by run with GenericTypographic (no padding, precise positioning)
    $runs = [System.Collections.ArrayList]::new()
    $i = 0
    while ($i -lt $Text.Length) {
        $isAlpha = ($Text[$i] -match '[a-zA-Z]')
        $start = $i
        if ($isAlpha) {
            while ($i -lt $Text.Length -and $Text[$i] -match '[a-zA-Z]') { $i++ }
        } else {
            while ($i -lt $Text.Length -and $Text[$i] -notmatch '[a-zA-Z]') { $i++ }
        }
        [void]$runs.Add(@{ Start = $start; Length = $i - $start; IsAlpha = $isAlpha })
    }

    $padX = $Font.Size / 6.0
    $currentX = $Rect.X + $padX
    foreach ($run in $runs) {
        $substr = $Text.Substring($run.Start, $run.Length)
        $size = $Graphics.MeasureString($substr, $Font, [System.Drawing.PointF]::new(0, 0), $sfTypo)
        $brush = if ($run.IsAlpha) { $AlphaBrush } else { $BaseBrush }
        $Graphics.DrawString($substr, $Font, $brush, $currentX, $Rect.Y, $sfTypo)
        $currentX += $size.Width
    }
    $sfTypo.Dispose()
}

function Render-ArtFrame {
    if ($null -eq $script:artBufferGraphics) { return }

    # Silent mode: draw a blank frame and skip all art logic
    if ($script:artSilent) {
        $script:artBufferGraphics.FillRectangle($script:artBgBrush, 0, 0,
            $script:artBufferBitmap.Width, $script:artBufferBitmap.Height)
        $artCanvas.Image = $script:artBufferBitmap
        return
    }

    $now = [DateTime]::Now
    $g = $script:artBufferGraphics
    $w = $script:artBufferBitmap.Width
    $h = $script:artBufferBitmap.Height

    # Consume trigger queue when waiting
    if ($script:artState -eq "waiting" -and $script:artTriggerQueue -gt 0) {
        $script:artTriggerQueue--
        Invoke-ArtTrigger
    }

    # Advance typing
    if ($script:artState -eq "typing" -and $script:artCurrentText.Length -gt 0 -and $script:artCursorPos -lt $script:artCurrentText.Length) {
        $elapsed = ($now - $script:artLastTypeTime).TotalMilliseconds
        if ($elapsed -ge $script:artTypeSpeed) {
            $charsToAdvance = [Math]::Max(1, [int]($elapsed / $script:artTypeSpeed))
            $script:artCursorPos = [Math]::Min($script:artCurrentText.Length, $script:artCursorPos + $charsToAdvance)
            $script:artLastTypeTime = $now
        }
    }
    elseif ($script:artState -eq "typing" -and $script:artCurrentText.Length -gt 0 -and $script:artCursorPos -ge $script:artCurrentText.Length) {
        # Sentence complete
        if ($script:artContinueOnSameLine -and $script:artDisplayLines.Count -gt 0) {
            $lastLine = $script:artDisplayLines[$script:artDisplayLines.Count - 1]
            $lastLine.Text += $script:artCurrentText
            $lastLine.Age = 0
        }
        else {
            $script:artDisplayLines.Add(@{
                Text  = $script:artCurrentText
                Color = "dim"
                Age   = 0
            }) | Out-Null
        }

        while ($script:artDisplayLines.Count -gt $script:ART_MAX_LINES) {
            $script:artDisplayLines.RemoveAt(0)
        }

        $script:artCurrentText = ""
        $script:artCursorPos = 0
        $script:artState = "waiting"
        $script:artContinueOnSameLine = $false
    }

    # Cursor blink
    if (($now - $script:artCursorBlinkTime).TotalMilliseconds -ge 500) {
        $script:artCursorVisible = -not $script:artCursorVisible
        $script:artCursorBlinkTime = $now
    }

    # --- Draw ---
    $g.FillRectangle($script:artBgBrush, 0, 0, $w, $h)

    # Flash overlay
    if ($script:artFlashFrames -gt 0) {
        $alpha = [Math]::Min(80, $script:artFlashFrames * 12)
        $flashBrush = New-Object System.Drawing.SolidBrush(
            [System.Drawing.Color]::FromArgb($alpha, $script:artFlashColor.R, $script:artFlashColor.G, $script:artFlashColor.B)
        )
        $g.FillRectangle($flashBrush, 0, 0, $w, $h)
        $flashBrush.Dispose()
        $script:artFlashFrames--
    }

    # String format for word wrap
    $sf = [System.Drawing.StringFormat]::GenericDefault
    $maxTextWidth = $w - 16
    $y = 8

    # Pre-calculate total height for scroll
    $totalHeights = @()
    for ($i = 0; $i -lt $script:artDisplayLines.Count; $i++) {
        $measured = $g.MeasureString($script:artDisplayLines[$i].Text, $script:artFont, $maxTextWidth, $sf)
        $totalHeights += [Math]::Max($script:artLineH, [int][Math]::Ceiling($measured.Height))
    }
    $reserveH = $script:artLineH * 2
    $availH = $h - 8 - $reserveH
    $cumH = 0
    $startIdx = $script:artDisplayLines.Count
    for ($i = $script:artDisplayLines.Count - 1; $i -ge 0; $i--) {
        $cumH += $totalHeights[$i]
        if ($cumH -gt $availH) { break }
        $startIdx = $i
    }

    # Draw completed lines with word wrap
    for ($i = $startIdx; $i -lt $script:artDisplayLines.Count; $i++) {
        $line = $script:artDisplayLines[$i]
        $line.Age++

        $brush      = if ($line.Age -lt 60) { $script:artDimGreenBrush } else { $script:artDimGrayBrush }
        $alphaBrush = if ($line.Age -lt 60) { $script:artAlphaDimGreenBrush } else { $script:artAlphaDimGrayBrush }

        $text = $line.Text
        if ($script:artGlitchFrames -gt 0 -and $script:artRng.Next(100) -lt 20) {
            $pos = $script:artRng.Next([Math]::Min($text.Length, [Math]::Max(1, $text.Length)))
            $glitchChar = $script:GLITCH_CHARS[$script:artRng.Next($script:GLITCH_CHARS.Count)]
            $text = $text.Substring(0, $pos) + $glitchChar + $text.Substring([Math]::Min($pos + 1, $text.Length))
            $brush = New-Object System.Drawing.SolidBrush($script:artGlitchWhite)
        }

        $rect = New-Object System.Drawing.RectangleF(8, $y, $maxTextWidth, ($h - $y))
        Draw-ColoredText -Graphics $g -Text $text -Font $script:artFont -BaseBrush $brush -AlphaBrush $alphaBrush -Rect $rect -SF $sf
        $measured = $g.MeasureString($text, $script:artFont, $maxTextWidth, $sf)
        $y += [Math]::Max($script:artLineH, [int][Math]::Ceiling($measured.Height))
    }

    # Draw current typing line or waiting cursor
    if ($script:artState -eq "typing" -and $script:artCurrentText.Length -gt 0) {
        $textStr = $script:artCurrentText.Substring(0, $script:artCursorPos)

        $xOffset = 8
        $drawWidth = $maxTextWidth
        if ($script:artContinueOnSameLine -and $script:artDisplayLines.Count -gt 0) {
            $lastLine = $script:artDisplayLines[$script:artDisplayLines.Count - 1]
            $lastMeasured = $g.MeasureString($lastLine.Text, $script:artFont, $maxTextWidth, $sf)
            $lastLineCount = [Math]::Max(1, [int][Math]::Ceiling($lastMeasured.Height / $script:artLineH))
            $lastLineWidth = $lastMeasured.Width
            $xOffset = 8 + $lastLineWidth - 6
            $drawWidth = $maxTextWidth - ($xOffset - 8)
            if ($drawWidth -lt 50) {
                $xOffset = 8
                $drawWidth = $maxTextWidth
            }
            else {
                $y -= $script:artLineH
            }
        }

        # Glitch on current char
        if ($script:artCursorPos -lt $script:artCurrentText.Length -and $script:artRng.Next(100) -lt 8) {
            $glitchChar = $script:GLITCH_CHARS[$script:artRng.Next($script:GLITCH_CHARS.Count)]
            $glitchText = $textStr + $glitchChar
            $glitchBrush = New-Object System.Drawing.SolidBrush($script:artGlitchWhite)
            $typeRect = New-Object System.Drawing.RectangleF($xOffset, $y, $drawWidth, ($h - $y))
            Draw-ColoredText -Graphics $g -Text $glitchText -Font $script:artFontBold -BaseBrush $glitchBrush -AlphaBrush $script:artAlphaBrush -Rect $typeRect -SF $sf
            $glitchBrush.Dispose()
        }
        else {
            $typeRect = New-Object System.Drawing.RectangleF($xOffset, $y, $drawWidth, ($h - $y))
            Draw-ColoredText -Graphics $g -Text $textStr -Font $script:artFontBold -BaseBrush $script:artCyanBrush -AlphaBrush $script:artAlphaBrush -Rect $typeRect -SF $sf

            if ($script:artCursorVisible -and $script:artCursorPos -lt $script:artCurrentText.Length) {
                $typeMeasured = $g.MeasureString($textStr, $script:artFontBold, $drawWidth, $sf)
                $typeLineCount = [Math]::Max(1, [int][Math]::Ceiling($typeMeasured.Height / $script:artLineH))
                $cursorY = $y + ($typeLineCount - 1) * $script:artLineH
                $cursorX = $xOffset + $typeMeasured.Width - 6
                if ($typeLineCount -gt 1) {
                    $cursorX = 8 + $typeMeasured.Width - 6
                }
                if ($cursorX -gt $w - 16) { $cursorX = 8 }
                $g.DrawString("_", $script:artFontBold, $script:artCyanBrush, $cursorX, $cursorY)
            }
        }
    }
    else {
        # Waiting state: blinking cursor only
        if ($script:artCursorVisible) {
            $g.DrawString("_", $script:artFontBold, $script:artCyanBrush, 8, $y)
        }
    }

    # Glitch countdown
    if ($script:artGlitchFrames -gt 0) { $script:artGlitchFrames-- }

    # Swap to screen
    $artCanvas.Image = $script:artBufferBitmap
}

# ========================================
# Timer Setup (1500ms interval)
# ========================================
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 1500
$timer.Add_Tick({ Update-StatusDisplay })
$timer.Start()

# Art Display render timer (40ms ~25fps)
$artRenderTimer = New-Object System.Windows.Forms.Timer
$artRenderTimer.Interval = $script:ART_RENDER_INTERVAL
$artRenderTimer.Add_Tick({ Render-ArtFrame })
$artRenderTimer.Start()

# Art Display pulse timer (200ms)
$artPulseTimer = New-Object System.Windows.Forms.Timer
$artPulseTimer.Interval = $script:ART_PULSE_INTERVAL
$artPulseTimer.Add_Tick({
    Poll-ArtSilenceFlag
    Poll-ArtPulseFile
    Poll-ArtStatusFile
})
$artPulseTimer.Start()

# ========================================
# Form Event Handlers
# ========================================
$form.Add_Shown({ Initialize-ArtBuffer })
$artCanvas.Add_Resize({ Initialize-ArtBuffer })

$form.Add_FormClosing({
    $timer.Stop()
    $timer.Dispose()
    $artRenderTimer.Stop()
    $artRenderTimer.Dispose()
    $artPulseTimer.Stop()
    $artPulseTimer.Dispose()
    if ($null -ne $script:artBufferGraphics) { $script:artBufferGraphics.Dispose() }
    if ($null -ne $script:artFont) { $script:artFont.Dispose() }
    if ($null -ne $script:artFontBold) { $script:artFontBold.Dispose() }
    if ($null -ne $script:artBgBrush) { $script:artBgBrush.Dispose() }
    if ($null -ne $script:artCyanBrush) { $script:artCyanBrush.Dispose() }
    if ($null -ne $script:artDimGreenBrush) { $script:artDimGreenBrush.Dispose() }
    if ($null -ne $script:artDimGrayBrush) { $script:artDimGrayBrush.Dispose() }
    if ($null -ne $script:artAlphaBrush) { $script:artAlphaBrush.Dispose() }
    if ($null -ne $script:artAlphaDimGreenBrush) { $script:artAlphaDimGreenBrush.Dispose() }
    if ($null -ne $script:artAlphaDimGrayBrush) { $script:artAlphaDimGrayBrush.Dispose() }
})

# --- Screenshot button click (gyotaq pattern: hide -> capture -> show) ---
$btnScreenshot.Add_Click({
    $savedLocation = $form.Location
    $savedSize = $form.Size
    $form.Hide()
    [System.Windows.Forms.Application]::DoEvents()
    Start-Sleep -Milliseconds 300

    $result = $null
    $errorMsg = $null
    try {
        $result = Save-Screenshot -BaseDir $script:gyotakuDir
    }
    catch {
        $errorMsg = $_.Exception.Message
    }

    $form.Location = $savedLocation
    $form.Size = $savedSize
    $form.Show()
    [System.Windows.Forms.Application]::DoEvents()

    if ($null -ne $result) {
        $statusLabel.ForeColor = $successGreen
        $statusLabel.Text = "Screenshot saved: $([System.IO.Path]::GetFileName($result))"
    }
    elseif ($errorMsg) {
        $statusLabel.ForeColor = $errorRed
        $statusLabel.Text = "Screenshot error: $errorMsg"
    }
    else {
        $statusLabel.ForeColor = $errorRed
        $statusLabel.Text = "Screenshot failed (Save-Screenshot returned null)"
    }
})

# --- Skip button click: write skip_request.flag for async monitor loop ---
# The flag is consumed by Invoke-SafeCommandAsync in the kernel process. If
# the currently running module was dispatched synchronously (no __ASYNC__
# marker in the profile), the flag is left in place as a stale file and
# will be cleared on the next async module start — it does not affect sync
# execution.
$btnSkipAsync.Add_Click({
    $flagPath = Join-Path $script:fabriqRoot "kernel\json\skip_request.flag"
    try {
        $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        "requested at $ts" | Out-File -FilePath $flagPath -Encoding UTF8 -Force
        $statusLabel.ForeColor = [System.Drawing.Color]::FromArgb(255, 170, 60)
        $statusLabel.Text = "Skip requested (takes effect only for async modules)"
    }
    catch {
        $statusLabel.ForeColor = $errorRed
        $statusLabel.Text = "Skip request failed: $($_.Exception.Message)"
    }
})

# ========================================
# Initial update and run
# ========================================
Update-StatusDisplay
[System.Windows.Forms.Application]::Run($form)
