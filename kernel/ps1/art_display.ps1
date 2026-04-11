# ========================================
# Fabriq Art Display - Terminal Typing Window
# ========================================
# Launched as a separate process alongside the status monitor.
# Displays sentences typed character-by-character, driven by
# module execution progress (status.json polling).
# Usage: powershell -NoProfile -ExecutionPolicy Unrestricted -File .\kernel\ps1\art_display.ps1 -StatusFilePath ".\kernel\json\status.json" -SentenceFilePath ".\kernel\txt\art_sentences.txt"

param(
    [string]$StatusFilePath = ".\kernel\json\status.json",
    [string]$SentenceFilePath = ".\kernel\txt\art_sentences.txt",
    [string]$PulseFilePath = ".\kernel\json\art_pulse.txt"
)

# ========================================
# DPI Awareness (must be set BEFORE any Forms/Drawing operations)
# ========================================
Add-Type -AssemblyName System.Drawing
Add-Type -TypeDefinition @"
    using System.Runtime.InteropServices;
    public class DPIUtil {
        [DllImport("user32.dll")]
        public static extern bool SetProcessDPIAware();
    }
"@ -ErrorAction SilentlyContinue
$null = [DPIUtil]::SetProcessDPIAware()

$tmpG = [System.Drawing.Graphics]::FromHwnd([IntPtr]::Zero)
$script:dpiScale = $tmpG.DpiX / 96.0
$tmpG.Dispose()

Add-Type -AssemblyName System.Windows.Forms

# Hide the PowerShell console window
Add-Type -Name Win32 -Namespace Native -MemberDefinition @'
    [DllImport("kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
'@

# NoActivateForm: does not steal focus
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
'@

$consoleHwnd = [Native.Win32]::GetConsoleWindow()
if ($consoleHwnd -ne [IntPtr]::Zero) {
    [Native.Win32]::ShowWindow($consoleHwnd, 0) | Out-Null
}

# ========================================
# Configuration
# ========================================
$script:FORM_W         = 580        # form width (base)
$script:FORM_H         = 380        # form height (base)
$script:RENDER_INTERVAL = 40        # render timer interval (ms) ~25fps
$script:POLL_INTERVAL   = 1500      # status.json poll interval (ms)
$script:PULSE_INTERVAL  = 200       # pulse file poll interval (ms)
$script:IDLE_SPEED      = 70        # ms per character (idle)
$script:BURST_SPEED     = 18        # ms per character (triggered)
$script:MAX_LINES       = 50        # max completed lines before trimming
$script:LINE_HEIGHT     = 20        # line height in pixels (base)
$script:GLITCH_CHARS    = @('_','#','@','!','^','~','`','|','{','}','[',']','<','>','/','?','+','=','*','0','1')

# ========================================
# Load Sentences
# ========================================
$script:sentences = @()
$resolvedSentencePath = $SentenceFilePath
if (-not [System.IO.Path]::IsPathRooted($SentenceFilePath)) {
    $resolvedSentencePath = Join-Path (Get-Location) $SentenceFilePath
}
if (Test-Path $resolvedSentencePath) {
    $script:sentences = @(Get-Content -Path $resolvedSentencePath -Encoding UTF8 | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
}
if ($script:sentences.Count -eq 0) {
    $script:sentences = @(
        "SYSTEM INITIALIZED.",
        "DATA STREAM ACTIVE.",
        "PROCESSING SIGNAL.",
        "KERNEL READY.",
        "CONFIGURATION LOADED.",
        "VERIFY SEQUENCE COMPLETE."
    )
}

# ========================================
# State
# ========================================
$script:rng = New-Object System.Random
$script:displayLines = [System.Collections.ArrayList]::new()  # completed lines: @{Text; Color}
$script:currentText = ""          # full text of sentence being typed
$script:cursorPos = 0             # characters typed so far
$script:state = "waiting"         # waiting / typing
$script:typeSpeed = $script:BURST_SPEED
$script:lastTypeTime = [DateTime]::Now
$script:cursorVisible = $true
$script:cursorBlinkTime = [DateTime]::Now
$script:glitchFrames = 0
$script:flashFrames = 0
$script:flashColor = $null
$script:continueOnSameLine = $false

# Status polling state
$script:lastStatusWriteTime = [DateTime]::MinValue
$script:lastDetailCount = 0
$script:currentPhase = "idle"
$script:lastPulseValue = 0
$script:triggerQueue = 0

# ========================================
# Form Setup
# ========================================
$form = New-Object NoActivateForm
$form.Text = "Fabriq"
$form.Size = New-Object System.Drawing.Size(
    [int]($script:FORM_W * $script:dpiScale),
    [int]($script:FORM_H * $script:dpiScale)
)
$form.StartPosition = "Manual"
$workArea = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
$form.Location = New-Object System.Drawing.Point(
    [int](20 * $script:dpiScale),
    ($workArea.Bottom - [int]($script:FORM_H * $script:dpiScale) - [int](20 * $script:dpiScale))
)
$form.FormBorderStyle = "None"
$script:transparencyColor = [System.Drawing.Color]::Magenta
$form.BackColor = $script:transparencyColor
$form.TransparencyKey = $script:transparencyColor
$form.TopMost = $true
$form.ShowInTaskbar = $false

# PictureBox for GDI+ rendering
$canvas = New-Object System.Windows.Forms.PictureBox
$canvas.Dock = [System.Windows.Forms.DockStyle]::Fill
$canvas.BackColor = $script:transparencyColor
$form.Controls.Add($canvas)

# Double-buffered bitmap
$script:bufferBitmap = $null
$script:bufferGraphics = $null
$script:font = New-Object System.Drawing.Font("Consolas", (10 * $script:dpiScale), [System.Drawing.FontStyle]::Regular)
$script:fontBold = New-Object System.Drawing.Font("Consolas", (10 * $script:dpiScale), [System.Drawing.FontStyle]::Bold)
$script:lineH = [int]($script:LINE_HEIGHT * $script:dpiScale)

# Colors
$script:bgColor       = [System.Drawing.Color]::FromArgb(10, 10, 14)
$script:cyanColor     = [System.Drawing.Color]::FromArgb(0, 255, 255)
$script:dimGreen      = [System.Drawing.Color]::FromArgb(0, 140, 100)
$script:dimGray       = [System.Drawing.Color]::FromArgb(60, 70, 60)
$script:glitchWhite   = [System.Drawing.Color]::FromArgb(255, 255, 255)

# Brushes (reusable)
$script:bgBrush       = New-Object System.Drawing.SolidBrush($script:transparencyColor)
$script:cyanBrush     = New-Object System.Drawing.SolidBrush($script:cyanColor)
$script:dimGreenBrush = New-Object System.Drawing.SolidBrush($script:dimGreen)
$script:dimGrayBrush  = New-Object System.Drawing.SolidBrush($script:dimGray)

function Initialize-Buffer {
    $w = $canvas.ClientSize.Width
    $h = $canvas.ClientSize.Height
    if ($w -le 0 -or $h -le 0) { return }

    if ($null -ne $script:bufferGraphics) { $script:bufferGraphics.Dispose() }
    if ($null -ne $script:bufferBitmap) { $script:bufferBitmap.Dispose() }
    $script:bufferBitmap = New-Object System.Drawing.Bitmap($w, $h)
    $script:bufferGraphics = [System.Drawing.Graphics]::FromImage($script:bufferBitmap)
    $script:bufferGraphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit
}

# ========================================
# Sentence Selection
# ========================================
function Select-NextSentence {
    $script:currentText = $script:sentences[$script:rng.Next($script:sentences.Count)]
    $script:cursorPos = 0
    $script:lastTypeTime = [DateTime]::Now
}

# ========================================
# Rendering
# ========================================
function Render-Frame {
    if ($null -eq $script:bufferGraphics) { return }

    $now = [DateTime]::Now
    $g = $script:bufferGraphics
    $w = $script:bufferBitmap.Width
    $h = $script:bufferBitmap.Height

    # Consume trigger queue when waiting
    if ($script:state -eq "waiting" -and $script:triggerQueue -gt 0) {
        $script:triggerQueue--
        Invoke-Trigger
    }

    # Advance typing only when in "typing" state
    if ($script:state -eq "typing" -and $script:currentText.Length -gt 0 -and $script:cursorPos -lt $script:currentText.Length) {
        $elapsed = ($now - $script:lastTypeTime).TotalMilliseconds
        if ($elapsed -ge $script:typeSpeed) {
            $charsToAdvance = [Math]::Max(1, [int]($elapsed / $script:typeSpeed))
            $script:cursorPos = [Math]::Min($script:currentText.Length, $script:cursorPos + $charsToAdvance)
            $script:lastTypeTime = $now
        }
    }
    elseif ($script:state -eq "typing" -and $script:currentText.Length -gt 0 -and $script:cursorPos -ge $script:currentText.Length) {
        # Sentence complete: append to last line or add new line
        if ($script:continueOnSameLine -and $script:displayLines.Count -gt 0) {
            $lastLine = $script:displayLines[$script:displayLines.Count - 1]
            $lastLine.Text += $script:currentText
            $lastLine.Age = 0
        }
        else {
            $script:displayLines.Add(@{
                Text  = $script:currentText
                Color = "dim"
                Age   = 0
            }) | Out-Null
        }

        # Trim old lines
        while ($script:displayLines.Count -gt $script:MAX_LINES) {
            $script:displayLines.RemoveAt(0)
        }

        # Return to waiting state until next trigger
        $script:currentText = ""
        $script:cursorPos = 0
        $script:state = "waiting"
        $script:continueOnSameLine = $false
    }

    # Cursor blink (toggle every 500ms)
    if (($now - $script:cursorBlinkTime).TotalMilliseconds -ge 500) {
        $script:cursorVisible = -not $script:cursorVisible
        $script:cursorBlinkTime = $now
    }

    # --- Draw ---
    $g.FillRectangle($script:bgBrush, 0, 0, $w, $h)

    # Flash overlay
    if ($script:flashFrames -gt 0) {
        $alpha = [Math]::Min(80, $script:flashFrames * 12)
        $flashBrush = New-Object System.Drawing.SolidBrush(
            [System.Drawing.Color]::FromArgb($alpha, $script:flashColor.R, $script:flashColor.G, $script:flashColor.B)
        )
        $g.FillRectangle($flashBrush, 0, 0, $w, $h)
        $flashBrush.Dispose()
        $script:flashFrames--
    }

    # String format for word wrap
    $sf = [System.Drawing.StringFormat]::GenericDefault
    $maxTextWidth = $w - 16
    $y = 8

    # Pre-calculate total height to determine start index (scroll from bottom)
    $totalHeights = @()
    for ($i = 0; $i -lt $script:displayLines.Count; $i++) {
        $measured = $g.MeasureString($script:displayLines[$i].Text, $script:font, $maxTextWidth, $sf)
        $totalHeights += [Math]::Max($script:lineH, [int][Math]::Ceiling($measured.Height))
    }
    # Reserve space for current typing line
    $reserveH = $script:lineH * 2
    $availH = $h - 8 - $reserveH
    $cumH = 0
    $startIdx = $script:displayLines.Count
    for ($i = $script:displayLines.Count - 1; $i -ge 0; $i--) {
        $cumH += $totalHeights[$i]
        if ($cumH -gt $availH) { break }
        $startIdx = $i
    }

    # Draw completed lines with word wrap
    for ($i = $startIdx; $i -lt $script:displayLines.Count; $i++) {
        $line = $script:displayLines[$i]
        $line.Age++

        # Age-based fade: recent = dimGreen, old = dimGray
        $brush = if ($line.Age -lt 60) { $script:dimGreenBrush } else { $script:dimGrayBrush }

        # Glitch: occasionally scramble a random char in old lines
        $text = $line.Text
        if ($script:glitchFrames -gt 0 -and $script:rng.Next(100) -lt 20) {
            $pos = $script:rng.Next([Math]::Min($text.Length, [Math]::Max(1, $text.Length)))
            $glitchChar = $script:GLITCH_CHARS[$script:rng.Next($script:GLITCH_CHARS.Count)]
            $text = $text.Substring(0, $pos) + $glitchChar + $text.Substring([Math]::Min($pos + 1, $text.Length))
            $brush = New-Object System.Drawing.SolidBrush($script:glitchWhite)
        }

        $rect = New-Object System.Drawing.RectangleF(8, $y, $maxTextWidth, ($h - $y))
        $g.DrawString($text, $script:font, $brush, $rect, $sf)
        $measured = $g.MeasureString($text, $script:font, $maxTextWidth, $sf)
        $y += [Math]::Max($script:lineH, [int][Math]::Ceiling($measured.Height))
    }

    # Draw current typing line or waiting cursor
    if ($script:state -eq "typing" -and $script:currentText.Length -gt 0) {
        $textStr = $script:currentText.Substring(0, $script:cursorPos)

        # Calculate X offset when continuing on same line
        $xOffset = 8
        $drawWidth = $maxTextWidth
        if ($script:continueOnSameLine -and $script:displayLines.Count -gt 0) {
            $lastLine = $script:displayLines[$script:displayLines.Count - 1]
            # Measure last visual line width (last line of wrapped text)
            $lastMeasured = $g.MeasureString($lastLine.Text, $script:font, $maxTextWidth, $sf)
            $lastLineCount = [Math]::Max(1, [int][Math]::Ceiling($lastMeasured.Height / $script:lineH))
            # Get width of the last visual line by measuring char-by-char from end
            $lastLineWidth = $lastMeasured.Width
            if ($lastLineCount -gt 1) {
                # Approximate: full lines are maxTextWidth, remainder is proportional
                $lastLineWidth = $lastMeasured.Width  # GDI reports total bounding width
            }
            $xOffset = 8 + $lastLineWidth - 6
            $drawWidth = $maxTextWidth - ($xOffset - 8)
            if ($drawWidth -lt 50) {
                # Not enough space, wrap to new line
                $xOffset = 8
                $drawWidth = $maxTextWidth
            }
            else {
                # Draw on same Y as last completed line's last visual row
                $y -= $script:lineH
            }
        }

        # Glitch on current char: randomly show wrong char for a frame
        if ($script:cursorPos -lt $script:currentText.Length -and $script:rng.Next(100) -lt 8) {
            $glitchChar = $script:GLITCH_CHARS[$script:rng.Next($script:GLITCH_CHARS.Count)]
            $glitchText = $textStr + $glitchChar
            $glitchBrush = New-Object System.Drawing.SolidBrush($script:glitchWhite)
            $typeRect = New-Object System.Drawing.RectangleF($xOffset, $y, $drawWidth, ($h - $y))
            $g.DrawString($glitchText, $script:fontBold, $glitchBrush, $typeRect, $sf)
            $glitchBrush.Dispose()
        }
        else {
            $typeRect = New-Object System.Drawing.RectangleF($xOffset, $y, $drawWidth, ($h - $y))
            $g.DrawString($textStr, $script:fontBold, $script:cyanBrush, $typeRect, $sf)

            # Cursor: position at end of typed text
            if ($script:cursorVisible -and $script:cursorPos -lt $script:currentText.Length) {
                $typeMeasured = $g.MeasureString($textStr, $script:fontBold, $drawWidth, $sf)
                $typeLineCount = [Math]::Max(1, [int][Math]::Ceiling($typeMeasured.Height / $script:lineH))
                $cursorY = $y + ($typeLineCount - 1) * $script:lineH
                # Estimate cursor X on the last visual line
                $cursorX = $xOffset + $typeMeasured.Width - 6
                if ($typeLineCount -gt 1) {
                    $cursorX = 8 + $typeMeasured.Width - 6
                }
                if ($cursorX -gt $w - 16) { $cursorX = 8 }
                $g.DrawString("_", $script:fontBold, $script:cyanBrush, $cursorX, $cursorY)
            }
        }
    }
    else {
        # Waiting state: blinking cursor only
        if ($script:cursorVisible) {
            $g.DrawString("_", $script:fontBold, $script:cyanBrush, 8, $y)
        }
    }

    # Glitch countdown
    if ($script:glitchFrames -gt 0) { $script:glitchFrames-- }

    # Swap to screen
    $canvas.Image = $script:bufferBitmap
}

# ========================================
# Status Polling
# ========================================
function Invoke-Trigger {
    $script:typeSpeed = $script:BURST_SPEED
    $script:glitchFrames = [Math]::Max($script:glitchFrames, 4)

    if ($script:state -eq "typing" -and $script:currentText.Length -gt 0 -and $script:cursorPos -lt $script:currentText.Length) {
        # Force-complete current sentence, then start new one
        $script:displayLines.Add(@{
            Text  = $script:currentText
            Color = "dim"
            Age   = 0
        }) | Out-Null
        while ($script:displayLines.Count -gt $script:MAX_LINES) {
            $script:displayLines.RemoveAt(0)
        }
    }

    # Random reset: 25% chance to clear all displayed lines
    if ($script:rng.Next(100) -lt 25) {
        $script:displayLines.Clear()
    }

    # Random line continuation: 40% chance to continue on same line
    $script:continueOnSameLine = ($script:displayLines.Count -gt 0 -and $script:rng.Next(100) -lt 40)

    # Select new sentence and begin typing
    Select-NextSentence
    $script:state = "typing"
}

function Poll-PulseFile {
    try {
        $resolvedPulsePath = $PulseFilePath
        if (-not [System.IO.Path]::IsPathRooted($PulseFilePath)) {
            $resolvedPulsePath = Join-Path (Get-Location) $PulseFilePath
        }
        if (Test-Path $resolvedPulsePath) {
            $val = [int][System.IO.File]::ReadAllText($resolvedPulsePath).Trim()
            if ($val -gt $script:lastPulseValue) {
                $diff = $val - $script:lastPulseValue
                $script:triggerQueue += $diff
                $script:lastPulseValue = $val
            }
        }
    }
    catch { }
}

function Poll-StatusFile {
    # --- Status file check (phase changes, error/success flash effects) ---
    $resolvedPath = $StatusFilePath
    if (-not [System.IO.Path]::IsPathRooted($StatusFilePath)) {
        $resolvedPath = Join-Path (Get-Location) $StatusFilePath
    }
    if (-not (Test-Path $resolvedPath)) { return }

    try {
        $fileInfo = Get-Item $resolvedPath -ErrorAction Stop
        if ($fileInfo.LastWriteTime -eq $script:lastStatusWriteTime) { return }
        $script:lastStatusWriteTime = $fileInfo.LastWriteTime

        $json = $null
        for ($retry = 0; $retry -lt 3; $retry++) {
            try {
                $json = [System.IO.File]::ReadAllText($resolvedPath, [System.Text.Encoding]::UTF8)
                break
            }
            catch { Start-Sleep -Milliseconds 50 }
        }
        if ([string]::IsNullOrWhiteSpace($json)) { return }

        $status = $json | ConvertFrom-Json

        # Detect phase changes
        $newPhase = $status.Execution.Phase

        if ($newPhase -ne $script:currentPhase) {
            $script:currentPhase = $newPhase

            if ($newPhase -eq "complete") {
                $script:flashColor = [System.Drawing.Color]::White
                $script:flashFrames = 8
            }
        }

        # Detect new execution details (for flash effects only)
        $detailCount = 0
        if ($null -ne $status.Execution.Details) {
            $detailCount = @($status.Execution.Details).Count
        }

        if ($detailCount -gt $script:lastDetailCount) {
            $latestDetail = $status.Execution.Details[-1]

            if ($null -ne $latestDetail -and $latestDetail.Status -eq "Error") {
                $script:flashColor = [System.Drawing.Color]::FromArgb(200, 30, 30)
                $script:flashFrames = 6
                $script:glitchFrames = 8
            }
            elseif ($null -ne $latestDetail -and $latestDetail.Status -eq "Success") {
                $script:flashColor = [System.Drawing.Color]::FromArgb(0, 200, 180)
                $script:flashFrames = 3
            }
        }
        $script:lastDetailCount = $detailCount
    }
    catch { }
}

# ========================================
# Form Events
# ========================================
$form.Add_Shown({ Initialize-Buffer })
$form.Add_Resize({ Initialize-Buffer })

$form.Add_FormClosing({
    $renderTimer.Stop()
    $renderTimer.Dispose()
    $pulseTimer.Stop()
    $pulseTimer.Dispose()
    $pollTimer.Stop()
    $pollTimer.Dispose()
    if ($null -ne $script:bufferGraphics) { $script:bufferGraphics.Dispose() }
    if ($null -ne $script:font) { $script:font.Dispose() }
    if ($null -ne $script:fontBold) { $script:fontBold.Dispose() }
    if ($null -ne $script:bgBrush) { $script:bgBrush.Dispose() }
    if ($null -ne $script:cyanBrush) { $script:cyanBrush.Dispose() }
    if ($null -ne $script:dimGreenBrush) { $script:dimGreenBrush.Dispose() }
    if ($null -ne $script:dimGrayBrush) { $script:dimGrayBrush.Dispose() }
})

# Right-click to close
$canvas.Add_MouseClick({
    param($sender, $e)
    if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Right) {
        $form.Close()
    }
})

# ========================================
# Timers
# ========================================
$renderTimer = New-Object System.Windows.Forms.Timer
$renderTimer.Interval = $script:RENDER_INTERVAL
$renderTimer.Add_Tick({ Render-Frame })
$renderTimer.Start()

$pulseTimer = New-Object System.Windows.Forms.Timer
$pulseTimer.Interval = $script:PULSE_INTERVAL
$pulseTimer.Add_Tick({ Poll-PulseFile })
$pulseTimer.Start()

$pollTimer = New-Object System.Windows.Forms.Timer
$pollTimer.Interval = $script:POLL_INTERVAL
$pollTimer.Add_Tick({ Poll-StatusFile })
$pollTimer.Start()

# ========================================
# Run
# ========================================
[System.Windows.Forms.Application]::Run($form)
