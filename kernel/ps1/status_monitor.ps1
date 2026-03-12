# ========================================
# Fabriq Status Monitor Window
# ========================================
# Launched as a separate process by main.ps1
# Reads status.json and displays PC info + execution status
# Usage: powershell -NoProfile -ExecutionPolicy Unrestricted -File .\kernel\ps1\status_monitor.ps1 -StatusFilePath ".\kernel\json\status.json"

param(
    [string]$StatusFilePath = ".\kernel\json\status.json"
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
# Form Setup
# ========================================
$form = New-Object NoActivateForm
$form.Text = "Fabriq - Status Monitor"
# Scale form dimensions by DPI factor (designed at 96 DPI / 100%)
$form.Size = New-Object System.Drawing.Size(
    [int](750 * $script:dpiScale),
    [int](600 * $script:dpiScale)
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
# Main Layout (TableLayoutPanel: 1 row, 2 columns)
# ========================================
$mainLayout = New-Object System.Windows.Forms.TableLayoutPanel
$mainLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
$mainLayout.RowCount = 1
$mainLayout.ColumnCount = 2
$mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
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

$execLabel = New-Object System.Windows.Forms.Label
$execLabel.Dock = [System.Windows.Forms.DockStyle]::Fill
$execLabel.ForeColor = $textWhite
$execLabel.Font = $fontNormal
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
$script:lastUpdatedAt = $null

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

        # ファイル変更チェック（変更なければ鮮度チェックのみ実行）
        $fileInfo = Get-Item $StatusFilePath -ErrorAction SilentlyContinue
        if ($null -eq $fileInfo) { return }
        if ($fileInfo.LastWriteTime -eq $script:lastWriteTime) {
            # ファイル未変更でも鮮度チェックは実行
            if ($null -ne $script:lastUpdatedAt) {
                $staleSeconds = ([datetime]::Now - $script:lastUpdatedAt).TotalSeconds
                if ($staleSeconds -gt 60) {
                    $staleDisplay = [Math]::Floor($staleSeconds)
                    $statusLabel.ForeColor = $warnYellow
                    $statusLabel.Text = "WARNING: Data stale (${staleDisplay}s) - main process may have exited"
                }
            }
            return
        }
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

        # UpdatedAt を datetime としてキャッシュ（鮮度チェック用）
        try {
            $script:lastUpdatedAt = [datetime]::ParseExact($status.UpdatedAt, "yyyy-MM-dd HH:mm:ss", $null)
        }
        catch {
            $script:lastUpdatedAt = $null
        }

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
                $maxShow = [Math]::Min($details.Count, 20)
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
                    if ($line.Length -gt 50) { $line = $line.Substring(0, 47) + "..." }
                    $execText += "$line`r`n"
                }
                if ($details.Count -gt 20) {
                    $remaining = $details.Count - 20
                    $execText += "... and $remaining more`r`n"
                }
            }
        }

        $execLabel.Text = $execText

        # --- データ鮮度チェック ---
        $staleSeconds = 0
        try {
            $updatedTime = [datetime]::ParseExact($status.UpdatedAt, "yyyy-MM-dd HH:mm:ss", $null)
            $staleSeconds = ([datetime]::Now - $updatedTime).TotalSeconds
        }
        catch {
            # パース失敗時はスキップ（通常の表示を継続）
            $staleSeconds = 0
        }

        if ($staleSeconds -gt 60) {
            # 60秒以上古い → 警告表示
            $staleDisplay = [Math]::Floor($staleSeconds)
            $statusLabel.ForeColor = $warnYellow
            $statusLabel.Text = "WARNING: Data stale (${staleDisplay}s) - main process may have exited"
        }
        else {
            # 正常
            $statusLabel.ForeColor = $textGray
            $statusLabel.Text = "Last update: $($status.UpdatedAt)"
        }
    }
    catch {
        $statusLabel.ForeColor = $warnYellow
        $statusLabel.Text = "Read error: $($_.Exception.Message)"
    }
}

# ========================================
# Timer Setup (1500ms interval)
# ========================================
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 1500
$timer.Add_Tick({ Update-StatusDisplay })
$timer.Start()

# ========================================
# Form Event Handlers
# ========================================
$form.Add_FormClosing({
    $timer.Stop()
    $timer.Dispose()
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

# ========================================
# Initial update and run
# ========================================
Update-StatusDisplay
[System.Windows.Forms.Application]::Run($form)
