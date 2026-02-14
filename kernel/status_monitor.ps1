# ========================================
# Fabriq Status Monitor Window
# ========================================
# Launched as a separate process by main.ps1
# Reads status.json and displays PC info + execution status
# Usage: powershell -NoProfile -ExecutionPolicy Unrestricted -File .\kernel\status_monitor.ps1 -StatusFilePath ".\kernel\status.json"

param(
    [string]$StatusFilePath = ".\kernel\status.json"
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Hide the PowerShell console window (keep only the Forms window visible)
Add-Type -Name Win32 -Namespace Native -MemberDefinition @'
    [DllImport("kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
'@
$consoleHwnd = [Native.Win32]::GetConsoleWindow()
if ($consoleHwnd -ne [IntPtr]::Zero) {
    [Native.Win32]::ShowWindow($consoleHwnd, 0) | Out-Null  # SW_HIDE = 0
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
$form = New-Object System.Windows.Forms.Form
$form.Text = "Fabriq - Status Monitor"
$form.Size = New-Object System.Drawing.Size(480, 780)
$form.StartPosition = "Manual"
$form.Location = New-Object System.Drawing.Point(
    ([System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Right - 500),
    50
)
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox = $false
$form.BackColor = $darkBg
$form.ForeColor = $textWhite
$form.Font = $fontNormal

# ========================================
# PC Info Panel (Top Section)
# ========================================
$pcInfoGroup = New-Object System.Windows.Forms.GroupBox
$pcInfoGroup.Text = " PC Info Comparison "
$pcInfoGroup.Location = New-Object System.Drawing.Point(10, 10)
$pcInfoGroup.Size = New-Object System.Drawing.Size(445, 340)
$pcInfoGroup.ForeColor = $accentCyan
$pcInfoGroup.Font = $fontBold
$form.Controls.Add($pcInfoGroup)

$pcInfoRtb = New-Object System.Windows.Forms.RichTextBox
$pcInfoRtb.Location = New-Object System.Drawing.Point(10, 22)
$pcInfoRtb.Size = New-Object System.Drawing.Size(425, 310)
$pcInfoRtb.ForeColor = $textWhite
$pcInfoRtb.BackColor = $darkBg
$pcInfoRtb.Font = $fontNormal
$pcInfoRtb.ReadOnly = $true
$pcInfoRtb.BorderStyle = "None"
$pcInfoRtb.TabStop = $false
$pcInfoRtb.Text = "Waiting for status data..."
$pcInfoGroup.Controls.Add($pcInfoRtb)

# ========================================
# Execution Summary Panel (Bottom Section)
# ========================================
$execGroup = New-Object System.Windows.Forms.GroupBox
$execGroup.Text = " Execution Summary "
$execGroup.Location = New-Object System.Drawing.Point(10, 360)
$execGroup.Size = New-Object System.Drawing.Size(445, 350)
$execGroup.ForeColor = $accentCyan
$execGroup.Font = $fontBold
$form.Controls.Add($execGroup)

$execLabel = New-Object System.Windows.Forms.Label
$execLabel.Location = New-Object System.Drawing.Point(10, 22)
$execLabel.Size = New-Object System.Drawing.Size(425, 320)
$execLabel.ForeColor = $textWhite
$execLabel.Font = $fontNormal
$execLabel.Text = "No execution data yet."
$execGroup.Controls.Add($execLabel)

# ========================================
# Status Bar
# ========================================
$statusBar = New-Object System.Windows.Forms.StatusStrip
$statusBar.BackColor = [System.Drawing.Color]::FromArgb(25, 25, 25)
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

# ========================================
# Initial update and run
# ========================================
Update-StatusDisplay
[System.Windows.Forms.Application]::Run($form)
