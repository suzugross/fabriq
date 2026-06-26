# ========================================
# Fabriq Operator - Execution Toolbar (Status Monitor recreation)
# ========================================
# In-process replacement for kernel/ps1/status_monitor.ps1 (retired in
# 3.4.0). Faithfully recreates the old monitor's visual identity (dark
# theme, Consolas, cyan GroupBox titles, [OK]/[!!]/[--] markers, art
# panel) while:
#
#   - Living on a dedicated STA Runspace inside the kernel powershell.exe
#     so Defender / ASR child-process restrictions do not apply.
#   - Driving PC Info from $FabriqToolbarShared.TargetHostInfo (pushed
#     by Set-SelectedHostEnvironment via Update-ExecutionToolbar) and
#     live OS queries (Get-NetIPAddress / Get-Printer / etc.), removing
#     the dependence on a written status.json for the comparison view.
#   - Keeping the original Surkitinisme art panel intact, reading
#     art_pulse.txt / art_sentences.txt / silence.flag / status.json
#     directly from disk just like the old monitor did.
#
# Public surface (called by kernel main.ps1 / common.ps1):
#   - Show-ExecutionToolbar
#   - Hide-ExecutionToolbar
#   - Update-ExecutionToolbar [-ExecutionState 'Idle'|'Running']
#                             [-ModuleName <s>]
#                             [-TargetHostInfo <hashtable>]
# ========================================

$script:ExecutionToolbarRunspace = $null
$script:ExecutionToolbarPS       = $null
$script:ExecutionToolbarHandle   = $null
$script:ExecutionToolbarShared   = $null

$script:ExecutionToolbarFabriqRoot = [System.IO.Path]::GetFullPath(
    (Join-Path $PSScriptRoot "..\..\..")
).TrimEnd('\')

function Get-FabriqHostInfoFromEnv {
    <#
    .SYNOPSIS
        Read the current SELECTED_* env vars and build the TargetHostInfo
        hashtable that the execution toolbar expects.
    .DESCRIPTION
        Called from two places: Set-SelectedHostEnvironment (kernel) on
        every host-selection change, and Show-ExecutionToolbar itself
        on startup to recover the snapshot when the toolbar comes up
        after env vars are already populated (fresh-start session init
        or resume).
    #>
    [CmdletBinding()]
    param()

    $printers = @()
    for ($i = 1; $i -le 10; $i++) {
        $pName = (Get-Item -Path "env:SELECTED_PRINTER_$($i)_NAME" -ErrorAction SilentlyContinue).Value
        if (-not [string]::IsNullOrWhiteSpace($pName)) {
            $printers += @{
                Name   = $pName
                Driver = (Get-Item -Path "env:SELECTED_PRINTER_$($i)_DRIVER" -ErrorAction SilentlyContinue).Value
                Port   = (Get-Item -Path "env:SELECTED_PRINTER_$($i)_PORT"   -ErrorAction SilentlyContinue).Value
            }
        }
    }
    $dns = @($env:SELECTED_DNS1, $env:SELECTED_DNS2, $env:SELECTED_DNS3, $env:SELECTED_DNS4) |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

    return @{
        Hostname    = $env:SELECTED_NEW_PCNAME
        KanriNo     = $env:SELECTED_KANRI_NO
        Pin         = $env:SELECTED_PIN
        EthIP       = $env:SELECTED_ETH_IP
        EthSubnet   = $env:SELECTED_ETH_SUBNET
        EthGateway  = $env:SELECTED_ETH_GATEWAY
        WifiIP      = $env:SELECTED_WIFI_IP
        WifiSubnet  = $env:SELECTED_WIFI_SUBNET
        WifiGateway = $env:SELECTED_WIFI_GATEWAY
        DNS         = @($dns)
        Printers    = $printers
    }
}

function Show-ExecutionToolbar {
    <#
    .SYNOPSIS
        Open the Status-Monitor-style floating panel on a dedicated STA Runspace.
    #>
    [CmdletBinding()]
    param()

    if ($null -ne $script:ExecutionToolbarHandle -and -not $script:ExecutionToolbarHandle.IsCompleted) {
        return
    }

    $shared = [hashtable]::Synchronized(@{
        State                 = 'Idle'
        ModuleName            = ''
        EvidenceBasePath      = if (-not [string]::IsNullOrWhiteSpace($global:FabriqEvidenceBasePath)) { $global:FabriqEvidenceBasePath } else { '' }
        CloseRequested        = $false
        FabriqRoot            = $script:ExecutionToolbarFabriqRoot
        TargetHostInfo        = $null
        TargetHostInfoVersion = 0
    })

    $rs = [runspacefactory]::CreateRunspace($Host)
    $rs.ApartmentState = "STA"
    $rs.ThreadOptions  = "ReuseThread"
    $rs.Open()
    $rs.SessionStateProxy.SetVariable('FabriqToolbarShared',     $shared)
    $rs.SessionStateProxy.SetVariable('FabriqToolbarCommonPath', (Join-Path $script:ExecutionToolbarFabriqRoot 'kernel\common.ps1'))
    $rs.SessionStateProxy.Path.SetLocation($script:ExecutionToolbarFabriqRoot) | Out-Null

    $ps = [PowerShell]::Create()
    $ps.Runspace = $rs
    [void]$ps.AddScript({
        try {
            Add-Type -AssemblyName System.Drawing
            Add-Type -AssemblyName System.Windows.Forms
            [System.Windows.Forms.Application]::EnableVisualStyles()

            # Last-resort net for the operator. Route any unhandled UI-thread
            # exception (e.g. a paint failure raised deep inside WinForms
            # during a display change / Explorer restart, which cannot be
            # wrapped in PowerShell try/catch) to a silent handler instead of
            # the .NET crash dialog. This honours the runspace-level
            # "invisible to the operator; swallow" contract. Must be set
            # before the first window on this thread is created.
            try {
                [System.Windows.Forms.Application]::SetUnhandledExceptionMode(
                    [System.Windows.Forms.UnhandledExceptionMode]::CatchException)
            } catch { }
            [System.Windows.Forms.Application]::add_ThreadException({ param($sender, $e) })

            . $FabriqToolbarCommonPath

            $fabriqRoot = $FabriqToolbarShared.FabriqRoot
            $gyotakuDir = Join-Path $fabriqRoot 'evidence\gyotaku'

            # Disk paths used by the art panel (resolved at startup so
            # CWD changes during module execution do not break them).
            $artPulseFilePath = Join-Path $fabriqRoot 'kernel\json\art_pulse.txt'
            $sentenceFilePath = Join-Path $fabriqRoot 'kernel\txt\art_sentences.txt'
            $silenceFlagPath  = Join-Path $fabriqRoot 'kernel\txt\silence.flag'
            $statusFilePath   = Join-Path $fabriqRoot 'kernel\json\status.json'

            # ----- DPI scale -----
            $script:dpiScale = 1.0
            try {
                $tmpG = [System.Drawing.Graphics]::FromHwnd([IntPtr]::Zero)
                $rawDpi = $tmpG.DpiX
                $tmpG.Dispose()
                if ($rawDpi -gt 0) { $script:dpiScale = $rawDpi / 96.0 }
            } catch { }
            if ($script:dpiScale -le 0) { $script:dpiScale = 1.0 }

            # ----- Colors (legacy Status Monitor palette) -----
            $darkBg       = [System.Drawing.Color]::FromArgb(30, 30, 30)
            $accentCyan   = [System.Drawing.Color]::FromArgb(0, 200, 200)
            $textWhite    = [System.Drawing.Color]::White
            $textGray     = [System.Drawing.Color]::FromArgb(160, 160, 160)
            $successGreen = [System.Drawing.Color]::FromArgb(80, 220, 80)
            $errorRed     = [System.Drawing.Color]::FromArgb(255, 80, 80)
            $orangeSkip   = [System.Drawing.Color]::FromArgb(255, 170, 60)
            $warnYellow   = [System.Drawing.Color]::FromArgb(230, 220, 90)

            # ----- Fonts -----
            $fontNormal = New-Object System.Drawing.Font("Consolas", 9)
            $fontBold   = New-Object System.Drawing.Font("Consolas", 9, [System.Drawing.FontStyle]::Bold)

            # ========================================
            # Art panel configuration (copied verbatim
            # from the retired status_monitor.ps1)
            # ========================================
            # Art render at 10fps (100ms). The retired Status Monitor
            # ran at 25fps (40ms) because it was in a separate process
            # and CPU was its own to burn. In-process we share the UI
            # thread with PC Info refresh and click handlers, so a more
            # conservative cadence is friendlier. Typing animation is
            # still smooth and cursor blink (500ms) has plenty of room.
            $script:ART_RENDER_INTERVAL = 100
            $script:ART_PULSE_INTERVAL  = 200
            $script:ART_BURST_SPEED     = 8
            $script:ART_IDLE_SPEED      = 35
            $script:ART_MAX_LINES       = 50
            $script:ART_LINE_HEIGHT     = 16
            $script:GLITCH_CHARS = @('_','#','@','!','^','~','`','|','{','}','[',']','<','>','/','?','+','=','*','0','1')

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
            $script:artLastStatusWriteTime = [DateTime]::MinValue
            $script:artLastDetailCount = 0
            # Independent status.json change-guard for the Execution panel
            # (kept separate from the art sync above so the two render paths
            # do not steal each other's "unchanged -> skip" short-circuit).
            $script:execLastStatusWriteTime = [DateTime]::MinValue
            $script:lastExecText = $null
            $script:artCurrentPhase = "idle"
            $script:artSilent = $false

            # Load art sentences (paragraph / sentence / clause pools)
            $script:artByParagraph = @()
            $script:artBySentence  = @()
            $script:artByClause    = @()
            if (Test-Path $sentenceFilePath) {
                $rawText = [System.IO.File]::ReadAllText($sentenceFilePath, [System.Text.Encoding]::UTF8)
                $script:artByParagraph = @($rawText -split '\r?\n' | ForEach-Object { $_.Trim() } | Where-Object { $_.Length -ge 8 })
                $script:artBySentence  = @($rawText -split '(?<=\u3002)' | ForEach-Object { $_.Trim() } | Where-Object { $_.Length -ge 8 })
                $script:artByClause    = @($rawText -split '(?<=\u3001)' | ForEach-Object { $_.Trim() } | Where-Object { $_.Length -ge 8 })
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
            # Form
            # ========================================
            $form = New-Object System.Windows.Forms.Form
            $form.Text = "Fabriq - Status Monitor"
            $formW = [int](600 * $script:dpiScale)
            $formH = [int](940 * $script:dpiScale)
            $form.StartPosition = "Manual"
            $workingArea = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
            # Clamp height to the working area so the 3-panel window never
            # runs off the bottom on short / low-resolution real displays.
            if ($formH -gt $workingArea.Height) { $formH = $workingArea.Height }
            $form.Size = New-Object System.Drawing.Size($formW, $formH)
            $form.Location = New-Object System.Drawing.Point(
                ($workingArea.Right - $formW - [int](20 * $script:dpiScale)),
                ([int](50 * $script:dpiScale))
            )
            $form.FormBorderStyle = "FixedSingle"
            $form.MaximizeBox = $false
            $form.MinimizeBox = $true
            $form.ShowInTaskbar = $false
            $form.BackColor = $darkBg
            $form.ForeColor = $textWhite
            $form.Font = $fontNormal
            $form.TopMost = $true

            # ----- Main layout: PC Info (42%) + Execution (33%) + Surkitinisme (25%) -----
            $mainLayout = New-Object System.Windows.Forms.TableLayoutPanel
            $mainLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
            $mainLayout.RowCount = 3
            $mainLayout.ColumnCount = 1
            $mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 42))) | Out-Null
            $mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 33))) | Out-Null
            $mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 25))) | Out-Null
            $mainLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
            $mainLayout.Padding = New-Object System.Windows.Forms.Padding(6, 6, 6, 0)
            $form.Controls.Add($mainLayout)

            # ----- PC Info GroupBox -----
            $pcInfoGroup = New-Object System.Windows.Forms.GroupBox
            $pcInfoGroup.Text = " PC Info Comparison "
            $pcInfoGroup.Dock = [System.Windows.Forms.DockStyle]::Fill
            $pcInfoGroup.ForeColor = $accentCyan
            $pcInfoGroup.Font = $fontBold
            $mainLayout.Controls.Add($pcInfoGroup, 0, 0)

            $pcInfoRtb = New-Object System.Windows.Forms.RichTextBox
            $pcInfoRtb.Dock = [System.Windows.Forms.DockStyle]::Fill
            $pcInfoRtb.ForeColor = $textWhite
            $pcInfoRtb.BackColor = $darkBg
            $pcInfoRtb.Font = $fontNormal
            $pcInfoRtb.ReadOnly = $true
            $pcInfoRtb.BorderStyle = "None"
            $pcInfoRtb.TabStop = $false
            $pcInfoRtb.Text = "Waiting for host selection..."
            $pcInfoGroup.Controls.Add($pcInfoRtb)

            # ----- Execution result GroupBox -----
            $execGroup = New-Object System.Windows.Forms.GroupBox
            $execGroup.Text = " Execution "
            $execGroup.Dock = [System.Windows.Forms.DockStyle]::Fill
            $execGroup.ForeColor = $accentCyan
            $execGroup.Font = $fontBold
            $mainLayout.Controls.Add($execGroup, 0, 1)

            $execRtb = New-Object System.Windows.Forms.RichTextBox
            $execRtb.Dock = [System.Windows.Forms.DockStyle]::Fill
            $execRtb.ForeColor = $textWhite
            $execRtb.BackColor = $darkBg
            $execRtb.Font = $fontNormal
            $execRtb.ReadOnly = $true
            $execRtb.BorderStyle = "None"
            $execRtb.TabStop = $false
            $execRtb.Text = "No execution data yet."
            $execGroup.Controls.Add($execRtb)

            # ----- Surkitinisme art GroupBox -----
            $artGroup = New-Object System.Windows.Forms.GroupBox
            $artGroup.Text = " Surkitinisme "
            $artGroup.Dock = [System.Windows.Forms.DockStyle]::Fill
            $artGroup.ForeColor = $accentCyan
            $artGroup.Font = $fontBold
            $mainLayout.Controls.Add($artGroup, 0, 2)

            $artCanvas = New-Object System.Windows.Forms.PictureBox
            $artCanvas.Dock = [System.Windows.Forms.DockStyle]::Fill
            $artCanvas.BackColor = $darkBg
            $artGroup.Controls.Add($artCanvas)

            # ----- Status bar -----
            $statusBar = New-Object System.Windows.Forms.StatusStrip
            $statusBar.BackColor = [System.Drawing.Color]::FromArgb(25, 25, 25)

            $btnGyotaq = New-Object System.Windows.Forms.ToolStripButton
            $btnGyotaq.Text = "Gyotaq"
            $btnGyotaq.ForeColor = $accentCyan
            $btnGyotaq.Font = New-Object System.Drawing.Font("Consolas", 9, [System.Drawing.FontStyle]::Bold)
            $btnGyotaq.Margin = New-Object System.Windows.Forms.Padding(4, 2, 8, 0)
            $btnGyotaq.Enabled = $false
            $statusBar.Items.Add($btnGyotaq) | Out-Null

            $btnSkip = New-Object System.Windows.Forms.ToolStripButton
            $btnSkip.Text = "Skip"
            $btnSkip.ForeColor = $orangeSkip
            $btnSkip.Font = New-Object System.Drawing.Font("Consolas", 9, [System.Drawing.FontStyle]::Bold)
            $btnSkip.Margin = New-Object System.Windows.Forms.Padding(0, 2, 8, 0)
            $btnSkip.ToolTipText = "Request async module skip. Only effective for modules running after __ASYNC__ marker."
            $btnSkip.Enabled = $false
            $statusBar.Items.Add($btnSkip) | Out-Null

            $statusSep = New-Object System.Windows.Forms.ToolStripSeparator
            $statusBar.Items.Add($statusSep) | Out-Null

            $statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
            $statusLabel.ForeColor = $textGray
            $statusLabel.Text = "Idle"
            $statusLabel.Spring = $true
            $statusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
            $statusBar.Items.Add($statusLabel) | Out-Null

            $emailLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
            $emailLabel.ForeColor = [System.Drawing.Color]::FromArgb(80, 80, 80)
            $emailLabel.Text = "yuki.suzuki@suzugross.com"
            $statusBar.Items.Add($emailLabel) | Out-Null

            $form.Controls.Add($statusBar)

            # ========================================
            # PC Info helpers
            # ========================================
            function Format-StatusLine {
                param([string]$Content, [string]$Marker, [int]$Width = 50)
                $padding = $Width - $Content.Length - $Marker.Length
                if ($padding -lt 1) { $padding = 1 }
                return "$Content$(' ' * $padding)$Marker"
            }

            # PrefixLength -> dotted-decimal subnet mask, mirroring the
            # conversion in Get-CurrentPCInfo (kernel/common.ps1). The
            # hostlist stores masks as dotted decimal (e.g. 255.255.0.0),
            # so the comparison is mask-vs-mask. Returns '' for an
            # out-of-range prefix so a bogus value can never yield a
            # spurious mask match. Defined here (before Update-PCInfoDisplay)
            # so it exists for the first display call at startup.
            function ConvertTo-SubnetMask {
                param([int]$PrefixLength)
                if ($PrefixLength -le 0 -or $PrefixLength -gt 32) { return '' }
                $maskInt = [uint32]([math]::Pow(2, 32) - [math]::Pow(2, 32 - $PrefixLength))
                return "{0}.{1}.{2}.{3}" -f `
                    (($maskInt -shr 24) -band 0xFF),
                    (($maskInt -shr 16) -band 0xFF),
                    (($maskInt -shr 8) -band 0xFF),
                    ($maskInt -band 0xFF)
            }

            function Set-ColorizedText {
                param(
                    [System.Windows.Forms.RichTextBox]$RichTextBox,
                    [string]$Text
                )
                # The 25fps art-render timer can starve WM_PAINT for the
                # RichTextBox, which leaves the internal text up-to-date
                # but the screen frozen on the previous frame. Suspending
                # layout around the bulk edit + forcing Refresh() at the
                # end gives the control a synchronous repaint window.
                $RichTextBox.SuspendLayout()
                $RichTextBox.Text = $Text
                $RichTextBox.SelectAll()
                $RichTextBox.SelectionFont = $fontNormal
                $RichTextBox.SelectionColor = $textWhite

                $markerMap = @(
                    @{ Token = "[OK]"; Color = $successGreen },
                    @{ Token = "[!!]"; Color = $errorRed     },
                    @{ Token = "[--]"; Color = $errorRed     },
                    @{ Token = "[ER]"; Color = $errorRed     },
                    @{ Token = "[SK]"; Color = $orangeSkip   },
                    @{ Token = "[CA]"; Color = $textGray     },
                    @{ Token = "[PT]"; Color = $warnYellow   },
                    @{ Token = "[WN]"; Color = $warnYellow   }
                )
                foreach ($m in $markerMap) {
                    $pos = 0
                    while (($idx = $RichTextBox.Text.IndexOf($m.Token, $pos)) -ge 0) {
                        $RichTextBox.Select($idx, $m.Token.Length)
                        $RichTextBox.SelectionColor = $m.Color
                        $pos = $idx + $m.Token.Length
                    }
                }
                $RichTextBox.Select(0, 0)
                $RichTextBox.ResumeLayout($true)
                $RichTextBox.Refresh()
            }

            $script:cachedCurrent = @{
                Hostname    = $env:COMPUTERNAME
                EthIPs      = @()
                WifiIPs     = @()
                EthGateway  = ''
                WifiGateway = ''
                DNS         = @()
                Printers    = @()
            }

            function Update-CurrentSnapshot {
                param([string]$Tier)

                if ($Tier -eq 'hostname' -or $Tier -eq 'all') {
                    $script:cachedCurrent.Hostname = $env:COMPUTERNAME
                }

                if ($Tier -eq 'network' -or $Tier -eq 'all') {
                    try {
                        # Adapter classification ported from Get-CurrentPCInfo
                        # (kernel/common.ps1): iterate PHYSICAL adapters and
                        # classify by InterfaceDescription (always English,
                        # locale-independent) instead of InterfaceAlias. The
                        # alias is localized on Japanese Windows (the default
                        # Ethernet connection name is not the English word
                        # "Ethernet"), and an unfiltered alias match would also
                        # over-match Hyper-V "vEthernet (...)". The default
                        # gateway is read per interface index so a down adapter
                        # cannot borrow a live adapter's gateway. Reset first so
                        # a link-down transition clears stale values.
                        $script:cachedCurrent.EthIPs      = @()
                        $script:cachedCurrent.WifiIPs     = @()
                        $script:cachedCurrent.EthGateway  = ''
                        $script:cachedCurrent.WifiGateway = ''
                        foreach ($ad in @(Get-NetAdapter -Physical -ErrorAction SilentlyContinue | Where-Object { $_.Status -ne 'Disabled' })) {
                            $ips = @(Get-NetIPAddress -InterfaceIndex $ad.ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue |
                                Where-Object { $_.PrefixOrigin -ne 'WellKnown' })
                            if ($ips.Count -eq 0) { continue }
                            $gw = ''
                            $cfg = Get-NetIPConfiguration -InterfaceIndex $ad.ifIndex -ErrorAction SilentlyContinue
                            if ($cfg -and $cfg.IPv4DefaultGateway) { $gw = [string]@($cfg.IPv4DefaultGateway.NextHop)[0] }
                            if ($ad.InterfaceDescription -match 'Wi-Fi|Wireless|WLAN|802\.11') {
                                $script:cachedCurrent.WifiIPs     = $ips
                                $script:cachedCurrent.WifiGateway = $gw
                            }
                            elseif ($script:cachedCurrent.EthIPs.Count -eq 0) {
                                $script:cachedCurrent.EthIPs     = $ips
                                $script:cachedCurrent.EthGateway = $gw
                            }
                        }

                        $script:cachedCurrent.DNS = @(Get-DnsClientServerAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue |
                            ForEach-Object { $_.ServerAddresses } |
                            Where-Object { $_ -and $_ -notlike 'fec0:*' } |
                            Select-Object -Unique)
                    } catch { }
                }

                if ($Tier -eq 'printer' -or $Tier -eq 'all') {
                    try {
                        $script:cachedCurrent.Printers = @(Get-Printer -ErrorAction SilentlyContinue)
                    } catch {
                        $script:cachedCurrent.Printers = @()
                    }
                }
            }

            function Update-PCInfoDisplay {
                $hi = $FabriqToolbarShared.TargetHostInfo
                $cur = $script:cachedCurrent

                $hasAnyTarget = $false
                if ($null -ne $hi) {
                    foreach ($key in @('Hostname','KanriNo','Pin','EthIP','EthSubnet','EthGateway','WifiIP','WifiSubnet','WifiGateway')) {
                        if (-not [string]::IsNullOrWhiteSpace([string]$hi.$key)) { $hasAnyTarget = $true; break }
                    }
                    if (-not $hasAnyTarget -and $hi.DNS -and $hi.DNS.Count -gt 0)           { $hasAnyTarget = $true }
                    if (-not $hasAnyTarget -and $hi.Printers -and $hi.Printers.Count -gt 0) { $hasAnyTarget = $true }
                }

                if (-not $hasAnyTarget) {
                    $placeholder = "Waiting for host selection..."
                    if ($script:lastPcInfoText -ne $placeholder) {
                        Set-ColorizedText -RichTextBox $pcInfoRtb -Text $placeholder
                        $script:lastPcInfoText = $placeholder
                    }
                    return
                }

                $pcText = ""

                if (-not [string]::IsNullOrWhiteSpace([string]$hi.KanriNo)) {
                    $pcText += "ID:        $($hi.KanriNo)`r`n"
                }
                if (-not [string]::IsNullOrWhiteSpace([string]$hi.Pin)) {
                    $pcText += "PIN:       $($hi.Pin)`r`n"
                }
                if ($pcText) { $pcText += "`r`n" }

                # PC Name
                if (-not [string]::IsNullOrWhiteSpace([string]$hi.Hostname)) {
                    $curName = $cur.Hostname
                    $tgtName = $hi.Hostname
                    if ($curName -eq $tgtName) {
                        $pcText += (Format-StatusLine "PC Name:   $curName" "[OK]") + "`r`n"
                    } else {
                        $pcText += (Format-StatusLine "PC Name:   $curName" "[!!]") + "`r`n"
                        $pcText += "           -> $tgtName`r`n"
                    }
                    $pcText += "`r`n"
                }

                # Ethernet
                $hasEth = (-not [string]::IsNullOrWhiteSpace([string]$hi.EthIP)) -or
                          (-not [string]::IsNullOrWhiteSpace([string]$hi.EthSubnet)) -or
                          (-not [string]::IsNullOrWhiteSpace([string]$hi.EthGateway))
                if ($hasEth) {
                    $pcText += "[Ethernet]`r`n"
                    $ethIps = @($cur.EthIPs | Select-Object -ExpandProperty IPAddress)
                    $ethPfx = @($cur.EthIPs | Select-Object -ExpandProperty PrefixLength)

                    if (-not [string]::IsNullOrWhiteSpace([string]$hi.EthIP)) {
                        $curVal = if ($ethIps.Count -gt 0) { $ethIps[0] } else { "(none)" }
                        if ($ethIps -contains $hi.EthIP) {
                            $pcText += (Format-StatusLine "  IP:      $curVal" "[OK]") + "`r`n"
                        } else {
                            $pcText += (Format-StatusLine "  IP:      $curVal" "[!!]") + "`r`n"
                            $pcText += "           -> $($hi.EthIP)`r`n"
                        }
                    }
                    if (-not [string]::IsNullOrWhiteSpace([string]$hi.EthSubnet)) {
                        $curMask = if ($ethPfx.Count -gt 0) { ConvertTo-SubnetMask $ethPfx[0] } else { '' }
                        $curDisp = if ($curMask) { $curMask } else { "(none)" }
                        if ($curMask -eq [string]$hi.EthSubnet) {
                            $pcText += (Format-StatusLine "  Subnet:  $curDisp" "[OK]") + "`r`n"
                        } else {
                            $pcText += (Format-StatusLine "  Subnet:  $curDisp" "[!!]") + "`r`n"
                            $pcText += "           -> $($hi.EthSubnet)`r`n"
                        }
                    }
                    if (-not [string]::IsNullOrWhiteSpace([string]$hi.EthGateway)) {
                        $curVal = if (-not [string]::IsNullOrWhiteSpace([string]$cur.EthGateway)) { $cur.EthGateway } else { "(none)" }
                        if ($cur.EthGateway -eq $hi.EthGateway) {
                            $pcText += (Format-StatusLine "  GW:      $curVal" "[OK]") + "`r`n"
                        } else {
                            $pcText += (Format-StatusLine "  GW:      $curVal" "[!!]") + "`r`n"
                            $pcText += "           -> $($hi.EthGateway)`r`n"
                        }
                    }
                    $pcText += "`r`n"
                }

                # Wi-Fi
                $hasWifi = (-not [string]::IsNullOrWhiteSpace([string]$hi.WifiIP)) -or
                           (-not [string]::IsNullOrWhiteSpace([string]$hi.WifiSubnet)) -or
                           (-not [string]::IsNullOrWhiteSpace([string]$hi.WifiGateway))
                if ($hasWifi) {
                    $pcText += "[Wi-Fi]`r`n"
                    $wifiIps = @($cur.WifiIPs | Select-Object -ExpandProperty IPAddress)
                    $wifiPfx = @($cur.WifiIPs | Select-Object -ExpandProperty PrefixLength)

                    if (-not [string]::IsNullOrWhiteSpace([string]$hi.WifiIP)) {
                        $curVal = if ($wifiIps.Count -gt 0) { $wifiIps[0] } else { "(none)" }
                        if ($wifiIps -contains $hi.WifiIP) {
                            $pcText += (Format-StatusLine "  IP:      $curVal" "[OK]") + "`r`n"
                        } else {
                            $pcText += (Format-StatusLine "  IP:      $curVal" "[!!]") + "`r`n"
                            $pcText += "           -> $($hi.WifiIP)`r`n"
                        }
                    }
                    if (-not [string]::IsNullOrWhiteSpace([string]$hi.WifiSubnet)) {
                        $curMask = if ($wifiPfx.Count -gt 0) { ConvertTo-SubnetMask $wifiPfx[0] } else { '' }
                        $curDisp = if ($curMask) { $curMask } else { "(none)" }
                        if ($curMask -eq [string]$hi.WifiSubnet) {
                            $pcText += (Format-StatusLine "  Subnet:  $curDisp" "[OK]") + "`r`n"
                        } else {
                            $pcText += (Format-StatusLine "  Subnet:  $curDisp" "[!!]") + "`r`n"
                            $pcText += "           -> $($hi.WifiSubnet)`r`n"
                        }
                    }
                    if (-not [string]::IsNullOrWhiteSpace([string]$hi.WifiGateway)) {
                        $curVal = if (-not [string]::IsNullOrWhiteSpace([string]$cur.WifiGateway)) { $cur.WifiGateway } else { "(none)" }
                        if ($cur.WifiGateway -eq $hi.WifiGateway) {
                            $pcText += (Format-StatusLine "  GW:      $curVal" "[OK]") + "`r`n"
                        } else {
                            $pcText += (Format-StatusLine "  GW:      $curVal" "[!!]") + "`r`n"
                            $pcText += "           -> $($hi.WifiGateway)`r`n"
                        }
                    }
                    $pcText += "`r`n"
                }

                # DNS
                if ($hi.DNS -and $hi.DNS.Count -gt 0) {
                    $targetDns  = @($hi.DNS | Where-Object { -not [string]::IsNullOrEmpty($_) } | Sort-Object)
                    $currentDns = @($cur.DNS | Where-Object { -not [string]::IsNullOrEmpty($_) } | Sort-Object)
                    $tgtStr = $targetDns -join ", "
                    $curStr = $currentDns -join ", "
                    $allMatched = $true
                    foreach ($d in $targetDns) {
                        if ($cur.DNS -notcontains $d) { $allMatched = $false; break }
                    }
                    if ($allMatched) {
                        $pcText += (Format-StatusLine "[DNS]  $tgtStr" "[OK]") + "`r`n"
                    } else {
                        $curDisplay = if ($curStr) { $curStr } else { "(none)" }
                        $pcText += (Format-StatusLine "[DNS]  $curDisplay" "[!!]") + "`r`n"
                        $pcText += "       -> $tgtStr`r`n"
                    }
                    $pcText += "`r`n"
                }

                # Printers
                if ($hi.Printers -and $hi.Printers.Count -gt 0) {
                    $pcText += "[Printers]`r`n"
                    foreach ($tp in $hi.Printers) {
                        $pName = $tp.Name
                        if ($pName.Length -gt 38) { $pName = $pName.Substring(0, 35) + "..." }
                        $found = $cur.Printers | Where-Object { $_.Name -eq $tp.Name } | Select-Object -First 1
                        if ($null -ne $found) {
                            $driverOk = [string]::IsNullOrWhiteSpace([string]$tp.Driver) -or ($found.DriverName -eq $tp.Driver)
                            $portOk = [string]::IsNullOrWhiteSpace([string]$tp.Port)
                            if (-not $portOk) {
                                # printer_driver_config names Standard TCP/IP ports "IP_<value>"
                                # (printer_config.ps1) while the hostlist stores the bare value;
                                # accept an exact match or the leading "IP_" stripped, mirroring the
                                # module's own post-apply check (expected PortName = "IP_<Port>").
                                $cp  = ([string]$found.PortName).Trim()
                                $tpp = ([string]$tp.Port).Trim()
                                $portOk = ($cp -eq $tpp) -or (($cp -replace '^IP_', '') -eq $tpp)
                            }
                            if ($driverOk -and $portOk) {
                                $pcText += (Format-StatusLine "  $pName" "[OK]") + "`r`n"
                            } else {
                                $detail = if (-not $driverOk -and -not $portOk) { "drv+port" }
                                          elseif (-not $driverOk)               { "drv" }
                                          else                                  { "port" }
                                $pcText += (Format-StatusLine "  $pName ($detail)" "[!!]") + "`r`n"
                            }
                        } else {
                            $pcText += (Format-StatusLine "  $pName" "[--]") + "`r`n"
                        }
                    }
                }

                if ($script:lastPcInfoText -ne $pcText) {
                    Set-ColorizedText -RichTextBox $pcInfoRtb -Text $pcText
                    $script:lastPcInfoText = $pcText
                }
            }

            # ========================================
            # Execution panel (ported from status_monitor.ps1
            # Update-StatusDisplay execution-summary block). Self-contained
            # status.json read with its own LastWriteTime change-guard.
            # ========================================
            function Update-ExecutionDisplay {
                if (-not (Test-Path $statusFilePath)) {
                    if ($script:lastExecText -ne '__WAITING__') {
                        Set-ColorizedText -RichTextBox $execRtb -Text "Waiting for status..."
                        $script:lastExecText = '__WAITING__'
                    }
                    return
                }
                try {
                    $fileInfo = Get-Item $statusFilePath -ErrorAction Stop
                    # Skip reparse when the file has not changed since last tick.
                    if ($fileInfo.LastWriteTime -eq $script:execLastStatusWriteTime) { return }
                    $script:execLastStatusWriteTime = $fileInfo.LastWriteTime

                    # Lock-free read with retry (kernel writes atomically via
                    # temp + Move-Item, so a partial read is never observed;
                    # the retry only covers the brief rename window).
                    $jsonText = $null
                    for ($retry = 0; $retry -lt 3; $retry++) {
                        try {
                            $jsonText = [System.IO.File]::ReadAllText($statusFilePath, [System.Text.Encoding]::UTF8)
                            break
                        } catch { Start-Sleep -Milliseconds 50 }
                    }
                    if ([string]::IsNullOrEmpty($jsonText)) { return }

                    $status = $jsonText | ConvertFrom-Json
                    $exec = $status.Execution

                    $execText = ""
                    if ($null -eq $exec -or ($exec.TotalCount -eq 0 -and $exec.Phase -eq "idle")) {
                        $execText = "No execution data yet."
                    }
                    else {
                        $phaseLabel = switch ($exec.Phase) {
                            "idle"      { "Idle" }
                            "executing" { ">> Running..." }
                            "complete"  { "Complete" }
                            default     { [string]$exec.Phase }
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
                            for ($i = 0; $i -lt $details.Count; $i++) {
                                $d = $details[$i]

                                # Session-boundary separator
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

                                # Prefix restored (pre-__RESTART__) entries with ^
                                $prefix = ""
                                if ($d.IsRestored -eq $true) { $prefix = "^ " }

                                $msg = if ($d.Message) { " $($d.Message)" } else { "" }
                                $line = "$prefix$icon $($d.Operation)$msg"
                                if ($line.Length -gt 70) { $line = $line.Substring(0, 67) + "..." }
                                $execText += "$line`r`n"
                            }
                        }
                    }

                    if ($script:lastExecText -ne $execText) {
                        Set-ColorizedText -RichTextBox $execRtb -Text $execText
                        $script:lastExecText = $execText
                        # Follow the tail so the newest module stays visible
                        # as Details grows during a long profile run.
                        $execRtb.SelectionStart = $execRtb.Text.Length
                        $execRtb.ScrollToCaret()
                    }
                }
                catch { }
            }

            # ========================================
            # Art panel (ported verbatim from status_monitor.ps1)
            # ========================================
            function Initialize-ArtBuffer {
                $w = $artCanvas.ClientSize.Width
                $h = $artCanvas.ClientSize.Height
                if ($w -le 0 -or $h -le 0) { return }

                # Detach the current back-buffer from the PictureBox BEFORE
                # disposing it. The PictureBox holds $artBufferBitmap through
                # its .Image property; disposing the bitmap while it is still
                # assigned leaves the control pointing at a dead GDI+ image.
                # The next WM_PAINT - which a display change / Explorer
                # restart triggers via the resize that brought us here - then
                # throws ArgumentException ("Parameter is not valid") inside
                # PictureBox.OnPaint and escapes as an unhandled UI-thread
                # exception (the crash dialog operators reported).
                $artCanvas.Image = $null

                if ($null -ne $script:artBufferGraphics) { $script:artBufferGraphics.Dispose() }
                if ($null -ne $script:artBufferBitmap)   { $script:artBufferBitmap.Dispose() }
                $script:artBufferBitmap   = New-Object System.Drawing.Bitmap($w, $h)
                $script:artBufferGraphics = [System.Drawing.Graphics]::FromImage($script:artBufferBitmap)
                $script:artBufferGraphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit
                # Reattach the fresh (blank) buffer immediately so the
                # PictureBox always has a live image to paint, even before
                # the next render tick reassigns it.
                $artCanvas.Image = $script:artBufferBitmap
            }

            function Select-NextArtSentence {
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
                    $script:artDisplayLines.Add(@{
                        Text  = $script:artCurrentText
                        Color = "dim"
                        Age   = 0
                    }) | Out-Null
                    while ($script:artDisplayLines.Count -gt $script:ART_MAX_LINES) {
                        $script:artDisplayLines.RemoveAt(0)
                    }
                }

                if ($script:artRng.Next(100) -lt 25) {
                    $script:artDisplayLines.Clear()
                }

                $script:artContinueOnSameLine = ($script:artDisplayLines.Count -gt 0 -and $script:artRng.Next(100) -lt 40)

                Select-NextArtSentence
                $script:artState = "typing"
            }

            function Sync-ArtSilenceFlag {
                try {
                    $present = Test-Path $silenceFlagPath
                    if ($present -eq $script:artSilent) { return }
                    if ($present) {
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
                        try {
                            if (Test-Path $artPulseFilePath) {
                                $val = [int][System.IO.File]::ReadAllText($artPulseFilePath).Trim()
                                $script:artLastPulseValue = $val
                            }
                        } catch { }
                    }
                    $script:artSilent = $present
                } catch { }
            }

            function Sync-ArtPulseFile {
                try {
                    if (Test-Path $artPulseFilePath) {
                        $val = [int][System.IO.File]::ReadAllText($artPulseFilePath).Trim()
                        if ($val -gt $script:artLastPulseValue) {
                            $diff = $val - $script:artLastPulseValue
                            if ($script:artCurrentPhase -eq "executing") {
                                $script:artTriggerQueue += $diff
                                $script:artLastPulseValue = $val
                            }
                        }
                    }
                } catch { }
            }

            function Sync-ArtStatusFile {
                if (-not (Test-Path $statusFilePath)) { return }
                try {
                    $fileInfo = Get-Item $statusFilePath -ErrorAction Stop
                    if ($fileInfo.LastWriteTime -eq $script:artLastStatusWriteTime) { return }
                    $script:artLastStatusWriteTime = $fileInfo.LastWriteTime

                    $json = $null
                    for ($retry = 0; $retry -lt 3; $retry++) {
                        try {
                            $json = [System.IO.File]::ReadAllText($statusFilePath, [System.Text.Encoding]::UTF8)
                            break
                        } catch { Start-Sleep -Milliseconds 50 }
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
                } catch { }
            }

            function Write-ColoredText {
                param(
                    [System.Drawing.Graphics]$Graphics,
                    [string]$Text,
                    [System.Drawing.Font]$Font,
                    [System.Drawing.Brush]$BaseBrush,
                    [System.Drawing.Brush]$AlphaBrush,
                    [System.Drawing.RectangleF]$Rect,
                    [System.Drawing.StringFormat]$SF
                )

                if ($Text -notmatch '[a-zA-Z]') {
                    $Graphics.DrawString($Text, $Font, $BaseBrush, $Rect, $SF)
                    return
                }
                if ($Text -match '^[a-zA-Z ]+$') {
                    $Graphics.DrawString($Text, $Font, $AlphaBrush, $Rect, $SF)
                    return
                }

                $sfTypo = [System.Drawing.StringFormat]::GenericTypographic.Clone()
                $sfTypo.FormatFlags = [System.Drawing.StringFormatFlags]::MeasureTrailingSpaces -bor [System.Drawing.StringFormatFlags]::NoWrap
                $fullSize = $Graphics.MeasureString($Text, $Font, [System.Drawing.PointF]::new(0, 0), $sfTypo)

                if ($fullSize.Width -gt $Rect.Width) {
                    $Graphics.DrawString($Text, $Font, $BaseBrush, $Rect, $SF)
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

            function Update-ArtFrame {
                if ($null -eq $script:artBufferGraphics) { return }

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

                if ($script:artState -eq "waiting" -and $script:artTriggerQueue -gt 0) {
                    $script:artTriggerQueue--
                    Invoke-ArtTrigger
                }

                if ($script:artState -eq "typing" -and $script:artCurrentText.Length -gt 0 -and $script:artCursorPos -lt $script:artCurrentText.Length) {
                    $elapsed = ($now - $script:artLastTypeTime).TotalMilliseconds
                    if ($elapsed -ge $script:artTypeSpeed) {
                        $charsToAdvance = [Math]::Max(1, [int]($elapsed / $script:artTypeSpeed))
                        $script:artCursorPos = [Math]::Min($script:artCurrentText.Length, $script:artCursorPos + $charsToAdvance)
                        $script:artLastTypeTime = $now
                    }
                }
                elseif ($script:artState -eq "typing" -and $script:artCurrentText.Length -gt 0 -and $script:artCursorPos -ge $script:artCurrentText.Length) {
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

                if (($now - $script:artCursorBlinkTime).TotalMilliseconds -ge 500) {
                    $script:artCursorVisible = -not $script:artCursorVisible
                    $script:artCursorBlinkTime = $now
                }

                $g.FillRectangle($script:artBgBrush, 0, 0, $w, $h)

                if ($script:artFlashFrames -gt 0) {
                    $alpha = [Math]::Min(80, $script:artFlashFrames * 12)
                    $flashBrush = New-Object System.Drawing.SolidBrush(
                        [System.Drawing.Color]::FromArgb($alpha, $script:artFlashColor.R, $script:artFlashColor.G, $script:artFlashColor.B)
                    )
                    $g.FillRectangle($flashBrush, 0, 0, $w, $h)
                    $flashBrush.Dispose()
                    $script:artFlashFrames--
                }

                $sf = [System.Drawing.StringFormat]::GenericDefault
                $maxTextWidth = $w - 16
                $y = 8

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

                for ($i = $startIdx; $i -lt $script:artDisplayLines.Count; $i++) {
                    $line = $script:artDisplayLines[$i]
                    $line.Age++

                    $brush      = if ($line.Age -lt 60) { $script:artDimGreenBrush }       else { $script:artDimGrayBrush }
                    $alphaBrush = if ($line.Age -lt 60) { $script:artAlphaDimGreenBrush } else { $script:artAlphaDimGrayBrush }

                    $text = $line.Text
                    if ($script:artGlitchFrames -gt 0 -and $script:artRng.Next(100) -lt 20) {
                        $pos = $script:artRng.Next([Math]::Min($text.Length, [Math]::Max(1, $text.Length)))
                        $glitchChar = $script:GLITCH_CHARS[$script:artRng.Next($script:GLITCH_CHARS.Count)]
                        $text = $text.Substring(0, $pos) + $glitchChar + $text.Substring([Math]::Min($pos + 1, $text.Length))
                        $brush = New-Object System.Drawing.SolidBrush($script:artGlitchWhite)
                    }

                    $rect = New-Object System.Drawing.RectangleF(8, $y, $maxTextWidth, ($h - $y))
                    Write-ColoredText -Graphics $g -Text $text -Font $script:artFont -BaseBrush $brush -AlphaBrush $alphaBrush -Rect $rect -SF $sf
                    $measured = $g.MeasureString($text, $script:artFont, $maxTextWidth, $sf)
                    $y += [Math]::Max($script:artLineH, [int][Math]::Ceiling($measured.Height))
                }

                if ($script:artState -eq "typing" -and $script:artCurrentText.Length -gt 0) {
                    $textStr = $script:artCurrentText.Substring(0, $script:artCursorPos)

                    $xOffset = 8
                    $drawWidth = $maxTextWidth
                    if ($script:artContinueOnSameLine -and $script:artDisplayLines.Count -gt 0) {
                        $lastLine = $script:artDisplayLines[$script:artDisplayLines.Count - 1]
                        $lastMeasured = $g.MeasureString($lastLine.Text, $script:artFont, $maxTextWidth, $sf)
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

                    if ($script:artCursorPos -lt $script:artCurrentText.Length -and $script:artRng.Next(100) -lt 8) {
                        $glitchChar = $script:GLITCH_CHARS[$script:artRng.Next($script:GLITCH_CHARS.Count)]
                        $glitchText = $textStr + $glitchChar
                        $glitchBrush = New-Object System.Drawing.SolidBrush($script:artGlitchWhite)
                        $typeRect = New-Object System.Drawing.RectangleF($xOffset, $y, $drawWidth, ($h - $y))
                        Write-ColoredText -Graphics $g -Text $glitchText -Font $script:artFontBold -BaseBrush $glitchBrush -AlphaBrush $script:artAlphaBrush -Rect $typeRect -SF $sf
                        $glitchBrush.Dispose()
                    }
                    else {
                        $typeRect = New-Object System.Drawing.RectangleF($xOffset, $y, $drawWidth, ($h - $y))
                        Write-ColoredText -Graphics $g -Text $textStr -Font $script:artFontBold -BaseBrush $script:artCyanBrush -AlphaBrush $script:artAlphaBrush -Rect $typeRect -SF $sf

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
                    if ($script:artCursorVisible) {
                        $g.DrawString("_", $script:artFontBold, $script:artCyanBrush, 8, $y)
                    }
                }

                if ($script:artGlitchFrames -gt 0) { $script:artGlitchFrames-- }

                $artCanvas.Image = $script:artBufferBitmap
            }

            # ========================================
            # Button handlers
            # ========================================
            $btnSkip.Add_Click({
                try {
                    $cfg      = Get-FabriqAsyncConfig
                    $flagPath = $cfg.SkipFlagPath
                    if (-not [System.IO.Path]::IsPathRooted($flagPath)) {
                        $flagPath = Join-Path $FabriqToolbarShared.FabriqRoot $flagPath
                    }
                    $flagPath = [System.IO.Path]::GetFullPath($flagPath)

                    $flagDir = Split-Path $flagPath -Parent
                    if (-not (Test-Path $flagDir)) {
                        New-Item -Path $flagDir -ItemType Directory -Force | Out-Null
                    }

                    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    "requested at $ts" | Out-File -FilePath $flagPath -Encoding UTF8 -Force

                    $statusLabel.ForeColor = $orangeSkip
                    $statusLabel.Text = "Skip requested (effective only for async modules)"
                }
                catch {
                    $statusLabel.ForeColor = $errorRed
                    $statusLabel.Text = "Skip request failed: $($_.Exception.Message)"
                }
            })

            $btnGyotaq.Add_Click({
                $wasEnabled = $btnGyotaq.Enabled
                $btnGyotaq.Enabled = $false

                $savedLocation = $form.Location
                $savedSize     = $form.Size
                $form.Hide()
                [System.Windows.Forms.Application]::DoEvents()
                Start-Sleep -Milliseconds 300

                $result   = $null
                $errorMsg = $null
                try {
                    $global:FabriqEvidenceBasePath = $FabriqToolbarShared.EvidenceBasePath
                    $baseDir = if (-not [string]::IsNullOrWhiteSpace($global:FabriqEvidenceBasePath)) {
                        Join-Path $global:FabriqEvidenceBasePath "gyotaku"
                    } else {
                        $gyotakuDir
                    }
                    $result = Save-Screenshot -BaseDir $baseDir
                }
                catch {
                    $errorMsg = $_.Exception.Message
                }
                finally {
                    $form.Location = $savedLocation
                    $form.Size     = $savedSize
                    $form.Show()
                    $btnGyotaq.Enabled = $wasEnabled
                    [System.Windows.Forms.Application]::DoEvents()
                }

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
            # Timers
            # ========================================
            $script:tickCount = 0

            # Cache of the most-recent rendered PC Info text. The 1s
            # tick regenerates this on every fire even when nothing has
            # changed; skipping the RichTextBox rewrite when content is
            # identical avoids needless invalidation churn that competes
            # with the art panel for paint cycles.
            $script:lastPcInfoText = $null

            $timer = New-Object System.Windows.Forms.Timer
            $timer.Interval = 1000
            $timer.Add_Tick({
                if ($FabriqToolbarShared.CloseRequested) {
                    $timer.Stop()
                    $form.Close()
                    return
                }

                $script:tickCount++

                Update-CurrentSnapshot -Tier 'hostname'
                if (($script:tickCount % 3) -eq 0) {
                    Update-CurrentSnapshot -Tier 'network'
                }
                if (($script:tickCount % 10) -eq 0) {
                    Update-CurrentSnapshot -Tier 'printer'
                }

                Update-PCInfoDisplay
                Update-ExecutionDisplay

                $running = ($FabriqToolbarShared.State -eq 'Running')
                $modName = $FabriqToolbarShared.ModuleName
                if ($running) {
                    $statusLabel.ForeColor = $textWhite
                    $statusLabel.Text = if ([string]::IsNullOrWhiteSpace($modName)) { "Running..." } else { "Running: $modName" }
                } else {
                    if ($statusLabel.ForeColor -ne $orangeSkip -and
                        $statusLabel.ForeColor -ne $successGreen -and
                        $statusLabel.ForeColor -ne $errorRed) {
                        $statusLabel.ForeColor = $textGray
                        $statusLabel.Text = "Idle"
                    }
                }
                if ($btnSkip.Enabled   -ne $running) { $btnSkip.Enabled   = $running }
                if ($btnGyotaq.Enabled -ne $running) { $btnGyotaq.Enabled = $running }
            })

            $artRenderTimer = New-Object System.Windows.Forms.Timer
            $artRenderTimer.Interval = $script:ART_RENDER_INTERVAL
            $artRenderTimer.Add_Tick({
                try { Update-ArtFrame }
                catch {
                    # A GDI+ failure during a display change can corrupt the
                    # back-buffer. Rebuild it so the panel recovers on the
                    # next tick instead of repainting a broken image forever.
                    try { Initialize-ArtBuffer } catch { }
                }
            })

            $artPulseTimer = New-Object System.Windows.Forms.Timer
            $artPulseTimer.Interval = $script:ART_PULSE_INTERVAL
            $artPulseTimer.Add_Tick({
                Sync-ArtSilenceFlag
                Sync-ArtPulseFile
                Sync-ArtStatusFile
            })

            $form.Add_Shown({ Initialize-ArtBuffer })
            $artCanvas.Add_Resize({ Initialize-ArtBuffer })

            $form.Add_FormClosing({
                try { $timer.Stop();          $timer.Dispose()          } catch { }
                try { $artRenderTimer.Stop(); $artRenderTimer.Dispose() } catch { }
                try { $artPulseTimer.Stop();  $artPulseTimer.Dispose()  } catch { }
                # Detach the buffer from the PictureBox before disposing it
                # so a final paint cannot land on a dead image (same hazard
                # as Initialize-ArtBuffer).
                try { $artCanvas.Image = $null } catch { }
                try { if ($null -ne $script:artBufferGraphics)     { $script:artBufferGraphics.Dispose() } } catch { }
                try { if ($null -ne $script:artBufferBitmap)       { $script:artBufferBitmap.Dispose() } } catch { }
                foreach ($b in @($script:artBgBrush, $script:artCyanBrush, $script:artDimGreenBrush, $script:artDimGrayBrush, $script:artAlphaBrush, $script:artAlphaDimGreenBrush, $script:artAlphaDimGrayBrush)) {
                    try { if ($null -ne $b) { $b.Dispose() } } catch { }
                }
                foreach ($f in @($script:artFont, $script:artFontBold, $fontNormal, $fontBold)) {
                    try { if ($null -ne $f) { $f.Dispose() } } catch { }
                }
            })

            Update-CurrentSnapshot -Tier 'all'
            Update-PCInfoDisplay
            Update-ExecutionDisplay

            $timer.Start()
            $artRenderTimer.Start()
            $artPulseTimer.Start()

            [System.Windows.Forms.Application]::Run($form)

            try { $form.Dispose() } catch { }
        }
        catch {
            # Toolbar runspace failures are invisible to the operator;
            # swallow so the BeginInvoke handle ends cleanly.
        }
    })

    $script:ExecutionToolbarHandle   = $ps.BeginInvoke()
    $script:ExecutionToolbarPS       = $ps
    $script:ExecutionToolbarRunspace = $rs
    $script:ExecutionToolbarShared   = $shared

    # Self-push host info from current SELECTED_* env vars. The kernel
    # may have already populated these (Set-SelectedHostEnvironment in
    # fresh-start, Restore-HostEnvironment in resume) BEFORE the toolbar
    # was started, so without this self-push we would sit at the
    # placeholder until the next host re-selection.
    try {
        Update-ExecutionToolbar -TargetHostInfo (Get-FabriqHostInfoFromEnv)
    }
    catch { }
}

function Hide-ExecutionToolbar {
    <#
    .SYNOPSIS
        Signal the toolbar runspace to shut down cleanly and dispose
        the runspace handles.
    #>
    [CmdletBinding()]
    param()

    if ($null -eq $script:ExecutionToolbarRunspace) { return }

    if ($null -ne $script:ExecutionToolbarShared) {
        try { $script:ExecutionToolbarShared.CloseRequested = $true } catch { }
    }

    if ($null -ne $script:ExecutionToolbarHandle) {
        $waitStart = Get-Date
        while (-not $script:ExecutionToolbarHandle.IsCompleted -and ((Get-Date) - $waitStart).TotalSeconds -lt 2) {
            Start-Sleep -Milliseconds 50
        }

        if (-not $script:ExecutionToolbarHandle.IsCompleted) {
            try { $script:ExecutionToolbarPS.Stop() } catch { }
        }
        else {
            try { [void]$script:ExecutionToolbarPS.EndInvoke($script:ExecutionToolbarHandle) } catch { }
        }
    }

    try { if ($null -ne $script:ExecutionToolbarPS)       { $script:ExecutionToolbarPS.Dispose() } } catch { }
    try {
        if ($null -ne $script:ExecutionToolbarRunspace) {
            $script:ExecutionToolbarRunspace.Close()
            $script:ExecutionToolbarRunspace.Dispose()
        }
    } catch { }

    $script:ExecutionToolbarRunspace = $null
    $script:ExecutionToolbarPS       = $null
    $script:ExecutionToolbarHandle   = $null
    $script:ExecutionToolbarShared   = $null
}

function Update-ExecutionToolbar {
    <#
    .SYNOPSIS
        Push state changes to the toolbar runspace via the shared
        hashtable. The runspace's main Timer applies them on next tick.

    .PARAMETER ModuleName
        Display name of the module currently executing.

    .PARAMETER ExecutionState
        'Idle'    -> Skip / Gyotaq buttons disabled
        'Running' -> Skip / Gyotaq buttons enabled

    .PARAMETER TargetHostInfo
        Hashtable of selected hostlist row values. See execution_toolbar
        block in Set-SelectedHostEnvironment (kernel/main.ps1) for the
        full key list. Empty / missing keys mean "field not set in
        hostlist" and the corresponding rows / sections are omitted.
    #>
    [CmdletBinding()]
    param(
        [string]$ModuleName = "",
        [ValidateSet('Idle', 'Running')]
        [string]$ExecutionState = 'Idle',
        [hashtable]$TargetHostInfo = $null
    )

    if ($null -eq $script:ExecutionToolbarShared) { return }

    try {
        $script:ExecutionToolbarShared.State      = $ExecutionState
        $script:ExecutionToolbarShared.ModuleName = $ModuleName
        if (-not [string]::IsNullOrWhiteSpace($global:FabriqEvidenceBasePath)) {
            $script:ExecutionToolbarShared.EvidenceBasePath = $global:FabriqEvidenceBasePath
        }
        if ($null -ne $TargetHostInfo) {
            $script:ExecutionToolbarShared.TargetHostInfo        = $TargetHostInfo
            $script:ExecutionToolbarShared.TargetHostInfoVersion = [int]$script:ExecutionToolbarShared.TargetHostInfoVersion + 1
        }
    } catch { }
}
