# ========================================
# Fabriq Operator - Execution Toolbar
# ========================================
# In-process floating toolbar that replaces the out-of-process
# Status Monitor (kernel/ps1/status_monitor.ps1, retired in 3.4.0).
#
# Provides operator touch surface while modules are executing:
#   - [Skip]   : write skip_request.flag (honored by Invoke-SafeCommandAsync)
#   - [Gyotaq] : on-demand screenshot via Save-Screenshot (the
#                evidence subfolder stays "gyotaku/" for path
#                compatibility with existing evidence_manager flows)
#   - Status label showing the currently running module
#
# Architecturally an in-process WinForms TopMost window hosted on a
# dedicated STA Runspace. Lives in the same powershell.exe as the
# kernel/dashboard, so Defender / ASR heuristics that block
# "powershell.exe spawning hidden powershell.exe children" do not
# apply (which is why Status Monitor failed to launch on recent
# Windows builds - see CHANGELOG 3.4.0).
#
# Why a dedicated runspace:
#   - The runspace owns its own message loop (Application::Run), so
#     the toolbar is responsive even when the kernel main thread is
#     blocked in Read-Host or a modal ShowDialog (e.g. FlexProfile
#     dashboard, which would otherwise input-disable any sibling
#     window on the SAME thread).
#   - Cross-thread state is a single synchronized hashtable
#     (`$FabriqToolbarShared`): kernel writes, toolbar's WinForms
#     Timer reads at 100ms cadence.
#   - Click handlers (Skip / Gyotaq) run inside the toolbar runspace
#     with common.ps1 + theme.ps1 dot-sourced, so Get-FabriqAsyncConfig
#     / Save-Screenshot / Show-* are available without marshalling.
#
# Public surface (called by kernel main.ps1):
#   - Show-ExecutionToolbar
#   - Hide-ExecutionToolbar
#   - Update-ExecutionToolbar -ExecutionState 'Idle'|'Running' [-ModuleName <s>]
# ========================================

# ----------------------------------------------------------------
# Module-level handles for the toolbar's dedicated runspace
# ----------------------------------------------------------------
$script:ExecutionToolbarRunspace = $null
$script:ExecutionToolbarPS       = $null
$script:ExecutionToolbarHandle   = $null
$script:ExecutionToolbarShared   = $null

# Fabriq root captured at dot-source time. Used to resolve paths
# (skip_request.flag, evidence/gyotaku/, common.ps1, theme.ps1)
# without depending on the current working directory, which modules
# may alter during execution. Derived from $PSScriptRoot
# (.../apps/fabriq_operator/lib) -> three levels up = fabriq root.
$script:ExecutionToolbarFabriqRoot = [System.IO.Path]::GetFullPath(
    (Join-Path $PSScriptRoot "..\..\..")
).TrimEnd('\')

# Header drag is implemented with pure-PowerShell MouseDown/Move/Up
# handlers using Cursor.Position deltas. Add-Type -TypeDefinition for
# WM_NCLBUTTONDOWN P/Invoke would work too, but spawns csc.exe at
# runspace startup, which is needless overhead here.

function Show-ExecutionToolbar {
    <#
    .SYNOPSIS
        Spawn the floating execution toolbar on a dedicated STA Runspace.
    .DESCRIPTION
        Idempotent: a second call while the toolbar runspace is alive
        is a no-op. The runspace runs its own message loop via
        [Application]::Run($form), so the toolbar stays responsive
        regardless of what the kernel main thread is doing.

        Cross-thread state is exchanged via $script:ExecutionToolbarShared
        (a synchronized hashtable). Update-ExecutionToolbar writes to it
        from the main thread; a WinForms Timer inside the runspace polls
        and applies changes to the form.
    #>
    [CmdletBinding()]
    param()

    if ($null -ne $script:ExecutionToolbarHandle -and -not $script:ExecutionToolbarHandle.IsCompleted) {
        return
    }

    # Synchronized cross-thread state. Operations on a Synchronized
    # hashtable serialize through an internal lock, so reads/writes
    # are atomic at the entry level (sufficient for our purposes -
    # all values are simple types, no nested mutation).
    $shared = [hashtable]::Synchronized(@{
        State            = 'Idle'
        ModuleName       = ''
        EvidenceBasePath = if (-not [string]::IsNullOrWhiteSpace($global:FabriqEvidenceBasePath)) { $global:FabriqEvidenceBasePath } else { '' }
        CloseRequested   = $false
        FabriqRoot       = $script:ExecutionToolbarFabriqRoot
    })

    $rs = [runspacefactory]::CreateRunspace($Host)
    $rs.ApartmentState = "STA"
    $rs.ThreadOptions  = "ReuseThread"
    $rs.Open()
    $rs.SessionStateProxy.SetVariable('FabriqToolbarShared',      $shared)
    $rs.SessionStateProxy.SetVariable('FabriqToolbarCommonPath',  (Join-Path $script:ExecutionToolbarFabriqRoot 'kernel\common.ps1'))
    $rs.SessionStateProxy.SetVariable('FabriqToolbarThemePath',   (Join-Path $script:ExecutionToolbarFabriqRoot 'apps\fabriq_operator\lib\theme.ps1'))
    $rs.SessionStateProxy.Path.SetLocation($script:ExecutionToolbarFabriqRoot) | Out-Null

    $ps = [PowerShell]::Create()
    $ps.Runspace = $rs
    [void]$ps.AddScript({
        try {
            Add-Type -AssemblyName System.Windows.Forms
            Add-Type -AssemblyName System.Drawing
            [System.Windows.Forms.Application]::EnableVisualStyles()

            # Dot-source fabriq dependencies into THIS runspace. common.ps1
            # provides Get-FabriqAsyncConfig / Save-Screenshot / Show-Info etc.
            # theme.ps1 provides $script:bg* colors + New-StyledButton.
            . $FabriqToolbarCommonPath
            . $FabriqToolbarThemePath

            # Drag state for the header strip. Lives in this runspace's
            # $script: scope; closures attached to MouseDown / Move / Up
            # bind by name at execution time. No Add-Type / csc.exe spawn
            # (see top of file comment).
            $script:dragState = @{
                Dragging = $false
                OffsetX  = 0
                OffsetY  = 0
            }

            # ---------- Form geometry ----------
            $headerH = 22
            $formW   = 210
            $formH   = 92

            $form = New-Object System.Windows.Forms.Form
            $form.Text            = "fabriq"
            $form.Size            = New-Object System.Drawing.Size($formW, $formH)
            $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
            $form.BackColor       = $script:bgForm
            $form.ForeColor       = $script:fgText
            $form.Font            = $script:fontNormal
            $form.TopMost         = $true
            $form.ShowInTaskbar   = $false
            $form.StartPosition   = [System.Windows.Forms.FormStartPosition]::Manual

            $workArea = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
            $form.Location = New-Object System.Drawing.Point(
                ($workArea.Right - $formW - 16),
                ($workArea.Top + 16)
            )

            # ---------- Header strip ----------
            $headerPanel = New-Object System.Windows.Forms.Panel
            $headerPanel.Location  = New-Object System.Drawing.Point(0, 0)
            $headerPanel.Size      = New-Object System.Drawing.Size($formW, $headerH)
            $headerPanel.BackColor = $script:bgPanel
            $headerPanel.Cursor    = [System.Windows.Forms.Cursors]::SizeAll
            $form.Controls.Add($headerPanel)

            $stripe = New-Object System.Windows.Forms.Panel
            $stripe.Location  = New-Object System.Drawing.Point(0, 0)
            $stripe.Size      = New-Object System.Drawing.Size(3, $headerH)
            $stripe.BackColor = $script:stripeBlue
            $headerPanel.Controls.Add($stripe)

            $headerLabel           = New-Object System.Windows.Forms.Label
            $headerLabel.Text      = "fabriq"
            $headerLabel.Location  = New-Object System.Drawing.Point(10, 3)
            $headerLabel.Size      = New-Object System.Drawing.Size(100, 16)
            $headerLabel.Font      = $script:fontSemiBold
            $headerLabel.ForeColor = $script:fgWhite
            $headerLabel.BackColor = $script:bgPanel
            $headerPanel.Controls.Add($headerLabel)

            # Pure-PowerShell drag handlers.
            # MouseDown: snapshot cursor-to-form offset (constant for the
            #            duration of the drag, regardless of which control
            #            within the header strip received the event).
            # MouseMove: move the form so the cursor maintains the same
            #            offset from the form's top-left.
            # MouseUp:   release.
            $dragMouseDown = {
                param($src, $e)
                if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
                    $f = $src.FindForm()
                    if ($null -ne $f) {
                        $cur = [System.Windows.Forms.Cursor]::Position
                        $script:dragState.Dragging = $true
                        $script:dragState.OffsetX  = $cur.X - $f.Location.X
                        $script:dragState.OffsetY  = $cur.Y - $f.Location.Y
                    }
                }
            }
            $dragMouseMove = {
                param($src, $e)
                if ($script:dragState.Dragging) {
                    $f = $src.FindForm()
                    if ($null -ne $f) {
                        $cur = [System.Windows.Forms.Cursor]::Position
                        $f.Location = New-Object System.Drawing.Point(
                            ($cur.X - $script:dragState.OffsetX),
                            ($cur.Y - $script:dragState.OffsetY)
                        )
                    }
                }
            }
            $dragMouseUp = {
                param($src, $e)
                $script:dragState.Dragging = $false
            }

            foreach ($dragCtl in @($headerPanel, $headerLabel, $stripe)) {
                $dragCtl.Add_MouseDown($dragMouseDown)
                $dragCtl.Add_MouseMove($dragMouseMove)
                $dragCtl.Add_MouseUp($dragMouseUp)
            }

            # ---------- Status label ----------
            $bodyTopY = $headerH + 6

            $lblStatus              = New-Object System.Windows.Forms.Label
            $lblStatus.Name         = "lblStatus"
            $lblStatus.Text         = "Idle"
            $lblStatus.Location     = New-Object System.Drawing.Point(10, $bodyTopY)
            $lblStatus.Size         = New-Object System.Drawing.Size(($formW - 20), 16)
            $lblStatus.Font         = $script:fontNormal
            $lblStatus.ForeColor    = $script:fgDim
            $lblStatus.BackColor    = $script:bgForm
            $lblStatus.AutoEllipsis = $true
            $form.Controls.Add($lblStatus)

            # ---------- Action buttons ----------
            # Skip is small (60x22) yellow to discourage accidental
            # interruption. Gyotaq is larger (108x26) green - benign
            # capture action.
            $btnRowY    = $bodyTopY + 24
            $btnSkipH   = 22
            $btnGyotaqH = 26

            $btnSkip = New-StyledButton -Text "Skip" -X 10 -Y ($btnRowY + ($btnGyotaqH - $btnSkipH) / 2) -Width 60 -Height $btnSkipH -BgColor $script:stripeYellow
            $btnSkip.Name    = "btnSkip"
            $btnSkip.Enabled = $false
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

                    Show-Info "[Toolbar] Skip requested (effective only for async modules)"
                }
                catch {
                    Show-Warning "[Toolbar] Skip request failed: $($_.Exception.Message)"
                }
            })
            $form.Controls.Add($btnSkip)

            $btnGyotaq = New-StyledButton -Text "Gyotaq" -X ($formW - 108 - 10) -Y $btnRowY -Width 108 -Height $btnGyotaqH -BgColor $script:bgAdd
            $btnGyotaq.Name    = "btnGyotaq"
            $btnGyotaq.Enabled = $false
            $btnGyotaq.Add_Click({
                $btnG       = $form.Controls["btnGyotaq"]
                $wasEnabled = $btnG.Enabled
                $btnG.Enabled = $false

                $savedLocation = $form.Location
                $savedSize     = $form.Size
                $form.Hide()
                [System.Windows.Forms.Application]::DoEvents()
                Start-Sleep -Milliseconds 300

                $result   = $null
                $errorMsg = $null
                try {
                    # Save-Screenshot internally reads $global:FabriqEvidenceBasePath
                    # to choose unified vs legacy layout. Sync the snapshot from
                    # $FabriqToolbarShared into this runspace's globals so the
                    # unified path is selected (the kernel updates the snapshot
                    # on every Update-ExecutionToolbar call).
                    $global:FabriqEvidenceBasePath = $FabriqToolbarShared.EvidenceBasePath

                    $baseDir = if (-not [string]::IsNullOrWhiteSpace($global:FabriqEvidenceBasePath)) {
                        Join-Path $global:FabriqEvidenceBasePath "gyotaku"
                    } else {
                        Join-Path $FabriqToolbarShared.FabriqRoot "evidence\gyotaku"
                    }
                    $result = Save-Screenshot -BaseDir $baseDir
                }
                catch {
                    $errorMsg = $_.Exception.Message
                }
                finally {
                    $form.Location   = $savedLocation
                    $form.Size       = $savedSize
                    $form.Show()
                    $btnG.Enabled    = $wasEnabled
                    [System.Windows.Forms.Application]::DoEvents()
                }

                if ($null -ne $result) {
                    Show-Success "[Toolbar] Screenshot saved: $([System.IO.Path]::GetFileName($result))"
                }
                elseif ($errorMsg) {
                    Show-Warning "[Toolbar] Screenshot error: $errorMsg"
                }
                else {
                    Show-Warning "[Toolbar] Screenshot failed (Save-Screenshot returned null)"
                }
            })
            $form.Controls.Add($btnGyotaq)

            # ---------- Timer: $shared -> form sync ----------
            # 100ms polling on the runspace's UI thread. Reads atomic
            # values from $FabriqToolbarShared and applies to the form.
            # Diff-checks prevent unnecessary repaints.
            $timer = New-Object System.Windows.Forms.Timer
            $timer.Interval = 100
            $timer.Add_Tick({
                if ($FabriqToolbarShared.CloseRequested) {
                    $timer.Stop()
                    $form.Close()
                    return
                }

                $running  = ($FabriqToolbarShared.State -eq 'Running')
                $newText  = if ($running) {
                    if ([string]::IsNullOrWhiteSpace($FabriqToolbarShared.ModuleName)) { "Running..." } else { $FabriqToolbarShared.ModuleName }
                } else { "Idle" }
                $newColor = if ($running) { $script:fgText } else { $script:fgDim }

                if ($lblStatus.Text -ne $newText)       { $lblStatus.Text      = $newText }
                if ($lblStatus.ForeColor -ne $newColor) { $lblStatus.ForeColor = $newColor }
                if ($btnSkip.Enabled    -ne $running)   { $btnSkip.Enabled     = $running }
                if ($btnGyotaq.Enabled  -ne $running)   { $btnGyotaq.Enabled   = $running }
            })
            $timer.Start()

            # ---------- Run dedicated message loop ----------
            # Blocks until $form.Close() (signalled by Timer when
            # CloseRequested goes true). Drag / clicks / repaints are
            # processed by THIS loop, independent of whatever the kernel
            # main thread is doing.
            [System.Windows.Forms.Application]::Run($form)

            $timer.Dispose()
            $form.Dispose()
        }
        catch {
            # The toolbar runspace cannot reliably write to the operator
            # console from a failure mode (cross-runspace Show-* may not
            # have loaded yet). Swallow to keep the BeginInvoke handle
            # well-formed; the operator notices via "no toolbar".
        }
    })

    $script:ExecutionToolbarHandle   = $ps.BeginInvoke()
    $script:ExecutionToolbarPS       = $ps
    $script:ExecutionToolbarRunspace = $rs
    $script:ExecutionToolbarShared   = $shared
}

function Hide-ExecutionToolbar {
    <#
    .SYNOPSIS
        Signal the toolbar runspace to shut down cleanly and dispose
        the runspace handles.
    .DESCRIPTION
        Idempotent: safe to call when toolbar is not running. Waits
        up to 2 seconds for the runspace's Application::Run to return
        gracefully (Timer detects CloseRequested -> Form.Close).
        Falls back to PowerShell.Stop() if the runspace hangs.
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
        hashtable. The runspace's Timer applies them within 100ms.
    .DESCRIPTION
        Non-blocking, safe to call when toolbar is not running (no-op).
        Also refreshes the EvidenceBasePath snapshot so Gyotaq picks
        up host-selection updates that happen after toolbar startup.

    .PARAMETER ModuleName
        Display name of the module currently executing. Empty string
        when no module is running.

    .PARAMETER ExecutionState
        'Idle'    -> buttons disabled, label shows 'Idle' (dim)
        'Running' -> buttons enabled, label shows ModuleName
                     (or 'Running...' when ModuleName is empty)
    #>
    [CmdletBinding()]
    param(
        [string]$ModuleName = "",
        [ValidateSet('Idle', 'Running')]
        [string]$ExecutionState = 'Idle'
    )

    if ($null -eq $script:ExecutionToolbarShared) { return }

    try {
        $script:ExecutionToolbarShared.State      = $ExecutionState
        $script:ExecutionToolbarShared.ModuleName = $ModuleName
        if (-not [string]::IsNullOrWhiteSpace($global:FabriqEvidenceBasePath)) {
            $script:ExecutionToolbarShared.EvidenceBasePath = $global:FabriqEvidenceBasePath
        }
    } catch { }
}
