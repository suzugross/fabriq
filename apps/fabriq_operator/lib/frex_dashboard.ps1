# ========================================
# Fabriq Operator - FrexProfile Dashboard
# ========================================
# Per-module state-aware execution dashboard. Lets the operator
# selectively re-run modules from a profile while preserving the
# Linear pipeline (Resolve-ProfileModules -> Invoke-BatchExecution
# -> Complete-ProfileExecution).
#
# This file only renders the form and returns the operator's intent.
# Actual execution dispatch happens in main.ps1 (P6 wiring).
# ========================================

function Show-FrexDashboard {
    param(
        [Parameter(Mandatory)][string]$ProfilePath,
        [Parameter(Mandatory)][string]$ProfileName,
        [Parameter(Mandatory)][array]$AllModules,
        # Optional: most-recent batch results from $script:LastBatchResults.
        # When present, overlays history.csv for the rows it covers (more
        # authoritative than history because the history CSV may not have
        # been re-imported since the last write).
        [array]$LastBatchResults = @(),
        # Optional: timestamp of last [Complete] press (display only)
        [string]$LastFinalizedAt = "",
        [string]$HostName = $env:SELECTED_NEW_PCNAME,
        [string]$WorkerName = $env:FABRIQ_WORKER_NAME
    )

    # ----------------------------------------
    # Load profile rows (incl. Enabled=0 ones)
    # ----------------------------------------
    $resolved = Resolve-ProfileModules -ProfileCsvPath $ProfilePath -AllModules $AllModules -IncludeDisabled
    $rows = @($resolved.ValidModules | Sort-Object Order)

    # ----------------------------------------
    # Build state map: Order -> { Status, Verified, Message }
    #
    # Precedence (later wins):
    #   1. Seed all rows as Pending
    #   2. Overlay history.csv (latest entry per MenuName, current SessionID)
    #   3. Overlay $LastBatchResults (most recent in-process record)
    # ----------------------------------------
    $stateMap = @{}
    foreach ($r in $rows) {
        $stateMap[[int]$r.Order] = @{ Status = 'Pending'; Verified = $null; Message = '' }
    }

    if (-not [string]::IsNullOrEmpty($env:SELECTED_KANRI_NO)) {
        $history = @(Import-ExecutionHistory -FilterKanriNo $env:SELECTED_KANRI_NO)
        # Filter to current session only; Import-ExecutionHistory returns descending
        # by Timestamp, so the first occurrence per key is the latest.
        $sessionHist = @($history | Where-Object { $_.SessionID -eq $script:SessionID })

        # Build two lookup tables:
        #   byOrder: keyed on int Order (preferred, per-row precision)
        #   byMenu : keyed on ModuleName (fallback for legacy entries
        #            without Order, and for non-Profile entries)
        $byOrder = @{}
        $byMenu  = @{}
        foreach ($h in $sessionHist) {
            $hasOrder = $h.PSObject.Properties.Name -contains 'Order' -and -not [string]::IsNullOrWhiteSpace($h.Order)
            if ($hasOrder) {
                try {
                    $hOrder = [int]$h.Order
                    if ($hOrder -gt 0 -and -not $byOrder.ContainsKey($hOrder)) {
                        $byOrder[$hOrder] = $h
                    }
                } catch { }
            }
            if (-not $byMenu.ContainsKey($h.ModuleName)) {
                $byMenu[$h.ModuleName] = $h
            }
        }

        foreach ($r in $rows) {
            $found = $null
            $rOrder = [int]$r.Order
            if ($byOrder.ContainsKey($rOrder)) {
                $found = $byOrder[$rOrder]
            }
            elseif ($byMenu.ContainsKey($r.MenuName)) {
                $found = $byMenu[$r.MenuName]
            }
            if ($null -ne $found) {
                $verified = if ($found.Verified -eq 'True') { $true }
                            elseif ($found.Verified -eq 'False') { $false }
                            else { $null }
                $stateMap[$rOrder] = @{
                    Status   = $found.Status
                    Verified = $verified
                    Message  = $found.Message
                }
            }
        }
    }

    foreach ($lbr in $LastBatchResults) {
        $ord = [int]$lbr.Order
        if ($stateMap.ContainsKey($ord)) {
            $stateMap[$ord] = @{
                Status   = $lbr.Status
                Verified = $lbr.Verified
                Message  = $lbr.Message
            }
        }
    }

    # ----------------------------------------
    # Result hashtable returned to caller
    # ----------------------------------------
    $result = @{
        Action           = "Close"
        SelectedOrders   = @()
        TargetOrder      = -1
        AutoPilot        = $false
        AutoPilotWaitSec = 3
        ResetTargetOrder = -1
    }

    # Internal dashboard state (separate from $result return value).
    # BulkUpdating: set $true while the AutoPilot toggle is bulk-setting
    # row checkboxes, so the per-cell CellValueChanged handler can
    # short-circuit its updateCounters call (which would otherwise fire
    # N times for N rows during a bulk operation). Hashtable is used so
    # event handler scriptblocks can mutate the value via reference.
    $frexState = @{ BulkUpdating = $false }

    # ----------------------------------------
    # Form scaffold
    # ----------------------------------------
    $form = New-Object System.Windows.Forms.Form
    Set-FormStyle -Form $form -Title "fabriq - FrexProfile: $ProfileName" -Width 900 -Height 660

    # Header bar
    $headerPanel = New-StyledPanel -X 0 -Y 0 -Width 900 -Height 50 -BgColor $script:bgPanel
    $form.Controls.Add($headerPanel)

    $titleLbl = New-StyledLabel -Text "FrexProfile: $ProfileName" -X 16 -Y 6 -Width 600 -Height 22 -Font $script:fontLarge -FgColor $script:fgWhite
    $headerPanel.Controls.Add($titleLbl)

    $finalizedTxt = if ([string]::IsNullOrEmpty($LastFinalizedAt)) {
        "Last finalized: -"
    } else {
        "Last finalized: $LastFinalizedAt"
    }
    $finalizedLbl = New-StyledLabel -Text $finalizedTxt -X 16 -Y 28 -Width 400 -Height 18 -FgColor ([System.Drawing.Color]::FromArgb(200, 200, 200))
    $headerPanel.Controls.Add($finalizedLbl)

    $hostLbl = New-StyledLabel -Text "$HostName  |  W: $WorkerName" -X 500 -Y 14 -Width 380 -Height 22 -FgColor ([System.Drawing.Color]::FromArgb(200, 200, 200))
    $hostLbl.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
    $headerPanel.Controls.Add($hostLbl)

    # ----------------------------------------
    # Module grid
    # ----------------------------------------
    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Location = New-Object System.Drawing.Point(10, 60)
    $grid.Size = New-Object System.Drawing.Size(870, 470)
    Set-GridStyle -Grid $grid

    # Set-GridStyle sets $grid.ReadOnly = $true, which in WinForms overrides
    # per-column ReadOnly settings — making cells uneditable regardless of
    # column-level configuration. Reset the grid-level ReadOnly so the
    # checkbox column can be toggled, and lock down the other columns
    # individually so only the checkbox is editable.
    $grid.ReadOnly = $false

    $colChk = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $colChk.Name = "Checked"
    $colChk.HeaderText = "*"
    $colChk.Width = 36
    $colChk.ReadOnly = $false
    $grid.Columns.Add($colChk) | Out-Null

    $colOrder = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colOrder.Name = "Order"
    $colOrder.HeaderText = "#"
    $colOrder.Width = 50
    $colOrder.DefaultCellStyle.Alignment = "MiddleRight"
    $colOrder.ReadOnly = $true
    $grid.Columns.Add($colOrder) | Out-Null

    $colMenu = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colMenu.Name = "MenuName"
    $colMenu.HeaderText = "Module"
    $colMenu.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
    $colMenu.ReadOnly = $true
    $grid.Columns.Add($colMenu) | Out-Null

    $colStatus = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colStatus.Name = "Status"
    $colStatus.HeaderText = "Status"
    $colStatus.Width = 90
    $colStatus.DefaultCellStyle.Alignment = "MiddleCenter"
    $colStatus.ReadOnly = $true
    $grid.Columns.Add($colStatus) | Out-Null

    $colVerified = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colVerified.Name = "Verified"
    $colVerified.HeaderText = "Verified"
    $colVerified.Width = 70
    $colVerified.DefaultCellStyle.Alignment = "MiddleCenter"
    $colVerified.ReadOnly = $true
    $grid.Columns.Add($colVerified) | Out-Null

    $form.Controls.Add($grid)

    # ----------------------------------------
    # Populate grid rows
    # ----------------------------------------
    foreach ($r in $rows) {
        $ord = [int]$r.Order
        $st  = $stateMap[$ord]
        $verifiedDisplay = if ($null -eq $st.Verified) { "-" }
                           elseif ($st.Verified)      { "PASS" }
                           else                       { "FAIL" }
        # Initial checkbox state is always unchecked (Frex philosophy:
        # default mode = pick from blank). The CSV's Enabled value is
        # captured into row.Tag.IsCheckedDefault so the AutoPilot toggle
        # can bulk-check Enabled=1 rows on demand without losing the
        # CSV-author's intent.
        $rowIndex = $grid.Rows.Add(
            $false,
            $ord,
            $r.MenuName,
            $st.Status,
            $verifiedDisplay
        )
        $row = $grid.Rows[$rowIndex]
        $row.Tag = @{
            Order            = $ord
            MenuName         = $r.MenuName
            RelativePath     = $r.RelativePath
            IsRestart        = [bool]$r._IsRestart
            IsReexplorer     = [bool]$r._IsReexplorer
            IsCheckedDefault = [bool]$r._IsCheckedDefault
            Message          = $st.Message
        }

        # Tooltip on Status cell shows the full Message
        if (-not [string]::IsNullOrWhiteSpace($st.Message)) {
            $row.Cells['Status'].ToolTipText = $st.Message
        }

        # Visual hint for special markers
        if ($r._IsRestart -or $r._IsReexplorer) {
            $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(232, 224, 200)
        }
    }

    # ----------------------------------------
    # Status color coding (CellFormatting)
    # ----------------------------------------
    $grid.Add_CellFormatting({
        param($s, $e)
        if ($e.RowIndex -lt 0) { return }
        $colName = $grid.Columns[$e.ColumnIndex].Name
        if ($colName -eq 'Status') {
            switch ($e.Value) {
                'Success'   { $e.CellStyle.ForeColor = [System.Drawing.Color]::FromArgb(34, 139, 34) }
                'Partial'   { $e.CellStyle.ForeColor = [System.Drawing.Color]::FromArgb(184, 134, 11) }
                'Error'     { $e.CellStyle.ForeColor = [System.Drawing.Color]::FromArgb(178, 34, 34) }
                'Skipped'   { $e.CellStyle.ForeColor = [System.Drawing.Color]::FromArgb(105, 105, 105) }
                'Cancelled' { $e.CellStyle.ForeColor = [System.Drawing.Color]::FromArgb(105, 105, 105) }
                'Pending'   { $e.CellStyle.ForeColor = [System.Drawing.Color]::FromArgb(140, 140, 140) }
            }
        }
        elseif ($colName -eq 'Verified') {
            switch ($e.Value) {
                'PASS' { $e.CellStyle.ForeColor = [System.Drawing.Color]::FromArgb(34, 139, 34) }
                'FAIL' { $e.CellStyle.ForeColor = [System.Drawing.Color]::FromArgb(178, 34, 34) }
            }
        }
    })

    # ----------------------------------------
    # Right-click context menu: Mark as Pending
    # ----------------------------------------
    $ctxMenu = New-Object System.Windows.Forms.ContextMenuStrip
    $ctxReset = $ctxMenu.Items.Add("Mark as Pending (reset state)")
    $grid.ContextMenuStrip = $ctxMenu

    $ctxReset.Add_Click({
        if ($grid.SelectedRows.Count -eq 0) { return }
        $row = $grid.SelectedRows[0]
        $tag = $row.Tag
        $result.Action = "ResetState"
        $result.ResetTargetOrder = [int]$tag.Order
        $form.Close()
    })

    # ----------------------------------------
    # Footer panel
    # ----------------------------------------
    $footerPanel = New-StyledPanel -X 10 -Y 540 -Width 870 -Height 80 -BgColor $script:bgTabPage
    $form.Controls.Add($footerPanel)

    # AutoPilot checkbox + WaitSec input (left)
    $chkAutoPilot = New-StyledCheckBox -Text "AutoPilot" -X 10 -Y 8 -Width 100 -Checked $false
    $footerPanel.Controls.Add($chkAutoPilot)

    $waitLbl = New-StyledLabel -Text "WaitSec:" -X 115 -Y 10 -Width 60 -Height 22
    $footerPanel.Controls.Add($waitLbl)

    $waitInput = New-Object System.Windows.Forms.NumericUpDown
    $waitInput.Location = New-Object System.Drawing.Point(175, 8)
    $waitInput.Size = New-Object System.Drawing.Size(60, 24)
    $waitInput.Minimum = 0
    $waitInput.Maximum = 60
    $waitInput.Value = 3
    $waitInput.BackColor = $script:bgInput
    $footerPanel.Controls.Add($waitInput)

    # Buttons (right side, two rows)
    # Row 1: Run Selected / Run This / Restart Now
    $btnRunSelected = New-StyledButton -Text "Run Selected (0)" -X 320 -Y 6 -Width 160 -Height 30 -BgColor $script:bgAccent
    $btnRunSelected.Font = $script:fontBold
    $footerPanel.Controls.Add($btnRunSelected)

    $btnRunThis = New-StyledButton -Text "Run This: -" -X 490 -Y 6 -Width 170 -Height 30
    $footerPanel.Controls.Add($btnRunThis)

    $btnRestartNow = New-StyledButton -Text "Restart Now" -X 670 -Y 6 -Width 130 -Height 30
    $footerPanel.Controls.Add($btnRestartNow)

    # Row 2: Complete / Back
    $btnComplete = New-StyledButton -Text "Complete" -X 320 -Y 42 -Width 340 -Height 30 -BgColor $script:bgAdd
    $btnComplete.Font = $script:fontBold
    $btnComplete.ForeColor = $script:fgWhite
    $footerPanel.Controls.Add($btnComplete)

    $btnBack = New-StyledButton -Text "Back" -X 670 -Y 42 -Width 130 -Height 30
    $footerPanel.Controls.Add($btnBack)

    # ----------------------------------------
    # Counter helpers (closed over $grid / $btnRunSelected / $btnComplete)
    # ----------------------------------------
    $updateCounters = {
        $checkedCount = 0
        $issueCount = 0
        foreach ($row in $grid.Rows) {
            $isChecked = [bool]$row.Cells['Checked'].Value
            if (-not $isChecked) { continue }
            $checkedCount++
            $st = "$($row.Cells['Status'].Value)"
            if ($st -in @('Error', 'Partial', 'Pending')) {
                $issueCount++
            }
        }
        $btnRunSelected.Text = "Run Selected ($checkedCount)"
        if ($issueCount -gt 0) {
            $btnComplete.Text = "Complete with $issueCount issue$(if ($issueCount -eq 1) { '' } else { 's' })"
            $btnComplete.BackColor = $script:stripeYellow
            $btnComplete.ForeColor = $script:fgText
        } else {
            $btnComplete.Text = "Complete"
            $btnComplete.BackColor = $script:bgAdd
            $btnComplete.ForeColor = $script:fgWhite
        }
    }

    $updateRunThisLabel = {
        if ($grid.SelectedRows.Count -gt 0) {
            $tag = $grid.SelectedRows[0].Tag
            $btnRunThis.Text = "Run This: $($tag.Order)"
            $btnRunThis.Enabled = $true
        } else {
            $btnRunThis.Text = "Run This: -"
            $btnRunThis.Enabled = $false
        }
    }

    # ----------------------------------------
    # Events
    # ----------------------------------------
    # Commit checkbox edit immediately so CellValueChanged fires on click
    $grid.Add_CurrentCellDirtyStateChanged({
        if ($grid.IsCurrentCellDirty) {
            $grid.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
        }
    })

    $grid.Add_CellValueChanged({
        param($s, $e)
        # During AutoPilot bulk update, suppress per-row counter recompute
        # to avoid N redundant updateCounters firings. The bulk handler
        # calls updateCounters once at its end.
        if ($frexState.BulkUpdating) { return }
        if ($e.RowIndex -ge 0 -and $grid.Columns[$e.ColumnIndex].Name -eq 'Checked') {
            & $updateCounters
        }
    })

    $grid.Add_SelectionChanged({
        & $updateRunThisLabel
    })

    # AutoPilot toggle: bulk-check Enabled=1 rows when turned ON.
    # Turning OFF preserves user's current checkbox state (intentional —
    # supports the workflow of "tick AutoPilot to bulk-select, untick to
    # switch off AutoPilot mode while keeping the bulk selection for a
    # manual-confirm batch run").
    # Only Enabled=1 rows are bulk-checked (CSV author's intent for
    # opt-in Enabled=0 rows is preserved; operators can manually check
    # them after the bulk operation).
    $chkAutoPilot.Add_CheckedChanged({
        if (-not $chkAutoPilot.Checked) {
            return
        }
        $frexState.BulkUpdating = $true
        try {
            foreach ($row in $grid.Rows) {
                $row.Cells['Checked'].Value = [bool]$row.Tag.IsCheckedDefault
            }
        }
        finally {
            $frexState.BulkUpdating = $false
        }
        & $updateCounters
    })

    $btnRunSelected.Add_Click({
        $checkedOrders = @()
        foreach ($row in $grid.Rows) {
            if ([bool]$row.Cells['Checked'].Value) {
                $checkedOrders += [int]$row.Tag.Order
            }
        }
        if ($checkedOrders.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show(
                "No modules are checked.",
                "FrexProfile",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
            return
        }
        $result.Action = "RunBatch"
        $result.SelectedOrders = $checkedOrders
        $result.AutoPilot = [bool]$chkAutoPilot.Checked
        $result.AutoPilotWaitSec = [int]$waitInput.Value
        $form.Close()
    })

    $btnRunThis.Add_Click({
        if ($grid.SelectedRows.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show(
                "Select a row first.",
                "FrexProfile",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
            return
        }
        $tag = $grid.SelectedRows[0].Tag
        $result.Action = "RunSingle"
        $result.TargetOrder = [int]$tag.Order
        $form.Close()
    })

    $btnComplete.Add_Click({
        # Build SelectedOrders snapshot for issue accounting / artifact scope
        $checkedOrders = @()
        foreach ($row in $grid.Rows) {
            if ([bool]$row.Cells['Checked'].Value) {
                $checkedOrders += [int]$row.Tag.Order
            }
        }

        # Soft warning when issues remain (Error / Partial / Pending among checked)
        $issueRows = @()
        foreach ($row in $grid.Rows) {
            if (-not [bool]$row.Cells['Checked'].Value) { continue }
            $st = "$($row.Cells['Status'].Value)"
            if ($st -in @('Error', 'Partial', 'Pending')) {
                $issueRows += "  Order $($row.Tag.Order): $($row.Tag.MenuName) [$st]"
            }
        }

        if ($issueRows.Count -gt 0) {
            $msg = "The following checked rows have unresolved status:`n`n" + ($issueRows -join "`n") + "`n`nMark this profile complete and export evidence?"
            $dlg = [System.Windows.Forms.MessageBox]::Show(
                $msg,
                "FrexProfile - Complete with issues",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            if ($dlg -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        }

        $result.Action = "Complete"
        $result.SelectedOrders = $checkedOrders
        $form.Close()
    })

    $btnRestartNow.Add_Click({
        $dlg = [System.Windows.Forms.MessageBox]::Show(
            "Restart the computer now?`n`nFrexProfile state (checkboxes / module status) will be saved and the dashboard will be restored after reboot.",
            "FrexProfile - Restart Now",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        if ($dlg -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        # Carry forward the current checkbox subset so post-reboot dashboard
        # opens with the same selection (orchestrated by main.ps1 in P6).
        $checkedOrders = @()
        foreach ($row in $grid.Rows) {
            if ([bool]$row.Cells['Checked'].Value) {
                $checkedOrders += [int]$row.Tag.Order
            }
        }
        $result.Action = "RestartNow"
        $result.SelectedOrders = $checkedOrders
        $form.Close()
    })

    $btnBack.Add_Click({
        $result.Action = "Close"
        $form.Close()
    })

    # Form close via X button = Close action
    $form.Add_FormClosing({
        if ($result.Action -eq "" -or $null -eq $result.Action) {
            $result.Action = "Close"
        }
    })

    # Initial counter sync
    & $updateCounters
    & $updateRunThisLabel

    # Show dialog
    $form.Add_Shown({ $form.Activate() })
    [void]$form.ShowDialog()
    $form.Dispose()

    return $result
}
