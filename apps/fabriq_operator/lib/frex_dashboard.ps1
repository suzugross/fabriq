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
        # When $true, the header shows a red "PENDING FINALIZE" badge in
        # place of the "Last finalized" line, and [Back] / X-button close
        # asks for confirmation. Caller (Invoke-FrexProfileLoop) tracks
        # this flag across dashboard reopens so it persists from batch
        # completion to operator's [Complete] press.
        [bool]$PendingFinalize = $false,
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
                # MenuName fallback is intentionally STRICT to handle the
                # case of two Profile rows sharing the same MenuName
                # (e.g., same module + same segment at Orders 80 and 110).
                # Without this strictness, an executed sibling row's
                # entry would leak into an unexecuted row's state.
                # Accept the candidate only when:
                #   (a) candidate has Order=0 (legacy CSV without Order
                #       column, or non-Profile entries like [RESTART NOW]),
                #       which cannot have a sibling-row identity, OR
                #   (b) candidate's Order matches the row's Order
                #       (defensive — should normally be caught by byOrder).
                $candidate = $byMenu[$r.MenuName]
                $candOrder = 0
                $candHasOrder = $candidate.PSObject.Properties.Name -contains 'Order' -and -not [string]::IsNullOrWhiteSpace($candidate.Order)
                if ($candHasOrder) {
                    try { $candOrder = [int]$candidate.Order } catch { $candOrder = 0 }
                }
                if ($candOrder -eq 0 -or $candOrder -eq $rOrder) {
                    $found = $candidate
                }
                # Otherwise: candidate belongs to a sibling Profile row
                # (different non-zero Order, same MenuName) — leave this
                # row Pending so its state is genuinely "not yet run".
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
    # Action values:
    #   Close / RunSingle / RunBatch / RunGroup / Complete /
    #   RestartNow / ResetState
    $result = @{
        Action           = "Close"
        SelectedOrders   = @()
        TargetOrder      = -1
        TargetGroup      = ""
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
    # Compute unique groups (preserves CSV appearance order)
    # ----------------------------------------
    # Groups bar is rendered only when at least one row has a non-empty
    # _Group value, so old profiles without the Group column have zero
    # layout impact (form / grid Y stays at the original position).
    $uniqueGroups = @()
    $seenGroups   = @{}
    foreach ($r in $rows) {
        $g = if ($r.PSObject.Properties.Name -contains '_Group') { "$($r._Group)".Trim() } else { "" }
        if (-not [string]::IsNullOrWhiteSpace($g) -and -not $seenGroups.ContainsKey($g)) {
            $seenGroups[$g] = $true
            $uniqueGroups += $g
        }
    }
    $hasGroups = ($uniqueGroups.Count -gt 0)

    # Layout offsets: with-Groups variants are 40px taller to fit the
    # Groups bar between header and grid.
    $groupsBarY = 52
    $groupsBarH = 40
    $gridY      = if ($hasGroups) { $groupsBarY + $groupsBarH + 4 } else { 60 }
    $gridH      = 470
    $footerY    = $gridY + $gridH + 10
    $formH      = $footerY + 80 + 20

    # ----------------------------------------
    # Form scaffold
    # ----------------------------------------
    $form = New-Object System.Windows.Forms.Form
    Set-FormStyle -Form $form -Title "fabriq - FrexProfile: $ProfileName" -Width 900 -Height $formH

    # Header bar
    $headerPanel = New-StyledPanel -X 0 -Y 0 -Width 900 -Height 50 -BgColor $script:bgPanel
    $form.Controls.Add($headerPanel)

    $titleLbl = New-StyledLabel -Text "FrexProfile: $ProfileName" -X 16 -Y 6 -Width 600 -Height 22 -Font $script:fontLarge -FgColor $script:fgWhite
    $headerPanel.Controls.Add($titleLbl)

    # Header sub-line: Pending Finalize badge (warning) takes precedence
    # over the neutral "Last finalized" timestamp. They communicate
    # mutually exclusive states — pending = batch ran but no Complete
    # yet, finalized = Complete pressed.
    if ($PendingFinalize) {
        $finalizedLbl = New-Object System.Windows.Forms.Label
        $finalizedLbl.Location = New-Object System.Drawing.Point(16, 28)
        $finalizedLbl.Size = New-Object System.Drawing.Size(180, 18)
        $finalizedLbl.Text = "  PENDING FINALIZE  "
        $finalizedLbl.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $finalizedLbl.Font = $script:fontBold
        $finalizedLbl.BackColor = $script:bgDelete
        $finalizedLbl.ForeColor = $script:fgWhite
    }
    else {
        $finalizedTxt = if ([string]::IsNullOrEmpty($LastFinalizedAt)) {
            "Last finalized: -"
        } else {
            "Last finalized: $LastFinalizedAt"
        }
        $finalizedLbl = New-StyledLabel -Text $finalizedTxt -X 16 -Y 28 -Width 400 -Height 18 -FgColor ([System.Drawing.Color]::FromArgb(200, 200, 200))
    }
    $headerPanel.Controls.Add($finalizedLbl)

    $hostLbl = New-StyledLabel -Text "$HostName  |  W: $WorkerName" -X 500 -Y 14 -Width 380 -Height 22 -FgColor ([System.Drawing.Color]::FromArgb(200, 200, 200))
    $hostLbl.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
    $headerPanel.Controls.Add($hostLbl)

    # ----------------------------------------
    # Groups bar (only when CSV has Group column with values)
    # ----------------------------------------
    # One [Run: <Group>] button per unique Group value, dispatched as
    # the new "RunGroup" action which translates into the same Invoke-
    # BatchExecution call as RunBatch — just sourced from the group's
    # Order list rather than the operator's checkbox state. Layout:
    # single row, ~135px per button. Profiles with 5-6 groups fit in
    # the 870px panel without clipping; more groups will extend past
    # the right edge (not auto-wrapped — kept simple per YAGNI).
    if ($hasGroups) {
        $groupsPanel = New-StyledPanel -X 10 -Y $groupsBarY -Width 870 -Height $groupsBarH -BgColor $script:bgTabPage
        $form.Controls.Add($groupsPanel)

        $groupsLbl = New-StyledLabel -Text "Groups:" -X 8 -Y 11 -Width 60 -Height 20 -Font $script:fontBold -FgColor $script:fgText
        $groupsPanel.Controls.Add($groupsLbl)

        $btnX = 72
        $btnW = 130
        foreach ($g in $uniqueGroups) {
            $btnGroup = New-StyledButton -Text "Run: $g" -X $btnX -Y 6 -Width $btnW -Height 28 -BgColor $script:bgAccent
            $btnGroup.Font = $script:fontBold
            $btnGroup.ForeColor = $script:fgWhite
            $btnGroup.Tag = $g

            $btnGroup.Add_Click({
                param($s, $e)
                $clickedGroup = $s.Tag
                $groupOrders = @()
                foreach ($row in $grid.Rows) {
                    if ("$($row.Tag.Group)" -eq "$clickedGroup") {
                        $groupOrders += [int]$row.Tag.Order
                    }
                }
                if ($groupOrders.Count -eq 0) {
                    [System.Windows.Forms.MessageBox]::Show(
                        "No modules in group '$clickedGroup'.",
                        "FrexProfile",
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Information
                    ) | Out-Null
                    return
                }
                $result.Action = "RunGroup"
                $result.TargetGroup = $clickedGroup
                $result.SelectedOrders = $groupOrders
                $result.AutoPilot = $true
                $result.AutoPilotWaitSec = [int]$waitInput.Value
                $form.Close()
            })
            $groupsPanel.Controls.Add($btnGroup)
            $btnX += ($btnW + 5)
        }
    }

    # ----------------------------------------
    # Module grid
    # ----------------------------------------
    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Location = New-Object System.Drawing.Point(10, $gridY)
    $grid.Size = New-Object System.Drawing.Size(870, $gridH)
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

    # Per-row Run button. Same dispatch as the legacy footer
    # [Run This: M] button (RunSingle action) — added at right end of
    # each row so single-execution becomes a 1-click action without
    # the "select row -> click footer" two-step. The footer button
    # was removed in 3.1.8 in favor of this per-row UI.
    $colRun = New-Object System.Windows.Forms.DataGridViewButtonColumn
    $colRun.Name = "RunBtn"
    $colRun.HeaderText = ""
    $colRun.Text = "Run"
    $colRun.UseColumnTextForButtonValue = $true
    $colRun.Width = 56
    $colRun.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $colRun.DefaultCellStyle.Alignment = "MiddleCenter"
    $grid.Columns.Add($colRun) | Out-Null

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
        $rowGroup = if ($r.PSObject.Properties.Name -contains '_Group') { "$($r._Group)".Trim() } else { "" }
        $row.Tag = @{
            Order            = $ord
            MenuName         = $r.MenuName
            RelativePath     = $r.RelativePath
            IsRestart        = [bool]$r._IsRestart
            IsReexplorer     = [bool]$r._IsReexplorer
            IsCheckedDefault = [bool]$r._IsCheckedDefault
            Group            = $rowGroup
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
    # Status / Verified cell badge styling (CellFormatting)
    # ----------------------------------------
    # 3.1.9: switched from text-color hints to background fill for
    # at-a-glance visibility. The full Status / Verified cell becomes
    # a colored badge with bold contrast text. SelectionBackColor /
    # SelectionForeColor are also set so the badge remains visible
    # when the row is highlighted by selection.
    $grid.Add_CellFormatting({
        param($s, $e)
        if ($e.RowIndex -lt 0) { return }
        $colName = $grid.Columns[$e.ColumnIndex].Name

        $bg = $null
        $fg = $null

        if ($colName -eq 'Status') {
            switch ($e.Value) {
                'Success'   { $bg = $script:bgAdd;                                 $fg = $script:fgWhite }
                'Partial'   { $bg = $script:stripeYellow;                          $fg = $script:fgText }
                'Error'     { $bg = $script:bgDelete;                              $fg = $script:fgWhite }
                'Skipped'   { $bg = [System.Drawing.Color]::FromArgb(130,130,130); $fg = $script:fgWhite }
                'Cancelled' { $bg = [System.Drawing.Color]::FromArgb(130,130,130); $fg = $script:fgWhite }
                'Pending'   { $bg = [System.Drawing.Color]::FromArgb(200,200,200); $fg = [System.Drawing.Color]::FromArgb(80,80,80) }
            }
        }
        elseif ($colName -eq 'Verified') {
            switch ($e.Value) {
                'PASS' { $bg = $script:bgAdd;    $fg = $script:fgWhite }
                'FAIL' { $bg = $script:bgDelete; $fg = $script:fgWhite }
            }
        }

        if ($null -ne $bg) {
            $e.CellStyle.BackColor          = $bg
            $e.CellStyle.SelectionBackColor = $bg
            $e.CellStyle.ForeColor          = $fg
            $e.CellStyle.SelectionForeColor = $fg
            $e.CellStyle.Font               = $script:fontBold
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
    $footerPanel = New-StyledPanel -X 10 -Y $footerY -Width 870 -Height 80 -BgColor $script:bgTabPage
    $form.Controls.Add($footerPanel)

    # Bulk-select buttons + WaitSec input (left).
    # AutoPilot checkbox was removed in 3.1.5 — execution is always
    # AutoPilot-style (unattended), and finalization is always manual
    # via [Complete]. [Select All] checks Enabled=1 rows (CSV author's
    # default set); [Clear All] unchecks everything.
    $btnSelectAll = New-StyledButton -Text "Select All" -X 10 -Y 8 -Width 90 -Height 24
    $footerPanel.Controls.Add($btnSelectAll)

    $btnClearAll = New-StyledButton -Text "Clear All" -X 105 -Y 8 -Width 90 -Height 24
    $footerPanel.Controls.Add($btnClearAll)

    $waitLbl = New-StyledLabel -Text "WaitSec:" -X 205 -Y 10 -Width 55 -Height 22
    $footerPanel.Controls.Add($waitLbl)

    $waitInput = New-Object System.Windows.Forms.NumericUpDown
    $waitInput.Location = New-Object System.Drawing.Point(260, 8)
    $waitInput.Size = New-Object System.Drawing.Size(55, 24)
    $waitInput.Minimum = 0
    $waitInput.Maximum = 60
    $waitInput.Value = 3
    $waitInput.BackColor = $script:bgInput
    $footerPanel.Controls.Add($waitInput)

    # Buttons (right side, two rows). 3.1.8: [Run This: M] removed in
    # favor of per-row [Run] buttons in the grid. [Run Selected] is
    # widened to fill the row 1 space and visually match the [Complete]
    # button in row 2.
    # Row 1: Run Selected (wide) / Restart Now
    $btnRunSelected = New-StyledButton -Text "Run Selected (0)" -X 320 -Y 6 -Width 340 -Height 30 -BgColor $script:bgAccent
    $btnRunSelected.Font = $script:fontBold
    $footerPanel.Controls.Add($btnRunSelected)

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
        # "hasExecuted" tracks whether at least one row has a non-Pending
        # status — i.e., something actually ran in this session. The
        # initial dashboard state (all rows Pending, none checked) would
        # otherwise produce $issueCount=0 and a misleading green
        # "Complete" button, allowing the operator to finalize an empty
        # checklist with no execution behind it.
        $hasExecuted = $false
        foreach ($row in $grid.Rows) {
            $isChecked = [bool]$row.Cells['Checked'].Value
            if ($isChecked) { $checkedCount++ }
            $st = "$($row.Cells['Status'].Value)"
            if ($st -ne 'Pending') { $hasExecuted = $true }
            # Errors and Partials are facts about an actual execution
            # outcome — count regardless of check state (operator cannot
            # mask a real failure by simply unchecking the row).
            # Pending is intent-based: count only when the row is checked
            # (operator has declared "I plan to run this but haven't yet").
            # Success / Skipped / Cancelled are not flagged.
            if ($st -eq 'Error' -or $st -eq 'Partial') {
                $issueCount++
            }
            elseif ($st -eq 'Pending' -and $isChecked) {
                $issueCount++
            }
        }
        $btnRunSelected.Text = "Run Selected ($checkedCount)"

        # Warning state: explicit issues OR nothing has been executed.
        # Two separate labels so the operator can distinguish "rows
        # have problems" from "you haven't run anything yet".
        if ($issueCount -gt 0) {
            $btnComplete.Text = "Complete with $issueCount issue$(if ($issueCount -eq 1) { '' } else { 's' })"
            $btnComplete.BackColor = $script:stripeYellow
            $btnComplete.ForeColor = $script:fgText
        }
        elseif (-not $hasExecuted -and $grid.Rows.Count -gt 0) {
            $btnComplete.Text = "Complete (nothing executed)"
            $btnComplete.BackColor = $script:stripeYellow
            $btnComplete.ForeColor = $script:fgText
        }
        else {
            $btnComplete.Text = "Complete"
            $btnComplete.BackColor = $script:bgAdd
            $btnComplete.ForeColor = $script:fgWhite
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

    # Per-row [Run] button click → dispatch RunSingle for that row.
    # Same action as the legacy footer [Run This: M] (replaced in 3.1.8).
    # Filter strictly on column name to avoid interfering with the
    # checkbox column's own click-and-toggle behavior.
    $grid.Add_CellContentClick({
        param($s, $e)
        if ($e.RowIndex -lt 0) { return }
        if ($grid.Columns[$e.ColumnIndex].Name -ne 'RunBtn') { return }
        $tag = $grid.Rows[$e.RowIndex].Tag
        if ($null -eq $tag) { return }
        $result.Action = "RunSingle"
        $result.TargetOrder = [int]$tag.Order
        $form.Close()
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

    # [Select All] bulk-checks Enabled=1 rows (CSV author's default
    # set). Operators can then individually uncheck rows they want to
    # exclude. Enabled=0 rows stay unchecked unless the operator
    # explicitly checks them.
    $btnSelectAll.Add_Click({
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

    # [Clear All] unchecks every row.
    $btnClearAll.Add_Click({
        $frexState.BulkUpdating = $true
        try {
            foreach ($row in $grid.Rows) {
                $row.Cells['Checked'].Value = $false
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
        # Execution mode is unconditionally AutoPilot since 3.1.5 —
        # the AutoPilot checkbox was removed and unattended batch
        # behavior is the default. Finalize is always manual via
        # the [Complete] button (handled in main.ps1).
        $result.AutoPilot = $true
        $result.AutoPilotWaitSec = [int]$waitInput.Value
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

        # Soft warning composition:
        #   (a) "nothing executed" — all rows still Pending in current
        #       session (initial state pressing Complete would generate
        #       an empty checklist, almost certainly a mistake).
        #   (b) "rows have unresolved status" — Error / Partial always,
        #       Pending only when checked (operator's declared intent).
        # Both conditions can fire simultaneously; the dialog combines them.
        $hasExecuted = $false
        $issueRows = @()
        foreach ($row in $grid.Rows) {
            $isChecked = [bool]$row.Cells['Checked'].Value
            $st = "$($row.Cells['Status'].Value)"
            if ($st -ne 'Pending') { $hasExecuted = $true }
            $isIssue = $false
            if ($st -eq 'Error' -or $st -eq 'Partial') {
                $isIssue = $true
            }
            elseif ($st -eq 'Pending' -and $isChecked) {
                $isIssue = $true
            }
            if ($isIssue) {
                $issueRows += "  Order $($row.Tag.Order): $($row.Tag.MenuName) [$st]"
            }
        }

        $warnSections = @()
        if (-not $hasExecuted -and $grid.Rows.Count -gt 0) {
            $warnSections += "No modules have been executed in this session yet."
        }
        if ($issueRows.Count -gt 0) {
            $warnSections += "The following rows have unresolved status:`n" + ($issueRows -join "`n")
        }

        if ($warnSections.Count -gt 0) {
            $msg = ($warnSections -join "`n`n") + "`n`nMark this profile complete and export evidence?"
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

    # Form close (Back button or X button). Warn the operator if a
    # batch ran but [Complete] has not been pressed yet — leaving in
    # this state means no HTML checklist is generated and no evidence
    # is uploaded to log_destinations. Active execution actions
    # (RunBatch / RunSingle / Complete / RestartNow / ResetState) set
    # $result.Action before calling form.Close, so they bypass the
    # warning. The warning fires only on the Close path.
    $form.Add_FormClosing({
        param($s, $e)
        $isCloseIntent = [string]::IsNullOrEmpty($result.Action) -or $result.Action -eq 'Close'
        if ($isCloseIntent -and $PendingFinalize) {
            $msg = "Batch results are pending finalize.`n`nPress [Complete] to generate the HTML checklist and upload evidence to log destinations.`n`nLeave the FrexProfile dashboard without finalizing anyway?"
            $dlg = [System.Windows.Forms.MessageBox]::Show(
                $msg,
                "FrexProfile - Pending Finalize",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            if ($dlg -ne [System.Windows.Forms.DialogResult]::Yes) {
                $e.Cancel = $true
                return
            }
        }
        if ([string]::IsNullOrEmpty($result.Action)) {
            $result.Action = "Close"
        }
    })

    # Initial counter sync
    & $updateCounters

    # Show dialog
    $form.Add_Shown({ $form.Activate() })
    [void]$form.ShowDialog()
    $form.Dispose()

    return $result
}
