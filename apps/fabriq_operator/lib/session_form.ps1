# ========================================
# Fabriq Operator - Session Setup Form
# ========================================
# Displays a GUI form for worker selection,
# host selection (with live search + sort),
# and master passphrase input.
# Returns a hashtable with session data or $null if cancelled.
# ========================================

function Show-SessionSetupForm {
    param(
        [array]$Workers,
        [array]$HostList,
        [string]$VerifyTokenPath,
        [string]$CurrentPCName = $env:COMPUTERNAME
    )

    $result = @{
        WorkerID         = ""
        WorkerName       = ""
        SelectedHost     = $null
        MasterPassphrase = ""
        Cancelled        = $true
    }

    $form = New-Object System.Windows.Forms.Form
    Set-FormStyle -Form $form -Title "fabriq operator - Session Setup" -Width 620 -Height 630

    $yPos = 15

    # ========================================
    # Title
    # ========================================
    $titleLabel = New-StyledLabel -Text "fabriq operator" -X 20 -Y $yPos -Width 560 -Height 32 -Font $script:fontTitle -FgColor $script:fgHeader
    $form.Controls.Add($titleLabel)
    $yPos += 40

    # ========================================
    # Worker Selection Section
    # ========================================
    $workerLabel = New-StyledLabel -Text "Worker" -X 20 -Y $yPos -Width 560 -Height 20 -Font $script:fontBold -FgColor $script:fgHeader
    $form.Controls.Add($workerLabel)
    $yPos += 24

    $workerGrid = New-Object System.Windows.Forms.DataGridView
    $workerGrid.Location = New-Object System.Drawing.Point(20, $yPos)
    $workerGrid.Size = New-Object System.Drawing.Size(560, 100)
    Set-GridStyle -Grid $workerGrid

    # Add columns (SortMode = Automatic so header click sorts)
    $colWID = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colWID.Name = "ID"
    $colWID.HeaderText = "ID"
    $colWID.Width = 80
    $colWID.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic
    $workerGrid.Columns.Add($colWID) | Out-Null

    $colWName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colWName.Name = "Name"
    $colWName.HeaderText = "Name"
    $colWName.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
    $colWName.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic
    $workerGrid.Columns.Add($colWName) | Out-Null

    # Populate workers
    if ($Workers -and $Workers.Count -gt 0) {
        foreach ($w in $Workers) {
            $workerGrid.Rows.Add($w.ID, $w.Name) | Out-Null
        }
        if ($workerGrid.Rows.Count -gt 0) {
            $workerGrid.Rows[0].Selected = $true
        }
    }

    $form.Controls.Add($workerGrid)
    $yPos += 106

    # Manual worker input (exclusive with grid selection)
    $manualWorkerLabel = New-StyledLabel -Text "or enter manually:" -X 20 -Y $yPos -Width 130 -Height 22 -FgColor $script:fgDim
    $form.Controls.Add($manualWorkerLabel)

    $manualWorkerBox = New-Object System.Windows.Forms.TextBox
    $manualWorkerBox.Location = New-Object System.Drawing.Point(155, $yPos)
    $manualWorkerBox.Size = New-Object System.Drawing.Size(260, 22)
    Set-TextBoxStyle -TextBox $manualWorkerBox
    $form.Controls.Add($manualWorkerBox)
    $yPos += 32

    # ========================================
    # Host Selection Section
    # ========================================
    $hostLabel = New-StyledLabel -Text "Target Host" -X 20 -Y $yPos -Width 560 -Height 20 -Font $script:fontBold -FgColor $script:fgHeader
    $form.Controls.Add($hostLabel)
    $yPos += 24

    # ----------------------------------------
    # Search / filter row
    # Scope: AdminID + NewPCName only (case-insensitive substring).
    # Other columns (OldPCName / EthernetIP / Pin) are intentionally not
    # searched: OldPCName is often factory-default and not what the
    # operator remembers, IP ranges vary per site, and Pin may contain
    # sensitive tokens.
    # ----------------------------------------
    $searchLabel = New-StyledLabel -Text "Search:" -X 20 -Y $yPos -Width 60 -Height 22 -FgColor $script:fgDim
    $form.Controls.Add($searchLabel)

    $searchBox = New-Object System.Windows.Forms.TextBox
    $searchBox.Location = New-Object System.Drawing.Point(85, $yPos)
    $searchBox.Size = New-Object System.Drawing.Size(380, 22)
    Set-TextBoxStyle -TextBox $searchBox
    $form.Controls.Add($searchBox)

    $totalHosts = if ($HostList) { $HostList.Count } else { 0 }
    $countLabel = New-StyledLabel -Text "$totalHosts / $totalHosts" -X 470 -Y $yPos -Width 110 -Height 22 -FgColor $script:fgDim
    $countLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
    $form.Controls.Add($countLabel)
    $yPos += 30

    # ----------------------------------------
    # Host Grid
    # ----------------------------------------
    $hostGrid = New-Object System.Windows.Forms.DataGridView
    $hostGrid.Location = New-Object System.Drawing.Point(20, $yPos)
    $hostGrid.Size = New-Object System.Drawing.Size(560, 140)
    Set-GridStyle -Grid $hostGrid

    $colHID = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colHID.Name = "AdminID"
    $colHID.HeaderText = "ID"
    $colHID.Width = 50
    $colHID.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic
    $hostGrid.Columns.Add($colHID) | Out-Null

    $colOld = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colOld.Name = "OldPCName"
    $colOld.HeaderText = "OldPC"
    $colOld.Width = 140
    $colOld.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic
    $hostGrid.Columns.Add($colOld) | Out-Null

    $colNew = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colNew.Name = "NewPCName"
    $colNew.HeaderText = "NewPC"
    $colNew.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
    $colNew.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic
    $hostGrid.Columns.Add($colNew) | Out-Null

    $colIP = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colIP.Name = "EthernetIP"
    $colIP.HeaderText = "IP"
    $colIP.Width = 130
    $colIP.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic
    $hostGrid.Columns.Add($colIP) | Out-Null

    $colPin = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colPin.Name = "Pin"
    $colPin.HeaderText = "Pin"
    $colPin.Width = 80
    $colPin.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic
    $hostGrid.Columns.Add($colPin) | Out-Null

    # Track the auto-detected host by OBJECT REFERENCE so it survives the
    # filter redraws (DataGridView row indices shift when Rows.Clear()+rebuild).
    $autoSelectedHost = $null
    if ($HostList -and $HostList.Count -gt 0) {
        # Pass 1: an explicit hostlist row whose NewPCName equals this PC wins.
        foreach ($h in $HostList) {
            if ($h.NewPCName -eq $CurrentPCName) {
                $autoSelectedHost = $h
                break
            }
        }
        # Pass 2: fall back to a self-referencing row (__SELF__ is by definition
        # "this PC"), so the operator does not have to hunt for it manually.
        if ($null -eq $autoSelectedHost) {
            foreach ($h in $HostList) {
                if ($h.NewPCName -eq '__SELF__' -or $h.OldPCName -eq '__SELF__') {
                    $autoSelectedHost = $h
                    break
                }
            }
        }
    }

    # ----------------------------------------
    # Filter scriptblock: rebuilds the Host Grid with rows that match the
    # search text. Stores the source host object in Row.Tag so the Start
    # handler can resolve selection back to the source host regardless of
    # sort order or filter state.
    # ----------------------------------------
    $refreshHostGrid = {
        param([string]$searchText)

        $hostGrid.Rows.Clear()
        $needle = ''
        if ($searchText) { $needle = $searchText.Trim().ToLowerInvariant() }
        $matched = 0

        if ($HostList) {
            foreach ($h in $HostList) {
                $visible = $true
                if ($needle) {
                    $adminId = "$($h.AdminID)".ToLowerInvariant()
                    $newPc   = "$($h.NewPCName)".ToLowerInvariant()
                    $visible = ($adminId.Contains($needle)) -or ($newPc.Contains($needle))
                }
                if ($visible) {
                    # Display __SELF__ cells as a friendly placeholder so the
                    # operator is not shown an opaque token; the raw object
                    # stays in Row.Tag, so Start-time resolution is unaffected.
                    $selfLabel = '(this PC)'
                    $dispOld = if ($h.OldPCName  -eq '__SELF__') { $selfLabel } else { $h.OldPCName }
                    $dispNew = if ($h.NewPCName  -eq '__SELF__') { $selfLabel } else { $h.NewPCName }
                    $dispEth = if ($h.EthernetIP -eq '__SELF__') { $selfLabel } else { $h.EthernetIP }
                    $rowIdx = $hostGrid.Rows.Add($h.AdminID, $dispOld, $dispNew, $dispEth, $h.Pin)
                    $hostGrid.Rows[$rowIdx].Tag = $h
                    $matched++
                }
            }
        }

        $countLabel.Text = "$matched / $totalHosts"

        if ($hostGrid.Rows.Count -eq 0) { return }
        $hostGrid.ClearSelection()

        # Empty search + prior auto-detected host => reselect that host.
        if (-not $needle -and $null -ne $autoSelectedHost) {
            foreach ($r in $hostGrid.Rows) {
                if ([object]::ReferenceEquals($r.Tag, $autoSelectedHost)) {
                    $r.Selected = $true
                    try { $hostGrid.FirstDisplayedScrollingRowIndex = $r.Index } catch { }
                    return
                }
            }
        }

        # Otherwise (1 match or many), select the first visible row.
        $hostGrid.Rows[0].Selected = $true
        try { $hostGrid.FirstDisplayedScrollingRowIndex = 0 } catch { }
    }

    # Initial populate
    & $refreshHostGrid ''

    $form.Controls.Add($hostGrid)
    $yPos += 146

    # Auto-detection hint (shown only when a match was found)
    if ($null -ne $autoSelectedHost) {
        $autoLabel = New-StyledLabel -Text "* Auto-detected: $CurrentPCName (matches current PC)" -X 20 -Y $yPos -Width 560 -Height 18 -FgColor ([System.Drawing.Color]::FromArgb(46, 125, 50))
        $form.Controls.Add($autoLabel)
    }
    $yPos += 24

    # ========================================
    # Master Passphrase Section
    # ========================================
    $ppLabel = New-StyledLabel -Text "Master Passphrase" -X 20 -Y $yPos -Width 560 -Height 20 -Font $script:fontBold -FgColor $script:fgHeader
    $form.Controls.Add($ppLabel)
    $yPos += 24

    $ppBox = New-Object System.Windows.Forms.TextBox
    $ppBox.Location = New-Object System.Drawing.Point(20, $yPos)
    $ppBox.Size = New-Object System.Drawing.Size(400, 24)
    $ppBox.UseSystemPasswordChar = $true
    Set-TextBoxStyle -TextBox $ppBox
    $form.Controls.Add($ppBox)
    $yPos += 34

    # Validation message label
    $msgLabel = New-StyledLabel -Text "" -X 20 -Y $yPos -Width 560 -Height 20 -FgColor ([System.Drawing.Color]::FromArgb(198, 40, 40))
    $form.Controls.Add($msgLabel)
    $yPos += 28

    # ========================================
    # Buttons
    # ========================================
    $startButton = New-StyledButton -Text "Start Session" -X 370 -Y $yPos -Width 140 -Height 34 -BgColor $script:bgAccent
    $startButton.Font = $script:fontBold
    $form.Controls.Add($startButton)

    $cancelButton = New-StyledButton -Text "Quit" -X 240 -Y $yPos -Width 120 -Height 34
    $form.Controls.Add($cancelButton)

    # ========================================
    # Event Handlers
    # ========================================

    # Manual vs grid-selected worker: exclusive inputs
    $manualWorkerBox.Add_TextChanged({
        if ($manualWorkerBox.Text.Length -gt 0) {
            $workerGrid.ClearSelection()
        }
    })

    $workerGrid.Add_CellClick({
        $manualWorkerBox.Text = ""
    })

    # Live host filter on every keystroke in the search box
    $searchBox.Add_TextChanged({
        $msgLabel.Text = ""
        & $refreshHostGrid $searchBox.Text
    })

    # Escape clears the search; Enter advances focus to passphrase
    $searchBox.Add_KeyDown({
        if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Escape) {
            $searchBox.Text = ""
            $_.Handled = $true
            $_.SuppressKeyPress = $true
        }
        elseif ($_.KeyCode -eq [System.Windows.Forms.Keys]::Return) {
            $ppBox.Focus()
            $_.Handled = $true
            $_.SuppressKeyPress = $true
        }
    })

    # Enter in passphrase box = Start Session
    $ppBox.Add_KeyDown({
        if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Return) {
            $startButton.PerformClick()
        }
    })

    # Start Session button
    $startButton.Add_Click({
        # Resolve worker
        if ($manualWorkerBox.Text.Length -gt 0) {
            $result.WorkerID = "MANUAL"
            $result.WorkerName = $manualWorkerBox.Text.Trim()
        }
        elseif ($workerGrid.SelectedRows.Count -gt 0) {
            $row = $workerGrid.SelectedRows[0]
            $result.WorkerID = $row.Cells["ID"].Value
            $result.WorkerName = $row.Cells["Name"].Value
        }
        else {
            $msgLabel.Text = "Please select a worker or enter a name manually."
            return
        }

        if ([string]::IsNullOrWhiteSpace($result.WorkerName)) {
            $msgLabel.Text = "Worker name cannot be empty."
            return
        }

        # Resolve host via Row.Tag (stable across sort and filter states)
        if ($hostGrid.SelectedRows.Count -eq 0) {
            $msgLabel.Text = "Please select a target host."
            return
        }
        $selectedHostObj = $hostGrid.SelectedRows[0].Tag
        if ($null -eq $selectedHostObj) {
            $msgLabel.Text = "Invalid host selection."
            return
        }
        $result.SelectedHost = $selectedHostObj

        # Validate passphrase
        $pp = $ppBox.Text
        if ([string]::IsNullOrWhiteSpace($pp)) {
            $msgLabel.Text = "Master passphrase is required."
            return
        }

        # Fail-closed: a missing verification token must BLOCK the session,
        # not skip verification — an unverified (possibly mistyped)
        # passphrase silently breaks every ENC: decryption downstream.
        if ([string]::IsNullOrWhiteSpace($VerifyTokenPath) -or -not (Test-Path $VerifyTokenPath)) {
            $msgLabel.Text = "Verification token not found - cannot verify passphrase. Run Fabriq Studio to generate it."
            return
        }
        if (-not (Test-MasterPassphrase -Passphrase $pp -VerifyTokenPath $VerifyTokenPath)) {
            $msgLabel.Text = "Passphrase verification failed. Please try again."
            $ppBox.SelectAll()
            $ppBox.Focus()
            return
        }

        $result.MasterPassphrase = $pp
        $result.Cancelled = $false
        $form.Close()
    })

    # Cancel button
    $cancelButton.Add_Click({
        $result.Cancelled = $true
        $form.Close()
    })

    # Show dialog
    $form.Add_Shown({ $form.Activate() })
    [void]$form.ShowDialog()
    $form.Dispose()

    return $result
}
