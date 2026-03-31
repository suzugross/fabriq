# ========================================
# Fabriq Operator - Session Setup Form
# ========================================
# Displays a GUI form for worker selection,
# host selection, and master passphrase input.
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
    Set-FormStyle -Form $form -Title "fabriq operator - Session Setup" -Width 620 -Height 600

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

    # Add columns
    $colWID = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colWID.Name = "ID"
    $colWID.HeaderText = "ID"
    $colWID.Width = 80
    $workerGrid.Columns.Add($colWID) | Out-Null

    $colWName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colWName.Name = "Name"
    $colWName.HeaderText = "Name"
    $colWName.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
    $workerGrid.Columns.Add($colWName) | Out-Null

    # Populate workers
    if ($Workers -and $Workers.Count -gt 0) {
        foreach ($w in $Workers) {
            $workerGrid.Rows.Add($w.ID, $w.Name) | Out-Null
        }
        # Select first row by default
        if ($workerGrid.Rows.Count -gt 0) {
            $workerGrid.Rows[0].Selected = $true
        }
    }

    $form.Controls.Add($workerGrid)
    $yPos += 106

    # Manual worker input
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

    $hostGrid = New-Object System.Windows.Forms.DataGridView
    $hostGrid.Location = New-Object System.Drawing.Point(20, $yPos)
    $hostGrid.Size = New-Object System.Drawing.Size(560, 140)
    Set-GridStyle -Grid $hostGrid

    # Add columns
    $colHID = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colHID.Name = "AdminID"
    $colHID.HeaderText = "ID"
    $colHID.Width = 50
    $hostGrid.Columns.Add($colHID) | Out-Null

    $colOld = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colOld.Name = "OldPCName"
    $colOld.HeaderText = "OldPC"
    $colOld.Width = 140
    $hostGrid.Columns.Add($colOld) | Out-Null

    $colNew = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colNew.Name = "NewPCName"
    $colNew.HeaderText = "NewPC"
    $colNew.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
    $hostGrid.Columns.Add($colNew) | Out-Null

    $colIP = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colIP.Name = "EthernetIP"
    $colIP.HeaderText = "IP"
    $colIP.Width = 130
    $hostGrid.Columns.Add($colIP) | Out-Null

    $colPin = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colPin.Name = "Pin"
    $colPin.HeaderText = "Pin"
    $colPin.Width = 80
    $hostGrid.Columns.Add($colPin) | Out-Null

    # Populate hosts and auto-select
    $autoSelectedIdx = -1
    if ($HostList -and $HostList.Count -gt 0) {
        for ($i = 0; $i -lt $HostList.Count; $i++) {
            $h = $HostList[$i]
            $hostGrid.Rows.Add($h.AdminID, $h.OldPCName, $h.NewPCName, $h.EthernetIP, $h.Pin) | Out-Null
            if ($h.NewPCName -eq $CurrentPCName) {
                $autoSelectedIdx = $i
            }
        }
        if ($autoSelectedIdx -ge 0) {
            $hostGrid.ClearSelection()
            $hostGrid.Rows[$autoSelectedIdx].Selected = $true
            $hostGrid.FirstDisplayedScrollingRowIndex = $autoSelectedIdx
        }
        elseif ($hostGrid.Rows.Count -gt 0) {
            $hostGrid.Rows[0].Selected = $true
        }
    }

    $form.Controls.Add($hostGrid)
    $yPos += 146

    # Auto-detection hint
    if ($autoSelectedIdx -ge 0) {
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

    # When manual worker text is entered, deselect grid
    $manualWorkerBox.Add_TextChanged({
        if ($manualWorkerBox.Text.Length -gt 0) {
            $workerGrid.ClearSelection()
        }
    })

    # When worker grid is clicked, clear manual input
    $workerGrid.Add_CellClick({
        $manualWorkerBox.Text = ""
    })

    # Enter key triggers Start
    $ppBox.Add_KeyDown({
        if ($_.KeyCode -eq "Return") {
            $startButton.PerformClick()
        }
    })

    # Start Session button
    $startButton.Add_Click({
        # Determine worker
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

        # Determine host
        if ($hostGrid.SelectedRows.Count -eq 0) {
            $msgLabel.Text = "Please select a target host."
            return
        }
        $hostIdx = $hostGrid.SelectedRows[0].Index
        if ($hostIdx -ge 0 -and $hostIdx -lt $HostList.Count) {
            $result.SelectedHost = $HostList[$hostIdx]
        }
        else {
            $msgLabel.Text = "Invalid host selection."
            return
        }

        # Validate passphrase
        $pp = $ppBox.Text
        if ([string]::IsNullOrWhiteSpace($pp)) {
            $msgLabel.Text = "Master passphrase is required."
            return
        }

        if (-not [string]::IsNullOrWhiteSpace($VerifyTokenPath) -and (Test-Path $VerifyTokenPath)) {
            if (-not (Test-MasterPassphrase -Passphrase $pp -VerifyTokenPath $VerifyTokenPath)) {
                $msgLabel.Text = "Passphrase verification failed. Please try again."
                $ppBox.SelectAll()
                $ppBox.Focus()
                return
            }
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
