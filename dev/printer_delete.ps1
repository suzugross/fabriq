# ========================================
# Fabriq Printer Manager
# ========================================
# GUI tool to list and remove printers.
# Based on Fabriq UI standards.
# ========================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

# ========================================
# Color Scheme (Fabriq Standard)
# ========================================
$bgDark       = [System.Drawing.Color]::FromArgb(30, 30, 30)
$bgPanel      = [System.Drawing.Color]::FromArgb(40, 40, 40)
$bgGrid       = [System.Drawing.Color]::FromArgb(35, 35, 35)
$bgCell       = [System.Drawing.Color]::FromArgb(45, 45, 45)
$bgHeader     = [System.Drawing.Color]::FromArgb(55, 55, 55)
$bgButton     = [System.Drawing.Color]::FromArgb(60, 60, 60)
$bgButtonHov  = [System.Drawing.Color]::FromArgb(80, 80, 80)
$bgDelete     = [System.Drawing.Color]::FromArgb(180, 40, 40)
$bgRefresh    = [System.Drawing.Color]::FromArgb(0, 120, 215)
$fgText       = [System.Drawing.Color]::FromArgb(220, 220, 220)
$fgDim        = [System.Drawing.Color]::FromArgb(150, 150, 150)
$fgHeader     = [System.Drawing.Color]::FromArgb(100, 180, 255)
$gridLine     = [System.Drawing.Color]::FromArgb(60, 60, 60)

# ========================================
# Helper Functions
# ========================================

function Get-InstalledPrinters {
    try {
        return @(Get-Printer -ErrorAction SilentlyContinue | Select-Object Name, DriverName, PortName, DriverType)
    }
    catch {
        return @()
    }
}

function Refresh-Grid {
    $script:dgv.Rows.Clear()
    $printers = Get-InstalledPrinters

    foreach ($p in $printers) {
        $idx = $script:dgv.Rows.Add()
        $row = $script:dgv.Rows[$idx]
        
        $row.Cells["Check"].Value = $false
        $row.Cells["Name"].Value = $p.Name
        $row.Cells["Driver"].Value = $p.DriverName
        $row.Cells["Port"].Value = $p.PortName
    }
    
    $script:statusLabel.Text = "Printers loaded: $($printers.Count)"
}

function Remove-SelectedPrinters {
    $rowsToDelete = @()
    foreach ($row in $script:dgv.Rows) {
        if ($row.Cells["Check"].Value -eq $true) {
            $rowsToDelete += $row
        }
    }

    if ($rowsToDelete.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "削除するプリンタが選択されていません。", "Information",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        return
    }

    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "$($rowsToDelete.Count) 台のプリンタを削除しますか？`nこの操作は取り消せません。",
        "Confirm Delete",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $successCount = 0
    $failCount = 0
    $log = ""

    foreach ($row in $rowsToDelete) {
        $pName = $row.Cells["Name"].Value
        try {
            Remove-Printer -Name $pName -ErrorAction Stop
            $log += "[Success] $pName`n"
            $successCount++
        }
        catch {
            $log += "[Failed] $pName : $($_.Exception.Message)`n"
            $failCount++
        }
    }

    # Refresh and show result
    Refresh-Grid

    $icon = if ($failCount -eq 0) { [System.Windows.Forms.MessageBoxIcon]::Information } else { [System.Windows.Forms.MessageBoxIcon]::Warning }
    [System.Windows.Forms.MessageBox]::Show(
        "削除完了`n成功: $successCount`n失敗: $failCount`n`n$log",
        "Result",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        $icon
    )
}

# ========================================
# Main Form
# ========================================
$form = New-Object System.Windows.Forms.Form
$form.Text = "Fabriq Printer Manager"
$form.Size = New-Object System.Drawing.Size(800, 600)
$form.StartPosition = "CenterScreen"
$form.BackColor = $bgDark
$form.ForeColor = $fgText
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

# ========================================
# Top Toolbar Panel
# ========================================
$toolPanel = New-Object System.Windows.Forms.Panel
$toolPanel.Dock = "Top"
$toolPanel.Height = 50
$toolPanel.BackColor = $bgPanel
$toolPanel.Padding = New-Object System.Windows.Forms.Padding(10)

function New-StyledButton {
    param([string]$Text, [int]$X, [int]$Width = 100, $BgColor = $bgButton)
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = $Text
    $btn.Location = New-Object System.Drawing.Point($X, 10)
    $btn.Size = New-Object System.Drawing.Size($Width, 30)
    $btn.FlatStyle = "Flat"
    $btn.FlatAppearance.BorderColor = $gridLine
    $btn.FlatAppearance.MouseOverBackColor = $bgButtonHov
    $btn.BackColor = $BgColor
    $btn.ForeColor = $fgText
    $btn.Cursor = [System.Windows.Forms.Cursors]::Hand
    return $btn
}

$btnRefresh = New-StyledButton -Text "再読み込み" -X 10 -Width 100 -BgColor $bgRefresh
$btnDelete  = New-StyledButton -Text "選択したプリンタを削除" -X 120 -Width 160 -BgColor $bgDelete

$btnSelectAll = New-StyledButton -Text "すべて選択" -X 650 -Width 100
$btnSelectAll.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
# Fix Anchor position manually since simple Anchor logic might be tricky in script
$form.Add_Resize({
    $btnSelectAll.Left = $form.ClientSize.Width - 110
})

$toolPanel.Controls.AddRange(@($btnRefresh, $btnDelete, $btnSelectAll))
$form.Controls.Add($toolPanel)

# ========================================
# Status Bar (Footer)
# ========================================
$statusPanel = New-Object System.Windows.Forms.Panel
$statusPanel.Dock = "Bottom"
$statusPanel.Height = 30
$statusPanel.BackColor = $bgPanel

$script:statusLabel = New-Object System.Windows.Forms.Label
$script:statusLabel.Text = "Ready"
$script:statusLabel.Location = New-Object System.Drawing.Point(10, 8)
$script:statusLabel.AutoSize = $true
$script:statusLabel.ForeColor = $fgDim

$statusPanel.Controls.Add($script:statusLabel)
$form.Controls.Add($statusPanel)

# ========================================
# DataGridView
# ========================================
$script:dgv = New-Object System.Windows.Forms.DataGridView
$script:dgv.Dock = "Fill"
$script:dgv.BackgroundColor = $bgGrid
$script:dgv.GridColor = $gridLine
$script:dgv.BorderStyle = "None"
$script:dgv.CellBorderStyle = "SingleHorizontal"
$script:dgv.RowHeadersVisible = $false
$script:dgv.AllowUserToAddRows = $false
$script:dgv.AllowUserToDeleteRows = $false
$script:dgv.AllowUserToResizeRows = $false
$script:dgv.SelectionMode = "FullRowSelect"
$script:dgv.MultiSelect = $true
$script:dgv.AutoSizeColumnsMode = "Fill"
$script:dgv.ColumnHeadersHeightSizeMode = "DisableResizing"
$script:dgv.ColumnHeadersHeight = 32
$script:dgv.RowTemplate.Height = 28
$script:dgv.DefaultCellStyle.BackColor = $bgCell
$script:dgv.DefaultCellStyle.ForeColor = $fgText
$script:dgv.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(0, 80, 140)
$script:dgv.DefaultCellStyle.SelectionForeColor = $fgText
$script:dgv.EnableHeadersVisualStyles = $false
$script:dgv.ColumnHeadersDefaultCellStyle.BackColor = $bgHeader
$script:dgv.ColumnHeadersDefaultCellStyle.ForeColor = $fgHeader
$script:dgv.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

# DoubleBuffered for smooth rendering
$dgvType = $script:dgv.GetType()
$pi = $dgvType.GetProperty("DoubleBuffered", [System.Reflection.BindingFlags]"Instance,NonPublic")
$pi.SetValue($script:dgv, $true, $null)

# --- Columns ---

# Checkbox
$colCheck = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$colCheck.Name = "Check"
$colCheck.HeaderText = "選択"
$colCheck.Width = 50
$colCheck.FillWeight = 10
$script:dgv.Columns.Add($colCheck)

# Printer Name
$colName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colName.Name = "Name"
$colName.HeaderText = "プリンタ名"
$colName.FillWeight = 40
$colName.ReadOnly = $true
$script:dgv.Columns.Add($colName)

# Driver
$colDriver = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colDriver.Name = "Driver"
$colDriver.HeaderText = "ドライバ"
$colDriver.FillWeight = 30
$colDriver.ReadOnly = $true
$script:dgv.Columns.Add($colDriver)

# Port
$colPort = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colPort.Name = "Port"
$colPort.HeaderText = "ポート"
$colPort.FillWeight = 20
$colPort.ReadOnly = $true
$script:dgv.Columns.Add($colPort)

$form.Controls.Add($script:dgv)

# ========================================
# Events
# ========================================

# Checkbox click handling (Single click toggle)
$script:dgv.Add_CellContentClick({
    param($sender, $e)
    if ($e.RowIndex -ge 0 -and $sender.Columns[$e.ColumnIndex].Name -eq "Check") {
        $sender.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
    }
})

$btnRefresh.Add_Click({
    Refresh-Grid
})

$btnDelete.Add_Click({
    Remove-SelectedPrinters
})

$btnSelectAll.Add_Click({
    foreach ($row in $script:dgv.Rows) {
        $row.Cells["Check"].Value = -not $row.Cells["Check"].Value
    }
})

# ========================================
# Layout Ordering
# ========================================
$statusPanel.BringToFront()
$toolPanel.BringToFront()
$script:dgv.BringToFront()

# ========================================
# Initial Load
# ========================================
Refresh-Grid

# ========================================
# Show Form
# ========================================
$form.ShowDialog() | Out-Null
$form.Dispose()