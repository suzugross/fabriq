# ========================================
# AutoKey Recipe Editor
# ========================================
# Visual recipe editor for AutoKey modules.
# Creates individual autokey modules from templates.
# ========================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic

[System.Windows.Forms.Application]::EnableVisualStyles()

# ========================================
# Constants
# ========================================
$script:Actions = @("Open", "WaitWin", "AppFocus", "Type", "Key", "Wait")
$script:FabriqRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -ErrorAction SilentlyContinue
if (-not $script:FabriqRoot) { $script:FabriqRoot = (Resolve-Path "..").Path }
$script:TemplatePath = Join-Path $PSScriptRoot "autokey_template"
$script:ExtendedPath = Join-Path $script:FabriqRoot "modules\extended"
$script:CurrentFilePath = $null
$script:DragRowIndex = -1
$script:IsLoading = $false

# Action default Wait values
$script:ActionDefaults = @{
    "Open"     = @{ Wait = "0" }
    "WaitWin"  = @{ Wait = "10000" }
    "AppFocus" = @{ Wait = "500" }
    "Type"     = @{ Wait = "500" }
    "Key"      = @{ Wait = "200" }
    "Wait"     = @{ Wait = "0"; Value = "1000" }
}

# ========================================
# Color Scheme
# ========================================
$bgDark       = [System.Drawing.Color]::FromArgb(30, 30, 30)
$bgPanel      = [System.Drawing.Color]::FromArgb(40, 40, 40)
$bgGrid       = [System.Drawing.Color]::FromArgb(35, 35, 35)
$bgCell        = [System.Drawing.Color]::FromArgb(45, 45, 45)
$bgHeader     = [System.Drawing.Color]::FromArgb(55, 55, 55)
$bgButton     = [System.Drawing.Color]::FromArgb(60, 60, 60)
$bgButtonHov  = [System.Drawing.Color]::FromArgb(80, 80, 80)
$bgExport     = [System.Drawing.Color]::FromArgb(0, 120, 215)
$bgTestRun    = [System.Drawing.Color]::FromArgb(200, 80, 0)
$fgText       = [System.Drawing.Color]::FromArgb(220, 220, 220)
$fgDim        = [System.Drawing.Color]::FromArgb(150, 150, 150)
$fgHeader     = [System.Drawing.Color]::FromArgb(100, 180, 255)
$fgFooter     = [System.Drawing.Color]::FromArgb(170, 170, 170)
$errorBg      = [System.Drawing.Color]::FromArgb(80, 30, 30)
$accentCyan   = [System.Drawing.Color]::FromArgb(0, 200, 200)
$gridLine     = [System.Drawing.Color]::FromArgb(60, 60, 60)

# ========================================
# Helper: Renumber Steps
# ========================================
function Update-StepNumbers {
    for ($i = 0; $i -lt $script:dgv.Rows.Count; $i++) {
        if (-not $script:dgv.Rows[$i].IsNewRow) {
            $script:dgv.Rows[$i].Cells["Step"].Value = ($i + 1).ToString()
        }
    }
}

# ========================================
# Helper: Validate Wait cell
# ========================================
function Test-WaitCell {
    param($row, $colIndex)
    $val = $row.Cells[$colIndex].Value
    if ([string]::IsNullOrWhiteSpace($val) -or $val -eq "0") {
        $row.Cells[$colIndex].Style.BackColor = $bgCell
        return $true
    }
    $parsed = 0
    if ([int]::TryParse($val, [ref]$parsed) -and $parsed -ge 0) {
        $row.Cells[$colIndex].Style.BackColor = $bgCell
        return $true
    }
    else {
        $row.Cells[$colIndex].Style.BackColor = $errorBg
        return $false
    }
}

# ========================================
# Helper: Validate all Wait cells
# ========================================
function Test-AllWaitCells {
    $allValid = $true
    $waitIdx = $script:dgv.Columns["Wait"].Index
    for ($i = 0; $i -lt $script:dgv.Rows.Count; $i++) {
        if (-not $script:dgv.Rows[$i].IsNewRow) {
            if (-not (Test-WaitCell -row $script:dgv.Rows[$i] -colIndex $waitIdx)) {
                $allValid = $false
            }
        }
    }
    return $allValid
}

# ========================================
# Helper: Grid data to CSV string
# ========================================
function Get-GridAsCsv {
    $lines = @("Step,Action,Value,Wait,Note")
    for ($i = 0; $i -lt $script:dgv.Rows.Count; $i++) {
        $row = $script:dgv.Rows[$i]
        if ($row.IsNewRow) { continue }
        $step   = ($i + 1).ToString()
        $action = if ($row.Cells["Action"].Value) { $row.Cells["Action"].Value } else { "" }
        $value  = if ($row.Cells["Value"].Value)  { $row.Cells["Value"].Value }  else { "" }
        $wait   = if ($row.Cells["Wait"].Value)   { $row.Cells["Wait"].Value }   else { "0" }
        $note   = if ($row.Cells["Note"].Value)   { $row.Cells["Note"].Value }   else { "" }

        # Escape CSV fields
        $value = $value -replace '"', '""'
        $note  = $note  -replace '"', '""'
        if ($value -match '[,"]') { $value = "`"$value`"" }
        if ($note  -match '[,"]') { $note  = "`"$note`""  }

        $lines += "$step,$action,$value,$wait,$note"
    }
    return ($lines -join "`r`n") + "`r`n"
}

# ========================================
# Helper: Load CSV into grid
# ========================================
function Import-RecipeCsv {
    param([string]$Path)
    try {
        $data = @(Import-Csv -Path $Path -Encoding Default)
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to load CSV: $_",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return
    }

    $script:IsLoading = $true

    $script:dgv.Rows.Clear()

    foreach ($item in $data) {
        if ([string]::IsNullOrWhiteSpace($item.Action)) { continue }
        $idx = $script:dgv.Rows.Add()
        $row = $script:dgv.Rows[$idx]
        $row.Cells["Step"].Value   = if ($item.Step) { $item.Step } else { "" }
        $row.Cells["Action"].Value = $item.Action
        $row.Cells["Value"].Value  = $item.Value
        $row.Cells["Wait"].Value   = if ($item.Wait) { $item.Wait } else { "0" }
        $row.Cells["Note"].Value   = $item.Note
    }

    Update-StepNumbers

    # Validate all Wait cells
    $waitIdx = $script:dgv.Columns["Wait"].Index
    for ($i = 0; $i -lt $script:dgv.Rows.Count; $i++) {
        if (-not $script:dgv.Rows[$i].IsNewRow) {
            $null = Test-WaitCell -row $script:dgv.Rows[$i] -colIndex $waitIdx
        }
    }

    $script:IsLoading = $false

    # Reset scroll to top
    if ($script:dgv.Rows.Count -gt 0) {
        $script:dgv.ClearSelection()
        $script:dgv.FirstDisplayedScrollingRowIndex = 0
        $script:dgv.Rows[0].Selected = $true
        $script:dgv.CurrentCell = $script:dgv.Rows[0].Cells["Action"]
    }
}

# ========================================
# Main Form
# ========================================
$form = New-Object System.Windows.Forms.Form
$form.Text = "AutoKey Recipe Editor"
$form.Size = New-Object System.Drawing.Size(960, 750)
$form.MinimumSize = New-Object System.Drawing.Size(800, 600)
$form.StartPosition = "CenterScreen"
$form.BackColor = $bgDark
$form.ForeColor = $fgText
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

# ========================================
# Top Toolbar Panel
# ========================================
$toolPanel = New-Object System.Windows.Forms.Panel
$toolPanel.Dock = "Top"
$toolPanel.Height = 40
$toolPanel.BackColor = $bgPanel
$toolPanel.Padding = New-Object System.Windows.Forms.Padding(8, 5, 8, 5)

# Helper: Create toolbar button
function New-ToolButton {
    param([string]$Text, [int]$X, [int]$Width = 80, $BgColor = $bgButton)
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = $Text
    $btn.Location = New-Object System.Drawing.Point($X, 5)
    $btn.Size = New-Object System.Drawing.Size($Width, 28)
    $btn.FlatStyle = "Flat"
    $btn.FlatAppearance.BorderColor = $gridLine
    $btn.FlatAppearance.MouseOverBackColor = $bgButtonHov
    $btn.BackColor = $BgColor
    $btn.ForeColor = $fgText
    $btn.Cursor = [System.Windows.Forms.Cursors]::Hand
    return $btn
}

$btnNew    = New-ToolButton -Text "New"         -X 8   -Width 60
$btnOpen   = New-ToolButton -Text "Open CSV"    -X 75  -Width 85
$btnSave   = New-ToolButton -Text "Save"        -X 167 -Width 60
$btnExport = New-ToolButton -Text "Export Module" -X 234 -Width 110 -BgColor $bgExport
$btnTest   = New-ToolButton -Text "Test Run"    -X 351 -Width 85  -BgColor $bgTestRun

$toolPanel.Controls.AddRange(@($btnNew, $btnOpen, $btnSave, $btnExport, $btnTest))
$form.Controls.Add($toolPanel)

# ========================================
# Module Name Panel
# ========================================
$namePanel = New-Object System.Windows.Forms.Panel
$namePanel.Dock = "Top"
$namePanel.Height = 35
$namePanel.BackColor = $bgPanel
$namePanel.Padding = New-Object System.Windows.Forms.Padding(8, 4, 8, 4)

$nameLabel = New-Object System.Windows.Forms.Label
$nameLabel.Text = "Module Name:"
$nameLabel.Location = New-Object System.Drawing.Point(10, 8)
$nameLabel.AutoSize = $true
$nameLabel.ForeColor = $fgDim

$script:txtModuleName = New-Object System.Windows.Forms.TextBox
$script:txtModuleName.Location = New-Object System.Drawing.Point(110, 5)
$script:txtModuleName.Size = New-Object System.Drawing.Size(300, 24)
$script:txtModuleName.BackColor = $bgCell
$script:txtModuleName.ForeColor = $fgText
$script:txtModuleName.BorderStyle = "FixedSingle"
$script:txtModuleName.Font = New-Object System.Drawing.Font("Segoe UI", 10)

$namePanel.Controls.AddRange(@($nameLabel, $script:txtModuleName))
$form.Controls.Add($namePanel)

# ========================================
# Footer Panel (Action Reference)
# ========================================
$footerPanel = New-Object System.Windows.Forms.Panel
$footerPanel.Dock = "Bottom"
$footerPanel.Height = 170
$footerPanel.BackColor = [System.Drawing.Color]::FromArgb(25, 25, 25)
$footerPanel.Padding = New-Object System.Windows.Forms.Padding(12, 6, 12, 6)

$footerText = @"
Action 解説:
  Open       アプリケーション、ファイル、URLを開きます  例: notepad, calc, https://www.google.com, ms-settings:windowsupdate
  WaitWin    指定したタイトル(部分一致)のウィンドウが表示されるまで待機します  Wait列=タイムアウト(ms) デフォルト: 10000ms(10秒)
  AppFocus   既に開いているウィンドウにフォーカスを移動します (タイトル部分一致)
  Type       アクティブなウィンドウにテキストを入力します (SendKeys経由)
  Key        特殊キーやショートカットを送信します (SendKeys記法)
               {ENTER} {TAB} {ESC} {BACKSPACE} {DELETE} {UP} {DOWN} {LEFT} {RIGHT} {F1}~{F12}
               %(Alt) ^(Ctrl) +(Shift)  例: %{F4}=Alt+F4  ^c=Ctrl+C  +{TAB}=Shift+Tab  ^a=全選択
  Wait       Value列に指定したミリ秒だけ処理を一時停止します  例: Value=2000 → 2秒待機
"@

$footerLabel = New-Object System.Windows.Forms.Label
$footerLabel.Dock = "Fill"
$footerLabel.Text = $footerText
$footerLabel.ForeColor = $fgFooter
$footerLabel.Font = New-Object System.Drawing.Font("Consolas", 8.5)
$footerLabel.AutoSize = $false

$footerPanel.Controls.Add($footerLabel)
$form.Controls.Add($footerPanel)

# ========================================
# Row Operation Panel
# ========================================
$rowPanel = New-Object System.Windows.Forms.Panel
$rowPanel.Dock = "Bottom"
$rowPanel.Height = 38
$rowPanel.BackColor = $bgPanel
$rowPanel.Padding = New-Object System.Windows.Forms.Padding(8, 4, 8, 4)

function New-RowButton {
    param([string]$Text, [int]$X, [int]$Width = 95)
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = $Text
    $btn.Location = New-Object System.Drawing.Point($X, 4)
    $btn.Size = New-Object System.Drawing.Size($Width, 28)
    $btn.FlatStyle = "Flat"
    $btn.FlatAppearance.BorderColor = $gridLine
    $btn.FlatAppearance.MouseOverBackColor = $bgButtonHov
    $btn.BackColor = $bgButton
    $btn.ForeColor = $fgText
    $btn.Cursor = [System.Windows.Forms.Cursors]::Hand
    return $btn
}

$btnAddRow      = New-RowButton -Text "+ Add Row"      -X 8
$btnInsertAbove = New-RowButton -Text "Insert Above"   -X 110 -Width 105
$btnDeleteRow   = New-RowButton -Text "Delete Row"     -X 222 -Width 95
$btnMoveUp      = New-RowButton -Text "Move Up"        -X 340 -Width 80
$btnMoveDown    = New-RowButton -Text "Move Down"      -X 427 -Width 85

$rowPanel.Controls.AddRange(@($btnAddRow, $btnInsertAbove, $btnDeleteRow, $btnMoveUp, $btnMoveDown))
$form.Controls.Add($rowPanel)

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
$script:dgv.MultiSelect = $false
$script:dgv.AllowDrop = $true
$script:dgv.AutoSizeColumnsMode = "Fill"
$script:dgv.ColumnHeadersHeightSizeMode = "DisableResizing"
$script:dgv.ColumnHeadersHeight = 32
$script:dgv.RowTemplate.Height = 28
$script:dgv.DefaultCellStyle.BackColor = $bgCell
$script:dgv.DefaultCellStyle.ForeColor = $fgText
$script:dgv.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(0, 80, 140)
$script:dgv.DefaultCellStyle.SelectionForeColor = $fgText
$script:dgv.DefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$script:dgv.ColumnHeadersDefaultCellStyle.BackColor = $bgHeader
$script:dgv.ColumnHeadersDefaultCellStyle.ForeColor = $fgHeader
$script:dgv.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$script:dgv.ColumnHeadersDefaultCellStyle.Alignment = "MiddleCenter"
$script:dgv.EnableHeadersVisualStyles = $false

# Enable DoubleBuffered for flicker-free rendering
$dgvType = $script:dgv.GetType()
$pi = $dgvType.GetProperty("DoubleBuffered", [System.Reflection.BindingFlags]"Instance,NonPublic")
$pi.SetValue($script:dgv, $true, $null)

# --- Columns ---

# Step (auto-number, read-only)
$colStep = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colStep.Name = "Step"
$colStep.HeaderText = "Step (自動連番)"
$colStep.Width = 50
$colStep.MinimumWidth = 40
$colStep.FillWeight = 8
$colStep.ReadOnly = $true
$colStep.DefaultCellStyle.Alignment = "MiddleCenter"
$colStep.SortMode = "NotSortable"
$null = $script:dgv.Columns.Add($colStep)

# Action (ComboBox)
$colAction = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
$colAction.Name = "Action"
$colAction.HeaderText = "Action (実行命令)"
$colAction.FillWeight = 15
$colAction.MinimumWidth = 90
$colAction.FlatStyle = "Flat"
$colAction.SortMode = "NotSortable"
foreach ($a in $script:Actions) { $null = $colAction.Items.Add($a) }
$null = $script:dgv.Columns.Add($colAction)

# Value
$colValue = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colValue.Name = "Value"
$colValue.HeaderText = "Value (対象パス/テキスト/キー)"
$colValue.FillWeight = 35
$colValue.MinimumWidth = 120
$colValue.SortMode = "NotSortable"
$null = $script:dgv.Columns.Add($colValue)

# Wait
$colWait = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colWait.Name = "Wait"
$colWait.HeaderText = "Wait (待機ms)"
$colWait.FillWeight = 12
$colWait.MinimumWidth = 60
$colWait.DefaultCellStyle.Alignment = "MiddleRight"
$colWait.SortMode = "NotSortable"
$null = $script:dgv.Columns.Add($colWait)

# Note
$colNote = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colNote.Name = "Note"
$colNote.HeaderText = "Note (メモ)"
$colNote.FillWeight = 30
$colNote.MinimumWidth = 100
$colNote.SortMode = "NotSortable"
$null = $script:dgv.Columns.Add($colNote)

$form.Controls.Add($script:dgv)

# ========================================
# Event: CellValueChanged (Action defaults)
# ========================================
$script:dgv.Add_CellValueChanged({
    param($sender, $e)
    if ($e.RowIndex -lt 0) { return }
    if ($script:IsLoading) { return }
    $colName = $sender.Columns[$e.ColumnIndex].Name

    if ($colName -eq "Action") {
        $row = $sender.Rows[$e.RowIndex]
        $action = $row.Cells["Action"].Value
        if ($script:ActionDefaults.ContainsKey($action)) {
            $defaults = $script:ActionDefaults[$action]
            # Only set Wait if current value is empty or "0"
            $currentWait = $row.Cells["Wait"].Value
            if ([string]::IsNullOrWhiteSpace($currentWait) -or $currentWait -eq "0") {
                $row.Cells["Wait"].Value = $defaults.Wait
            }
            # Set Value default for Wait action
            if ($defaults.ContainsKey("Value")) {
                $currentValue = $row.Cells["Value"].Value
                if ([string]::IsNullOrWhiteSpace($currentValue)) {
                    $row.Cells["Value"].Value = $defaults.Value
                }
            }
        }
        # Validate Wait cell
        $waitIdx = $sender.Columns["Wait"].Index
        $null = Test-WaitCell -row $row -colIndex $waitIdx
    }

    if ($colName -eq "Wait") {
        $row = $sender.Rows[$e.RowIndex]
        $waitIdx = $e.ColumnIndex
        $null = Test-WaitCell -row $row -colIndex $waitIdx
    }
})

# Need to commit edit immediately for ComboBox changes
$script:dgv.Add_CurrentCellDirtyStateChanged({
    param($sender, $e)
    if ($script:IsLoading) { return }
    if ($sender.IsCurrentCellDirty) {
        $null = $sender.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
    }
})

# ========================================
# Event: CellDoubleClick (Open → file dialog)
# ========================================
$script:dgv.Add_CellDoubleClick({
    param($sender, $e)
    if ($e.RowIndex -lt 0) { return }
    $colName = $sender.Columns[$e.ColumnIndex].Name
    if ($colName -ne "Value") { return }

    $row = $sender.Rows[$e.RowIndex]
    $action = $row.Cells["Action"].Value
    if ($action -ne "Open") { return }

    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Title = "Select file or application"
    $dlg.Filter = "All Files (*.*)|*.*|Executables (*.exe)|*.exe|Scripts (*.ps1;*.bat;*.cmd)|*.ps1;*.bat;*.cmd"
    $dlg.FilterIndex = 1

    # Pre-fill if current value is a valid path
    $currentVal = $row.Cells["Value"].Value
    if ($currentVal -and (Test-Path (Split-Path $currentVal -Parent -ErrorAction SilentlyContinue) -ErrorAction SilentlyContinue)) {
        $dlg.InitialDirectory = Split-Path $currentVal -Parent
    }

    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $row.Cells["Value"].Value = $dlg.FileName
    }
    $dlg.Dispose()
})

# ========================================
# Event: Drag & Drop row reordering
# ========================================
$script:dgv.Add_MouseDown({
    param($sender, $e)
    if ($e.Button -ne [System.Windows.Forms.MouseButtons]::Left) { return }
    $hitTest = $sender.HitTest($e.X, $e.Y)
    if ($hitTest.Type -eq [System.Windows.Forms.DataGridViewHitTestType]::Cell -or
        $hitTest.Type -eq [System.Windows.Forms.DataGridViewHitTestType]::RowHeader) {
        $script:DragRowIndex = $hitTest.RowIndex
    }
})

$script:dgv.Add_MouseMove({
    param($sender, $e)
    if ($e.Button -ne [System.Windows.Forms.MouseButtons]::Left) { return }
    if ($script:DragRowIndex -lt 0) { return }

    # Only start drag if moved enough
    $threshold = [System.Windows.Forms.SystemInformation]::DragSize
    $dragRect = New-Object System.Drawing.Rectangle(
        ([System.Windows.Forms.Control]::MousePosition.X - $threshold.Width / 2),
        ([System.Windows.Forms.Control]::MousePosition.Y - $threshold.Height / 2),
        $threshold.Width, $threshold.Height
    )

    if (-not $sender.Dragging) {
        $null = $sender.DoDragDrop($script:DragRowIndex, [System.Windows.Forms.DragDropEffects]::Move)
    }
})

$script:dgv.Add_DragOver({
    param($sender, $e)
    $e.Effect = [System.Windows.Forms.DragDropEffects]::Move
})

$script:dgv.Add_DragDrop({
    param($sender, $e)
    $pt = [System.Drawing.Point]::new($e.X, $e.Y)
    $clientPoint = $sender.PointToClient($pt)
    $hitTest = $sender.HitTest($clientPoint.X, $clientPoint.Y)
    $targetIndex = $hitTest.RowIndex

    if ($targetIndex -lt 0 -or $targetIndex -ge $sender.Rows.Count) { return }
    if ($script:DragRowIndex -lt 0 -or $script:DragRowIndex -eq $targetIndex) {
        $script:DragRowIndex = -1
        return
    }
    if ($sender.Rows[$script:DragRowIndex].IsNewRow) {
        $script:DragRowIndex = -1
        return
    }

    # Save source row data
    $srcRow = $sender.Rows[$script:DragRowIndex]
    $data = @{
        Action = $srcRow.Cells["Action"].Value
        Value  = $srcRow.Cells["Value"].Value
        Wait   = $srcRow.Cells["Wait"].Value
        Note   = $srcRow.Cells["Note"].Value
    }

    # Remove source row
    $sender.Rows.RemoveAt($script:DragRowIndex)

    # Adjust target index if needed
    if ($script:DragRowIndex -lt $targetIndex) { $targetIndex-- }

    # Insert at new position
    $null = $sender.Rows.Insert($targetIndex, 1)
    $newRow = $sender.Rows[$targetIndex]
    $newRow.Cells["Action"].Value = $data.Action
    $newRow.Cells["Value"].Value  = $data.Value
    $newRow.Cells["Wait"].Value   = $data.Wait
    $newRow.Cells["Note"].Value   = $data.Note

    $sender.ClearSelection()
    $sender.Rows[$targetIndex].Selected = $true
    $sender.CurrentCell = $sender.Rows[$targetIndex].Cells["Action"]

    Update-StepNumbers
    $script:DragRowIndex = -1
})

# ========================================
# Button: Add Row
# ========================================
$btnAddRow.Add_Click({
    $idx = $script:dgv.Rows.Add()
    $script:dgv.Rows[$idx].Cells["Wait"].Value = "0"
    Update-StepNumbers
    $script:dgv.ClearSelection()
    $script:dgv.Rows[$idx].Selected = $true
    $script:dgv.CurrentCell = $script:dgv.Rows[$idx].Cells["Action"]
})

# ========================================
# Button: Insert Above
# ========================================
$btnInsertAbove.Add_Click({
    $selIdx = if ($script:dgv.CurrentRow) { $script:dgv.CurrentRow.Index } else { -1 }
    if ($selIdx -lt 0) {
        $selIdx = 0
    }
    $null = $script:dgv.Rows.Insert($selIdx, 1)
    $script:dgv.Rows[$selIdx].Cells["Wait"].Value = "0"
    Update-StepNumbers
    $script:dgv.ClearSelection()
    $script:dgv.Rows[$selIdx].Selected = $true
    $script:dgv.CurrentCell = $script:dgv.Rows[$selIdx].Cells["Action"]
})

# ========================================
# Button: Delete Row
# ========================================
$btnDeleteRow.Add_Click({
    if ($script:dgv.CurrentRow -and -not $script:dgv.CurrentRow.IsNewRow) {
        $idx = $script:dgv.CurrentRow.Index
        $script:dgv.Rows.RemoveAt($idx)
        Update-StepNumbers
    }
})

# ========================================
# Button: Move Up
# ========================================
$btnMoveUp.Add_Click({
    if (-not $script:dgv.CurrentRow) { return }
    $idx = $script:dgv.CurrentRow.Index
    if ($idx -le 0) { return }

    $row = $script:dgv.Rows[$idx]
    $data = @{
        Action = $row.Cells["Action"].Value
        Value  = $row.Cells["Value"].Value
        Wait   = $row.Cells["Wait"].Value
        Note   = $row.Cells["Note"].Value
    }
    $script:dgv.Rows.RemoveAt($idx)
    $newIdx = $idx - 1
    $null = $script:dgv.Rows.Insert($newIdx, 1)
    $newRow = $script:dgv.Rows[$newIdx]
    $newRow.Cells["Action"].Value = $data.Action
    $newRow.Cells["Value"].Value  = $data.Value
    $newRow.Cells["Wait"].Value   = $data.Wait
    $newRow.Cells["Note"].Value   = $data.Note

    $script:dgv.ClearSelection()
    $script:dgv.Rows[$newIdx].Selected = $true
    $script:dgv.CurrentCell = $script:dgv.Rows[$newIdx].Cells["Action"]
    Update-StepNumbers
})

# ========================================
# Button: Move Down
# ========================================
$btnMoveDown.Add_Click({
    if (-not $script:dgv.CurrentRow) { return }
    $idx = $script:dgv.CurrentRow.Index
    if ($idx -ge ($script:dgv.Rows.Count - 1)) { return }

    $row = $script:dgv.Rows[$idx]
    $data = @{
        Action = $row.Cells["Action"].Value
        Value  = $row.Cells["Value"].Value
        Wait   = $row.Cells["Wait"].Value
        Note   = $row.Cells["Note"].Value
    }
    $script:dgv.Rows.RemoveAt($idx)
    $newIdx = $idx + 1
    if ($newIdx -gt $script:dgv.Rows.Count) { $newIdx = $script:dgv.Rows.Count }
    $null = $script:dgv.Rows.Insert($newIdx, 1)
    $newRow = $script:dgv.Rows[$newIdx]
    $newRow.Cells["Action"].Value = $data.Action
    $newRow.Cells["Value"].Value  = $data.Value
    $newRow.Cells["Wait"].Value   = $data.Wait
    $newRow.Cells["Note"].Value   = $data.Note

    $script:dgv.ClearSelection()
    $script:dgv.Rows[$newIdx].Selected = $true
    $script:dgv.CurrentCell = $script:dgv.Rows[$newIdx].Cells["Action"]
    Update-StepNumbers
})

# ========================================
# Button: New
# ========================================
$btnNew.Add_Click({
    $result = [System.Windows.Forms.MessageBox]::Show(
        "Clear current recipe and start new?",
        "New Recipe",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    if ($result -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $script:dgv.Rows.Clear()
    $script:txtModuleName.Text = ""
    $script:CurrentFilePath = $null
    $form.Text = "AutoKey Recipe Editor"

    # Add one empty row
    $idx = $script:dgv.Rows.Add()
    $script:dgv.Rows[$idx].Cells["Wait"].Value = "0"
    Update-StepNumbers
})

# ========================================
# Button: Open CSV
# ========================================
$btnOpen.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Title = "Open Recipe CSV"
    $dlg.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $dlg.FilterIndex = 1

    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        Import-RecipeCsv -Path $dlg.FileName
        $script:CurrentFilePath = $dlg.FileName
        $form.Text = "AutoKey Recipe Editor - $([System.IO.Path]::GetFileName($dlg.FileName))"
    }
    $dlg.Dispose()
})

# ========================================
# Button: Save
# ========================================
$btnSave.Add_Click({
    # Validate
    if (-not (Test-AllWaitCells)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Wait column contains non-numeric values. Please fix the highlighted cells.",
            "Validation Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    $csvContent = Get-GridAsCsv

    if ($script:CurrentFilePath) {
        # Overwrite existing file
        try {
            [System.IO.File]::WriteAllText($script:CurrentFilePath, $csvContent, [System.Text.Encoding]::Default)
            [System.Windows.Forms.MessageBox]::Show(
                "Saved: $($script:CurrentFilePath)",
                "Save",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Failed to save: $_",
                "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
    }
    else {
        # SaveFileDialog
        $dlg = New-Object System.Windows.Forms.SaveFileDialog
        $dlg.Title = "Save Recipe CSV"
        $dlg.Filter = "CSV Files (*.csv)|*.csv"
        $dlg.FileName = "recipe.csv"

        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            try {
                [System.IO.File]::WriteAllText($dlg.FileName, $csvContent, [System.Text.Encoding]::Default)
                $script:CurrentFilePath = $dlg.FileName
                $form.Text = "AutoKey Recipe Editor - $([System.IO.Path]::GetFileName($dlg.FileName))"
                [System.Windows.Forms.MessageBox]::Show(
                    "Saved: $($dlg.FileName)",
                    "Save",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information
                )
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show(
                    "Failed to save: $_",
                    "Error",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error
                )
            }
        }
        $dlg.Dispose()
    }
})

# ========================================
# Button: Export Module
# ========================================
$btnExport.Add_Click({
    # Validate module name
    $moduleName = $script:txtModuleName.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($moduleName)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please enter a Module Name.",
            "Export",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        $script:txtModuleName.Focus()
        return
    }

    # Sanitize folder name
    $safeName = $moduleName -replace '[\\/:*?"<>|\s]', '_'
    $safeName = $safeName.ToLower()

    # Validate Wait cells
    if (-not (Test-AllWaitCells)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Wait column contains non-numeric values. Please fix the highlighted cells.",
            "Validation Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    # Check if grid has data
    $dataRows = 0
    for ($i = 0; $i -lt $script:dgv.Rows.Count; $i++) {
        if (-not $script:dgv.Rows[$i].IsNewRow -and $script:dgv.Rows[$i].Cells["Action"].Value) {
            $dataRows++
        }
    }
    if ($dataRows -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Recipe is empty. Add at least one step before exporting.",
            "Export",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    # Check template exists
    $templateScript = Join-Path $script:TemplatePath "autokey_config.ps1"
    if (-not (Test-Path $templateScript)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Template not found: $templateScript",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return
    }

    # Target directory
    $targetDir = Join-Path $script:ExtendedPath $safeName

    # Check if already exists
    if (Test-Path $targetDir) {
        $overwrite = [System.Windows.Forms.MessageBox]::Show(
            "Module '$safeName' already exists.`nOverwrite?",
            "Export",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        if ($overwrite -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    }

    try {
        # Create directory
        if (-not (Test-Path $targetDir)) {
            $null = New-Item -ItemType Directory -Path $targetDir -Force
        }

        # Copy autokey_config.ps1 → {name}_autokey_config.ps1
        $scriptName = "${safeName}_autokey_config.ps1"
        Copy-Item -Path $templateScript -Destination (Join-Path $targetDir $scriptName) -Force

        # Generate module.csv
        $moduleCsv = "MenuName,Category,Script,Order,Enabled`r`n"
        $moduleCsv += "$moduleName,Automation,$scriptName,50,0`r`n"
        [System.IO.File]::WriteAllText(
            (Join-Path $targetDir "module.csv"),
            $moduleCsv,
            [System.Text.Encoding]::Default
        )

        # Generate recipe.csv
        $recipeCsv = Get-GridAsCsv
        [System.IO.File]::WriteAllText(
            (Join-Path $targetDir "recipe.csv"),
            $recipeCsv,
            [System.Text.Encoding]::Default
        )

        $script:CurrentFilePath = Join-Path $targetDir "recipe.csv"
        $form.Text = "AutoKey Recipe Editor - $safeName/recipe.csv"

        [System.Windows.Forms.MessageBox]::Show(
            "Module exported successfully!`n`nLocation: $targetDir`n`nFiles:`n  - $scriptName`n  - module.csv`n  - recipe.csv`n`nNote: Enabled=0 (disabled by default).`nChange to 1 in module.csv or CSV Editor to enable.",
            "Export Complete",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Export failed: $_",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
})

# ========================================
# Button: Test Run
# ========================================
$btnTest.Add_Click({
    # Validate Wait cells
    if (-not (Test-AllWaitCells)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Wait column contains non-numeric values. Please fix the highlighted cells.",
            "Validation Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    # Check if grid has data
    $dataRows = 0
    for ($i = 0; $i -lt $script:dgv.Rows.Count; $i++) {
        if (-not $script:dgv.Rows[$i].IsNewRow -and $script:dgv.Rows[$i].Cells["Action"].Value) {
            $dataRows++
        }
    }
    if ($dataRows -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Recipe is empty.",
            "Test Run",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    # Save recipe to temp file
    $tempCsv = Join-Path $env:TEMP "fabriq_recipe_test.csv"
    $csvContent = Get-GridAsCsv
    [System.IO.File]::WriteAllText($tempCsv, $csvContent, [System.Text.Encoding]::Default)

    # Build test runner script (self-contained, no Fabriq dependencies)
    $testScript = Join-Path $env:TEMP "fabriq_recipe_test_runner.ps1"
    $testContent = @'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic

$csvPath = Join-Path $env:TEMP "fabriq_recipe_test.csv"
$steps = @(Import-Csv -Path $csvPath -Encoding Default | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Action) })

$WsShell = New-Object -ComObject WScript.Shell

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "AutoKey Test Run ($($steps.Count) steps)" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Starting in 3 seconds..." -ForegroundColor Yellow
Start-Sleep -Seconds 3
Write-Host ""

$total = $steps.Count
$current = 0

foreach ($step in $steps) {
    $current++
    $stepNum = if ($step.Step) { $step.Step } else { $current }
    Write-Host "[$current/$total] Step $stepNum - $($step.Action): $($step.Note)" -ForegroundColor Cyan

    $ok = $false
    switch ($step.Action) {
        "Open" {
            try {
                $val = $step.Value
                if ($val -match " ") {
                    $parts = $val -split " ", 2
                    try {
                        $null = Start-Process -FilePath $parts[0] -ArgumentList $parts[1] -ErrorAction Stop -PassThru
                    } catch {
                        $null = Start-Process -FilePath "cmd" -ArgumentList "/c start $($parts[0]) $($parts[1])" -WindowStyle Hidden
                    }
                } else {
                    $null = Start-Process $val -ErrorAction Stop -PassThru
                }
                Write-Host "  [OK] Opened: $val" -ForegroundColor Green
                $ok = $true
            } catch {
                Write-Host "  [ERROR] Failed: $_" -ForegroundColor Red
            }
        }
        "WaitWin" {
            $timeout = 10000
            if ($step.Wait) { $p = 0; if ([int]::TryParse($step.Wait, [ref]$p) -and $p -gt 0) { $timeout = $p } }
            $elapsed = 0
            Write-Host "  Waiting for '$($step.Value)' (max $($timeout/1000)s)..." -NoNewline -ForegroundColor Gray
            while ($elapsed -lt $timeout) {
                $proc = Get-Process | Where-Object { $_.MainWindowTitle -like "*$($step.Value)*" } | Select-Object -First 1
                if ($proc) {
                    Write-Host " Found!" -ForegroundColor Green
                    try { $WsShell.AppActivate($proc.Id) | Out-Null; Start-Sleep -Milliseconds 500 } catch {}
                    $ok = $true; break
                }
                Start-Sleep -Milliseconds 500; $elapsed += 500
                Write-Host "." -NoNewline
            }
            if (-not $ok) { Write-Host " Timeout!" -ForegroundColor Yellow }
        }
        "AppFocus" {
            try { $ok = [bool]$WsShell.AppActivate($step.Value) } catch {}
            if ($ok) { Write-Host "  [OK] Focused: $($step.Value)" -ForegroundColor Green }
            else { Write-Host "  [WARN] Focus failed: $($step.Value)" -ForegroundColor Yellow }
        }
        "Type" {
            try { [System.Windows.Forms.SendKeys]::SendWait($step.Value); $ok = $true } catch {}
            if ($ok) { Write-Host "  [OK] Typed text" -ForegroundColor Green }
            else { Write-Host "  [ERROR] Type failed" -ForegroundColor Red }
        }
        "Key" {
            try { [System.Windows.Forms.SendKeys]::SendWait($step.Value); $ok = $true } catch {}
            if ($ok) { Write-Host "  [OK] Key: $($step.Value)" -ForegroundColor Green }
            else { Write-Host "  [ERROR] Key failed" -ForegroundColor Red }
        }
        "Wait" {
            $ms = 0; if ($step.Value) { $null = [int]::TryParse($step.Value, [ref]$ms) }
            if ($ms -gt 0) { Start-Sleep -Milliseconds $ms }
            Write-Host "  [OK] Waited ${ms}ms" -ForegroundColor Gray
            $ok = $true
        }
    }

    if ($step.Action -ne "WaitWin" -and $step.Action -ne "Wait") {
        $waitMs = 0; if ($step.Wait) { $null = [int]::TryParse($step.Wait, [ref]$waitMs) }
        if ($waitMs -gt 0) { Start-Sleep -Milliseconds $waitMs }
    }
    Write-Host ""
}

try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($WsShell) | Out-Null } catch {}

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Test Run Complete" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Press any key to close..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
'@

    [System.IO.File]::WriteAllText($testScript, $testContent, [System.Text.Encoding]::Default)

    # Launch in separate PowerShell window
    Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Unrestricted -File `"$testScript`""
})

# ========================================
# ========================================
# Fix Dock Layout Order
# ========================================
# WinForms docks controls from highest child index first.
# DataGridView (Fill) must have the LOWEST index so it is
# docked LAST and receives the remaining space.
# Without this fix, the Fill control is docked first and
# the Top/Bottom panels overlap its first/last rows.
$footerPanel.BringToFront()
$rowPanel.BringToFront()
$toolPanel.BringToFront()
$namePanel.BringToFront()
$script:dgv.BringToFront()

# ========================================
# Initial state: load template recipe or add empty row
# ========================================
$defaultRecipe = Join-Path $script:TemplatePath "recipe.csv"
if (Test-Path $defaultRecipe) {
    Import-RecipeCsv -Path $defaultRecipe
    $form.Text = "AutoKey Recipe Editor - (Template loaded)"
}
else {
    $idx = $script:dgv.Rows.Add()
    $script:dgv.Rows[$idx].Cells["Wait"].Value = "0"
    Update-StepNumbers
}

# ========================================
# Show Form
# ========================================
$null = $form.ShowDialog()
$form.Dispose()
