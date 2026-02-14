# ========================================
# Digital Gyotaq Editor
# ========================================
# Visual editor for Digital Gyotaq task_list.csv.
# Create evidence capture task lists with
# command selection from material/ shortcuts
# and URI definitions.
# ========================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

# ========================================
# Constants
# ========================================
$script:FabriqRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -ErrorAction SilentlyContinue
if (-not $script:FabriqRoot) { $script:FabriqRoot = (Resolve-Path "..\..\").Path }
$script:MaterialDir = Join-Path $PSScriptRoot "material"
$script:TemplateDir = Join-Path $PSScriptRoot "gyotaq_template"
$script:ModulesDir = Join-Path $script:FabriqRoot "modules"
$script:CurrentFilePath = $null
$script:IsLoading = $false
$script:DragRowIndex = -1

# ========================================
# Color Scheme (same as Profile Editor)
# ========================================
$bgDark       = [System.Drawing.Color]::FromArgb(30, 30, 30)
$bgPanel      = [System.Drawing.Color]::FromArgb(40, 40, 40)
$bgGrid       = [System.Drawing.Color]::FromArgb(35, 35, 35)
$bgCell       = [System.Drawing.Color]::FromArgb(45, 45, 45)
$bgHeader     = [System.Drawing.Color]::FromArgb(55, 55, 55)
$bgButton     = [System.Drawing.Color]::FromArgb(60, 60, 60)
$bgButtonHov  = [System.Drawing.Color]::FromArgb(80, 80, 80)
$bgExport     = [System.Drawing.Color]::FromArgb(0, 120, 215)
$bgSelect     = [System.Drawing.Color]::FromArgb(0, 140, 100)
$fgText       = [System.Drawing.Color]::FromArgb(220, 220, 220)
$fgDim        = [System.Drawing.Color]::FromArgb(150, 150, 150)
$fgHeader     = [System.Drawing.Color]::FromArgb(100, 180, 255)
$fgFooter     = [System.Drawing.Color]::FromArgb(170, 170, 170)
$gridLine     = [System.Drawing.Color]::FromArgb(60, 60, 60)

# ========================================
# Load URI List from material/uri_list.csv
# ========================================
$script:UriCommands = @()
$uriListPath = Join-Path $script:MaterialDir "uri_list.csv"
if (Test-Path $uriListPath) {
    try {
        $script:UriCommands = @(Import-Csv -Path $uriListPath -Encoding Default | Where-Object {
            -not [string]::IsNullOrWhiteSpace($_.Name)
        })
    }
    catch { }
}

# ========================================
# Scan material/ for .lnk shortcuts
# ========================================
$script:ShortcutCommands = @()

function Get-ShortcutInfo {
    param([string]$LnkPath)
    try {
        $wsShell = New-Object -ComObject WScript.Shell
        $shortcut = $wsShell.CreateShortcut($LnkPath)
        $result = [PSCustomObject]@{
            TargetPath = $shortcut.TargetPath
            Arguments  = $shortcut.Arguments
            Name       = [System.IO.Path]::GetFileNameWithoutExtension((Split-Path $LnkPath -Leaf)) -replace ' - ショートカット$', ''
        }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wsShell) | Out-Null
        return $result
    }
    catch { return $null }
}

if (Test-Path $script:MaterialDir) {
    $lnkFiles = @(Get-ChildItem -Path $script:MaterialDir -Filter "*.lnk" -Recurse -File -ErrorAction SilentlyContinue)
    foreach ($lnk in $lnkFiles) {
        $info = Get-ShortcutInfo -LnkPath $lnk.FullName
        if ($null -eq $info) { continue }

        # Category = subdirectory name relative to material/
        $relDir = $lnk.DirectoryName.Substring($script:MaterialDir.Length).TrimStart("\")
        $category = if ([string]::IsNullOrEmpty($relDir)) { "General" } else { $relDir }

        # Use .lnk file path directly as OpenCommand (works with Start-Process on Windows)
        # This avoids issues where system shortcuts have empty/complex TargetPath
        $script:ShortcutCommands += [PSCustomObject]@{
            Category         = $category
            Name             = $info.Name
            OpenCommand      = $lnk.FullName
            OpenArgs         = ""
            DefaultTitle     = $info.Name
            DefaultInstruction = "$($info.Name)の設定を確認してください"
            Source           = "Shortcut"
            LnkPath          = $lnk.FullName
        }
    }
}

# ========================================
# Build combined command list
# ========================================
$script:AllCommands = @()

foreach ($u in $script:UriCommands) {
    $script:AllCommands += [PSCustomObject]@{
        Category         = $u.Category
        Name             = $u.Name
        OpenCommand      = $u.OpenCommand
        OpenArgs         = $u.OpenArgs
        DefaultTitle     = $u.DefaultTitle
        DefaultInstruction = $u.DefaultInstruction
        Source           = "URI"
        LnkPath          = ""
    }
}
foreach ($s in $script:ShortcutCommands) {
    $script:AllCommands += $s
}

# ========================================
# Helper: Auto-generate TaskID
# ========================================
function Get-NextTaskID {
    $maxNum = 0
    for ($i = 0; $i -lt $script:dgv.Rows.Count; $i++) {
        $row = $script:dgv.Rows[$i]
        if ($row.IsNewRow) { continue }
        $tid = $row.Cells["TaskID"].Value
        if ($tid -match '^T(\d+)$') {
            $num = [int]$matches[1]
            if ($num -gt $maxNum) { $maxNum = $num }
        }
    }
    return "T" + ($maxNum + 1).ToString("000")
}

# ========================================
# Helper: Renumber TaskIDs
# ========================================
function Update-TaskIDs {
    $num = 1
    for ($i = 0; $i -lt $script:dgv.Rows.Count; $i++) {
        if (-not $script:dgv.Rows[$i].IsNewRow) {
            $script:dgv.Rows[$i].Cells["TaskID"].Value = "T" + $num.ToString("000")
            $num++
        }
    }
}

# ========================================
# Helper: Grid data to CSV string
# ========================================
function Get-GridAsCsv {
    $lines = @("Enabled,TaskID,TaskTitle,Instruction,OpenCommand,OpenArgs")
    for ($i = 0; $i -lt $script:dgv.Rows.Count; $i++) {
        $row = $script:dgv.Rows[$i]
        if ($row.IsNewRow) { continue }

        $enabled = if ($row.Cells["Enabled"].Value) { $row.Cells["Enabled"].Value } else { "1" }
        $taskId  = if ($row.Cells["TaskID"].Value) { $row.Cells["TaskID"].Value } else { Get-NextTaskID }
        $title   = if ($row.Cells["TaskTitle"].Value) { $row.Cells["TaskTitle"].Value } else { "" }
        $instr   = if ($row.Cells["Instruction"].Value) { $row.Cells["Instruction"].Value } else { "" }
        $cmd     = if ($row.Cells["OpenCommand"].Value) { $row.Cells["OpenCommand"].Value } else { "" }
        $args    = if ($row.Cells["OpenArgs"].Value) { $row.Cells["OpenArgs"].Value } else { "" }

        # Escape CSV fields
        foreach ($fieldName in @('title', 'instr', 'cmd', 'args')) {
            $val = (Get-Variable $fieldName).Value
            $val = $val -replace '"', '""'
            if ($val -match '[,"]') { $val = "`"$val`"" }
            Set-Variable $fieldName $val
        }

        $lines += "$enabled,$taskId,$title,$instr,$cmd,$args"
    }
    return ($lines -join "`r`n") + "`r`n"
}

# ========================================
# Helper: Load CSV into grid
# ========================================
function Import-TaskCsv {
    param([string]$Path)
    try {
        $data = @(Import-Csv -Path $Path -Encoding Default)
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to load CSV: $_", "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return
    }

    $script:IsLoading = $true
    $script:dgv.Rows.Clear()

    foreach ($item in $data) {
        $idx = $script:dgv.Rows.Add()
        $row = $script:dgv.Rows[$idx]

        $row.Cells["Enabled"].Value     = if ($item.Enabled) { $item.Enabled } else { "1" }
        $row.Cells["TaskID"].Value      = $item.TaskID
        $row.Cells["TaskTitle"].Value   = $item.TaskTitle
        $row.Cells["Instruction"].Value = $item.Instruction
        $row.Cells["OpenCommand"].Value = $item.OpenCommand
        $row.Cells["OpenArgs"].Value    = $item.OpenArgs
    }

    $script:IsLoading = $false

    if ($script:dgv.Rows.Count -gt 0) {
        $script:dgv.ClearSelection()
        $script:dgv.FirstDisplayedScrollingRowIndex = 0
        $script:dgv.Rows[0].Selected = $true
        $script:dgv.CurrentCell = $script:dgv.Rows[0].Cells["TaskTitle"]
    }
}

# ========================================
# Command Selection Dialog
# ========================================
function Show-CommandSelector {
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Select Command"
    $dlg.Size = New-Object System.Drawing.Size(720, 560)
    $dlg.MinimumSize = New-Object System.Drawing.Size(600, 400)
    $dlg.StartPosition = "CenterParent"
    $dlg.BackColor = $bgDark
    $dlg.ForeColor = $fgText
    $dlg.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false

    # --- Search Panel ---
    $searchPanel = New-Object System.Windows.Forms.Panel
    $searchPanel.Dock = "Top"
    $searchPanel.Height = 38
    $searchPanel.BackColor = $bgPanel
    $searchPanel.Padding = New-Object System.Windows.Forms.Padding(8, 6, 8, 6)

    $lblSearch = New-Object System.Windows.Forms.Label
    $lblSearch.Text = "Search:"
    $lblSearch.Location = New-Object System.Drawing.Point(10, 9)
    $lblSearch.AutoSize = $true
    $lblSearch.ForeColor = $fgDim

    $txtSearch = New-Object System.Windows.Forms.TextBox
    $txtSearch.Location = New-Object System.Drawing.Point(70, 6)
    $txtSearch.Size = New-Object System.Drawing.Size(300, 24)
    $txtSearch.BackColor = $bgCell
    $txtSearch.ForeColor = $fgText
    $txtSearch.BorderStyle = "FixedSingle"

    # Category filter
    $lblCat = New-Object System.Windows.Forms.Label
    $lblCat.Text = "Category:"
    $lblCat.Location = New-Object System.Drawing.Point(390, 9)
    $lblCat.AutoSize = $true
    $lblCat.ForeColor = $fgDim

    $cmbCategory = New-Object System.Windows.Forms.ComboBox
    $cmbCategory.Location = New-Object System.Drawing.Point(460, 6)
    $cmbCategory.Size = New-Object System.Drawing.Size(220, 24)
    $cmbCategory.BackColor = $bgCell
    $cmbCategory.ForeColor = $fgText
    $cmbCategory.DropDownStyle = "DropDownList"
    $cmbCategory.FlatStyle = "Flat"
    $null = $cmbCategory.Items.Add("(All)")
    $categories = @($script:AllCommands | ForEach-Object { $_.Category } | Sort-Object -Unique)
    foreach ($cat in $categories) { $null = $cmbCategory.Items.Add($cat) }
    $cmbCategory.SelectedIndex = 0

    $searchPanel.Controls.AddRange(@($lblSearch, $txtSearch, $lblCat, $cmbCategory))
    $dlg.Controls.Add($searchPanel)

    # --- Button Panel ---
    $btnPanel = New-Object System.Windows.Forms.Panel
    $btnPanel.Dock = "Bottom"
    $btnPanel.Height = 45
    $btnPanel.BackColor = $bgPanel

    $btnPreview = New-Object System.Windows.Forms.Button
    $btnPreview.Text = "Open"
    $btnPreview.Size = New-Object System.Drawing.Size(80, 30)
    $btnPreview.Location = New-Object System.Drawing.Point(10, 7)
    $btnPreview.FlatStyle = "Flat"
    $btnPreview.FlatAppearance.BorderColor = $gridLine
    $btnPreview.FlatAppearance.MouseOverBackColor = $bgButtonHov
    $btnPreview.BackColor = [System.Drawing.Color]::FromArgb(80, 60, 20)
    $btnPreview.ForeColor = $fgText
    $btnPreview.Cursor = [System.Windows.Forms.Cursors]::Hand

    $btnPreview.Add_Click({
        if ($lv.SelectedItems.Count -eq 0) { return }
        $cmd = $lv.SelectedItems[0].Tag

        # For shortcuts, launch .lnk file directly (handles system shortcuts with empty TargetPath)
        $launchPath = if ($cmd.Source -eq "Shortcut" -and -not [string]::IsNullOrWhiteSpace($cmd.LnkPath)) {
            $cmd.LnkPath
        } else {
            $cmd.OpenCommand
        }

        if ([string]::IsNullOrWhiteSpace($launchPath)) { return }
        try {
            if ($cmd.Source -ne "Shortcut" -and -not [string]::IsNullOrWhiteSpace($cmd.OpenArgs)) {
                Start-Process -FilePath $launchPath -ArgumentList $cmd.OpenArgs
            }
            else {
                Start-Process -FilePath $launchPath
            }
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Failed to open: $launchPath`n$_", "Open Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
        }
    })

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "Select"
    $btnOK.Size = New-Object System.Drawing.Size(100, 30)
    $btnOK.Location = New-Object System.Drawing.Point(480, 7)
    $btnOK.FlatStyle = "Flat"
    $btnOK.FlatAppearance.BorderColor = $gridLine
    $btnOK.BackColor = $bgSelect
    $btnOK.ForeColor = $fgText
    $btnOK.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Size = New-Object System.Drawing.Size(100, 30)
    $btnCancel.Location = New-Object System.Drawing.Point(590, 7)
    $btnCancel.FlatStyle = "Flat"
    $btnCancel.FlatAppearance.BorderColor = $gridLine
    $btnCancel.BackColor = $bgButton
    $btnCancel.ForeColor = $fgText
    $btnCancel.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    $btnPanel.Controls.AddRange(@($btnPreview, $btnOK, $btnCancel))
    $dlg.Controls.Add($btnPanel)
    $dlg.AcceptButton = $btnOK
    $dlg.CancelButton = $btnCancel

    # --- ListView ---
    $lv = New-Object System.Windows.Forms.ListView
    $lv.Dock = "Fill"
    $lv.View = "Details"
    $lv.FullRowSelect = $true
    $lv.HideSelection = $false
    $lv.MultiSelect = $false
    $lv.BackColor = $bgGrid
    $lv.ForeColor = $fgText
    $lv.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lv.BorderStyle = "None"
    $lv.HeaderStyle = "Nonclickable"

    $null = $lv.Columns.Add("Category", 130)
    $null = $lv.Columns.Add("Name", 220)
    $null = $lv.Columns.Add("Command", 200)
    $null = $lv.Columns.Add("Source", 60)

    # DoubleBuffered
    $lvType = $lv.GetType()
    $lvProp = $lvType.GetProperty("DoubleBuffered", [System.Reflection.BindingFlags]"Instance,NonPublic")
    $lvProp.SetValue($lv, $true, $null)

    $dlg.Controls.Add($lv)

    # --- Populate function ---
    $populateList = {
        $lv.Items.Clear()
        $searchText = $txtSearch.Text.Trim()
        $filterCat = $cmbCategory.SelectedItem
        if ($filterCat -eq "(All)") { $filterCat = "" }

        foreach ($cmd in $script:AllCommands) {
            # Category filter
            if ($filterCat -and $cmd.Category -ne $filterCat) { continue }

            # Search filter
            if ($searchText) {
                $match = ($cmd.Name -like "*$searchText*") -or
                         ($cmd.Category -like "*$searchText*") -or
                         ($cmd.OpenCommand -like "*$searchText*") -or
                         ($cmd.DefaultTitle -like "*$searchText*")
                if (-not $match) { continue }
            }

            $item = New-Object System.Windows.Forms.ListViewItem($cmd.Category)
            $null = $item.SubItems.Add($cmd.Name)
            $cmdDisplay = $cmd.OpenCommand
            if ($cmd.OpenArgs) { $cmdDisplay += " $($cmd.OpenArgs)" }
            $null = $item.SubItems.Add($cmdDisplay)
            $null = $item.SubItems.Add($cmd.Source)
            $item.Tag = $cmd
            $null = $lv.Items.Add($item)
        }

        if ($lv.Items.Count -gt 0) {
            $lv.Items[0].Selected = $true
            $lv.Items[0].Focused = $true
        }
    }

    # Initial populate
    & $populateList

    # Events
    $txtSearch.Add_TextChanged({ & $populateList })
    $cmbCategory.Add_SelectedIndexChanged({ & $populateList })

    # Double-click to select
    $lv.Add_DoubleClick({
        if ($lv.SelectedItems.Count -gt 0) {
            $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $dlg.Close()
        }
    })

    # Fix dock order
    $btnPanel.BringToFront()
    $searchPanel.BringToFront()
    $lv.BringToFront()

    # Show
    $selectedCmd = $null
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        if ($lv.SelectedItems.Count -gt 0) {
            $selectedCmd = $lv.SelectedItems[0].Tag
        }
    }
    $dlg.Dispose()
    return $selectedCmd
}

# ========================================
# Main Form
# ========================================
$form = New-Object System.Windows.Forms.Form
$form.Text = "Digital Gyotaq Editor"
$form.Size = New-Object System.Drawing.Size(1100, 700)
$form.MinimumSize = New-Object System.Drawing.Size(850, 500)
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

$btnNew    = New-ToolButton -Text "New"          -X 8   -Width 60
$btnOpen   = New-ToolButton -Text "Open"         -X 75  -Width 65
$btnSave   = New-ToolButton -Text "Save"         -X 147 -Width 60
$btnSaveAs = New-ToolButton -Text "Save As"      -X 214 -Width 75
$btnExport = New-ToolButton -Text "Export Module" -X 330 -Width 110 -BgColor $bgExport

$toolPanel.Controls.AddRange(@($btnNew, $btnOpen, $btnSave, $btnSaveAs, $btnExport))
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
$script:txtModuleName.Size = New-Object System.Drawing.Size(350, 24)
$script:txtModuleName.BackColor = $bgCell
$script:txtModuleName.ForeColor = $fgText
$script:txtModuleName.BorderStyle = "FixedSingle"
$script:txtModuleName.Font = New-Object System.Drawing.Font("Segoe UI", 10)

$namePanel.Controls.AddRange(@($nameLabel, $script:txtModuleName))
$form.Controls.Add($namePanel)

# ========================================
# Footer Panel
# ========================================
$footerPanel = New-Object System.Windows.Forms.Panel
$footerPanel.Dock = "Bottom"
$footerPanel.Height = 65
$footerPanel.BackColor = [System.Drawing.Color]::FromArgb(25, 25, 25)
$footerPanel.Padding = New-Object System.Windows.Forms.Padding(12, 6, 12, 6)

$footerText = @"
Digital Gyotaq task_list.csv 解説:
  Enabled: 1=有効 0=無効   TaskID: タスク識別子 (自動連番)   TaskTitle: タスク名 (ファイル名にも使用)
  Instruction: 作業者への指示   OpenCommand: 自動起動コマンド   OpenArgs: コマンド引数
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
    param([string]$Text, [int]$X, [int]$Width = 95, $BgColor = $bgButton)
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = $Text
    $btn.Location = New-Object System.Drawing.Point($X, 4)
    $btn.Size = New-Object System.Drawing.Size($Width, 28)
    $btn.FlatStyle = "Flat"
    $btn.FlatAppearance.BorderColor = $gridLine
    $btn.FlatAppearance.MouseOverBackColor = $bgButtonHov
    $btn.BackColor = $BgColor
    $btn.ForeColor = $fgText
    $btn.Cursor = [System.Windows.Forms.Cursors]::Hand
    return $btn
}

$btnAddRow   = New-RowButton -Text "+ Add Task"      -X 8   -Width 100
$btnSelect   = New-RowButton -Text "Select Command"  -X 115 -Width 130 -BgColor $bgSelect
$btnInsert   = New-RowButton -Text "Insert Above"    -X 252 -Width 105
$btnDelete   = New-RowButton -Text "Delete Row"      -X 364 -Width 95
$btnMoveUp   = New-RowButton -Text "Move Up"         -X 482 -Width 80
$btnMoveDown = New-RowButton -Text "Move Down"       -X 569 -Width 85

$rowPanel.Controls.AddRange(@($btnAddRow, $btnSelect, $btnInsert, $btnDelete, $btnMoveUp, $btnMoveDown))
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

# DoubleBuffered
$dgvType = $script:dgv.GetType()
$pi = $dgvType.GetProperty("DoubleBuffered", [System.Reflection.BindingFlags]"Instance,NonPublic")
$pi.SetValue($script:dgv, $true, $null)

# --- Columns ---

# Enabled (ComboBox: 1/0)
$colEnabled = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
$colEnabled.Name = "Enabled"
$colEnabled.HeaderText = "Enabled"
$colEnabled.FillWeight = 6
$colEnabled.MinimumWidth = 60
$colEnabled.FlatStyle = "Flat"
$colEnabled.SortMode = "NotSortable"
$colEnabled.DefaultCellStyle.Alignment = "MiddleCenter"
$null = $colEnabled.Items.Add("1")
$null = $colEnabled.Items.Add("0")
$null = $script:dgv.Columns.Add($colEnabled)

# TaskID (auto, read-only)
$colTaskID = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colTaskID.Name = "TaskID"
$colTaskID.HeaderText = "TaskID"
$colTaskID.FillWeight = 6
$colTaskID.MinimumWidth = 55
$colTaskID.ReadOnly = $true
$colTaskID.DefaultCellStyle.Alignment = "MiddleCenter"
$colTaskID.DefaultCellStyle.ForeColor = $fgDim
$colTaskID.SortMode = "NotSortable"
$null = $script:dgv.Columns.Add($colTaskID)

# TaskTitle
$colTitle = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colTitle.Name = "TaskTitle"
$colTitle.HeaderText = "TaskTitle (タスク名)"
$colTitle.FillWeight = 20
$colTitle.MinimumWidth = 120
$colTitle.SortMode = "NotSortable"
$null = $script:dgv.Columns.Add($colTitle)

# Instruction
$colInstr = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colInstr.Name = "Instruction"
$colInstr.HeaderText = "Instruction (作業指示)"
$colInstr.FillWeight = 28
$colInstr.MinimumWidth = 150
$colInstr.SortMode = "NotSortable"
$null = $script:dgv.Columns.Add($colInstr)

# OpenCommand
$colCmd = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colCmd.Name = "OpenCommand"
$colCmd.HeaderText = "OpenCommand"
$colCmd.FillWeight = 20
$colCmd.MinimumWidth = 100
$colCmd.SortMode = "NotSortable"
$null = $script:dgv.Columns.Add($colCmd)

# OpenArgs
$colArgs = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colArgs.Name = "OpenArgs"
$colArgs.HeaderText = "OpenArgs"
$colArgs.FillWeight = 20
$colArgs.MinimumWidth = 80
$colArgs.SortMode = "NotSortable"
$null = $script:dgv.Columns.Add($colArgs)

$form.Controls.Add($script:dgv)

# ========================================
# Event: Commit ComboBox edits immediately
# ========================================
$script:dgv.Add_CurrentCellDirtyStateChanged({
    param($sender, $e)
    if ($script:IsLoading) { return }
    if ($sender.IsCurrentCellDirty) {
        $null = $sender.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
    }
})

# Suppress DataError
$script:dgv.Add_DataError({
    param($sender, $e)
    $e.ThrowException = $false
})

# ========================================
# Event: Drag & Drop
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
    $null = $sender.DoDragDrop($script:DragRowIndex, [System.Windows.Forms.DragDropEffects]::Move)
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

    $srcRow = $sender.Rows[$script:DragRowIndex]
    $data = @{
        Enabled     = $srcRow.Cells["Enabled"].Value
        TaskTitle   = $srcRow.Cells["TaskTitle"].Value
        Instruction = $srcRow.Cells["Instruction"].Value
        OpenCommand = $srcRow.Cells["OpenCommand"].Value
        OpenArgs    = $srcRow.Cells["OpenArgs"].Value
    }
    $sender.Rows.RemoveAt($script:DragRowIndex)
    if ($script:DragRowIndex -lt $targetIndex) { $targetIndex-- }

    $script:IsLoading = $true
    $null = $sender.Rows.Insert($targetIndex, 1)
    $newRow = $sender.Rows[$targetIndex]
    $newRow.Cells["Enabled"].Value     = $data.Enabled
    $newRow.Cells["TaskTitle"].Value   = $data.TaskTitle
    $newRow.Cells["Instruction"].Value = $data.Instruction
    $newRow.Cells["OpenCommand"].Value = $data.OpenCommand
    $newRow.Cells["OpenArgs"].Value    = $data.OpenArgs
    $script:IsLoading = $false

    $sender.ClearSelection()
    $sender.Rows[$targetIndex].Selected = $true
    $sender.CurrentCell = $sender.Rows[$targetIndex].Cells["TaskTitle"]

    Update-TaskIDs
    $script:DragRowIndex = -1
})

# ========================================
# Helper: Add a row with data
# ========================================
function Add-TaskRow {
    param(
        [string]$Enabled = "1",
        [string]$TaskTitle = "",
        [string]$Instruction = "",
        [string]$OpenCommand = "",
        [string]$OpenArgs = "",
        [int]$InsertAt = -1
    )
    $script:IsLoading = $true
    if ($InsertAt -ge 0) {
        $null = $script:dgv.Rows.Insert($InsertAt, 1)
        $idx = $InsertAt
    }
    else {
        $idx = $script:dgv.Rows.Add()
    }
    $row = $script:dgv.Rows[$idx]
    $row.Cells["Enabled"].Value     = $Enabled
    $row.Cells["TaskTitle"].Value   = $TaskTitle
    $row.Cells["Instruction"].Value = $Instruction
    $row.Cells["OpenCommand"].Value = $OpenCommand
    $row.Cells["OpenArgs"].Value    = $OpenArgs
    $script:IsLoading = $false

    Update-TaskIDs

    $script:dgv.ClearSelection()
    $script:dgv.Rows[$idx].Selected = $true
    $script:dgv.CurrentCell = $script:dgv.Rows[$idx].Cells["TaskTitle"]
}

# ========================================
# Button: + Add Task
# ========================================
$btnAddRow.Add_Click({
    Add-TaskRow -Enabled "1"
})

# ========================================
# Button: Select Command
# ========================================
$btnSelect.Add_Click({
    $selected = Show-CommandSelector
    if ($null -eq $selected) { return }

    # If current row is empty, fill it; otherwise add new row
    $currentRow = $script:dgv.CurrentRow
    $isEmptyRow = $false
    if ($currentRow -and -not $currentRow.IsNewRow) {
        $title = $currentRow.Cells["TaskTitle"].Value
        $cmd   = $currentRow.Cells["OpenCommand"].Value
        if ([string]::IsNullOrWhiteSpace($title) -and [string]::IsNullOrWhiteSpace($cmd)) {
            $isEmptyRow = $true
        }
    }

    if ($isEmptyRow) {
        $currentRow.Cells["TaskTitle"].Value   = $selected.DefaultTitle
        $currentRow.Cells["Instruction"].Value = $selected.DefaultInstruction
        $currentRow.Cells["OpenCommand"].Value = $selected.OpenCommand
        $currentRow.Cells["OpenArgs"].Value    = $selected.OpenArgs
    }
    else {
        Add-TaskRow -Enabled "1" `
            -TaskTitle $selected.DefaultTitle `
            -Instruction $selected.DefaultInstruction `
            -OpenCommand $selected.OpenCommand `
            -OpenArgs $selected.OpenArgs
    }
})

# ========================================
# Button: Insert Above
# ========================================
$btnInsert.Add_Click({
    $selIdx = if ($script:dgv.CurrentRow) { $script:dgv.CurrentRow.Index } else { 0 }
    Add-TaskRow -Enabled "1" -InsertAt $selIdx
})

# ========================================
# Button: Delete Row
# ========================================
$btnDelete.Add_Click({
    if ($script:dgv.CurrentRow -and -not $script:dgv.CurrentRow.IsNewRow) {
        $script:dgv.Rows.RemoveAt($script:dgv.CurrentRow.Index)
        Update-TaskIDs
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
        Enabled     = $row.Cells["Enabled"].Value
        TaskTitle   = $row.Cells["TaskTitle"].Value
        Instruction = $row.Cells["Instruction"].Value
        OpenCommand = $row.Cells["OpenCommand"].Value
        OpenArgs    = $row.Cells["OpenArgs"].Value
    }
    $script:dgv.Rows.RemoveAt($idx)
    $newIdx = $idx - 1

    $script:IsLoading = $true
    $null = $script:dgv.Rows.Insert($newIdx, 1)
    $newRow = $script:dgv.Rows[$newIdx]
    $newRow.Cells["Enabled"].Value     = $data.Enabled
    $newRow.Cells["TaskTitle"].Value   = $data.TaskTitle
    $newRow.Cells["Instruction"].Value = $data.Instruction
    $newRow.Cells["OpenCommand"].Value = $data.OpenCommand
    $newRow.Cells["OpenArgs"].Value    = $data.OpenArgs
    $script:IsLoading = $false

    $script:dgv.ClearSelection()
    $script:dgv.Rows[$newIdx].Selected = $true
    $script:dgv.CurrentCell = $script:dgv.Rows[$newIdx].Cells["TaskTitle"]
    Update-TaskIDs
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
        Enabled     = $row.Cells["Enabled"].Value
        TaskTitle   = $row.Cells["TaskTitle"].Value
        Instruction = $row.Cells["Instruction"].Value
        OpenCommand = $row.Cells["OpenCommand"].Value
        OpenArgs    = $row.Cells["OpenArgs"].Value
    }
    $script:dgv.Rows.RemoveAt($idx)
    $newIdx = $idx + 1
    if ($newIdx -gt $script:dgv.Rows.Count) { $newIdx = $script:dgv.Rows.Count }

    $script:IsLoading = $true
    $null = $script:dgv.Rows.Insert($newIdx, 1)
    $newRow = $script:dgv.Rows[$newIdx]
    $newRow.Cells["Enabled"].Value     = $data.Enabled
    $newRow.Cells["TaskTitle"].Value   = $data.TaskTitle
    $newRow.Cells["Instruction"].Value = $data.Instruction
    $newRow.Cells["OpenCommand"].Value = $data.OpenCommand
    $newRow.Cells["OpenArgs"].Value    = $data.OpenArgs
    $script:IsLoading = $false

    $script:dgv.ClearSelection()
    $script:dgv.Rows[$newIdx].Selected = $true
    $script:dgv.CurrentCell = $script:dgv.Rows[$newIdx].Cells["TaskTitle"]
    Update-TaskIDs
})

# ========================================
# Button: New
# ========================================
$btnNew.Add_Click({
    $result = [System.Windows.Forms.MessageBox]::Show(
        "Clear current task list and start new?",
        "New Task List",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    if ($result -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $script:dgv.Rows.Clear()
    $script:txtModuleName.Text = ""
    $script:CurrentFilePath = $null
    $form.Text = "Digital Gyotaq Editor"

    Add-TaskRow -Enabled "1"
})

# ========================================
# Button: Open
# ========================================
$btnOpen.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Title = "Open Task List CSV"
    $dlg.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $dlg.InitialDirectory = $PSScriptRoot

    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        Import-TaskCsv -Path $dlg.FileName
        $script:CurrentFilePath = $dlg.FileName
        $dirName = Split-Path (Split-Path $dlg.FileName -Parent) -Leaf
        $script:txtModuleName.Text = $dirName
        $form.Text = "Digital Gyotaq Editor - $dirName"
    }
    $dlg.Dispose()
})

# ========================================
# Helper: Save to file
# ========================================
function Save-TaskListToFile {
    param([string]$FilePath)
    $csvContent = Get-GridAsCsv
    try {
        [System.IO.File]::WriteAllText($FilePath, $csvContent, [System.Text.Encoding]::UTF8)
        $script:CurrentFilePath = $FilePath
        $form.Text = "Digital Gyotaq Editor - $(Split-Path (Split-Path $FilePath -Parent) -Leaf)"
        [System.Windows.Forms.MessageBox]::Show(
            "Saved: $FilePath", "Save",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to save: $_", "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
}

# ========================================
# Button: Save
# ========================================
$btnSave.Add_Click({
    if ($script:CurrentFilePath) {
        Save-TaskListToFile -FilePath $script:CurrentFilePath
    }
    else {
        # No file yet - use Save As logic
        $dlg = New-Object System.Windows.Forms.SaveFileDialog
        $dlg.Title = "Save Task List CSV"
        $dlg.Filter = "CSV Files (*.csv)|*.csv"
        $dlg.FileName = "task_list.csv"
        $dlg.InitialDirectory = $PSScriptRoot

        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            Save-TaskListToFile -FilePath $dlg.FileName
        }
        $dlg.Dispose()
    }
})

# ========================================
# Button: Save As
# ========================================
$btnSaveAs.Add_Click({
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Title = "Save Task List CSV"
    $dlg.Filter = "CSV Files (*.csv)|*.csv"
    $dlg.FileName = "task_list.csv"
    if ($script:CurrentFilePath) {
        $dlg.InitialDirectory = Split-Path $script:CurrentFilePath -Parent
    }
    else {
        $dlg.InitialDirectory = $PSScriptRoot
    }

    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        Save-TaskListToFile -FilePath $dlg.FileName
    }
    $dlg.Dispose()
})

# ========================================
# Button: Export Module
# ========================================
$btnExport.Add_Click({
    $moduleName = $script:txtModuleName.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($moduleName)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please enter a Module Name.", "Export Module",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        $script:txtModuleName.Focus()
        return
    }

    if ($script:dgv.Rows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "No tasks to export.", "Export Module",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    # Sanitize directory name
    $safeName = $moduleName -replace '[\\/:*?"<>|]', '_'

    # Target directory
    $targetDir = Join-Path $script:ModulesDir "extended\$safeName"

    if (Test-Path $targetDir) {
        $ow = [System.Windows.Forms.MessageBox]::Show(
            "Module '$safeName' already exists at:`n$targetDir`n`nOverwrite?",
            "Export Module",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        if ($ow -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    }

    try {
        # Create directory
        if (-not (Test-Path $targetDir)) {
            $null = New-Item -ItemType Directory -Path $targetDir -Force
        }

        # Copy gyotaq_config.ps1 from template
        $templateScript = Join-Path $script:TemplateDir "gyotaq_config.ps1"
        if (Test-Path $templateScript) {
            Copy-Item -Path $templateScript -Destination (Join-Path $targetDir "gyotaq_config.ps1") -Force
        }
        else {
            # Fallback: try gyotaku_template from modules
            $fallbackScript = Join-Path $script:ModulesDir "standard\gyotaku_template\gyotaku_config.ps1"
            if (Test-Path $fallbackScript) {
                Copy-Item -Path $fallbackScript -Destination (Join-Path $targetDir "gyotaq_config.ps1") -Force
            }
        }

        # Copy README if exists
        $templateReadme = Join-Path $script:TemplateDir "README.txt"
        if (Test-Path $templateReadme) {
            Copy-Item -Path $templateReadme -Destination (Join-Path $targetDir "README.txt") -Force
        }

        # Generate module.csv
        $moduleCsvContent = "MenuName,Category,Script,Order,Enabled`r`n$moduleName,ManualWorks,gyotaq_config.ps1,95,0`r`n"
        [System.IO.File]::WriteAllText(
            (Join-Path $targetDir "module.csv"),
            $moduleCsvContent,
            [System.Text.Encoding]::UTF8
        )

        # Generate task_list.csv from grid
        $csvContent = Get-GridAsCsv
        $taskListPath = Join-Path $targetDir "task_list.csv"
        [System.IO.File]::WriteAllText($taskListPath, $csvContent, [System.Text.Encoding]::UTF8)

        $script:CurrentFilePath = $taskListPath
        $form.Text = "Digital Gyotaq Editor - $moduleName"

        [System.Windows.Forms.MessageBox]::Show(
            "Module exported successfully!`n`nLocation: $targetDir`n`nFiles:`n  - module.csv`n  - gyotaq_config.ps1`n  - task_list.csv`n`nNote: Enabled=0 by default. Enable in module.csv or CSV Editor.",
            "Export Module",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Export failed: $_", "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
})

# ========================================
# Fix Dock Layout Order
# ========================================
$footerPanel.BringToFront()
$rowPanel.BringToFront()
$toolPanel.BringToFront()
$namePanel.BringToFront()
$script:dgv.BringToFront()

# ========================================
# Initial state
# ========================================
Add-TaskRow -Enabled "1"

# ========================================
# Show Form
# ========================================
$null = $form.ShowDialog()
$form.Dispose()
