# ========================================
# Profile Editor
# ========================================
# Visual editor for Fabriq profile CSVs.
# Discover all modules and build execution
# profiles with drag & drop ordering.
# ========================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

# ========================================
# Constants
# ========================================
$script:FabriqRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -ErrorAction SilentlyContinue
if (-not $script:FabriqRoot) { $script:FabriqRoot = (Resolve-Path "..\..\").Path }
$script:ModulesDir = Join-Path $script:FabriqRoot "modules"
$script:ProfilesDir = Join-Path $script:FabriqRoot "profiles"
$script:CurrentFilePath = $null
$script:IsLoading = $false
$script:DragRowIndex = -1

# ========================================
# Color Scheme (same as AutoKey editor)
# ========================================
$bgDark       = [System.Drawing.Color]::FromArgb(30, 30, 30)
$bgPanel      = [System.Drawing.Color]::FromArgb(40, 40, 40)
$bgGrid       = [System.Drawing.Color]::FromArgb(35, 35, 35)
$bgCell       = [System.Drawing.Color]::FromArgb(45, 45, 45)
$bgHeader     = [System.Drawing.Color]::FromArgb(55, 55, 55)
$bgButton     = [System.Drawing.Color]::FromArgb(60, 60, 60)
$bgButtonHov  = [System.Drawing.Color]::FromArgb(80, 80, 80)
$bgExport     = [System.Drawing.Color]::FromArgb(0, 120, 215)
$fgText       = [System.Drawing.Color]::FromArgb(220, 220, 220)
$fgDim        = [System.Drawing.Color]::FromArgb(150, 150, 150)
$fgHeader     = [System.Drawing.Color]::FromArgb(100, 180, 255)
$fgFooter     = [System.Drawing.Color]::FromArgb(170, 170, 170)
$gridLine     = [System.Drawing.Color]::FromArgb(60, 60, 60)
$specialBg    = [System.Drawing.Color]::FromArgb(80, 60, 0)
$specialFg    = [System.Drawing.Color]::FromArgb(255, 200, 50)
$restartBg    = $specialBg
$restartFg    = $specialFg
$autoPilotBg  = [System.Drawing.Color]::FromArgb(60, 0, 80)
$autoPilotFg  = [System.Drawing.Color]::FromArgb(220, 130, 255)
$autoLogonBg  = [System.Drawing.Color]::FromArgb(0, 60, 70)
$autoLogonFg  = [System.Drawing.Color]::FromArgb(80, 210, 190)

# Special marker definitions
$script:SpecialMarkers = [ordered]@{
    "--- AUTOPILOT ---"  = @{ Path = "__AUTOPILOT__";  Desc = "WaitSec=3" }
    "--- RESTART ---"    = @{ Path = "__RESTART__";    Desc = "Restart" }
    "--- REEXPLORER ---" = @{ Path = "__REEXPLORER__"; Desc = "Restart Explorer" }
    "--- STOPLOG ---"    = @{ Path = "__STOPLOG__";    Desc = "Stop Transcript" }
    "--- STARTLOG ---"   = @{ Path = "__STARTLOG__";   Desc = "Start Transcript" }
    "--- SHUTDOWN ---"   = @{ Path = "__SHUTDOWN__";   Desc = "Shutdown" }
    "--- PAUSE ---"      = @{ Path = "__PAUSE__";      Desc = "Pause (Enter wait)" }
}
# Reverse lookup: path -> display name
$script:MarkerPathToDisplay = @{}
foreach ($key in $script:SpecialMarkers.Keys) {
    $script:MarkerPathToDisplay[$script:SpecialMarkers[$key].Path] = $key
}

# ========================================
# AutoLogon entries from autologon_list.csv
# ========================================
function Get-AutoLogonEntries {
    $result = [ordered]@{}
    $csvPath = Join-Path $script:FabriqRoot "modules\standard\autologon_config\autologon_list.csv"
    if (-not (Test-Path $csvPath)) { return $result }
    try {
        $rows = @(Import-Csv -Path $csvPath -Encoding Default |
                  Where-Object { $_.Enabled -eq "1" -and -not [string]::IsNullOrWhiteSpace($_.User) })
        foreach ($row in $rows) {
            $displayName = "--- AUTO: $($row.User) ---"
            $result[$displayName] = @{
                Path = "__AUTO_to_$($row.User)__"
                Desc = if ($row.Description) { $row.Description } else { "AutoLogon as $($row.User)" }
            }
        }
    }
    catch { }
    return $result
}

# ========================================
# Discover all available modules
# ========================================
function Get-AllAvailableModules {
    $modules = @()

    foreach ($moduleType in @("standard", "extended")) {
        $basePath = Join-Path $script:ModulesDir $moduleType
        if (-not (Test-Path $basePath)) { continue }

        $dirs = @(Get-ChildItem -Path $basePath -Directory -ErrorAction SilentlyContinue)
        foreach ($dir in $dirs) {
            $moduleCsv = Join-Path $dir.FullName "module.csv"
            if (-not (Test-Path $moduleCsv)) { continue }

            try {
                $entries = @(Import-Csv -Path $moduleCsv -Encoding Default)
            }
            catch { continue }

            foreach ($entry in $entries) {
                if ([string]::IsNullOrWhiteSpace($entry.Script)) { continue }
                $scriptPath = "$moduleType\$($dir.Name)\$($entry.Script)"
                $modules += [PSCustomObject]@{
                    ScriptPath  = $scriptPath
                    MenuName    = $entry.MenuName
                    Category    = $entry.Category
                    ModuleType  = $moduleType
                    DisplayName = "[$($entry.Category)] $($entry.MenuName)"
                }
            }
        }
    }

    return ($modules | Sort-Object Category, MenuName)
}

$script:AllModules = @(Get-AllAvailableModules)

# Build lookup: ScriptPath -> DisplayName
$script:PathToDisplay = @{}
$script:DisplayToPath = @{}
$script:PathToDescription = @{}
foreach ($m in $script:AllModules) {
    $script:PathToDisplay[$m.ScriptPath] = $m.DisplayName
    $script:DisplayToPath[$m.DisplayName] = $m.ScriptPath
    $script:PathToDescription[$m.ScriptPath] = $m.MenuName
}

# Build AutoLogon entries and add to lookup tables
$script:AutoLogonEntries = Get-AutoLogonEntries
foreach ($key in $script:AutoLogonEntries.Keys) {
    $path = $script:AutoLogonEntries[$key].Path
    $script:PathToDisplay[$path] = $key
    $script:DisplayToPath[$key]  = $path
    $script:MarkerPathToDisplay[$path] = $key
}

# Build display list for ComboBox
$script:ModuleDisplayList = @($script:AllModules | ForEach-Object { $_.DisplayName })

# ========================================
# Helper: Renumber Order
# ========================================
function Update-OrderNumbers {
    for ($i = 0; $i -lt $script:dgv.Rows.Count; $i++) {
        if (-not $script:dgv.Rows[$i].IsNewRow) {
            $script:dgv.Rows[$i].Cells["Order"].Value = (($i + 1) * 10).ToString()
        }
    }
}

# ========================================
# Helper: Style restart rows
# ========================================
function Update-RowStyles {
    for ($i = 0; $i -lt $script:dgv.Rows.Count; $i++) {
        $row = $script:dgv.Rows[$i]
        if ($row.IsNewRow) { continue }
        $module = $row.Cells["Module"].Value
        if ($module -eq "--- AUTOPILOT ---") {
            $row.DefaultCellStyle.BackColor = $autoPilotBg
            $row.DefaultCellStyle.ForeColor = $autoPilotFg
        }
        elseif ($script:AutoLogonEntries.Contains($module)) {
            $row.DefaultCellStyle.BackColor = $autoLogonBg
            $row.DefaultCellStyle.ForeColor = $autoLogonFg
        }
        elseif ($script:SpecialMarkers.Contains($module)) {
            $row.DefaultCellStyle.BackColor = $specialBg
            $row.DefaultCellStyle.ForeColor = $specialFg
        }
        else {
            $row.DefaultCellStyle.BackColor = $bgCell
            $row.DefaultCellStyle.ForeColor = $fgText
        }
    }
}

# ========================================
# Helper: Grid data to CSV string
# ========================================
function Get-GridAsCsv {
    $lines = @("Order,ScriptPath,Enabled,Description")
    for ($i = 0; $i -lt $script:dgv.Rows.Count; $i++) {
        $row = $script:dgv.Rows[$i]
        if ($row.IsNewRow) { continue }

        $order   = if ($row.Cells["Order"].Value) { $row.Cells["Order"].Value } else { (($i + 1) * 10).ToString() }
        $module  = $row.Cells["Module"].Value
        $enabled = if ($row.Cells["Enabled"].Value) { $row.Cells["Enabled"].Value } else { "1" }
        $desc    = if ($row.Cells["Description"].Value) { $row.Cells["Description"].Value } else { "" }

        # Resolve display name to script path
        $scriptPath = ""
        if ($script:SpecialMarkers.Contains($module)) {
            $scriptPath = $script:SpecialMarkers[$module].Path
        }
        elseif ($script:DisplayToPath.ContainsKey($module)) {
            $scriptPath = $script:DisplayToPath[$module]
        }
        else {
            $scriptPath = $module
        }

        # Escape CSV
        $desc = $desc -replace '"', '""'
        if ($desc -match '[,"]') { $desc = "`"$desc`"" }
        if ($scriptPath -match '[,"]') { $scriptPath = "`"$scriptPath`"" }

        $lines += "$order,$scriptPath,$enabled,$desc"
    }
    return ($lines -join "`r`n") + "`r`n"
}

# ========================================
# Helper: Load CSV into grid
# ========================================
function Import-ProfileCsv {
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

        $row.Cells["Order"].Value = if ($item.Order) { $item.Order } else { "" }

        # Resolve ScriptPath to display name
        $sp = if ($item.ScriptPath) { $item.ScriptPath.Trim().Replace("/", "\") } else { "" }
        if ($script:MarkerPathToDisplay.ContainsKey($sp)) {
            $row.Cells["Module"].Value = $script:MarkerPathToDisplay[$sp]
        }
        elseif ($script:PathToDisplay.ContainsKey($sp)) {
            $row.Cells["Module"].Value = $script:PathToDisplay[$sp]
        }
        else {
            # Unknown path - show raw
            $row.Cells["Module"].Value = $sp
        }

        $row.Cells["Enabled"].Value = if ($item.Enabled) { $item.Enabled } else { "1" }
        $row.Cells["Description"].Value = $item.Description
    }

    Update-OrderNumbers
    Update-RowStyles
    $script:IsLoading = $false

    if ($script:dgv.Rows.Count -gt 0) {
        $script:dgv.ClearSelection()
        $script:dgv.FirstDisplayedScrollingRowIndex = 0
        $script:dgv.Rows[0].Selected = $true
        $script:dgv.CurrentCell = $script:dgv.Rows[0].Cells["Module"]
    }
}

# ========================================
# Main Form
# ========================================
$form = New-Object System.Windows.Forms.Form
$form.Text = "Profile Editor"
$form.Size = New-Object System.Drawing.Size(960, 700)
$form.MinimumSize = New-Object System.Drawing.Size(750, 500)
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

$btnNew    = New-ToolButton -Text "New"       -X 8   -Width 60
$btnOpen   = New-ToolButton -Text "Open"      -X 75  -Width 65
$btnSave   = New-ToolButton -Text "Save"      -X 147 -Width 60
$btnSaveAs = New-ToolButton -Text "Save As"   -X 214 -Width 75 -BgColor $bgExport

$toolPanel.Controls.AddRange(@($btnNew, $btnOpen, $btnSave, $btnSaveAs))
$form.Controls.Add($toolPanel)

# ========================================
# Profile Name Panel
# ========================================
$namePanel = New-Object System.Windows.Forms.Panel
$namePanel.Dock = "Top"
$namePanel.Height = 35
$namePanel.BackColor = $bgPanel
$namePanel.Padding = New-Object System.Windows.Forms.Padding(8, 4, 8, 4)

$nameLabel = New-Object System.Windows.Forms.Label
$nameLabel.Text = "Profile Name:"
$nameLabel.Location = New-Object System.Drawing.Point(10, 8)
$nameLabel.AutoSize = $true
$nameLabel.ForeColor = $fgDim

$script:txtProfileName = New-Object System.Windows.Forms.TextBox
$script:txtProfileName.Location = New-Object System.Drawing.Point(110, 5)
$script:txtProfileName.Size = New-Object System.Drawing.Size(300, 24)
$script:txtProfileName.BackColor = $bgCell
$script:txtProfileName.ForeColor = $fgText
$script:txtProfileName.BorderStyle = "FixedSingle"
$script:txtProfileName.Font = New-Object System.Drawing.Font("Segoe UI", 10)

$namePanel.Controls.AddRange(@($nameLabel, $script:txtProfileName))
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
Profile CSV 解説:
  Order: 実行順序 (自動連番・D&Dで並替可)   Module: モジュール or 特殊コマンドを選択
  Enabled: 1=有効 0=無効   特殊: AUTOPILOT/RESTART/REEXPLORER/STOPLOG/STARTLOG/PAUSE/SHUTDOWN
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

$btnAddModule = New-RowButton -Text "+ Add Module"   -X 8   -Width 110
$btnSpecial   = New-RowButton -Text "+ Special"      -X 125 -Width 90 -BgColor ([System.Drawing.Color]::FromArgb(120, 90, 0))
$btnInsert    = New-RowButton -Text "Insert Above"   -X 222 -Width 105
$btnDelete    = New-RowButton -Text "Delete Row"     -X 334 -Width 95
$btnMoveUp    = New-RowButton -Text "Move Up"        -X 452 -Width 80
$btnMoveDown  = New-RowButton -Text "Move Down"      -X 539 -Width 85

$rowPanel.Controls.AddRange(@($btnAddModule, $btnSpecial, $btnInsert, $btnDelete, $btnMoveUp, $btnMoveDown))
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

# Order (auto, read-only)
$colOrder = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colOrder.Name = "Order"
$colOrder.HeaderText = "Order (実行順)"
$colOrder.Width = 50
$colOrder.MinimumWidth = 50
$colOrder.FillWeight = 8
$colOrder.ReadOnly = $true
$colOrder.DefaultCellStyle.Alignment = "MiddleCenter"
$colOrder.SortMode = "NotSortable"
$null = $script:dgv.Columns.Add($colOrder)

# Module (ComboBox)
$colModule = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
$colModule.Name = "Module"
$colModule.HeaderText = "Module (実行モジュール)"
$colModule.FillWeight = 40
$colModule.MinimumWidth = 200
$colModule.FlatStyle = "Flat"
$colModule.SortMode = "NotSortable"
# Add special markers
foreach ($markerName in $script:SpecialMarkers.Keys) {
    $null = $colModule.Items.Add($markerName)
}
# Add AutoLogon entries
foreach ($key in $script:AutoLogonEntries.Keys) {
    $null = $colModule.Items.Add($key)
}
# Add all discovered modules
foreach ($m in $script:ModuleDisplayList) {
    $null = $colModule.Items.Add($m)
}
$null = $script:dgv.Columns.Add($colModule)

# Enabled (ComboBox: 1/0)
$colEnabled = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
$colEnabled.Name = "Enabled"
$colEnabled.HeaderText = "Enabled (有効)"
$colEnabled.FillWeight = 10
$colEnabled.MinimumWidth = 70
$colEnabled.FlatStyle = "Flat"
$colEnabled.SortMode = "NotSortable"
$colEnabled.DefaultCellStyle.Alignment = "MiddleCenter"
$null = $colEnabled.Items.Add("1")
$null = $colEnabled.Items.Add("0")
$null = $script:dgv.Columns.Add($colEnabled)

# Description
$colDesc = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colDesc.Name = "Description"
$colDesc.HeaderText = "Description (説明)"
$colDesc.FillWeight = 30
$colDesc.MinimumWidth = 100
$colDesc.SortMode = "NotSortable"
$null = $script:dgv.Columns.Add($colDesc)

$form.Controls.Add($script:dgv)

# ========================================
# Event: CellValueChanged
# ========================================
$script:dgv.Add_CellValueChanged({
    param($sender, $e)
    if ($e.RowIndex -lt 0) { return }
    if ($script:IsLoading) { return }
    $colName = $sender.Columns[$e.ColumnIndex].Name

    if ($colName -eq "Module") {
        $row = $sender.Rows[$e.RowIndex]
        $module = $row.Cells["Module"].Value
        if ($module -eq "--- AUTOPILOT ---") {
            $row.Cells["Description"].Value = $script:SpecialMarkers[$module].Desc
            $row.DefaultCellStyle.BackColor = $autoPilotBg
            $row.DefaultCellStyle.ForeColor = $autoPilotFg
        }
        elseif ($script:AutoLogonEntries.Contains($module)) {
            $row.Cells["Description"].Value = $script:AutoLogonEntries[$module].Desc
            $row.DefaultCellStyle.BackColor = $autoLogonBg
            $row.DefaultCellStyle.ForeColor = $autoLogonFg
        }
        elseif ($script:SpecialMarkers.Contains($module)) {
            $row.Cells["Description"].Value = $script:SpecialMarkers[$module].Desc
            $row.DefaultCellStyle.BackColor = $specialBg
            $row.DefaultCellStyle.ForeColor = $specialFg
        }
        else {
            $row.DefaultCellStyle.BackColor = $bgCell
            $row.DefaultCellStyle.ForeColor = $fgText
            # Auto-fill description from module name
            if ($script:DisplayToPath.ContainsKey($module)) {
                $sp = $script:DisplayToPath[$module]
                if ($script:PathToDescription.ContainsKey($sp)) {
                    $row.Cells["Description"].Value = $script:PathToDescription[$sp]
                }
            }
        }
    }
})

# Commit ComboBox edits immediately
$script:dgv.Add_CurrentCellDirtyStateChanged({
    param($sender, $e)
    if ($script:IsLoading) { return }
    if ($sender.IsCurrentCellDirty) {
        $null = $sender.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
    }
})

# Suppress DataError for ComboBox mismatches (e.g. raw paths not in list)
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
        Module      = $srcRow.Cells["Module"].Value
        Enabled     = $srcRow.Cells["Enabled"].Value
        Description = $srcRow.Cells["Description"].Value
    }
    $sender.Rows.RemoveAt($script:DragRowIndex)
    if ($script:DragRowIndex -lt $targetIndex) { $targetIndex-- }

    $script:IsLoading = $true
    $null = $sender.Rows.Insert($targetIndex, 1)
    $newRow = $sender.Rows[$targetIndex]
    $newRow.Cells["Module"].Value      = $data.Module
    $newRow.Cells["Enabled"].Value     = $data.Enabled
    $newRow.Cells["Description"].Value = $data.Description
    $script:IsLoading = $false

    $sender.ClearSelection()
    $sender.Rows[$targetIndex].Selected = $true
    $sender.CurrentCell = $sender.Rows[$targetIndex].Cells["Module"]

    Update-OrderNumbers
    Update-RowStyles
    $script:DragRowIndex = -1
})

# ========================================
# Helper: Add a row with data
# ========================================
function Add-ModuleRow {
    param(
        [string]$Module = "",
        [string]$Enabled = "1",
        [string]$Description = "",
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
    $row.Cells["Module"].Value = $Module
    $row.Cells["Enabled"].Value = $Enabled
    $row.Cells["Description"].Value = $Description
    $script:IsLoading = $false

    Update-OrderNumbers
    Update-RowStyles

    $script:dgv.ClearSelection()
    $script:dgv.Rows[$idx].Selected = $true
    $script:dgv.CurrentCell = $script:dgv.Rows[$idx].Cells["Module"]
}

# ========================================
# Button: + Add Module
# ========================================
$btnAddModule.Add_Click({
    Add-ModuleRow -Enabled "1"
})

# ========================================
# Button: + Special (context menu with all markers)
# ========================================
$specialMenu = New-Object System.Windows.Forms.ContextMenuStrip
$specialMenu.BackColor = [System.Drawing.Color]::FromArgb(50, 50, 50)
$specialMenu.ForeColor = $fgText
$specialMenu.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$specialMenu.Renderer = New-Object System.Windows.Forms.ToolStripProfessionalRenderer(
    (New-Object System.Windows.Forms.ProfessionalColorTable)
)

foreach ($markerName in $script:SpecialMarkers.Keys) {
    $desc = $script:SpecialMarkers[$markerName].Desc
    $menuItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $menuItem.Text = "$markerName  ($desc)"
    $menuItem.Tag = $markerName
    $menuItem.ForeColor = $specialFg
    $menuItem.BackColor = [System.Drawing.Color]::FromArgb(50, 50, 50)
    $menuItem.Add_Click({
        param($sender, $e)
        $name = $sender.Tag
        $d = $script:SpecialMarkers[$name].Desc
        Add-ModuleRow -Module $name -Enabled "1" -Description $d
    })
    $null = $specialMenu.Items.Add($menuItem)
}

$btnSpecial.Add_Click({
    $specialMenu.Show($btnSpecial, (New-Object System.Drawing.Point(0, $btnSpecial.Height)))
})

# ========================================
# Button: Insert Above
# ========================================
$btnInsert.Add_Click({
    $selIdx = if ($script:dgv.CurrentRow) { $script:dgv.CurrentRow.Index } else { 0 }
    Add-ModuleRow -Enabled "1" -InsertAt $selIdx
})

# ========================================
# Button: Delete Row
# ========================================
$btnDelete.Add_Click({
    if ($script:dgv.CurrentRow -and -not $script:dgv.CurrentRow.IsNewRow) {
        $script:dgv.Rows.RemoveAt($script:dgv.CurrentRow.Index)
        Update-OrderNumbers
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
        Module      = $row.Cells["Module"].Value
        Enabled     = $row.Cells["Enabled"].Value
        Description = $row.Cells["Description"].Value
    }
    $script:dgv.Rows.RemoveAt($idx)
    $newIdx = $idx - 1

    $script:IsLoading = $true
    $null = $script:dgv.Rows.Insert($newIdx, 1)
    $newRow = $script:dgv.Rows[$newIdx]
    $newRow.Cells["Module"].Value      = $data.Module
    $newRow.Cells["Enabled"].Value     = $data.Enabled
    $newRow.Cells["Description"].Value = $data.Description
    $script:IsLoading = $false

    $script:dgv.ClearSelection()
    $script:dgv.Rows[$newIdx].Selected = $true
    $script:dgv.CurrentCell = $script:dgv.Rows[$newIdx].Cells["Module"]
    Update-OrderNumbers
    Update-RowStyles
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
        Module      = $row.Cells["Module"].Value
        Enabled     = $row.Cells["Enabled"].Value
        Description = $row.Cells["Description"].Value
    }
    $script:dgv.Rows.RemoveAt($idx)
    $newIdx = $idx + 1
    if ($newIdx -gt $script:dgv.Rows.Count) { $newIdx = $script:dgv.Rows.Count }

    $script:IsLoading = $true
    $null = $script:dgv.Rows.Insert($newIdx, 1)
    $newRow = $script:dgv.Rows[$newIdx]
    $newRow.Cells["Module"].Value      = $data.Module
    $newRow.Cells["Enabled"].Value     = $data.Enabled
    $newRow.Cells["Description"].Value = $data.Description
    $script:IsLoading = $false

    $script:dgv.ClearSelection()
    $script:dgv.Rows[$newIdx].Selected = $true
    $script:dgv.CurrentCell = $script:dgv.Rows[$newIdx].Cells["Module"]
    Update-OrderNumbers
    Update-RowStyles
})

# ========================================
# Button: New
# ========================================
$btnNew.Add_Click({
    $result = [System.Windows.Forms.MessageBox]::Show(
        "Clear current profile and start new?",
        "New Profile",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    if ($result -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $script:dgv.Rows.Clear()
    $script:txtProfileName.Text = ""
    $script:CurrentFilePath = $null
    $form.Text = "Profile Editor"

    Add-ModuleRow -Enabled "1"
})

# ========================================
# Button: Open
# ========================================
$btnOpen.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Title = "Open Profile CSV"
    $dlg.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    if (Test-Path $script:ProfilesDir) {
        $dlg.InitialDirectory = $script:ProfilesDir
    }

    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        Import-ProfileCsv -Path $dlg.FileName
        $script:CurrentFilePath = $dlg.FileName
        $profileName = [System.IO.Path]::GetFileNameWithoutExtension($dlg.FileName)
        $script:txtProfileName.Text = $profileName
        $form.Text = "Profile Editor - $profileName"
    }
    $dlg.Dispose()
})

# ========================================
# Helper: Save to file
# ========================================
function Save-ProfileToFile {
    param([string]$FilePath)
    $csvContent = Get-GridAsCsv
    try {
        [System.IO.File]::WriteAllText($FilePath, $csvContent, [System.Text.Encoding]::Default)
        $script:CurrentFilePath = $FilePath
        $profileName = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
        $script:txtProfileName.Text = $profileName
        $form.Text = "Profile Editor - $profileName"
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
        Save-ProfileToFile -FilePath $script:CurrentFilePath
    }
    else {
        # No file yet - use Save As logic
        $profileName = $script:txtProfileName.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($profileName)) {
            [System.Windows.Forms.MessageBox]::Show(
                "Please enter a Profile Name.", "Save",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            $script:txtProfileName.Focus()
            return
        }
        $targetPath = Join-Path $script:ProfilesDir "$profileName.csv"
        if (Test-Path $targetPath) {
            $ow = [System.Windows.Forms.MessageBox]::Show(
                "'$profileName.csv' already exists. Overwrite?", "Save",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )
            if ($ow -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        }
        if (-not (Test-Path $script:ProfilesDir)) {
            $null = New-Item -ItemType Directory -Path $script:ProfilesDir -Force
        }
        Save-ProfileToFile -FilePath $targetPath
    }
})

# ========================================
# Button: Save As
# ========================================
$btnSaveAs.Add_Click({
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Title = "Save Profile CSV"
    $dlg.Filter = "CSV Files (*.csv)|*.csv"
    if (Test-Path $script:ProfilesDir) {
        $dlg.InitialDirectory = $script:ProfilesDir
    }
    $profileName = $script:txtProfileName.Text.Trim()
    if ($profileName) {
        $dlg.FileName = "$profileName.csv"
    }
    else {
        $dlg.FileName = "New Profile.csv"
    }

    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        Save-ProfileToFile -FilePath $dlg.FileName
    }
    $dlg.Dispose()
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
Add-ModuleRow -Enabled "1"

# ========================================
# Show Form
# ========================================
$null = $form.ShowDialog()
$form.Dispose()
