# =============================================================================
# AutoKey Recipe Editor (Prototype)
# =============================================================================
# Purpose: GUI tool to create/edit recipe.csv for AutoKey module
# Features: Grid editing, Action dropdowns, Row reordering, CSV I/O
# Compatibility: PowerShell 5.1+, Windows Forms
# =============================================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# =============================================================================
# Constants & Configuration
# =============================================================================
# Actions defined in autokey_config.ps1
$ActionList = @("Open", "WaitWin", "AppFocus", "Type", "Key", "Wait")
$InitialDir = Get-Location

# =============================================================================
# Main Form Setup
# =============================================================================
$form = New-Object System.Windows.Forms.Form
$form.Text = "AutoKey Recipe Editor"
$form.Size = New-Object System.Drawing.Size(900, 600)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.MinimizeBox = $true
$form.MaximizeBox = $true

# =============================================================================
# UI: Top Toolbar (FlowLayout)
# =============================================================================
$toolbar = New-Object System.Windows.Forms.FlowLayoutPanel
$toolbar.Dock = "Top"
$toolbar.Height = 45
$toolbar.Padding = New-Object System.Windows.Forms.Padding(10)
$toolbar.FlowDirection = "LeftToRight"
$toolbar.AutoSize = $false
$form.Controls.Add($toolbar)

function Create-ToolbarButton($text, $handler) {
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = $text
    $btn.AutoSize = $true
    $btn.MinimumSize = New-Object System.Drawing.Size(80, 28)
    $btn.Height = 28
    $btn.FlatStyle = "Standard"
    $btn.Margin = New-Object System.Windows.Forms.Padding(0, 0, 5, 0)
    $btn.Add_Click($handler)
    $toolbar.Controls.Add($btn)
}

function Create-Separator {
    $sep = New-Object System.Windows.Forms.Label
    $sep.Text = "|"
    $sep.AutoSize = $true
    $sep.ForeColor = [System.Drawing.Color]::Gray
    $sep.Margin = New-Object System.Windows.Forms.Padding(5, 5, 5, 5)
    $toolbar.Controls.Add($sep)
}

# --- File Operations ---
Create-ToolbarButton "Load CSV" { Load-Csv }
Create-ToolbarButton "Save CSV" { Save-Csv }

Create-Separator

# --- Row Operations ---
Create-ToolbarButton "Add Row"    { Add-Row }
Create-ToolbarButton "Remove"     { Remove-Row }
Create-ToolbarButton "Move Up"    { Move-Row -1 }
Create-ToolbarButton "Move Down"  { Move-Row 1 }

# =============================================================================
# UI: Data Grid
# =============================================================================
$grid = New-Object System.Windows.Forms.DataGridView
$grid.Dock = "Fill"
$grid.AutoSizeColumnsMode = "Fill"
$grid.AllowUserToAddRows = $false  # Controlled via buttons
$grid.MultiSelect = $false
$grid.SelectionMode = "FullRowSelect"
$grid.RowHeadersVisible = $true
$grid.BackgroundColor = [System.Drawing.SystemColors]::ControlLight
$grid.BorderStyle = "Fixed3D"
$form.Controls.Add($grid)

# Define Columns
function Add-Column($name, $header, $type, $fillWeight, $readOnly) {
    if ($type -eq "Combo") {
        $col = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
        $ActionList | ForEach-Object { $col.Items.Add($_) } | Out-Null
    } else {
        $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    }
    $col.Name = $name
    $col.HeaderText = $header
    $col.FillWeight = $fillWeight
    $col.ReadOnly = $readOnly
    $grid.Columns.Add($col) | Out-Null
    return $col
}

# Column Setup
$colStep = Add-Column "Step" "Step" "Text" 8 $true
$colStep.DefaultCellStyle.BackColor = [System.Drawing.Color]::WhiteSmoke
$colStep.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Gray

Add-Column "Action" "Action" "Combo" 15 $false
Add-Column "Value"  "Value (Path/Text)" "Text" 40 $false
Add-Column "Wait"   "Wait (ms)" "Text" 10 $false
Add-Column "Note"   "Note" "Text" 27 $false

# =============================================================================
# Logic Functions
# =============================================================================

function Recalculate-Steps {
    # Updates the Step number based on row index
    for ($i = 0; $i -lt $grid.Rows.Count; $i++) {
        $grid.Rows[$i].Cells["Step"].Value = ($i + 1).ToString()
    }
}

function Add-Row {
    $idx = $grid.Rows.Add()
    $row = $grid.Rows[$idx]
    
    # Default values
    $row.Cells["Step"].Value   = ($idx + 1).ToString()
    $row.Cells["Action"].Value = "Wait"
    $row.Cells["Wait"].Value   = "0"
    $row.Cells["Value"].Value  = ""
    $row.Cells["Note"].Value   = ""

    # Focus new row
    $grid.ClearSelection()
    $row.Selected = $true
    $grid.FirstDisplayedScrollingRowIndex = $idx
}

function Remove-Row {
    if ($grid.SelectedRows.Count -gt 0) {
        $row = $grid.SelectedRows[0]
        if (-not $row.IsNewRow) {
            $grid.Rows.Remove($row)
            Recalculate-Steps
        }
    }
}

function Move-Row($direction) {
    # Direction: -1 (Up), 1 (Down)
    if ($grid.SelectedRows.Count -eq 0) { return }
    
    $selRow = $grid.SelectedRows[0]
    $idx = $selRow.Index
    $newIdx = $idx + $direction

    # Boundary check
    if ($newIdx -lt 0 -or $newIdx -ge $grid.Rows.Count) { return }

    # Move logic: Create new row copy, insert, delete old
    $sourceRow = $grid.Rows[$idx]
    $data = @()
    foreach ($cell in $sourceRow.Cells) { $data += $cell.Value }
    
    $grid.Rows.RemoveAt($idx)
    $grid.Rows.Insert($newIdx, $data)
    
    # Restore selection
    $grid.ClearSelection()
    $grid.Rows[$newIdx].Selected = $true
    Recalculate-Steps
}

function Load-Csv {
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $dlg.InitialDirectory = $InitialDir
    $dlg.Title = "Load Recipe CSV"
    
    if ($dlg.ShowDialog() -eq "OK") {
        try {
            # Use Default encoding to match main.ps1 behavior
            $data = Import-Csv $dlg.FileName -Encoding Default
            $grid.Rows.Clear()
            
            foreach ($row in $data) {
                $idx = $grid.Rows.Add()
                $gRow = $grid.Rows[$idx]
                
                $gRow.Cells["Step"].Value = $row.Step
                
                # Validate Action
                if ($ActionList -contains $row.Action) {
                    $gRow.Cells["Action"].Value = $row.Action
                } else {
                    $gRow.Cells["Action"].Value = "Wait"
                }
                
                $gRow.Cells["Value"].Value = $row.Value
                $gRow.Cells["Wait"].Value  = $row.Wait
                $gRow.Cells["Note"].Value  = $row.Note
            }
            Recalculate-Steps
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Error loading file:`n$_", "Error", 0, 16)
        }
    }
}

function Save-Csv {
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $dlg.FileName = "recipe.csv"
    $dlg.InitialDirectory = $InitialDir
    $dlg.Title = "Save Recipe CSV"
    
    if ($dlg.ShowDialog() -eq "OK") {
        try {
            Recalculate-Steps
            
            $exportData = @()
            foreach ($row in $grid.Rows) {
                if ($row.IsNewRow) { continue }
                
                $obj = New-Object PSObject
                $obj | Add-Member -MemberType NoteProperty -Name "Step"   -Value $row.Cells["Step"].Value
                $obj | Add-Member -MemberType NoteProperty -Name "Action" -Value $row.Cells["Action"].Value
                $obj | Add-Member -MemberType NoteProperty -Name "Value"  -Value $row.Cells["Value"].Value
                $obj | Add-Member -MemberType NoteProperty -Name "Wait"   -Value $row.Cells["Wait"].Value
                $obj | Add-Member -MemberType NoteProperty -Name "Note"   -Value $row.Cells["Note"].Value
                
                $exportData += $obj
            }
            
            # Export with Default encoding for compatibility
            $exportData | Export-Csv -Path $dlg.FileName -NoTypeInformation -Encoding Default
            [System.Windows.Forms.MessageBox]::Show("File saved successfully.", "Success", 0, 64)
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Error saving file:`n$_", "Error", 0, 16)
        }
    }
}

# =============================================================================
# Startup Logic
# =============================================================================
# Add an initial empty row if grid is empty
if ($grid.Rows.Count -eq 0) {
    Add-Row
}

# Show Form
[void]$form.ShowDialog()