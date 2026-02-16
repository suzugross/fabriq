# ========================================
# Printer Delete (GUI)
# ========================================
# Displays installed printers in a GUI and
# allows batch deletion via checkboxes.
# If printer_delete.csv exists, matching
# printers are pre-checked automatically.
# ========================================

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Printer Delete (GUI)" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

# ========================================
# CSV Mode Detection
# ========================================
$csvPath = Join-Path $PSScriptRoot "printer_delete.csv"
$csvMode = $false
$csvPrinterNames = @()

if (Test-Path $csvPath) {
    $csvData = Import-CsvSafe -Path $csvPath -Description "printer_delete.csv"
    if ($null -ne $csvData) {
        if (Test-CsvColumns -CsvData $csvData -RequiredColumns @("Enabled", "PrinterName") -CsvName "printer_delete.csv") {
            $csvPrinterNames = @($csvData | Where-Object { $_.Enabled -eq "1" -and -not [string]::IsNullOrWhiteSpace($_.PrinterName) } | ForEach-Object { $_.PrinterName })
            if ($csvPrinterNames.Count -gt 0) {
                $csvMode = $true
                Write-Host "[INFO] CSV mode: $($csvPrinterNames.Count) printer(s) targeted for deletion" -ForegroundColor Cyan
            }
            else {
                Write-Host "[INFO] CSV found but no enabled entries. Switching to manual mode." -ForegroundColor Yellow
            }
        }
    }
}
else {
    Write-Host "[INFO] No printer_delete.csv found. Manual selection mode." -ForegroundColor Yellow
}

Write-Host ""

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
# Result Tracking
# ========================================
$script:successCount = 0
$script:failCount = 0
$script:deleted = $false

# ========================================
# Helper Functions
# ========================================

function Get-InstalledPrinters {
    try {
        return @(Get-Printer -ErrorAction SilentlyContinue | Select-Object Name, DriverName, PortName)
    }
    catch {
        return @()
    }
}

function Refresh-Grid {
    $script:dgv.Rows.Clear()
    $printers = Get-InstalledPrinters

    # Track CSV printers not found on system
    $csvNotFound = @()

    foreach ($p in $printers) {
        $idx = $script:dgv.Rows.Add()
        $row = $script:dgv.Rows[$idx]

        # Auto-check if CSV mode and printer matches
        $checked = $false
        if ($csvMode -and ($p.Name -in $csvPrinterNames)) {
            $checked = $true
        }

        $row.Cells["Check"].Value = $checked
        $row.Cells["Name"].Value = $p.Name
        $row.Cells["Driver"].Value = $p.DriverName
        $row.Cells["Port"].Value = $p.PortName
    }

    $modeText = if ($csvMode) { "CSV mode" } else { "Manual mode" }
    $statusText = "Printers: $($printers.Count) | $modeText"

    # Check for CSV entries not found on system
    if ($csvMode) {
        $installedNames = @($printers | ForEach-Object { $_.Name })
        $csvNotFound = @($csvPrinterNames | Where-Object { $_ -notin $installedNames })
        if ($csvNotFound.Count -gt 0) {
            $statusText += " | Not found: $($csvNotFound -join ', ')"
        }

        $checkedCount = @($csvPrinterNames | Where-Object { $_ -in $installedNames }).Count
        $statusText += " | Pre-checked: $checkedCount"
    }

    $script:statusLabel.Text = $statusText
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
            "No printers selected for deletion.",
            "Information",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
        return
    }

    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Delete $($rowsToDelete.Count) printer(s)?`nThis operation cannot be undone.",
        "Confirm Delete",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $log = ""

    foreach ($row in $rowsToDelete) {
        $pName = $row.Cells["Name"].Value
        try {
            Remove-Printer -Name $pName -ErrorAction Stop
            $log += "[SUCCESS] $pName`n"
            $script:successCount++
        }
        catch {
            $log += "[FAILED] $pName : $($_.Exception.Message)`n"
            $script:failCount++
        }
    }

    $script:deleted = $true

    # Refresh grid
    Refresh-Grid

    $icon = if ($script:failCount -eq 0) {
        [System.Windows.Forms.MessageBoxIcon]::Information
    } else {
        [System.Windows.Forms.MessageBoxIcon]::Warning
    }

    [System.Windows.Forms.MessageBox]::Show(
        "Deletion Complete`nSuccess: $($script:successCount)`nFailed: $($script:failCount)`n`n$log",
        "Result",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        $icon
    ) | Out-Null
}

# ========================================
# Main Form
# ========================================
$form = New-Object System.Windows.Forms.Form
$form.Text = "Fabriq Printer Delete"
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

$btnRefresh   = New-StyledButton -Text "Refresh" -X 10 -Width 90 -BgColor $bgRefresh
$btnDelete    = New-StyledButton -Text "Delete Selected" -X 110 -Width 130 -BgColor $bgDelete
$btnSelectAll = New-StyledButton -Text "Select All" -X 650 -Width 100

$btnSelectAll.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
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
$colCheck.HeaderText = "Select"
$colCheck.Width = 50
$colCheck.FillWeight = 10
$null = $script:dgv.Columns.Add($colCheck)

# Printer Name
$colName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colName.Name = "Name"
$colName.HeaderText = "Printer Name"
$colName.FillWeight = 40
$colName.ReadOnly = $true
$null = $script:dgv.Columns.Add($colName)

# Driver
$colDriver = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colDriver.Name = "Driver"
$colDriver.HeaderText = "Driver"
$colDriver.FillWeight = 30
$colDriver.ReadOnly = $true
$null = $script:dgv.Columns.Add($colDriver)

# Port
$colPort = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colPort.Name = "Port"
$colPort.HeaderText = "Port"
$colPort.FillWeight = 20
$colPort.ReadOnly = $true
$null = $script:dgv.Columns.Add($colPort)

$form.Controls.Add($script:dgv)

# ========================================
# Events
# ========================================

# Checkbox click handling (single click toggle)
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

$script:allSelected = $false
$btnSelectAll.Add_Click({
    $script:allSelected = -not $script:allSelected
    foreach ($row in $script:dgv.Rows) {
        $row.Cells["Check"].Value = $script:allSelected
    }
    $btnSelectAll.Text = if ($script:allSelected) { "Deselect All" } else { "Select All" }
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
Write-Host "[INFO] Opening Printer Delete GUI..." -ForegroundColor Cyan
Write-Host ""

$form.ShowDialog() | Out-Null
$form.Dispose()

# ========================================
# Console Summary & ModuleResult
# ========================================
Write-Host ""

if (-not $script:deleted) {
    Write-Host "[INFO] No printers were deleted" -ForegroundColor Yellow
    Write-Host ""
    return (New-ModuleResult -Status "Cancelled" -Message "No deletions performed")
}

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Printer Delete Results" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
if ($script:successCount -gt 0) {
    Write-Host "  Success: $($script:successCount) printer(s)" -ForegroundColor Green
}
if ($script:failCount -gt 0) {
    Write-Host "  Failed:  $($script:failCount) printer(s)" -ForegroundColor Red
}
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

$overallStatus = if ($script:failCount -eq 0 -and $script:successCount -gt 0) { "Success" }
    elseif ($script:successCount -gt 0 -and $script:failCount -gt 0) { "Partial" }
    elseif ($script:failCount -gt 0) { "Error" }
    else { "Cancelled" }

return (New-ModuleResult -Status $overallStatus -Message "Success: $($script:successCount), Fail: $($script:failCount)")
