# ========================================
# Printer Delete
# ========================================
# [PURPOSE]
# Delete printers in three coexisting modes:
#   1. KeepList mode   : Delete every TCP/IP printer NOT in the keep list.
#                        Keep list = hostlist Printer1..10Name env vars
#                                   + printer_driver_config/printer_list.csv
#                                     (TargetHost matched rows).
#                        Virtual printers (Print to PDF, OneNote, XPS, Fax)
#                        are NEVER touched by this mode (safety).
#   2. Explicit mode   : Delete printers listed in printer_delete.csv
#                        (TargetHost matched rows). Can target virtual
#                        printers by exact name.
#   3. Manual mode     : Fallback when neither source is active.
#                        Shows GUI with no pre-check.
#
# AutoPilot: GUI is skipped; delete candidates are auto-removed.
# If no candidates, result is Skipped.
#
# Post-Apply Verification reads back Get-Printer and reports -Verified.
# ========================================

Write-Host ""
Show-Separator
Write-Host "Printer Delete" -ForegroundColor Cyan
Show-Separator
Write-Host ""

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ========================================
# Step 1: Collect Keep List
# ========================================
# Sources:
#   (a) Hostlist environment variables SELECTED_PRINTER_1..10_NAME
#   (b) printer_driver_config/printer_list.csv, TargetHost matched rows
Show-Info "Collecting keep list..."

$keepNames = @()
$currentHost = [Environment]::GetEnvironmentVariable("SELECTED_NEW_PCNAME")

# (a) Hostlist env vars
for ($i = 1; $i -le 10; $i++) {
    $name = [Environment]::GetEnvironmentVariable("SELECTED_PRINTER_${i}_NAME")
    if (-not [string]::IsNullOrWhiteSpace($name)) {
        $trimmed = $name.Trim()
        if ($trimmed -notin $keepNames) {
            $keepNames += $trimmed
        }
    }
}

# (b) Cross-module reference to printer_driver_config/printer_list.csv
$printerListCsv = Join-Path $PSScriptRoot "..\printer_driver_config\printer_list.csv"
if (Test-Path $printerListCsv) {
    $csvItems = Import-ModuleCsv -Path $printerListCsv -FilterEnabled `
        -RequiredColumns @("Enabled", "TargetHost", "PrinterName")

    if ($null -ne $csvItems) {
        $addedFromCsv = 0
        foreach ($row in @($csvItems)) {
            $targetHost = if ($null -ne $row.TargetHost) { $row.TargetHost.Trim() } else { "" }
            $isAllHosts = [string]::IsNullOrEmpty($targetHost)
            $isMatch = $isAllHosts -or ($targetHost -ieq $currentHost)

            if (-not $isMatch) { continue }

            $pname = $row.PrinterName
            if (-not [string]::IsNullOrWhiteSpace($pname)) {
                $trimmed = $pname.Trim()
                if ($trimmed -notin $keepNames) {
                    $keepNames += $trimmed
                    $addedFromCsv++
                }
            }
        }
        if ($addedFromCsv -gt 0) {
            Show-Info "printer_list.csv: $addedFromCsv name(s) added to keep list"
        }
    }
}

Show-Info "Keep list: $($keepNames.Count) printer name(s)"

# ========================================
# Step 2: Collect Explicit Delete List
# ========================================
# Sources: printer_delete.csv (module-local), TargetHost matched rows.
# Backward compatible with old schema that lacks the TargetHost column.
$explicitDeletes = @()
$csvPath = Join-Path $PSScriptRoot "printer_delete.csv"

if (Test-Path $csvPath) {
    $deleteItems = Import-ModuleCsv -Path $csvPath -FilterEnabled `
        -RequiredColumns @("Enabled", "PrinterName")

    if ($null -ne $deleteItems) {
        $deleteItemsArray = @($deleteItems)
        $hasTargetHostColumn = $deleteItemsArray[0].PSObject.Properties.Name -contains 'TargetHost'
        $explicitMatched = 0

        foreach ($row in $deleteItemsArray) {
            $targetHost = if ($hasTargetHostColumn -and $null -ne $row.TargetHost) {
                $row.TargetHost.Trim()
            } else {
                ""
            }
            $isAllHosts = [string]::IsNullOrEmpty($targetHost)
            $isMatch = $isAllHosts -or ($targetHost -ieq $currentHost)

            if (-not $isMatch) { continue }

            $pname = $row.PrinterName
            if (-not [string]::IsNullOrWhiteSpace($pname)) {
                $trimmed = $pname.Trim()
                if ($trimmed -notin $explicitDeletes) {
                    $explicitDeletes += $trimmed
                    $explicitMatched++
                }
            }
        }

        if ($explicitMatched -gt 0) {
            Show-Info "printer_delete.csv: $explicitMatched explicit delete target(s)"
        }
    }
}
else {
    Show-Info "printer_delete.csv not found (explicit list empty)"
}

# ========================================
# Step 3: Enumerate Installed Printers & TCP/IP Ports
# ========================================
Show-Info "Enumerating installed printers..."

$installed = @()
try {
    $installed = @(Get-Printer -ErrorAction Stop | Select-Object Name, DriverName, PortName)
}
catch {
    Show-Error "Failed to enumerate installed printers: $_"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Failed to enumerate installed printers")
}

Show-Info "Installed printers: $($installed.Count)"

# Identify TCP/IP port names (ports with PrinterHostAddress)
$tcpPortNames = @()
try {
    $tcpPortNames = @(Get-PrinterPort -ErrorAction Stop |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_.PrinterHostAddress) } |
        ForEach-Object { $_.Name })
}
catch {
    Show-Warning "Failed to enumerate printer ports: $_"
}

# ========================================
# Step 4: Compute Delete Candidates
# ========================================
# KeepList mode:
#   For each installed TCP/IP printer, if its name is NOT in keepNames,
#   it becomes a KeepList delete candidate.
#   Virtual printers (non-TCP/IP) are never touched here.
#
# Explicit mode:
#   For each name in explicitDeletes, if the printer exists, mark it.
#   Can target virtual printers.
#   Overrides Keep (merges Reason).
$printerRows = @()  # Per installed printer: { Name, DriverName, PortName, IsTcp, Reason, PreCheck }

foreach ($p in $installed) {
    $isTcp = $p.PortName -in $tcpPortNames
    $inKeepList = $p.Name -in $keepNames
    $inExplicit = $p.Name -in $explicitDeletes

    $reason = "(Manual)"
    $preCheck = $false

    if ($inExplicit) {
        if ($keepNames.Count -gt 0 -and $isTcp -and -not $inKeepList) {
            $reason = "KeepList+Explicit"
        }
        else {
            $reason = "Explicit"
        }
        $preCheck = $true
    }
    elseif ($keepNames.Count -gt 0 -and $isTcp) {
        if ($inKeepList) {
            $reason = "Keep"
            $preCheck = $false
        }
        else {
            $reason = "KeepList"
            $preCheck = $true
        }
    }

    $printerRows += [PSCustomObject]@{
        Name       = $p.Name
        DriverName = $p.DriverName
        PortName   = $p.PortName
        IsTcp      = $isTcp
        Reason     = $reason
        PreCheck   = $preCheck
    }
}

$deleteCandidates = @($printerRows | Where-Object { $_.PreCheck })
$keepCount = @($printerRows | Where-Object { $_.Reason -eq "Keep" }).Count

# Explicit entries that are NOT found in installed list
$notFound = @($explicitDeletes | Where-Object { $_ -notin @($installed | ForEach-Object { $_.Name }) })

Show-Info "Delete candidates: $($deleteCandidates.Count) | Keep: $keepCount"
if ($notFound.Count -gt 0) {
    Show-Warning "Explicit names not found on this system: $($notFound -join ', ')"
}

# Determine active mode label for display
$modeLabel = if ($keepNames.Count -gt 0 -and $explicitDeletes.Count -gt 0) {
    "KeepList + Explicit"
} elseif ($keepNames.Count -gt 0) {
    "KeepList"
} elseif ($explicitDeletes.Count -gt 0) {
    "Explicit"
} else {
    "Manual"
}

Write-Host ""

# ========================================
# Step 5: Branch on AutoPilot / Manual
# ========================================
$script:successCount = 0
$script:failCount = 0
$script:deleted = $false
$script:deletedNames = @()

function Invoke-PrinterDelete {
    param([array]$Targets)

    if ($null -eq $Targets -or $Targets.Count -eq 0) { return }

    foreach ($t in $Targets) {
        try {
            Remove-Printer -Name $t.Name -ErrorAction Stop
            Show-Success "Deleted: $($t.Name) [$($t.Reason)]"
            $script:successCount++
            $script:deletedNames += $t.Name
        }
        catch {
            Show-Error "Failed to delete: $($t.Name) - $_"
            $script:failCount++
        }
    }

    $script:deleted = $true
}

if ($global:AutoPilotMode) {
    # AutoPilot path: skip GUI, auto-delete all candidates
    Show-Info "[AUTOPILOT] Mode: $modeLabel"

    if ($deleteCandidates.Count -eq 0) {
        Show-Info "No delete candidates (nothing to do)"
        Write-Host ""
        return (New-ModuleResult -Status "Skipped" -Message "No delete candidates")
    }

    Write-Host ""
    Write-Host "========================================" -ForegroundColor Magenta
    Write-Host "AutoPilot: deleting $($deleteCandidates.Count) printer(s)" -ForegroundColor Magenta
    Write-Host "========================================" -ForegroundColor Magenta
    foreach ($c in $deleteCandidates) {
        Write-Host "  - $($c.Name) [$($c.Reason)]" -ForegroundColor Yellow
    }
    Write-Host ""

    Invoke-PrinterDelete -Targets $deleteCandidates
    Write-Host ""
}
else {
    # ========================================
    # Manual / Interactive GUI path
    # ========================================
    Show-Info "Mode: $modeLabel (GUI)"
    Write-Host ""

    # ----- Color Scheme (Fabriq Standard) -----
    $bgDark       = [System.Drawing.Color]::FromArgb(30, 30, 30)
    $bgPanel      = [System.Drawing.Color]::FromArgb(40, 40, 40)
    $bgGrid       = [System.Drawing.Color]::FromArgb(35, 35, 35)
    $bgCell       = [System.Drawing.Color]::FromArgb(45, 45, 45)
    $bgHeader     = [System.Drawing.Color]::FromArgb(55, 55, 55)
    $bgButton     = [System.Drawing.Color]::FromArgb(60, 60, 60)
    $bgButtonHov  = [System.Drawing.Color]::FromArgb(80, 80, 80)
    $bgDelete     = [System.Drawing.Color]::FromArgb(180, 40, 40)
    $fgText       = [System.Drawing.Color]::FromArgb(220, 220, 220)
    $fgDim        = [System.Drawing.Color]::FromArgb(150, 150, 150)
    $fgHeader     = [System.Drawing.Color]::FromArgb(100, 180, 255)
    $gridLine     = [System.Drawing.Color]::FromArgb(60, 60, 60)

    # ----- Helper Functions -----
    function Update-Grid {

        $script:dgv.Rows.Clear()
        foreach ($row in $script:printerRows) {
            $idx = $script:dgv.Rows.Add()
            $r = $script:dgv.Rows[$idx]
            $r.Cells["Check"].Value  = [bool]$row.PreCheck
            $r.Cells["Name"].Value   = $row.Name
            $r.Cells["Driver"].Value = $row.DriverName
            $r.Cells["Port"].Value   = $row.PortName
            $r.Cells["Reason"].Value = $row.Reason
        }

        $statusText = "Printers: $($script:printerRows.Count) | Mode: $script:modeLabel"
        $statusText += " | Candidates: $($script:deleteCandidates.Count) | Keep: $script:keepCount"
        if ($script:notFound.Count -gt 0) {
            $statusText += " | Not found: $($script:notFound -join ', ')"
        }
        $script:statusLabel.Text = $statusText
    }

    function Remove-SelectedPrinters {
        $rowsToDelete = @()
        foreach ($row in $script:dgv.Rows) {
            if ($row.Cells["Check"].Value -eq $true) {
                $rowsToDelete += [PSCustomObject]@{
                    Name   = $row.Cells["Name"].Value
                    Reason = $row.Cells["Reason"].Value
                }
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
        foreach ($t in $rowsToDelete) {
            try {
                Remove-Printer -Name $t.Name -ErrorAction Stop
                $log += "[SUCCESS] $($t.Name) [$($t.Reason)]`n"
                $script:successCount++
                $script:deletedNames += $t.Name
            }
            catch {
                $log += "[FAILED] $($t.Name) : $($_.Exception.Message)`n"
                $script:failCount++
            }
        }

        $script:deleted = $true

        # Re-enumerate so the GUI reflects the post-delete state.
        # The remaining list uses the current reasons (no re-computation).
        $remainingNames = @((Get-Printer -ErrorAction SilentlyContinue) | ForEach-Object { $_.Name })
        $script:printerRows = @($script:printerRows | Where-Object { $_.Name -in $remainingNames })
        Update-Grid

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

    # Expose state to GUI closures
    $script:printerRows     = $printerRows
    $script:deleteCandidates = $deleteCandidates
    $script:keepCount        = $keepCount
    $script:notFound         = $notFound
    $script:modeLabel        = $modeLabel

    # ----- Main Form -----
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Fabriq Printer Delete"
    $form.Size = New-Object System.Drawing.Size(900, 600)
    $form.StartPosition = "CenterScreen"
    $form.BackColor = $bgDark
    $form.ForeColor = $fgText
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    # ----- Top Toolbar Panel -----
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

    $btnDelete    = New-StyledButton -Text "Delete Selected" -X 10 -Width 130 -BgColor $bgDelete
    $btnSelectAll = New-StyledButton -Text "Select All" -X 750 -Width 100

    $btnSelectAll.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
    $form.Add_Resize({
        $btnSelectAll.Left = $form.ClientSize.Width - 110
    })

    $toolPanel.Controls.AddRange(@($btnDelete, $btnSelectAll))
    $form.Controls.Add($toolPanel)

    # ----- Status Bar -----
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

    # ----- DataGridView -----
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

    $dgvType = $script:dgv.GetType()
    $pi = $dgvType.GetProperty("DoubleBuffered", [System.Reflection.BindingFlags]"Instance,NonPublic")
    $pi.SetValue($script:dgv, $true, $null)

    # Columns
    $colCheck = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $colCheck.Name = "Check"
    $colCheck.HeaderText = "Select"
    $colCheck.FillWeight = 8
    $null = $script:dgv.Columns.Add($colCheck)

    $colName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colName.Name = "Name"
    $colName.HeaderText = "Printer Name"
    $colName.FillWeight = 32
    $colName.ReadOnly = $true
    $null = $script:dgv.Columns.Add($colName)

    $colDriver = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colDriver.Name = "Driver"
    $colDriver.HeaderText = "Driver"
    $colDriver.FillWeight = 22
    $colDriver.ReadOnly = $true
    $null = $script:dgv.Columns.Add($colDriver)

    $colPort = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colPort.Name = "Port"
    $colPort.HeaderText = "Port"
    $colPort.FillWeight = 18
    $colPort.ReadOnly = $true
    $null = $script:dgv.Columns.Add($colPort)

    $colReason = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colReason.Name = "Reason"
    $colReason.HeaderText = "Reason"
    $colReason.FillWeight = 20
    $colReason.ReadOnly = $true
    $null = $script:dgv.Columns.Add($colReason)

    $form.Controls.Add($script:dgv)

    # ----- Events -----
    $script:dgv.Add_CellContentClick({
        param($src, $e)
        if ($e.RowIndex -ge 0 -and $src.Columns[$e.ColumnIndex].Name -eq "Check") {
            $src.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
        }
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

    # ----- Layout & Show -----
    $statusPanel.BringToFront()
    $toolPanel.BringToFront()
    $script:dgv.BringToFront()

    Update-Grid

    Show-Info "Opening Printer Delete GUI..."
    Write-Host ""

    $form.ShowDialog() | Out-Null
    $form.Dispose()
}

# ========================================
# Step 5.5: Post-Apply Verification
# ========================================
Write-Host ""

if (-not $script:deleted) {
    Show-Info "No printers were deleted"
    Write-Host ""
    return (New-ModuleResult -Status "Cancelled" -Message "No deletions performed")
}

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Post-Apply Verification" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

$verifyPass = 0
$verifyFail = 0

foreach ($name in $script:deletedNames) {
    $actual = Get-Printer -Name $name -ErrorAction SilentlyContinue
    if ($null -eq $actual) {
        Write-Host "  [VERIFIED] $name removed" -ForegroundColor Green
        $verifyPass++
    }
    else {
        Write-Host "  [VERIFY FAILED] $name still exists" -ForegroundColor Red
        $verifyFail++
    }
}

Write-Host ""
$verified = ($verifyFail -eq 0)

# ========================================
# Step 6: Result Summary
# ========================================
Show-Separator
Write-Host "Printer Delete Results" -ForegroundColor Cyan
Show-Separator
if ($script:successCount -gt 0) {
    Write-Host "  Success: $($script:successCount) printer(s)" -ForegroundColor Green
}
if ($script:failCount -gt 0) {
    Write-Host "  Failed:  $($script:failCount) printer(s)" -ForegroundColor Red
}
if ($verified) {
    Write-Host "  Verified: PASS" -ForegroundColor Green
}
else {
    Write-Host "  Verified: FAIL" -ForegroundColor Red
}
Show-Separator
Write-Host ""

$overallStatus = if ($script:failCount -eq 0 -and $script:successCount -gt 0) { "Success" }
    elseif ($script:successCount -gt 0 -and $script:failCount -gt 0) { "Partial" }
    elseif ($script:failCount -gt 0) { "Error" }
    else { "Cancelled" }

return (New-ModuleResult -Status $overallStatus `
    -Message "Success: $($script:successCount), Fail: $($script:failCount)" `
    -Verified $verified)
