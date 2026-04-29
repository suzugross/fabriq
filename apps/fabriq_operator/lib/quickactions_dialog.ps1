# ========================================
# Fabriq Operator - And More Dialog
# ========================================
# Sub-dialog for the "And More..." button on the Settings tab's
# Quick Actions section. Lists actions that are not on the front
# row but are still part of the operator's regular workflow.
# Returns a hashtable with .Action set to one of the existing main.ps1
# switch cases (Restart / HistoryExport / RegenerateChecklist /
# AppsMode), so no new kernel handlers are required for this dialog.
# ========================================

function Show-AndMoreDialog {
    $result = @{
        Action  = "Cancel"
        AppName = ""
    }

    # Items shown in the And More list. Action values must match
    # cases in kernel/main.ps1's switch ($guiSelection.Action).
    $items = @(
        [PSCustomObject]@{ Display = "Restart PC";           Action = "Restart" }
        [PSCustomObject]@{ Display = "Export History";       Action = "HistoryExport" }
        [PSCustomObject]@{ Display = "Regenerate Checklist"; Action = "RegenerateChecklist" }
        [PSCustomObject]@{ Display = "FabriqApps";           Action = "AppsMode" }
    )

    $form = New-Object System.Windows.Forms.Form
    Set-FormStyle -Form $form -Title "And More" -Width 420 -Height 320

    $titleLabel = New-StyledLabel -Text "And More" -X 16 -Y 12 -Width 380 -Height 28 `
                                  -Font $script:fontLarge -FgColor $script:fgHeader
    $form.Controls.Add($titleLabel)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Location = New-Object System.Drawing.Point(16, 46)
    $grid.Size = New-Object System.Drawing.Size(380, 180)
    Set-GridStyle -Grid $grid

    $colNum = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colNum.Name = "Num"
    $colNum.HeaderText = "#"
    $colNum.Width = 40
    $colNum.DefaultCellStyle.Alignment = "MiddleCenter"
    $grid.Columns.Add($colNum) | Out-Null

    $colItem = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colItem.Name = "Item"
    $colItem.HeaderText = "Action"
    $colItem.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
    $grid.Columns.Add($colItem) | Out-Null

    # Hidden column for the Action token returned to the dashboard form.
    $colAction = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colAction.Name = "ActionName"
    $colAction.Visible = $false
    $grid.Columns.Add($colAction) | Out-Null

    for ($i = 0; $i -lt $items.Count; $i++) {
        $grid.Rows.Add(($i + 1), $items[$i].Display, $items[$i].Action) | Out-Null
    }
    if ($grid.Rows.Count -gt 0) {
        $grid.Rows[0].Selected = $true
    }
    $form.Controls.Add($grid)

    $btnLaunch = New-StyledButton -Text "Launch" -X 266 -Y 240 -Width 130 -Height 32 -BgColor $script:bgAccent
    $btnLaunch.Font = $script:fontBold
    $form.Controls.Add($btnLaunch)

    $btnClose = New-StyledButton -Text "Close" -X 128 -Y 240 -Width 130 -Height 32
    $form.Controls.Add($btnClose)

    $btnLaunch.Add_Click({
        if ($grid.SelectedRows.Count -gt 0) {
            $idx = $grid.SelectedRows[0].Index
            $result.Action = $grid.Rows[$idx].Cells["ActionName"].Value
            $form.Close()
        }
    })

    $grid.Add_CellDoubleClick({
        $idx = $_.RowIndex
        if ($idx -ge 0 -and $idx -lt $items.Count) {
            $result.Action = $grid.Rows[$idx].Cells["ActionName"].Value
            $form.Close()
        }
    })

    $btnClose.Add_Click({
        $result.Action = "Cancel"
        $form.Close()
    })

    $form.Add_Shown({ $form.Activate() })
    [void]$form.ShowDialog()
    $form.Dispose()

    return $result
}
