# ========================================
# Fabriq Operator - Apps Launcher Dialog
# ========================================
# Displays a list of available fabriq apps
# discovered from the apps/ directory.
# Returns a hashtable describing the selected app.
# ========================================

function Show-AppsDialog {
    param(
        [string]$AppsDir = ".\apps"
    )

    $result = @{
        Action  = "Cancel"
        AppName = ""
        AppPath = ""
    }

    # Discover apps: each subdirectory containing a .ps1 with the same name
    # Exclude fabriq_operator itself
    $apps = @()
    if (Test-Path $AppsDir) {
        $appDirs = @(Get-ChildItem -Path $AppsDir -Directory | Sort-Object Name)
        foreach ($dir in $appDirs) {
            $entryScript = Join-Path $dir.FullName "$($dir.Name).ps1"
            if ((Test-Path $entryScript) -and $dir.Name -notin @("fabriq_operator","fabriq_ios")) {
                $apps += [PSCustomObject]@{
                    Name = $dir.Name
                    Path = $entryScript
                }
            }
        }
    }

    # Build form
    $form = New-Object System.Windows.Forms.Form
    Set-FormStyle -Form $form -Title "FabriqApps" -Width 450 -Height 400

    # Title
    $titleLabel = New-StyledLabel -Text "FabriqApps" -X 16 -Y 12 -Width 400 -Height 28 -Font $script:fontLarge -FgColor $script:fgHeader
    $form.Controls.Add($titleLabel)

    # App grid
    $appGrid = New-Object System.Windows.Forms.DataGridView
    $appGrid.Location = New-Object System.Drawing.Point(16, 46)
    $appGrid.Size = New-Object System.Drawing.Size(400, 260)
    Set-GridStyle -Grid $appGrid

    $colNum = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colNum.Name = "Num"
    $colNum.HeaderText = "#"
    $colNum.Width = 40
    $colNum.DefaultCellStyle.Alignment = "MiddleCenter"
    $appGrid.Columns.Add($colNum) | Out-Null

    $colName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colName.Name = "AppName"
    $colName.HeaderText = "App"
    $colName.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
    $appGrid.Columns.Add($colName) | Out-Null

    # Hidden column for path
    $colPath = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colPath.Name = "AppPath"
    $colPath.Visible = $false
    $appGrid.Columns.Add($colPath) | Out-Null

    # Populate grid
    for ($i = 0; $i -lt $apps.Count; $i++) {
        $appGrid.Rows.Add(($i + 1), $apps[$i].Name, $apps[$i].Path) | Out-Null
    }
    if ($appGrid.Rows.Count -gt 0) {
        $appGrid.Rows[0].Selected = $true
    }

    $form.Controls.Add($appGrid)

    # Buttons
    $btnLaunch = New-StyledButton -Text "Launch" -X 286 -Y 318 -Width 130 -Height 32 -BgColor $script:bgAccent
    $btnLaunch.Font = $script:fontBold
    $form.Controls.Add($btnLaunch)

    $btnClose = New-StyledButton -Text "Close" -X 148 -Y 318 -Width 130 -Height 32
    $form.Controls.Add($btnClose)

    # No apps message
    if ($apps.Count -eq 0) {
        $btnLaunch.Enabled = $false
        $noAppsLabel = New-StyledLabel -Text "No apps found in $AppsDir" -X 16 -Y 150 -Width 400 -Height 20 -FgColor $script:fgDim
        $noAppsLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $form.Controls.Add($noAppsLabel)
    }

    # ========================================
    # Event Handlers
    # ========================================

    # Launch button
    $btnLaunch.Add_Click({
        if ($appGrid.SelectedRows.Count -gt 0) {
            $idx = $appGrid.SelectedRows[0].Index
            $result.Action = "Launch"
            $result.AppName = $appGrid.Rows[$idx].Cells["AppName"].Value
            $result.AppPath = $appGrid.Rows[$idx].Cells["AppPath"].Value
            $form.Close()
        }
    })

    # Double-click to launch
    $appGrid.Add_CellDoubleClick({
        $idx = $_.RowIndex
        if ($idx -ge 0 -and $idx -lt $apps.Count) {
            $result.Action = "Launch"
            $result.AppName = $appGrid.Rows[$idx].Cells["AppName"].Value
            $result.AppPath = $appGrid.Rows[$idx].Cells["AppPath"].Value
            $form.Close()
        }
    })

    # Close button
    $btnClose.Add_Click({
        $result.Action = "Cancel"
        $form.Close()
    })

    # Show dialog
    $form.Add_Shown({ $form.Activate() })
    [void]$form.ShowDialog()
    $form.Dispose()

    return $result
}
