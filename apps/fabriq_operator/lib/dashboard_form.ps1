# ========================================
# Fabriq Operator - Main Dashboard Form
# ========================================
# Displays a tabbed dashboard for profile execution,
# module selection, and settings management.
# Returns a hashtable describing the user's selected action.
# ========================================

function Show-OperatorDashboard {
    param(
        [array]$AllModules,
        [array]$GroupedModules,
        [string]$HostName = $env:SELECTED_NEW_PCNAME,
        [string]$WorkerName = $env:FABRIQ_WORKER_NAME,
        [string]$LastResultSummary = ""
    )

    $result = @{
        Action             = "Quit"          # ExecuteProfile, FlexProfile, ExecuteModules, NewSession, OpenCsvEditor, OpenEvidence, WindowsUpdate, Restart, Refabriq, HistoryExport, RegenerateChecklist, Manifesto, LaunchApp, AppsMode, Quit
        ProfilePath        = ""
        ProfileName        = ""
        SelectedModules    = @()
        AutoPilot          = $false
        AutoPilotWaitSec   = 3
        AppName            = ""              # used with Action = LaunchApp
    }

    $form = New-Object System.Windows.Forms.Form
    Set-FormStyle -Form $form -Title "fabriq operator" -Width 700 -Height 560

    # ========================================
    # Header Bar
    # ========================================
    $headerPanel = New-StyledPanel -X 0 -Y 0 -Width 700 -Height 44 -BgColor $script:bgPanel
    $form.Controls.Add($headerPanel)

    $headerTitle = New-StyledLabel -Text "FABRIQ" -X 16 -Y 4 -Width 100 -Height 28 -Font $script:fontLarge -FgColor $script:fgWhite
    $headerPanel.Controls.Add($headerTitle)

    $headerSubtitle = New-StyledLabel -Text "- Manifeste du Surkitinisme -" -X 110 -Y 10 -Width 250 -Height 20 -Font $script:fontNormal -FgColor ([System.Drawing.Color]::FromArgb(160, 160, 160))
    $headerPanel.Controls.Add($headerSubtitle)

    # CentreCOM-style accent stripe (blue / yellow / red)
    $stripePanel = New-Object System.Windows.Forms.Panel
    $stripePanel.Location = New-Object System.Drawing.Point(0, 36)
    $stripePanel.Size = New-Object System.Drawing.Size(700, 8)
    $stripePanel.BackColor = $script:bgPanel
    $headerPanel.Controls.Add($stripePanel)

    $stripeBlue = New-Object System.Windows.Forms.Panel
    $stripeBlue.Location = New-Object System.Drawing.Point(0, 0)
    $stripeBlue.Size = New-Object System.Drawing.Size(234, 8)
    $stripeBlue.BackColor = $script:stripeBlue
    $stripePanel.Controls.Add($stripeBlue)

    $stripeYellow = New-Object System.Windows.Forms.Panel
    $stripeYellow.Location = New-Object System.Drawing.Point(234, 0)
    $stripeYellow.Size = New-Object System.Drawing.Size(233, 8)
    $stripeYellow.BackColor = $script:stripeYellow
    $stripePanel.Controls.Add($stripeYellow)

    $stripeRed = New-Object System.Windows.Forms.Panel
    $stripeRed.Location = New-Object System.Drawing.Point(467, 0)
    $stripeRed.Size = New-Object System.Drawing.Size(233, 8)
    $stripeRed.BackColor = $script:stripeRed
    $stripePanel.Controls.Add($stripeRed)

    $headerInfo = New-StyledLabel -Text "$HostName  |  W: $WorkerName" -X 300 -Y 10 -Width 370 -Height 22 -FgColor ([System.Drawing.Color]::FromArgb(200, 200, 200))
    $headerInfo.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
    $headerPanel.Controls.Add($headerInfo)

    # ========================================
    # Tab Control
    # ========================================
    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Location = New-Object System.Drawing.Point(10, 50)
    $tabControl.Size = New-Object System.Drawing.Size(664, 430)
    $tabControl.Font = $script:fontNormal
    $tabControl.BackColor = $script:bgForm
    $form.Controls.Add($tabControl)

    # ========================================
    # TAB 1: Profiles
    # ========================================
    $tabProfiles = New-Object System.Windows.Forms.TabPage
    $tabProfiles.Text = "Profiles"
    $tabProfiles.BackColor = $script:bgTabPage
    $tabControl.TabPages.Add($tabProfiles)

    # Profile grid
    $profileGrid = New-Object System.Windows.Forms.DataGridView
    $profileGrid.Location = New-Object System.Drawing.Point(10, 10)
    $profileGrid.Size = New-Object System.Drawing.Size(638, 260)
    Set-GridStyle -Grid $profileGrid

    $colPName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colPName.Name = "ProfileName"
    $colPName.HeaderText = "Profile"
    $colPName.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
    $profileGrid.Columns.Add($colPName) | Out-Null

    $colPModules = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colPModules.Name = "ModuleCount"
    $colPModules.HeaderText = "Modules"
    $colPModules.Width = 80
    $colPModules.DefaultCellStyle.Alignment = "MiddleCenter"
    $profileGrid.Columns.Add($colPModules) | Out-Null

    $colPTotal = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colPTotal.Name = "TotalCount"
    $colPTotal.HeaderText = "Total"
    $colPTotal.Width = 70
    $colPTotal.DefaultCellStyle.Alignment = "MiddleCenter"
    $profileGrid.Columns.Add($colPTotal) | Out-Null

    $colPPath = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colPPath.Name = "FilePath"
    $colPPath.HeaderText = "Path"
    $colPPath.Width = 120
    $colPPath.Visible = $false
    $profileGrid.Columns.Add($colPPath) | Out-Null

    # Load profiles
    $profiles = @(Load-Profiles -AllModules $AllModules)
    foreach ($p in $profiles) {
        $profileGrid.Rows.Add($p.ProfileName, $p.ModuleCount, $p.TotalCount, $p.FilePath) | Out-Null
    }
    if ($profileGrid.Rows.Count -gt 0) {
        $profileGrid.Rows[0].Selected = $true
    }

    $tabProfiles.Controls.Add($profileGrid)

    # Profile options
    $chkAutoPilot = New-StyledCheckBox -Text "AutoPilot" -X 10 -Y 280 -Width 120 -Checked $true
    $tabProfiles.Controls.Add($chkAutoPilot)

    # Execute Profile button (Linear path; preserved at original
    # position for operator muscle memory until the FlexProfile path
    # is fully validated and Linear is sunset).
    $btnExecProfile = New-StyledButton -Text "Execute Profile" -X 498 -Y 278 -Width 150 -Height 32 -BgColor $script:bgAccent
    $btnExecProfile.Font = $script:fontBold
    $tabProfiles.Controls.Add($btnExecProfile)

    # Execute (Flex) button — FlexProfile state-aware execution.
    # Placed between [View Details] and [Execute Profile] with the
    # success-green accent so the new flexible path is visually
    # distinguishable from the traditional Linear button.
    $btnExecFlex = New-StyledButton -Text "Execute (Flex)" -X 300 -Y 278 -Width 190 -Height 32 -BgColor $script:bgAdd
    $btnExecFlex.Font = $script:fontBold
    $btnExecFlex.ForeColor = $script:fgWhite
    $tabProfiles.Controls.Add($btnExecFlex)

    # View Details button (shifted left to make room for Execute (Flex))
    $btnViewDetails = New-StyledButton -Text "View Details" -X 200 -Y 278 -Width 92 -Height 32
    $tabProfiles.Controls.Add($btnViewDetails)

    # Profile detail display area
    $profileDetailBox = New-Object System.Windows.Forms.RichTextBox
    $profileDetailBox.Location = New-Object System.Drawing.Point(10, 318)
    $profileDetailBox.Size = New-Object System.Drawing.Size(638, 75)
    $profileDetailBox.BackColor = $script:bgPreview
    $profileDetailBox.ForeColor = $script:fgText
    $profileDetailBox.Font = $script:fontMono
    $profileDetailBox.ReadOnly = $true
    $profileDetailBox.BorderStyle = "None"
    $profileDetailBox.ScrollBars = "Vertical"
    $tabProfiles.Controls.Add($profileDetailBox)

    # ========================================
    # TAB 2: Modules
    # ========================================
    $tabModules = New-Object System.Windows.Forms.TabPage
    $tabModules.Text = "Modules"
    $tabModules.BackColor = $script:bgTabPage
    $tabControl.TabPages.Add($tabModules)

    # Category filter
    $catLabel = New-StyledLabel -Text "Category:" -X 10 -Y 12 -Width 65 -Height 22
    $tabModules.Controls.Add($catLabel)

    $catCombo = New-StyledComboBox -X 78 -Y 10 -Width 180
    $catCombo.Items.Add("All") | Out-Null
    if ($GroupedModules) {
        foreach ($cat in $GroupedModules) {
            $catCombo.Items.Add($cat.Name) | Out-Null
        }
    }
    $catCombo.SelectedIndex = 0
    $tabModules.Controls.Add($catCombo)

    # Search box
    $searchLabel = New-StyledLabel -Text "Search:" -X 280 -Y 12 -Width 50 -Height 22
    $tabModules.Controls.Add($searchLabel)

    $searchBox = New-Object System.Windows.Forms.TextBox
    $searchBox.Location = New-Object System.Drawing.Point(335, 10)
    $searchBox.Size = New-Object System.Drawing.Size(313, 22)
    Set-TextBoxStyle -TextBox $searchBox
    $tabModules.Controls.Add($searchBox)

    # Module grid
    $moduleGrid = New-Object System.Windows.Forms.DataGridView
    $moduleGrid.Location = New-Object System.Drawing.Point(10, 40)
    $moduleGrid.Size = New-Object System.Drawing.Size(638, 310)
    Set-GridStyle -Grid $moduleGrid

    $colMNum = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colMNum.Name = "Num"
    $colMNum.HeaderText = "#"
    $colMNum.Width = 35
    $colMNum.ReadOnly = $true
    $colMNum.DefaultCellStyle.Alignment = "MiddleCenter"
    $moduleGrid.Columns.Add($colMNum) | Out-Null

    $colMName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colMName.Name = "MenuName"
    $colMName.HeaderText = "Module"
    $colMName.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
    $colMName.ReadOnly = $true
    $moduleGrid.Columns.Add($colMName) | Out-Null

    $colMCat = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colMCat.Name = "Category"
    $colMCat.HeaderText = "Category"
    $colMCat.Width = 130
    $colMCat.ReadOnly = $true
    $moduleGrid.Columns.Add($colMCat) | Out-Null

    # Hidden columns for data
    $colMScript = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colMScript.Name = "Script"
    $colMScript.Visible = $false
    $moduleGrid.Columns.Add($colMScript) | Out-Null

    $colMOrder = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colMOrder.Name = "Order"
    $colMOrder.Visible = $false
    $moduleGrid.Columns.Add($colMOrder) | Out-Null

    # Store all modules for filtering
    $script:allModuleData = @()

    # Populate module grid
    function Update-ModuleGrid {
        param([string]$CategoryFilter = "All", [string]$SearchText = "")
        $moduleGrid.Rows.Clear()
        $num = 1
        $script:allModuleData = @()

        foreach ($cat in $GroupedModules) {
            if ($CategoryFilter -ne "All" -and $cat.Name -ne $CategoryFilter) { continue }
            $items = $cat.Group | Sort-Object Order
            foreach ($item in $items) {
                if ($SearchText.Length -gt 0 -and $item.MenuName -notlike "*$SearchText*") { continue }
                $moduleGrid.Rows.Add($num, $item.MenuName, $cat.Name, $item.Script, $item.Order) | Out-Null
                $script:allModuleData += $item
                $num++
            }
        }
    }

    Update-ModuleGrid

    $tabModules.Controls.Add($moduleGrid)

    # Execute button
    $btnExecModules = New-StyledButton -Text "Execute" -X 498 -Y 358 -Width 150 -Height 32 -BgColor $script:bgAccent
    $btnExecModules.Font = $script:fontBold
    $tabModules.Controls.Add($btnExecModules)

    # ========================================
    # TAB 3: Settings
    # ========================================
    $tabSettings = New-Object System.Windows.Forms.TabPage
    $tabSettings.Text = "Settings"
    $tabSettings.BackColor = $script:bgTabPage
    $tabControl.TabPages.Add($tabSettings)

    $settY = 16

    # Evidence path
    $evLabel = New-StyledLabel -Text "Evidence Output Path" -X 16 -Y $settY -Width 600 -Height 20 -Font $script:fontBold -FgColor $script:fgHeader
    $tabSettings.Controls.Add($evLabel)
    $settY += 24

    $evPathText = if (-not [string]::IsNullOrEmpty($global:FabriqEvidenceBasePath)) { $global:FabriqEvidenceBasePath } else { "(not initialized)" }
    $evPathLabel = New-StyledLabel -Text $evPathText -X 16 -Y $settY -Width 520 -Height 20 -FgColor $script:fgDim
    $tabSettings.Controls.Add($evPathLabel)

    $btnOpenEvidence = New-StyledButton -Text "Open Folder" -X 538 -Y ($settY - 2) -Width 110 -Height 26
    $tabSettings.Controls.Add($btnOpenEvidence)
    $settY += 36

    # Separator
    $sepLine = New-Object System.Windows.Forms.Label
    $sepLine.Location = New-Object System.Drawing.Point(16, $settY)
    $sepLine.Size = New-Object System.Drawing.Size(632, 1)
    $sepLine.BackColor = $script:gridLine
    $tabSettings.Controls.Add($sepLine)
    $settY += 16

    # Quick Actions - 5 front-row shortcuts plus an "And More..."
    # button that opens a dialog with the secondary actions
    # (Restart PC / Export History / Regenerate Checklist / FabriqApps).
    $qaLabel = New-StyledLabel -Text "Quick Actions" -X 16 -Y $settY -Width 600 -Height 20 -Font $script:fontBold -FgColor $script:fgHeader
    $tabSettings.Controls.Add($qaLabel)
    $settY += 30

    $btnOpenCsvEditor = New-StyledButton -Text "Open CSV Editor" -X 16 -Y $settY -Width 200 -Height 30
    $tabSettings.Controls.Add($btnOpenCsvEditor)

    $btnWindowsUpdate = New-StyledButton -Text "Windows Update" -X 226 -Y $settY -Width 200 -Height 30
    $tabSettings.Controls.Add($btnWindowsUpdate)

    $btnRefabriq = New-StyledButton -Text "Refabriq" -X 436 -Y $settY -Width 212 -Height 30
    $tabSettings.Controls.Add($btnRefabriq)
    $settY += 44

    $btnSystemLauncher = New-StyledButton -Text "System Launcher" -X 16 -Y $settY -Width 200 -Height 30
    $tabSettings.Controls.Add($btnSystemLauncher)

    $btnAndMore = New-StyledButton -Text "And More..." -X 226 -Y $settY -Width 212 -Height 30
    $tabSettings.Controls.Add($btnAndMore)
    $settY += 50

    # Separator
    $sepLine2 = New-Object System.Windows.Forms.Label
    $sepLine2.Location = New-Object System.Drawing.Point(16, $settY)
    $sepLine2.Size = New-Object System.Drawing.Size(632, 1)
    $sepLine2.BackColor = $script:gridLine
    $tabSettings.Controls.Add($sepLine2)
    $settY += 16

    # Session info
    $sessLabel = New-StyledLabel -Text "Session" -X 16 -Y $settY -Width 600 -Height 20 -Font $script:fontBold -FgColor $script:fgHeader
    $tabSettings.Controls.Add($sessLabel)
    $settY += 26

    $sessInfoText = "Worker: $WorkerName  |  Host: $HostName  ($env:SELECTED_KANRI_NO)"
    $sessInfoLabel = New-StyledLabel -Text $sessInfoText -X 16 -Y $settY -Width 400 -Height 20 -FgColor $script:fgDim
    $tabSettings.Controls.Add($sessInfoLabel)

    $btnNewSession = New-StyledButton -Text "New Session" -X 518 -Y ($settY - 2) -Width 130 -Height 28
    $tabSettings.Controls.Add($btnNewSession)
    $settY += 36

    # Manifesto
    $btnManifesto = New-StyledButton -Text "Manifeste du Surkitinisme" -X 16 -Y $settY -Width 200 -Height 28
    $tabSettings.Controls.Add($btnManifesto)

    # ========================================
    # Status Bar
    # ========================================
    $statusBar = New-Object System.Windows.Forms.StatusStrip
    $statusBar.BackColor = $script:bgHeader
    $statusBar.ForeColor = $script:fgText
    $statusItem = New-Object System.Windows.Forms.ToolStripStatusLabel
    $statusItem.Text = if ($LastResultSummary) { $LastResultSummary } else { "Ready" }
    $statusItem.ForeColor = $script:fgText
    $statusBar.Items.Add($statusItem) | Out-Null
    $form.Controls.Add($statusBar)

    # ========================================
    # Event Handlers
    # ========================================

    # Profile grid selection changed -> update detail
    $profileGrid.Add_SelectionChanged({
        if ($profileGrid.SelectedRows.Count -gt 0) {
            $idx = $profileGrid.SelectedRows[0].Index
            $filePath = $profileGrid.Rows[$idx].Cells["FilePath"].Value
            if ($filePath -and (Test-Path $filePath)) {
                try {
                    $entries = @(Import-Csv $filePath -Encoding Default | Where-Object { $_.Enabled -eq "1" } | Sort-Object { [int]$_.Order })
                    $lines = @()
                    foreach ($e in $entries) {
                        $desc = if ($e.Description) { $e.Description } else { $e.ScriptPath }
                        $lines += "  $($e.Order.PadLeft(3))  $desc"
                    }
                    $profileDetailBox.Text = ($lines -join "`r`n")
                }
                catch {
                    $profileDetailBox.Text = "Failed to read profile: $_"
                }
            }
        }
    })

    # Trigger initial detail load
    if ($profileGrid.Rows.Count -gt 0) {
        $profileGrid.Rows[0].Selected = $false
        $profileGrid.Rows[0].Selected = $true
    }

    # View Details button
    $btnViewDetails.Add_Click({
        if ($profileGrid.SelectedRows.Count -gt 0) {
            $idx = $profileGrid.SelectedRows[0].Index
            $filePath = $profileGrid.Rows[$idx].Cells["FilePath"].Value
            if ($filePath -and (Test-Path $filePath)) {
                Start-Process "explorer.exe" -ArgumentList "/select,`"$filePath`""
            }
        }
    })

    # Execute Profile button (Linear)
    $btnExecProfile.Add_Click({
        if ($profileGrid.SelectedRows.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Please select a profile.", "fabriq operator", "OK", "Warning") | Out-Null
            return
        }
        $idx = $profileGrid.SelectedRows[0].Index
        $result.Action = "ExecuteProfile"
        $result.ProfilePath = $profileGrid.Rows[$idx].Cells["FilePath"].Value
        $result.ProfileName = $profileGrid.Rows[$idx].Cells["ProfileName"].Value
        $result.AutoPilot = $chkAutoPilot.Checked
        $form.Close()
    })

    # Execute (Flex) button — opens the FlexProfile state-aware dashboard.
    # AutoPilot intentionally does NOT propagate from this dashboard's
    # checkbox; FlexProfile dashboard has its own AutoPilot toggle that
    # defaults OFF (per design). main.ps1's "FlexProfile" handler ignores
    # $result.AutoPilot.
    $btnExecFlex.Add_Click({
        if ($profileGrid.SelectedRows.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Please select a profile.", "fabriq operator", "OK", "Warning") | Out-Null
            return
        }
        $idx = $profileGrid.SelectedRows[0].Index
        $result.Action = "FlexProfile"
        $result.ProfilePath = $profileGrid.Rows[$idx].Cells["FilePath"].Value
        $result.ProfileName = $profileGrid.Rows[$idx].Cells["ProfileName"].Value
        $form.Close()
    })

    # Category filter changed
    $catCombo.Add_SelectedIndexChanged({
        Update-ModuleGrid -CategoryFilter $catCombo.SelectedItem -SearchText $searchBox.Text
    })

    # Search text changed
    $searchBox.Add_TextChanged({
        Update-ModuleGrid -CategoryFilter $catCombo.SelectedItem -SearchText $searchBox.Text
    })

    # Execute single highlighted module
    $btnExecModules.Add_Click({
        if ($moduleGrid.SelectedRows.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Please select a module.", "fabriq operator", "OK", "Warning") | Out-Null
            return
        }
        $idx = $moduleGrid.SelectedRows[0].Index
        if ($idx -ge 0 -and $idx -lt $script:allModuleData.Count) {
            $result.Action = "ExecuteModules"
            $result.SelectedModules = @($script:allModuleData[$idx])
            $form.Close()
        }
    })

    # Double-click on module row to execute
    $moduleGrid.Add_CellDoubleClick({
        $idx = $_.RowIndex
        if ($idx -ge 0 -and $idx -lt $script:allModuleData.Count) {
            $result.Action = "ExecuteModules"
            $result.SelectedModules = @($script:allModuleData[$idx])
            $form.Close()
        }
    })

    # Settings tab buttons
    $btnOpenEvidence.Add_Click({
        $evPath = $global:FabriqEvidenceBasePath
        if ($evPath -and (Test-Path $evPath)) {
            Start-Process "explorer.exe" -ArgumentList "`"$evPath`""
        }
        elseif ($evPath) {
            # Open parent if base path not yet created
            $parent = Split-Path $evPath -Parent
            if ($parent -and (Test-Path $parent)) {
                Start-Process "explorer.exe" -ArgumentList "`"$parent`""
            }
            else {
                Start-Process "explorer.exe" -ArgumentList "`".\evidence`""
            }
        }
        else {
            Start-Process "explorer.exe" -ArgumentList "`".\evidence`""
        }
    })

    $btnOpenCsvEditor.Add_Click({
        $result.Action = "OpenCsvEditor"
        $form.Close()
    })

    $btnWindowsUpdate.Add_Click({
        $result.Action = "WindowsUpdate"
        $form.Close()
    })

    $btnRefabriq.Add_Click({
        $result.Action = "Refabriq"
        $form.Close()
    })

    $btnSystemLauncher.Add_Click({
        $result.Action = "SystemLauncher"
        $form.Close()
    })

    $btnAndMore.Add_Click({
        # Show the And More sub-dialog modally over the dashboard.
        # On selection, propagate the chosen Action up to the
        # dashboard's $result so main.ps1 dispatches via its existing
        # switch case (Restart / HistoryExport / RegenerateChecklist /
        # AppsMode). Cancel keeps the dashboard open.
        $subResult = Show-AndMoreDialog
        if ($subResult.Action -and $subResult.Action -ne "Cancel") {
            $result.Action = $subResult.Action
            $form.Close()
        }
    })

    $btnNewSession.Add_Click({
        $result.Action = "NewSession"
        $form.Close()
    })

    $btnManifesto.Add_Click({
        $result.Action = "Manifesto"
        $form.Close()
    })

    # Form close via X button = Quit
    $form.Add_FormClosing({
        if ($result.Action -eq "Quit" -or $result.Action -eq "") {
            $result.Action = "Quit"
        }
    })

    # Show dialog
    $form.Add_Shown({ $form.Activate() })
    [void]$form.ShowDialog()
    $form.Dispose()

    return $result
}
