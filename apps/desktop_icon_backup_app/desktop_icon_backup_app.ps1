# ========================================
# Desktop Icon Layout Backup - Fabriq
# ========================================
# Standalone GUI app that exports the desktop
# icon layout registry key to a .reg file.
# Backup is stored in the shared backup folder
# used by the desktop_icon_config restore module.
#
# Registry key:
#   HKCU\Software\Microsoft\Windows\Shell\Bags\1\Desktop
#
# Backup location:
#   modules\extended\desktop_icon_config\backup\
# ========================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

# ========================================
# Color Scheme (Fabriq Standard Dark)
# ========================================
$bgDark      = [System.Drawing.Color]::FromArgb(30,  30,  30)
$bgAccent    = [System.Drawing.Color]::FromArgb(0,   120, 215)
$bgAccentHov = [System.Drawing.Color]::FromArgb(0,   100, 190)
$borderColor = [System.Drawing.Color]::FromArgb(60,  60,  60)
$fgText      = [System.Drawing.Color]::FromArgb(220, 220, 220)
$fgDim       = [System.Drawing.Color]::FromArgb(150, 150, 150)
$fgHeader    = [System.Drawing.Color]::FromArgb(100, 180, 255)
$fgSuccess   = [System.Drawing.Color]::FromArgb(80,  200, 80)
$fgWarning   = [System.Drawing.Color]::FromArgb(255, 160, 50)
$fgError     = [System.Drawing.Color]::FromArgb(255, 80,  80)

# ========================================
# Paths
# ========================================
$script:regKeyPs    = 'HKCU:\Software\Microsoft\Windows\Shell\Bags\1\Desktop'
$script:regKeyExport = 'HKEY_CURRENT_USER\Software\Microsoft\Windows\Shell\Bags\1\Desktop'
$script:backupDir   = Join-Path $PSScriptRoot "..\..\modules\extended\desktop_icon_config\backup"
$script:backupDir   = [System.IO.Path]::GetFullPath($script:backupDir)

# ========================================
# Helper: Separator line
# ========================================
function New-Separator {
    param([int]$Y)
    $sep = New-Object System.Windows.Forms.Panel
    $sep.Location  = New-Object System.Drawing.Point(20, $Y)
    $sep.Size      = New-Object System.Drawing.Size(420, 1)
    $sep.BackColor = $borderColor
    return $sep
}

# ========================================
# Helper: Label
# ========================================
function New-Label {
    param(
        [string]$Text,
        [int]$X, [int]$Y, [int]$W, [int]$H,
        $Color = $fgText,
        $Font  = $null,
        [string]$Align = "MiddleLeft"
    )
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text      = $Text
    $lbl.Location  = New-Object System.Drawing.Point($X, $Y)
    $lbl.Size      = New-Object System.Drawing.Size($W, $H)
    $lbl.ForeColor = $Color
    $lbl.TextAlign = $Align
    if ($Font) { $lbl.Font = $Font }
    return $lbl
}

# ========================================
# Main Form
# ========================================
$form = New-Object System.Windows.Forms.Form
$form.Text            = "Desktop Icon Backup - Fabriq"
$form.Size            = New-Object System.Drawing.Size(480, 420)
$form.StartPosition   = "CenterScreen"
$form.BackColor       = $bgDark
$form.ForeColor       = $fgText
$form.Font            = New-Object System.Drawing.Font("Segoe UI", 9)
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox     = $false

# ========================================
# SECTION 1: Title
# ========================================
$lblTitle = New-Label -Text "Desktop Icon Layout Backup" `
    -X 20 -Y 20 -W 420 -H 28 -Color $fgHeader `
    -Font (New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)) `
    -Align "MiddleLeft"
$form.Controls.Add($lblTitle)

$form.Controls.Add((New-Separator -Y 54))

# ========================================
# SECTION 2: Instruction
# ========================================
$lblInstruction = New-Label -Text "Arrange your desktop icons in the desired layout,`nthen click ""Run Backup"" to save the current positions." `
    -X 20 -Y 66 -W 420 -H 44 -Color $fgText -Align "TopLeft"
$form.Controls.Add($lblInstruction)

$form.Controls.Add((New-Separator -Y 118))

# ========================================
# SECTION 3: Registry key status
# ========================================
$lblRegCaption = New-Label -Text "Registry key:" -X 20 -Y 132 -W 130 -H 22 -Color $fgDim
$form.Controls.Add($lblRegCaption)

$lblRegStatus = New-Label -Text "Checking..." `
    -X 155 -Y 132 -W 285 -H 22 -Color $fgDim `
    -Font (New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold))
$form.Controls.Add($lblRegStatus)

$form.Controls.Add((New-Separator -Y 162))

# ========================================
# SECTION 4: Run Backup button
# ========================================
$btnBackup = New-Object System.Windows.Forms.Button
$btnBackup.Text      = "  Run Backup"
$btnBackup.Location  = New-Object System.Drawing.Point(130, 180)
$btnBackup.Size      = New-Object System.Drawing.Size(200, 44)
$btnBackup.FlatStyle = "Flat"
$btnBackup.FlatAppearance.BorderSize            = 0
$btnBackup.FlatAppearance.MouseOverBackColor    = $bgAccentHov
$btnBackup.BackColor = $bgAccent
$btnBackup.ForeColor = [System.Drawing.Color]::White
$btnBackup.Font      = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$btnBackup.Cursor    = [System.Windows.Forms.Cursors]::Hand
$form.Controls.Add($btnBackup)

$form.Controls.Add((New-Separator -Y 240))

# ========================================
# SECTION 5: Result area
# ========================================
$lblResultFile = New-Label -Text "" -X 20 -Y 254 -W 420 -H 20 -Color $fgText -Align "TopLeft"
$form.Controls.Add($lblResultFile)

$lblResultSize = New-Label -Text "" -X 20 -Y 276 -W 420 -H 20 -Color $fgText -Align "TopLeft"
$form.Controls.Add($lblResultSize)

$lblResultLocation = New-Label -Text "" -X 20 -Y 298 -W 420 -H 20 -Color $fgDim -Align "TopLeft"
$form.Controls.Add($lblResultLocation)

$form.Controls.Add((New-Separator -Y 328))

# ========================================
# SECTION 6: Status
# ========================================
$lblStatus = New-Label -Text "Ready" -X 20 -Y 342 -W 420 -H 22 -Color $fgDim
$form.Controls.Add($lblStatus)

# ========================================
# Logic: Update registry key status indicator
# ========================================
function Update-RegStatus {
    $exists = Test-Path $script:regKeyPs -ErrorAction SilentlyContinue
    if ($exists) {
        $lblRegStatus.Text      = "[Found]"
        $lblRegStatus.ForeColor = $fgSuccess
    }
    else {
        $lblRegStatus.Text      = "[Not Found]"
        $lblRegStatus.ForeColor = $fgWarning
    }
    return $exists
}

# ========================================
# Events
# ========================================

# --- Form Load: initial registry check ---
$form.Add_Load({
    Update-RegStatus | Out-Null
})

# --- Run Backup ---
$btnBackup.Add_Click({

    # Re-check registry key
    $keyExists = Update-RegStatus
    if (-not $keyExists) {
        [System.Windows.Forms.MessageBox]::Show(
            "The registry key was not found:`n$($script:regKeyPs)`n`nTry arranging an icon on the desktop first (drag it to a new position), then retry.",
            "Registry Key Not Found",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        $lblStatus.Text      = "Aborted: registry key not found"
        $lblStatus.ForeColor = $fgWarning
        return
    }

    # Clear previous result
    $lblResultFile.Text     = ""
    $lblResultSize.Text     = ""
    $lblResultLocation.Text = ""
    $lblStatus.Text         = "Running backup..."
    $lblStatus.ForeColor    = $fgDim
    $form.Refresh()

    # Ensure backup directory exists
    if (-not (Test-Path $script:backupDir)) {
        try {
            $null = New-Item -ItemType Directory -Path $script:backupDir -Force
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Failed to create backup directory:`n$($script:backupDir)`n`n$($_.Exception.Message)",
                "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null
            $lblStatus.Text      = "Error: could not create backup directory"
            $lblStatus.ForeColor = $fgError
            return
        }
    }

    # Build export file path
    $timestamp  = Get-Date -Format "yyyyMMdd_HHmmss"
    $exportFile = Join-Path $script:backupDir "DesktopIcons_${timestamp}.reg"

    # Execute reg.exe export
    try {
        $proc = Start-Process reg.exe `
            -ArgumentList "export `"$($script:regKeyExport)`" `"$exportFile`" /y" `
            -Wait -PassThru -NoNewWindow -ErrorAction Stop

        if ($proc.ExitCode -eq 0) {
            $fileInfo = Get-Item $exportFile -ErrorAction SilentlyContinue
            $sizeStr  = if ($fileInfo.Length -gt 1KB) {
                "{0:N1} KB" -f ($fileInfo.Length / 1KB)
            }
            else {
                "$($fileInfo.Length) bytes"
            }

            $lblResultFile.Text      = "File:     $([System.IO.Path]::GetFileName($exportFile))"
            $lblResultFile.ForeColor = $fgText
            $lblResultSize.Text      = "Size:     $sizeStr"
            $lblResultSize.ForeColor = $fgText
            $lblResultLocation.Text  = "Location: $($script:backupDir)"

            $lblStatus.Text      = "Backup completed successfully"
            $lblStatus.ForeColor = $fgSuccess
        }
        else {
            [System.Windows.Forms.MessageBox]::Show(
                "reg.exe exited with code $($proc.ExitCode).`nThe backup may not have been saved correctly.",
                "Backup Failed",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null
            $lblStatus.Text      = "Error: reg.exe exit code $($proc.ExitCode)"
            $lblStatus.ForeColor = $fgError
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to run reg.exe:`n$($_.Exception.Message)",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        $lblStatus.Text      = "Error: $($_.Exception.Message)"
        $lblStatus.ForeColor = $fgError
    }
})

# ========================================
# Show Form
# ========================================
$form.ShowDialog() | Out-Null
$form.Dispose()
