# ========================================
# Function: Show Manifesto GUI
# ========================================
Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue

# ========================================
# Parse manifesto.csv (multiline body)
# Format: Enabled,title,main(multiline),id
# Each entry ends with a line ",<id>"
# ========================================
function Load-ManifestoEntries {
    $csvPath = ".\kernel\csv\manifesto.csv"
    if (-not (Test-Path $csvPath)) { return @() }

    $lines = @(Get-Content -Path $csvPath -Encoding Default)
    if ($lines.Count -lt 2) { return @() }

    $entries = @()
    $currentTitle = ""
    $currentBody = ""
    $currentEnabled = ""
    $inEntry = $false

    # Skip header (line 0)
    for ($i = 1; $i -lt $lines.Count; $i++) {
        $line = $lines[$i]

        # Detect end-of-entry: line is just ",<number>"
        if ($line -match '^,(\d+)\s*$') {
            if ($inEntry) {
                $entries += [PSCustomObject]@{
                    Enabled = $currentEnabled
                    Title   = $currentTitle.Trim()
                    Main    = $currentBody.Trim()
                    Id      = $Matches[1]
                }
            }
            $inEntry = $false
            $currentTitle = ""
            $currentBody = ""
            continue
        }

        # Detect start of new entry: line begins with "1," or "0,"
        if (-not $inEntry -and $line -match '^([01]),(.*)') {
            $inEntry = $true
            $currentEnabled = $Matches[1]
            $rest = $Matches[2]

            # Split title and body start at the "――," boundary
            $splitIdx = $rest.IndexOf([string]([char]0x2015 + [char]0x2015 + ","))
            if ($splitIdx -ge 0) {
                $currentTitle = $rest.Substring(0, $splitIdx + 2)
                $currentBody = $rest.Substring($splitIdx + 3)
            }
            else {
                # Fallback: title ends at "――" without comma (entry 1 case)
                $endIdx = $rest.IndexOf([string]([char]0x2015 + [char]0x2015))
                if ($endIdx -ge 0) {
                    $currentTitle = $rest.Substring(0, $endIdx + 2)
                    $currentBody = $rest.Substring($endIdx + 2)
                }
                else {
                    $currentTitle = $rest
                    $currentBody = ""
                }
            }
            continue
        }

        # Continuation line (body text)
        if ($inEntry) {
            $currentBody += "`n" + $line
        }
    }

    # Return only enabled entries
    return @($entries | Where-Object { $_.Enabled -eq "1" })
}

# ========================================
# Show Manifesto GUI
# ========================================
function Show-Manifesto {
    # Load entries from CSV
    $entries = @(Load-ManifestoEntries)

    if ($entries.Count -eq 0) {
        Write-Host "[WARN] No manifesto entries found" -ForegroundColor Yellow
        return
    }

    # Random selection: pick one entry each time
    $selected = $entries | Get-Random

    $manifestoTitle = $selected.Title
    $manifestoText  = $selected.Main

    # Borderless form with paper-white theme
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Manifeste du Surkitinisme"
    $form.Size = New-Object System.Drawing.Size(900, 680)
    $form.StartPosition = "CenterScreen"
    $form.BackColor = [System.Drawing.Color]::FromArgb(250, 248, 244)
    $form.FormBorderStyle = "None"

    # Drop shadow border (thin gray line around borderless form)
    $form.Add_Paint({
        param($s, $e)
        $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(200, 200, 200), 1)
        $e.Graphics.DrawRectangle($pen, 0, 0, $s.Width - 1, $s.Height - 1)
        $pen.Dispose()
    })

    # Font setup
    $fontTitle = New-Object System.Drawing.Font("Meiryo UI", 18, [System.Drawing.FontStyle]::Bold)
    $fontSub   = New-Object System.Drawing.Font("Meiryo UI", 12, [System.Drawing.FontStyle]::Italic)
    $fontBody  = New-Object System.Drawing.Font("Meiryo UI", 11, [System.Drawing.FontStyle]::Regular)

    # Window drag logic (WM_NCLBUTTONDOWN)
    $dragAction = {
        if ($_.Button -eq 'Left') {
            $form.Capture = $false
            $msg = [System.Windows.Forms.Message]::Create($form.Handle, 0xA1, [IntPtr]2, [IntPtr]0)
            $form.DefWndProc([ref]$msg)
        }
    }

    # --- Header area ---
    $pnlHeader = New-Object System.Windows.Forms.Panel
    $pnlHeader.Dock = "Top"
    $pnlHeader.Height = 100
    $pnlHeader.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 65)
    $pnlHeader.Add_MouseDown($dragAction)
    $form.Controls.Add($pnlHeader)

    # Title label (fixed)
    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text = "Manifeste du Surkitinisme"
    $lblTitle.ForeColor = [System.Drawing.Color]::FromArgb(0, 220, 220)
    $lblTitle.Font = $fontTitle
    $lblTitle.AutoSize = $false
    $lblTitle.TextAlign = "MiddleCenter"
    $lblTitle.Size = New-Object System.Drawing.Size(900, 50)
    $lblTitle.Location = New-Object System.Drawing.Point(0, 15)
    $lblTitle.BackColor = [System.Drawing.Color]::Transparent
    $lblTitle.Add_MouseDown($dragAction)
    $pnlHeader.Controls.Add($lblTitle)

    # Subtitle (from CSV title field)
    $lblSub = New-Object System.Windows.Forms.Label
    $lblSub.Text = $manifestoTitle
    $lblSub.ForeColor = [System.Drawing.Color]::FromArgb(190, 190, 190)
    $lblSub.Font = $fontSub
    $lblSub.AutoSize = $false
    $lblSub.TextAlign = "MiddleCenter"
    $lblSub.Size = New-Object System.Drawing.Size(900, 30)
    $lblSub.Location = New-Object System.Drawing.Point(0, 60)
    $lblSub.BackColor = [System.Drawing.Color]::Transparent
    $lblSub.Add_MouseDown($dragAction)
    $pnlHeader.Controls.Add($lblSub)

    # --- Body area (from CSV main field) ---
    $txtContent = New-Object System.Windows.Forms.RichTextBox
    $txtContent.Text = $manifestoText
    $txtContent.Font = $fontBody
    $txtContent.ForeColor = [System.Drawing.Color]::FromArgb(40, 40, 40)
    $txtContent.BackColor = [System.Drawing.Color]::FromArgb(250, 248, 244)
    $txtContent.BorderStyle = "None"
    $txtContent.ReadOnly = $true
    $txtContent.ScrollBars = "Vertical"
    $txtContent.Location = New-Object System.Drawing.Point(80, 130)
    $txtContent.Size = New-Object System.Drawing.Size(800, 450)
    $txtContent.RightMargin = 680
    $txtContent.Cursor = [System.Windows.Forms.Cursors]::Default
    $form.Controls.Add($txtContent)

    # --- Footer area ---
    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = "Close"
    $btnClose.Font = New-Object System.Drawing.Font("Meiryo UI", 10)
    $btnClose.FlatStyle = "Flat"
    $btnClose.FlatAppearance.BorderSize = 1
    $btnClose.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(180, 180, 180)
    $btnClose.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(230, 230, 230)
    $btnClose.ForeColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
    $btnClose.BackColor = [System.Drawing.Color]::FromArgb(240, 238, 234)
    $btnClose.Size = New-Object System.Drawing.Size(200, 45)
    $btnClose.Location = New-Object System.Drawing.Point(350, 600)
    $btnClose.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnClose.Add_Click({ $form.Close() })
    $form.Controls.Add($btnClose)

    # Entry ID display (bottom-right, subtle)
    $lblId = New-Object System.Windows.Forms.Label
    $lblId.Text = "#$($selected.Id)"
    $lblId.ForeColor = [System.Drawing.Color]::FromArgb(200, 200, 200)
    $lblId.Font = New-Object System.Drawing.Font("Meiryo UI", 8)
    $lblId.AutoSize = $true
    $lblId.Location = New-Object System.Drawing.Point(840, 655)
    $form.Controls.Add($lblId)

    # Keyboard shortcut (ESC to close)
    $form.KeyPreview = $true
    $form.Add_KeyDown({
        if ($_.KeyCode -eq "Escape") { $form.Close() }
    })

    # Show form
    $form.Add_Shown({ $form.Activate(); $btnClose.Focus() })
    [void]$form.ShowDialog()
    $form.Dispose()
}
