# ========================================
# Digital Gyotaku - Screenshot Evidence Tool
# ========================================
# Template module for manual task evidence capture.
# Copy this directory and edit task_list.csv to create
# custom evidence capture modules.
# ========================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ========================================
# Helper Functions
# ========================================

function Sanitize-FileName {
    param([string]$Name, [int]$MaxLength = 40)
    $sanitized = $Name -replace '[\\/:*?"<>|\s]', '_'
    if ($sanitized.Length -gt $MaxLength) {
        $sanitized = $sanitized.Substring(0, $MaxLength)
    }
    return $sanitized
}

function Take-FullScreenshot {
    param([string]$SavePath)
    $bounds = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
    $bitmap = New-Object System.Drawing.Bitmap($bounds.Width, $bounds.Height)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $graphics.CopyFromScreen($bounds.Location, [System.Drawing.Point]::Empty, $bounds.Size)
    $bitmap.Save($SavePath, [System.Drawing.Imaging.ImageFormat]::Png)
    $graphics.Dispose()
    $bitmap.Dispose()
}

# ========================================
# Header
# ========================================
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Digital Gyotaku - Screenshot Evidence" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# ========================================
# Load CSV
# ========================================
$csvPath = Join-Path $PSScriptRoot "task_list.csv"

if (-not (Test-Path $csvPath)) {
    Write-Host "[ERROR] task_list.csv not found: $csvPath" -ForegroundColor Red
    return (New-ModuleResult -Status "Error" -Message "task_list.csv not found")
}

try {
    $allTasks = @(Import-Csv -Path $csvPath -Encoding UTF8)
}
catch {
    Write-Host "[ERROR] Failed to load task_list.csv: $_" -ForegroundColor Red
    return (New-ModuleResult -Status "Error" -Message "Failed to load task_list.csv: $_")
}

# Filter enabled tasks
$tasks = @($allTasks | Where-Object { $_.Enabled -eq '1' })

if ($tasks.Count -eq 0) {
    Write-Host "[INFO] No enabled tasks found in task_list.csv" -ForegroundColor Yellow
    return (New-ModuleResult -Status "Skipped" -Message "No enabled tasks")
}

# Validate required columns
$requiredColumns = @('Enabled', 'TaskID', 'TaskTitle', 'Instruction', 'OpenCommand', 'OpenArgs')
$csvColumns = $allTasks[0].PSObject.Properties.Name
$missingColumns = @($requiredColumns | Where-Object { $_ -notin $csvColumns })
if ($missingColumns.Count -gt 0) {
    Write-Host "[ERROR] Missing CSV columns: $($missingColumns -join ', ')" -ForegroundColor Red
    return (New-ModuleResult -Status "Error" -Message "Missing columns: $($missingColumns -join ', ')")
}

Write-Host "[INFO] Loaded $($tasks.Count) tasks (of $($allTasks.Count) total)" -ForegroundColor Cyan
Write-Host ""

# ========================================
# Determine PC Name and Output Directory
# ========================================
$pcName = if (-not [string]::IsNullOrEmpty($env:SELECTED_NEW_PCNAME)) {
    $env:SELECTED_NEW_PCNAME
} else {
    $env:COMPUTERNAME
}

$fabriqRoot = (Resolve-Path (Join-Path $PSScriptRoot "..\..\..")).Path
$outputDir = Join-Path $fabriqRoot "evidence\gyotaku\$pcName"

Write-Host "[INFO] PC Name: $pcName" -ForegroundColor Cyan
Write-Host "[INFO] Output:  $outputDir" -ForegroundColor Cyan
Write-Host ""

# ========================================
# Task List Display
# ========================================
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host "Task List" -ForegroundColor Cyan
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""

$index = 1
foreach ($task in $tasks) {
    $openInfo = if (-not [string]::IsNullOrEmpty($task.OpenCommand)) {
        " [Auto-open: $($task.OpenCommand)]"
    } else { "" }

    Write-Host "  [$index/$($tasks.Count)] $($task.TaskID): $($task.TaskTitle)$openInfo" -ForegroundColor Yellow
    Write-Host "    $($task.Instruction)" -ForegroundColor Gray
    Write-Host ""
    $index++
}

Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""

# ========================================
# Confirmation
# ========================================
if (-not (Confirm-Execution -Message "Start evidence capture for $($tasks.Count) tasks?")) {
    Write-Host ""
    Write-Host "[INFO] Canceled" -ForegroundColor Yellow
    Write-Host ""
    return (New-ModuleResult -Status "Cancelled" -Message "User canceled")
}

Write-Host ""

# ========================================
# Create Output Directory
# ========================================
if (-not (Test-Path $outputDir)) {
    try {
        $null = New-Item -ItemType Directory -Path $outputDir -Force
        Write-Host "[INFO] Created output directory: $outputDir" -ForegroundColor Cyan
    }
    catch {
        Write-Host "[ERROR] Failed to create output directory: $_" -ForegroundColor Red
        return (New-ModuleResult -Status "Error" -Message "Failed to create output directory: $_")
    }
}

Write-Host ""

# ========================================
# Build WinForms Dialog
# ========================================

# Dark theme colors (matching status_monitor.ps1)
$darkBg       = [System.Drawing.Color]::FromArgb(30, 30, 30)
$panelBg      = [System.Drawing.Color]::FromArgb(45, 45, 45)
$accentCyan   = [System.Drawing.Color]::FromArgb(0, 200, 200)
$textWhite    = [System.Drawing.Color]::White
$textGray     = [System.Drawing.Color]::FromArgb(160, 160, 160)
$successGreen = [System.Drawing.Color]::FromArgb(80, 220, 80)
$warnYellow   = [System.Drawing.Color]::FromArgb(255, 200, 0)
$errorRed     = [System.Drawing.Color]::FromArgb(255, 80, 80)
$btnGreenBg   = [System.Drawing.Color]::FromArgb(30, 100, 30)
$btnYellowBg  = [System.Drawing.Color]::FromArgb(100, 85, 20)
$btnRedBg     = [System.Drawing.Color]::FromArgb(100, 30, 30)
$btnCyanBg    = [System.Drawing.Color]::FromArgb(20, 60, 60)

$fontNormal = New-Object System.Drawing.Font("Consolas", 9)
$fontBold   = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Bold)
$fontButton = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Bold)
$fontSmall  = New-Object System.Drawing.Font("Consolas", 9)

# Action state
$script:UserAction = $null   # "capture", "skip", "cancel", "complete"
$script:ShotCount = 0        # screenshots taken for current task

# --- Form ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "Digital Gyotaku"
$form.Size = New-Object System.Drawing.Size(440, 500)
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedToolWindow
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
$form.BackColor = $darkBg
$form.TopMost = $true
$form.ShowInTaskbar = $true
$form.KeyPreview = $true

# Position: bottom-right of primary screen
$screen = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
$form.Location = New-Object System.Drawing.Point(
    ($screen.Right - $form.Width - 20),
    ($screen.Bottom - $form.Height - 20)
)

# --- Title Label ---
$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Location = New-Object System.Drawing.Point(15, 12)
$lblTitle.Size = New-Object System.Drawing.Size(400, 24)
$lblTitle.Font = $fontBold
$lblTitle.ForeColor = $accentCyan
$lblTitle.BackColor = $darkBg
$lblTitle.Text = ""
$null = $form.Controls.Add($lblTitle)

# --- Instruction TextBox (read-only, multiline) ---
$txtInstruction = New-Object System.Windows.Forms.TextBox
$txtInstruction.Location = New-Object System.Drawing.Point(15, 42)
$txtInstruction.Size = New-Object System.Drawing.Size(400, 130)
$txtInstruction.Font = $fontNormal
$txtInstruction.ForeColor = $textWhite
$txtInstruction.BackColor = $panelBg
$txtInstruction.ReadOnly = $true
$txtInstruction.Multiline = $true
$txtInstruction.WordWrap = $true
$txtInstruction.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$txtInstruction.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$txtInstruction.TabStop = $false
$null = $form.Controls.Add($txtInstruction)

# --- Status Label (shows auto-open info or last result) ---
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Location = New-Object System.Drawing.Point(15, 178)
$lblStatus.Size = New-Object System.Drawing.Size(400, 18)
$lblStatus.Font = $fontSmall
$lblStatus.ForeColor = $textGray
$lblStatus.BackColor = $darkBg
$lblStatus.Text = ""
$null = $form.Controls.Add($lblStatus)

# --- Shot Log TextBox (read-only, shows captured file list) ---
$txtShotLog = New-Object System.Windows.Forms.TextBox
$txtShotLog.Location = New-Object System.Drawing.Point(15, 200)
$txtShotLog.Size = New-Object System.Drawing.Size(400, 150)
$txtShotLog.Font = $fontSmall
$txtShotLog.ForeColor = $successGreen
$txtShotLog.BackColor = $panelBg
$txtShotLog.ReadOnly = $true
$txtShotLog.Multiline = $true
$txtShotLog.WordWrap = $false
$txtShotLog.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$txtShotLog.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$txtShotLog.TabStop = $false
$txtShotLog.Visible = $false
$null = $form.Controls.Add($txtShotLog)

# --- Capture Button (dynamic text: "Evidence Capture" / "Next Shot") ---
$btnCapture = New-Object System.Windows.Forms.Button
$btnCapture.Location = New-Object System.Drawing.Point(15, 360)
$btnCapture.Size = New-Object System.Drawing.Size(400, 42)
$btnCapture.Font = $fontButton
$btnCapture.Text = "Evidence Capture (Enter)"
$btnCapture.ForeColor = $successGreen
$btnCapture.BackColor = $btnGreenBg
$btnCapture.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnCapture.FlatAppearance.BorderColor = $successGreen
$btnCapture.FlatAppearance.BorderSize = 1
$btnCapture.Cursor = [System.Windows.Forms.Cursors]::Hand
$null = $form.Controls.Add($btnCapture)

# --- Complete Button (visible only after first shot) ---
$btnComplete = New-Object System.Windows.Forms.Button
$btnComplete.Location = New-Object System.Drawing.Point(223, 360)
$btnComplete.Size = New-Object System.Drawing.Size(192, 42)
$btnComplete.Font = $fontButton
$btnComplete.Text = "Complete (F2)"
$btnComplete.ForeColor = $accentCyan
$btnComplete.BackColor = $btnCyanBg
$btnComplete.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnComplete.FlatAppearance.BorderColor = $accentCyan
$btnComplete.FlatAppearance.BorderSize = 1
$btnComplete.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnComplete.Visible = $false
$null = $form.Controls.Add($btnComplete)

# --- Skip Button ---
$btnSkip = New-Object System.Windows.Forms.Button
$btnSkip.Location = New-Object System.Drawing.Point(15, 410)
$btnSkip.Size = New-Object System.Drawing.Size(192, 34)
$btnSkip.Font = $fontSmall
$btnSkip.Text = "Skip (Esc)"
$btnSkip.ForeColor = $warnYellow
$btnSkip.BackColor = $btnYellowBg
$btnSkip.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnSkip.FlatAppearance.BorderColor = $warnYellow
$btnSkip.FlatAppearance.BorderSize = 1
$btnSkip.Cursor = [System.Windows.Forms.Cursors]::Hand
$null = $form.Controls.Add($btnSkip)

# --- Cancel All Button ---
$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Location = New-Object System.Drawing.Point(223, 410)
$btnCancel.Size = New-Object System.Drawing.Size(192, 34)
$btnCancel.Font = $fontSmall
$btnCancel.Text = "Cancel All"
$btnCancel.ForeColor = $errorRed
$btnCancel.BackColor = $btnRedBg
$btnCancel.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnCancel.FlatAppearance.BorderColor = $errorRed
$btnCancel.FlatAppearance.BorderSize = 1
$btnCancel.Cursor = [System.Windows.Forms.Cursors]::Hand
$null = $form.Controls.Add($btnCancel)

# --- Button Events ---
$btnCapture.Add_Click({ $script:UserAction = "capture" })
$btnComplete.Add_Click({ $script:UserAction = "complete" })
$btnSkip.Add_Click({ $script:UserAction = "skip" })
$btnCancel.Add_Click({ $script:UserAction = "cancel" })

# --- Keyboard Shortcuts ---
$form.Add_KeyDown({
    param($sender, $e)
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Return -or $e.KeyCode -eq [System.Windows.Forms.Keys]::F5) {
        $script:UserAction = "capture"
        $e.Handled = $true
        $e.SuppressKeyPress = $true
    }
    elseif ($e.KeyCode -eq [System.Windows.Forms.Keys]::F2 -and $script:ShotCount -gt 0) {
        $script:UserAction = "complete"
        $e.Handled = $true
        $e.SuppressKeyPress = $true
    }
    elseif ($e.KeyCode -eq [System.Windows.Forms.Keys]::Escape) {
        $script:UserAction = "skip"
        $e.Handled = $true
        $e.SuppressKeyPress = $true
    }
})

# Handle form close via X button as cancel
$form.Add_FormClosing({
    param($sender, $e)
    if ($script:UserAction -ne "done") {
        $e.Cancel = $true
        $script:UserAction = "cancel"
    }
})

# Show the form (modeless)
$form.Show()

# ========================================
# Task Execution Loop (multi-shot per task)
# ========================================
$script:Results = @()
$captureCount = 0
$taskCompleteCount = 0
$skipCount = 0
$failCount = 0
$dateStr = (Get-Date).ToString("yyyy_MM_dd")
$cancelled = $false

# Helper: update GUI state based on shot count
$script:UpdateShotUI = {
    if ($script:ShotCount -eq 0) {
        # State 1: waiting for first capture
        $btnCapture.Text = "Evidence Capture (Enter)"
        $btnCapture.Size = New-Object System.Drawing.Size(400, 42)
        $btnComplete.Visible = $false
    } else {
        # State 3: has shots, can take more or complete
        $btnCapture.Text = "Next Shot (Enter)"
        $btnCapture.Size = New-Object System.Drawing.Size(192, 42)
        $btnComplete.Visible = $true
    }
}

$current = 0
foreach ($task in $tasks) {
    $current++
    $script:ShotCount = 0

    # Update dialog contents and reset shot log
    $lblTitle.Text = "[$current/$($tasks.Count)] $($task.TaskTitle)"
    $txtInstruction.Text = $task.Instruction
    $lblStatus.Text = ""
    $lblStatus.ForeColor = $textGray
    $txtShotLog.Text = ""
    $txtShotLog.Visible = $false
    & $script:UpdateShotUI
    $script:UserAction = $null

    Write-Host "[$current/$($tasks.Count)] $($task.TaskID): $($task.TaskTitle)" -ForegroundColor Cyan

    # Auto-open settings screen if specified
    if (-not [string]::IsNullOrEmpty($task.OpenCommand)) {
        $lblStatus.Text = "Opening: $($task.OpenCommand)"
        try {
            if (-not [string]::IsNullOrEmpty($task.OpenArgs)) {
                $null = Start-Process $task.OpenCommand -ArgumentList $task.OpenArgs -PassThru
            } else {
                $null = Start-Process $task.OpenCommand -PassThru
            }
            Start-Sleep -Milliseconds 500
        }
        catch {
            $lblStatus.Text = "Could not auto-open: $($task.OpenCommand)"
            $lblStatus.ForeColor = $warnYellow
            Write-Host "  [WARN] Could not auto-open: $($task.OpenCommand) ($_)" -ForegroundColor Yellow
        }
    }

    # --- Multi-shot task loop ---
    $taskDone = $false
    while (-not $taskDone) {
        # Wait for user action via DoEvents polling
        $script:UserAction = $null
        while ($null -eq $script:UserAction) {
            [System.Windows.Forms.Application]::DoEvents()
            Start-Sleep -Milliseconds 50
        }

        switch ($script:UserAction) {
            "capture" {
                $shotNum = $script:ShotCount + 1
                $sanitizedTitle = Sanitize-FileName $task.TaskTitle
                $fileName = "${dateStr}_$($task.TaskID)_${sanitizedTitle}_${pcName}_$($shotNum.ToString('000')).png"
                $savePath = Join-Path $outputDir $fileName

                # Hide form, capture screenshot, show form
                $formLocation = $form.Location
                $form.Hide()
                [System.Windows.Forms.Application]::DoEvents()
                Start-Sleep -Milliseconds 300

                $captureOk = $false
                try {
                    Take-FullScreenshot -SavePath $savePath
                    $captureOk = $true
                }
                catch {
                    Write-Host "  [ERROR] Screenshot failed: $_" -ForegroundColor Red
                }

                $form.Location = $formLocation
                $form.Show()
                [System.Windows.Forms.Application]::DoEvents()

                if ($captureOk) {
                    Write-Host "  [SUCCESS] Saved: $fileName" -ForegroundColor Green
                    $captureCount++
                    $script:ShotCount++
                    $script:Results += [PSCustomObject]@{
                        Timestamp      = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                        TaskID         = $task.TaskID
                        TaskTitle      = $task.TaskTitle
                        Status         = "Captured"
                        ScreenshotFile = $fileName
                    }

                    # Update shot log
                    $logEntry = "#$($shotNum.ToString('000'))  $fileName"
                    if ($txtShotLog.Text.Length -gt 0) {
                        $txtShotLog.Text += "`r`n$logEntry"
                    } else {
                        $txtShotLog.Text = $logEntry
                    }
                    $txtShotLog.Visible = $true
                    $txtShotLog.SelectionStart = $txtShotLog.Text.Length
                    $txtShotLog.ScrollToCaret()

                    # Update title with shot count
                    $lblTitle.Text = "[$current/$($tasks.Count)] $($task.TaskTitle) ($($script:ShotCount) shots)"
                    $lblStatus.Text = "Captured: shot #$($shotNum.ToString('000'))"
                    $lblStatus.ForeColor = $successGreen

                    # Update button layout for multi-shot state
                    & $script:UpdateShotUI
                }
                else {
                    $lblStatus.Text = "ERROR: Screenshot failed"
                    $lblStatus.ForeColor = $errorRed
                    $failCount++
                    $script:Results += [PSCustomObject]@{
                        Timestamp      = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                        TaskID         = $task.TaskID
                        TaskTitle      = $task.TaskTitle
                        Status         = "Failed"
                        ScreenshotFile = ""
                    }
                }

                Start-Sleep -Milliseconds 300
            }
            "complete" {
                # Complete task (only when shotCount >= 1)
                if ($script:ShotCount -gt 0) {
                    $taskCompleteCount++
                    $taskDone = $true
                }
            }
            "skip" {
                if ($script:ShotCount -eq 0) {
                    # No shots taken — record as skipped
                    Write-Host "  [SKIP] Task skipped" -ForegroundColor Yellow
                    $skipCount++
                    $script:Results += [PSCustomObject]@{
                        Timestamp      = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                        TaskID         = $task.TaskID
                        TaskTitle      = $task.TaskTitle
                        Status         = "Skipped"
                        ScreenshotFile = ""
                    }
                }
                # If shots were already taken, skip just moves to next task
                $taskDone = $true
            }
            "cancel" {
                Write-Host "  [CANCEL] All remaining tasks cancelled" -ForegroundColor Yellow

                # Record current task as skipped only if no shots taken
                if ($script:ShotCount -eq 0) {
                    $script:Results += [PSCustomObject]@{
                        Timestamp      = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                        TaskID         = $task.TaskID
                        TaskTitle      = $task.TaskTitle
                        Status         = "Skipped"
                        ScreenshotFile = ""
                    }
                    $skipCount++
                }

                # Record remaining tasks as skipped
                for ($i = $current; $i -lt $tasks.Count; $i++) {
                    $script:Results += [PSCustomObject]@{
                        Timestamp      = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                        TaskID         = $tasks[$i].TaskID
                        TaskTitle      = $tasks[$i].TaskTitle
                        Status         = "Skipped"
                        ScreenshotFile = ""
                    }
                }

                $skipCount += ($tasks.Count - $current)
                $cancelled = $true
                $taskDone = $true
            }
        }
    }

    Write-Host ""

    if ($cancelled) { break }
}

# Close the form
$script:UserAction = "done"
$form.Close()
$form.Dispose()

# ========================================
# Write Manifest CSV
# ========================================
if ($script:Results.Count -gt 0) {
    $manifestDate = (Get-Date).ToString("yyyy_MM_dd_HHmmss")
    $manifestPath = Join-Path $outputDir "${manifestDate}_manifest.csv"

    try {
        $operator = if (-not [string]::IsNullOrEmpty($env:SELECTED_KANRI_NO)) {
            $env:SELECTED_KANRI_NO
        } else {
            $env:USERNAME
        }

        $manifestData = $script:Results | ForEach-Object {
            [PSCustomObject]@{
                Timestamp      = $_.Timestamp
                TaskID         = $_.TaskID
                TaskTitle      = $_.TaskTitle
                Status         = $_.Status
                ScreenshotFile = $_.ScreenshotFile
                Operator       = $operator
                PCName         = $pcName
            }
        }

        $null = $manifestData | Export-Csv -Path $manifestPath -NoTypeInformation -Encoding UTF8
        Write-Host "[INFO] Manifest saved: $manifestPath" -ForegroundColor Cyan
    }
    catch {
        Write-Host "[WARN] Failed to save manifest CSV: $_" -ForegroundColor Yellow
    }

    Write-Host ""
}

# ========================================
# Result Summary
# ========================================
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Digital Gyotaku Results" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Captured: $captureCount screenshots ($taskCompleteCount tasks)" -ForegroundColor Green
Write-Host "  Skipped:  $skipCount tasks" -ForegroundColor Yellow
Write-Host "  Failed:   $failCount items" -ForegroundColor $(if ($failCount -gt 0) { "Red" } else { "Green" })
Write-Host "========================================" -ForegroundColor Cyan

# Detail list
if ($script:Results.Count -gt 0) {
    Write-Host "  Details:" -ForegroundColor White
    foreach ($r in $script:Results) {
        $icon = switch ($r.Status) {
            "Captured" { "[OK]  " }
            "Skipped"  { "[SKIP]" }
            "Failed"   { "[FAIL]" }
            default    { "[??]  " }
        }
        $color = switch ($r.Status) {
            "Captured" { "Green" }
            "Skipped"  { "Yellow" }
            "Failed"   { "Red" }
            default    { "Gray" }
        }
        $fileInfo = if (-not [string]::IsNullOrEmpty($r.ScreenshotFile)) { " ($($r.ScreenshotFile))" } else { "" }
        Write-Host "    $icon $($r.TaskID) $($r.TaskTitle)$fileInfo" -ForegroundColor $color
    }
    Write-Host "========================================" -ForegroundColor Cyan
}

Write-Host ""
Write-Host "  Output: $outputDir" -ForegroundColor Cyan
Write-Host ""

# Return ModuleResult
$overallStatus = if ($cancelled -and $captureCount -eq 0) { "Cancelled" }
    elseif ($failCount -eq 0 -and $captureCount -gt 0) { "Success" }
    elseif ($captureCount -gt 0 -and $failCount -gt 0) { "Partial" }
    elseif ($captureCount -eq 0 -and $skipCount -gt 0) { "Skipped" }
    else { "Error" }

return (New-ModuleResult -Status $overallStatus -Message "Captured: $captureCount shots ($taskCompleteCount tasks), Skip: $skipCount, Fail: $failCount")
