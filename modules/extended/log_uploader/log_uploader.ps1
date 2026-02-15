# ========================================
# Log Uploader Module
# logs/ と evidence/ を指定先へ一括コピー
# ========================================

Show-Separator
Write-Host "Log Upload" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# ========================================
# Load configuration
# ========================================
$configPath = ".\kernel\log_destinations.csv"

if (-not (Test-Path $configPath)) {
    Show-Error "log_destinations.csv not found: $configPath"
    return (New-ModuleResult -Status Error -Message "log_destinations.csv not found")
}

$destinations = @(Import-Csv -Path $configPath -Encoding Default | Where-Object { $_.Enabled -eq "1" })

if ($destinations.Count -eq 0) {
    Show-Warning "No enabled destinations in log_destinations.csv"
    return (New-ModuleResult -Status Skipped -Message "No enabled destinations")
}

# ========================================
# Load session info for folder naming
# ========================================
$sessionPath = ".\kernel\session.json"
$mediaSerial = "UNKNOWN"

if (Test-Path $sessionPath) {
    try {
        $session = Get-Content $sessionPath -Raw -Encoding UTF8 | ConvertFrom-Json
        if (-not [string]::IsNullOrWhiteSpace($session.MediaSerial)) {
            $mediaSerial = $session.MediaSerial
        }
    }
    catch {
        Show-Warning "Failed to read session.json: $_"
    }
}

# ========================================
# Generate destination folder name
# ========================================
$hostname = $env:COMPUTERNAME
$dateStr = Get-Date -Format "yyyy_MM_dd_HHmmss"
$folderName = "${dateStr}_${hostname}_${mediaSerial}"

Show-Info "Upload folder: $folderName"
Write-Host ""

# ========================================
# Determine source directories
# ========================================
$logsDir = ".\logs"
$evidenceDir = ".\evidence"

$hasLogs = (Test-Path $logsDir) -and @(Get-ChildItem $logsDir -Recurse -File -ErrorAction SilentlyContinue).Count -gt 0
$hasEvidence = (Test-Path $evidenceDir) -and @(Get-ChildItem $evidenceDir -Recurse -File -ErrorAction SilentlyContinue).Count -gt 0

if (-not $hasLogs -and -not $hasEvidence) {
    Show-Warning "No log or evidence files to upload"
    return (New-ModuleResult -Status Skipped -Message "No files to upload")
}

# ========================================
# Stop transcript temporarily (unlock log file)
# ========================================
$transcriptWasStopped = $false
$transcriptPath = $global:FabriqTranscriptPath

if (-not [string]::IsNullOrWhiteSpace($transcriptPath)) {
    try {
        Stop-Transcript -ErrorAction Stop | Out-Null
        $transcriptWasStopped = $true
    }
    catch {
        # Transcript might not be running
    }
}

# ========================================
# Copy to each destination
# ========================================
$successCount = 0
$failCount = 0
$details = @()

foreach ($dest in $destinations) {
    $destBase = Join-Path $dest.Path $folderName
    $destDesc = if (-not [string]::IsNullOrWhiteSpace($dest.Description)) { $dest.Description } else { $dest.Path }

    Show-Info "Uploading to: $destDesc"
    Write-Host "  Path: $destBase" -ForegroundColor DarkGray

    try {
        # Test destination reachability
        $parentPath = $dest.Path
        if (-not (Test-Path $parentPath)) {
            # Try creating the parent (works for local paths)
            if ($dest.Type -eq "Local") {
                $null = New-Item -ItemType Directory -Path $parentPath -Force -ErrorAction Stop
            }
            else {
                throw "Destination not reachable: $parentPath"
            }
        }

        # Create destination folder
        $null = New-Item -ItemType Directory -Path $destBase -Force -ErrorAction Stop

        # Copy logs/
        if ($hasLogs) {
            $destLogs = Join-Path $destBase "logs"
            $null = New-Item -ItemType Directory -Path $destLogs -Force
            $copyResult = robocopy $logsDir $destLogs /E /NJH /NJS /NDL /NP /R:2 /W:1 2>&1
            Write-Host "  [OK] logs/ copied" -ForegroundColor Green
        }

        # Copy evidence/
        if ($hasEvidence) {
            $destEvidence = Join-Path $destBase "evidence"
            $null = New-Item -ItemType Directory -Path $destEvidence -Force
            $copyResult = robocopy $evidenceDir $destEvidence /E /NJH /NJS /NDL /NP /R:2 /W:1 2>&1
            Write-Host "  [OK] evidence/ copied" -ForegroundColor Green
        }

        # Copy session.json as metadata
        if (Test-Path $sessionPath) {
            Copy-Item $sessionPath (Join-Path $destBase "session.json") -Force -ErrorAction SilentlyContinue
        }

        $successCount++
        $details += "OK: $destDesc"
        Show-Success "Upload complete: $destDesc"
    }
    catch {
        $failCount++
        $errMsg = $_.Exception.Message
        $details += "FAIL: $destDesc - $errMsg"
        Show-Error "Upload failed: $destDesc - $errMsg"
    }

    Write-Host ""
}

# ========================================
# Resume transcript
# ========================================
if ($transcriptWasStopped -and -not [string]::IsNullOrWhiteSpace($transcriptPath)) {
    try {
        Start-Transcript -Path $transcriptPath -Append | Out-Null
    }
    catch {
        Show-Warning "Failed to resume transcript: $_"
    }
}

# ========================================
# Summary
# ========================================
Write-Host ""
Show-Separator
Write-Host "Upload Summary" -ForegroundColor Cyan
Show-Separator
Write-Host "  Success: $successCount / $($destinations.Count)" -ForegroundColor $(if ($successCount -gt 0) { "Green" } else { "Gray" })
Write-Host "  Failed:  $failCount / $($destinations.Count)" -ForegroundColor $(if ($failCount -gt 0) { "Red" } else { "Gray" })
Show-Separator

if ($failCount -gt 0 -and $successCount -eq 0) {
    return (New-ModuleResult -Status Error -Message "All uploads failed" -Details $details)
}
elseif ($failCount -gt 0) {
    return (New-ModuleResult -Status Partial -Message "$successCount/$($destinations.Count) succeeded" -Details $details)
}
else {
    return (New-ModuleResult -Status Success -Message "$successCount destination(s) uploaded" -Details $details)
}
