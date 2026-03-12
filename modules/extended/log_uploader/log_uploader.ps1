# ========================================
# Log Uploader Module
# logs/ と evidence/ を指定先へ一括コピー
# UNC認証対応（AuthUser/AuthPass in CSV）
# ========================================

Show-Separator
Write-Host "Log Upload" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# ========================================
# Load configuration
# ========================================
$configPath = ".\kernel\csv\log_destinations.csv"

$allDestinations = Import-ModuleCsv -Path $configPath
if ($null -eq $allDestinations -or $allDestinations.Count -eq 0) {
    return (New-ModuleResult -Status Error -Message "Failed to load log_destinations.csv")
}

$destinations = @($allDestinations | Where-Object { $_.Enabled -eq "1" })

if ($destinations.Count -eq 0) {
    Show-Warning "No enabled destinations in log_destinations.csv"
    return (New-ModuleResult -Status Skipped -Message "No enabled destinations")
}

# ========================================
# Load session info for folder naming
# ========================================
$sessionPath = ".\kernel\json\session.json"
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
# Generate destination folder name & evidence source
# ========================================
$logsDir = ".\logs"
$useUnifiedPath = -not [string]::IsNullOrWhiteSpace($global:FabriqEvidenceBasePath)

if ($useUnifiedPath) {
    # Unified mode: use evidence base path directory name as upload folder
    $folderName  = Split-Path $global:FabriqEvidenceBasePath -Leaf
    $evidenceDir = $global:FabriqEvidenceBasePath
}
else {
    # Fallback: legacy naming and full evidence directory
    $hostname = $env:COMPUTERNAME
    $dateStr = Get-Date -Format "yyyy_MM_dd_HHmmss"
    $folderName  = "${dateStr}_${hostname}_${mediaSerial}"
    $evidenceDir = ".\evidence"
}

Show-Info "Upload folder: $folderName"
if ($useUnifiedPath) {
    Show-Info "Evidence source: $evidenceDir (session only)"
}
Write-Host ""

# ========================================
# Check source content
# ========================================
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

    # ----------------------------------------
    # UNC authentication: establish net use session
    # ----------------------------------------
    $connectedShare = $null
    $useAuth = -not [string]::IsNullOrWhiteSpace($dest.AuthUser) -and
               -not [string]::IsNullOrWhiteSpace($dest.AuthPass)

    if ($useAuth) {
        if ("$($dest.Path)" -match '^(\\\\[^\\]+\\[^\\]+)') {
            $uncShare = $Matches[1]
            Show-Info "Authenticating: $uncShare (User: $($dest.AuthUser))"

            $null = & net use $uncShare "$($dest.AuthPass)" /user:"$($dest.AuthUser)" 2>&1
            $netExitCode = $LASTEXITCODE

            if ($netExitCode -ne 0) {
                Show-Error "net use failed (ExitCode=$netExitCode): $uncShare"
                $failCount++
                $details += "FAIL: $destDesc - net use authentication failed"
                $dest.AuthPass = $null
                Write-Host ""
                continue
            }

            Show-Success "Connected: $uncShare"
            $connectedShare = $uncShare
        }
    }

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
            Show-Success "logs/ copied"
        }

        # Copy evidence/
        if ($hasEvidence) {
            if ($useUnifiedPath) {
                # Unified mode: copy session evidence directly into destBase/evidence/
                $destEvidence = Join-Path $destBase "evidence"
                $null = New-Item -ItemType Directory -Path $destEvidence -Force
                $copyResult = robocopy $evidenceDir $destEvidence /E /NJH /NJS /NDL /NP /R:2 /W:1 2>&1
                Show-Success "evidence/ copied (session only)"
            }
            else {
                # Fallback: copy entire evidence/ directory
                $destEvidence = Join-Path $destBase "evidence"
                $null = New-Item -ItemType Directory -Path $destEvidence -Force
                $copyResult = robocopy $evidenceDir $destEvidence /E /NJH /NJS /NDL /NP /R:2 /W:1 2>&1
                Show-Success "evidence/ copied"
            }
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
    finally {
        # ----------------------------------------
        # UNC cleanup: disconnect share (silent)
        # ----------------------------------------
        if ($null -ne $connectedShare) {
            & net use $connectedShare /delete /y 2>&1 | Out-Null
        }

        # Clear credential variables from memory
        if ($useAuth) {
            $dest.AuthPass = $null
        }
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
