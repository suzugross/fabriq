# Use UTF-8 encoding
$utf8 = [System.Text.Encoding]::UTF8
$lines = [System.IO.File]::ReadAllLines('main.ps1', $utf8)

$newLines = @()
$logAdded = $false

for ($i = 0; $i -lt $lines.Count; $i++) {
    $line = $lines[$i]

    # Find "# Main Process" section and add log code after it
    # Note: Matching English text as 'main.ps1' is assumed to be translated
    if ($line -match 'Main Process' -and $i -gt 200 -and -not $logAdded) {
        $newLines += $line

        # Add next 2 lines (separator and empty line)
        if ($i + 1 -lt $lines.Count) {
            $newLines += $lines[$i + 1]
            $i++
        }
        if ($i + 1 -lt $lines.Count) {
            $newLines += $lines[$i + 1]
            $i++
        }

        # Add log start code
        $newLines += ''
        $newLines += '# ========================================'
        $newLines += '# Start Logging'
        $newLines += '# ========================================'
        $newLines += '$logDir = ".\logs"'
        $newLines += 'if (-not (Test-Path $logDir)) {'
        $newLines += '    New-Item -ItemType Directory -Path $logDir -Force | Out-Null'
        $newLines += '}'
        $newLines += '$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"'
        $newLines += '$logFile = Join-Path $logDir "log_$timestamp.txt"'
        $newLines += 'Start-Transcript -Path $logFile -Append | Out-Null'
        $newLines += 'Write-Host "[INFO] Log file: $logFile" -ForegroundColor Cyan'
        $newLines += 'Write-Host ""'

        $logAdded = $true
        continue
    }

    # Add log stop code after the last Show-Separator
    if ($line -match '^Show-Separator$' -and $i -eq ($lines.Count - 1)) {
        $newLines += $line
        $newLines += ''
        $newLines += '# Stop Logging'
        $newLines += 'Stop-Transcript | Out-Null'
        continue
    }

    $newLines += $line
}

# Save as UTF-8
[System.IO.File]::WriteAllLines('main.ps1', $newLines, $utf8)

Write-Host "Successfully added logging functionality"