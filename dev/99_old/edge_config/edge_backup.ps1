<#
 Script: Backup Edge Settings (Native Execution Fix)
 Description: Mirrors User Data to a 'backup' folder in the script directory.
#>

# --- CONFIGURATION ---------------------------------------------------
$ScriptDir = $PSScriptRoot
$BackupDir = Join-Path -Path $ScriptDir -ChildPath "backup"
$SourceDir = "$env:LOCALAPPDATA\Microsoft\Edge\User Data"
# ---------------------------------------------------------------------

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "      BACKUP EDGE SETTINGS" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan

# Check if Edge is running
if (Get-Process -Name "msedge" -ErrorAction SilentlyContinue) {
    Write-Host "Error: Edge is running. Please run the KILL script first." -ForegroundColor Red
    Read-Host "Press Enter to exit..."
    exit
}

# Create backup directory if it doesn't exist
if (!(Test-Path -Path $BackupDir)) {
    try {
        New-Item -ItemType Directory -Path $BackupDir -Force | Out-Null
        Write-Host "Created backup directory: $BackupDir" -ForegroundColor Gray
    }
    catch {
        Write-Host "Error: Failed to create backup directory. Check permissions." -ForegroundColor Red
        Read-Host "Press Enter to exit..."
        exit
    }
}

Write-Host "Source:      $SourceDir"
Write-Host "Destination: $BackupDir"
Write-Host "Mode:        MIRROR (Exact Copy)"
Write-Host "------------------------------------------"

# Execute Robocopy directly using the Call Operator (&)
# This handles spaces in variables correctly without complex escaping.
& robocopy.exe "$SourceDir" "$BackupDir" /MIR /XJ /MT /R:1 /W:1 /NFL /NDL

# Capture Exit Code from the last command
$exitCode = $LASTEXITCODE

# Robocopy Exit Codes: 0-7 are success (0=No Change, 1=Copy Successful, etc.)
if ($exitCode -lt 8) {
    Write-Host "`n[SUCCESS] Backup completed successfully." -ForegroundColor Green
} else {
    Write-Host "`n[ERROR] Backup failed. Exit Code: $exitCode" -ForegroundColor Red
}

Write-Host "`nPress Enter to exit..."
Read-Host