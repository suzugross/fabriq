<#
 Script: Restore Edge Settings (Native Execution Fix)
 Description: Restores User Data from the 'backup' folder in the script directory.
#>

# --- CONFIGURATION ---------------------------------------------------
$ScriptDir = $PSScriptRoot
$BackupDir = Join-Path -Path $ScriptDir -ChildPath "backup"
$TargetDir = "$env:LOCALAPPDATA\Microsoft\Edge\User Data"
# ---------------------------------------------------------------------

Write-Host "==========================================" -ForegroundColor Magenta
Write-Host "      RESTORE EDGE SETTINGS" -ForegroundColor Magenta
Write-Host "==========================================" -ForegroundColor Magenta

# Check if Edge is running
if (Get-Process -Name "msedge" -ErrorAction SilentlyContinue) {
    Write-Host "Error: Edge is running. Please run the KILL script first." -ForegroundColor Red
    Read-Host "Press Enter to exit..."
    exit
}

# Check if backup exists
if (!(Test-Path -Path $BackupDir)) {
    Write-Host "Error: Backup folder not found at: $BackupDir" -ForegroundColor Red
    Write-Host "Make sure the 'backup' folder is in the same directory as this script."
    Read-Host "Press Enter to exit..."
    exit
}

Write-Host "Source:      $BackupDir"
Write-Host "Destination: $TargetDir"
Write-Host "WARNING: This will OVERWRITE your current Edge settings." -ForegroundColor Yellow
$confirm = Read-Host "Type 'YES' to continue"

if ($confirm -eq 'YES') {
    Write-Host "`nRestoring..."
    
    # Execute Robocopy directly using the Call Operator (&)
    # Quotes around variables ensure paths with spaces are treated as single arguments.
    & robocopy.exe "$BackupDir" "$TargetDir" /MIR /XJ /MT /R:1 /W:1 /NFL /NDL

    $exitCode = $LASTEXITCODE
    
    if ($exitCode -lt 8) {
        Write-Host "`n[SUCCESS] Restore completed successfully." -ForegroundColor Green
    } else {
        Write-Host "`n[ERROR] Restore failed. Exit Code: $exitCode" -ForegroundColor Red
    }
} else {
    Write-Host "Restore cancelled." -ForegroundColor Gray
}

Write-Host "`nPress Enter to exit..."
Read-Host