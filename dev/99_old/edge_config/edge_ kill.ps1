<#
 Script: Kill Microsoft Edge
 Description: Forcefully terminates all Edge processes.
#>

$processName = "msedge"

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "      KILL EDGE PROCESSES" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan

Write-Host "Checking for running Edge processes..."

$running = Get-Process -Name $processName -ErrorAction SilentlyContinue

if ($running) {
    Write-Host "Found $($running.Count) active processes. Terminating..." -ForegroundColor Yellow
    
    try {
        # Force kill process tree
        Stop-Process -Name $processName -Force -ErrorAction Stop
        Start-Sleep -Seconds 2
        Write-Host "Success: Edge has been terminated." -ForegroundColor Green
    }
    catch {
        Write-Host "Error: Could not kill Edge. Please close it manually." -ForegroundColor Red
    }
} else {
    Write-Host "No active Edge processes found." -ForegroundColor Green
}

Write-Host "`nPress Enter to exit..."
Read-Host