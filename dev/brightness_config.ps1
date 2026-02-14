# ==========================================
# CONFIGURATION
# Set the desired brightness level (0 - 100)
# ==========================================
$TargetBrightness = 80

# ==========================================
# FUNCTION: Set-Brightness (WMI)
# ==========================================
function Set-Brightness {
    param(
        [int]$Level
    )

    # Validate input range
    if ($Level -lt 0 -or $Level -gt 100) {
        Write-Host " [Error] Value must be between 0 and 100." -ForegroundColor Red
        return
    }

    try {
        # Get the WMI class for brightness control
        # Namespace: root\wmi
        # Class: WmiMonitorBrightnessMethods
        $monitor = Get-WmiObject -Namespace root\wmi -Class WmiMonitorBrightnessMethods -ErrorAction Stop
        
        if ($monitor) {
            # Method: WmiSetBrightness(Timeout, Brightness)
            # Timeout is set to 1 second.
            $monitor.WmiSetBrightness(1, $Level)
            
            Write-Host " [OK] Brightness successfully set to ${Level}%" -ForegroundColor Green
        }
    }
    catch {
        # This error usually occurs on Desktop PCs with external monitors
        # because they do not support WMI brightness control.
        Write-Host " [Skip] Could not set brightness. This device might not support WMI brightness control (e.g., Desktop PC)." -ForegroundColor Yellow
        # Uncomment the line below to see the detailed error message for debugging:
        # Write-Host " Debug Info: $_" -ForegroundColor DarkGray
    }
}

# ==========================================
# MAIN EXECUTION
# ==========================================
Write-Host "=== Starting Brightness Configuration ===" -ForegroundColor Cyan

Write-Host "Target Level: ${TargetBrightness}%"
Set-Brightness -Level $TargetBrightness

Write-Host "=== Configuration Complete ===" -ForegroundColor Cyan