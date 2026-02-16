# ========================================
# Resolution Configuration Script (Live)
# ========================================
# Changes display resolution immediately using
# Win32 ChangeDisplaySettings API.
# No restart required.
# ========================================

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Resolution Configuration (Live)" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# ========================================
# Load C# Resolution Handler
# ========================================
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

public class ResolutionHandler {
    [DllImport("user32.dll")]
    public static extern int ChangeDisplaySettings(ref DEVMODE devMode, int flags);

    [DllImport("user32.dll")]
    public static extern int EnumDisplaySettings(string deviceName, int modeNum, ref DEVMODE devMode);

    [StructLayout(LayoutKind.Sequential)]
    public struct DEVMODE {
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
        public string dmDeviceName;
        public short dmSpecVersion;
        public short dmDriverVersion;
        public short dmSize;
        public short dmDriverExtra;
        public int dmFields;
        public int dmPositionX;
        public int dmPositionY;
        public int dmDisplayOrientation;
        public int dmDisplayFixedOutput;
        public short dmColor;
        public short dmDuplex;
        public short dmYResolution;
        public short dmTTOption;
        public short dmCollate;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
        public string dmFormName;
        public short dmLogPixels;
        public int dmBitsPerPel;
        public int dmPelsWidth;
        public int dmPelsHeight;
        public int dmDisplayFlags;
        public int dmDisplayFrequency;
        public int dmICMMethod;
        public int dmICMIntent;
        public int dmMediaType;
        public int dmDitherType;
        public int dmReserved1;
        public int dmReserved2;
        public int dmPanningWidth;
        public int dmPanningHeight;
    }

    public const int ENUM_CURRENT_SETTINGS = -1;
    public const int CDS_UPDATEREGISTRY = 0x01;
    public const int CDS_TEST = 0x02;
    public const int DISP_CHANGE_SUCCESSFUL = 0;
    public const int DISP_CHANGE_RESTART = 1;
    public const int DISP_CHANGE_FAILED = -1;
    public const int DM_PELSWIDTH = 0x80000;
    public const int DM_PELSHEIGHT = 0x100000;

    public static int[] GetCurrentResolution() {
        DEVMODE dm = new DEVMODE();
        dm.dmDeviceName = new String(new char[32]);
        dm.dmFormName = new String(new char[32]);
        dm.dmSize = (short)Marshal.SizeOf(dm);

        if (EnumDisplaySettings(null, ENUM_CURRENT_SETTINGS, ref dm) != 0) {
            return new int[] { dm.dmPelsWidth, dm.dmPelsHeight };
        }
        return new int[] { 0, 0 };
    }

    public static int ChangeRes(int width, int height) {
        DEVMODE dm = new DEVMODE();
        dm.dmDeviceName = new String(new char[32]);
        dm.dmFormName = new String(new char[32]);
        dm.dmSize = (short)Marshal.SizeOf(dm);

        dm.dmPelsWidth = width;
        dm.dmPelsHeight = height;
        dm.dmFields = DM_PELSWIDTH | DM_PELSHEIGHT;

        return ChangeDisplaySettings(ref dm, CDS_UPDATEREGISTRY);
    }
}
'@ -ErrorAction SilentlyContinue

# ========================================
# Load resolution_list.csv
# ========================================
$csvPath = Join-Path $PSScriptRoot "resolution_list.csv"

$csvData = Import-CsvSafe -Path $csvPath -Description "resolution_list.csv" -RequiredColumns @("Enabled", "Width", "Height")
if ($null -eq $csvData) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load resolution_list.csv")
}

$enabledItems = @($csvData | Where-Object { $_.Enabled -eq "1" })

if ($enabledItems.Count -eq 0) {
    Write-Host "[INFO] No enabled entries in resolution_list.csv" -ForegroundColor Yellow
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "No enabled entries")
}

# ========================================
# Show Current Resolution
# ========================================
$current = [ResolutionHandler]::GetCurrentResolution()
$currentW = $current[0]
$currentH = $current[1]

Write-Host "[INFO] Current resolution: $currentW x $currentH" -ForegroundColor Cyan
Write-Host ""

# ========================================
# Show Target Resolutions
# ========================================
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "Target Resolutions" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

$successCount = 0
$skipCount = 0
$errorCount = 0

$hasChanges = $false

foreach ($item in $enabledItems) {
    $targetW = [int]$item.Width
    $targetH = [int]$item.Height
    $desc    = if ($item.Description) { $item.Description } else { "" }

    if ($targetW -eq $currentW -and $targetH -eq $currentH) {
        Write-Host "  [SKIP] $targetW x $targetH  $desc (already set)" -ForegroundColor Gray
    }
    else {
        Write-Host "  [CHANGE] $targetW x $targetH  $desc" -ForegroundColor White
        $hasChanges = $true
    }
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

if (-not $hasChanges) {
    Write-Host "[INFO] All resolutions already match current settings" -ForegroundColor Green
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "All resolutions already match")
}

# ========================================
# Confirmation
# ========================================
if (-not (Confirm-Execution -Message "Apply the above resolution settings?")) {
    Write-Host ""
    Write-Host "[INFO] Canceled" -ForegroundColor Cyan
    Write-Host ""
    return (New-ModuleResult -Status "Cancelled" -Message "User canceled")
}

Write-Host ""

# ========================================
# Apply Resolution Changes
# ========================================
Write-Host "--- Applying Resolution Settings ---" -ForegroundColor Cyan
Write-Host ""

foreach ($item in $enabledItems) {
    $targetW = [int]$item.Width
    $targetH = [int]$item.Height
    $desc    = if ($item.Description) { " ($($item.Description))" } else { "" }

    if ($targetW -eq $currentW -and $targetH -eq $currentH) {
        Write-Host "[SKIP] $targetW x $targetH$desc - already set" -ForegroundColor Gray
        $skipCount++
        continue
    }

    Write-Host "[INFO] Changing resolution to $targetW x $targetH$desc..." -ForegroundColor Cyan

    try {
        $result = [ResolutionHandler]::ChangeRes($targetW, $targetH)

        switch ($result) {
            ([ResolutionHandler]::DISP_CHANGE_SUCCESSFUL) {
                Write-Host "[SUCCESS] Resolution changed to $targetW x $targetH" -ForegroundColor Green
                $successCount++

                # Update current resolution for subsequent checks
                $currentW = $targetW
                $currentH = $targetH
            }
            ([ResolutionHandler]::DISP_CHANGE_RESTART) {
                Write-Host "[WARNING] Resolution set but restart required to take effect" -ForegroundColor Yellow
                $successCount++
            }
            default {
                Write-Host "[ERROR] Failed to change resolution - unsupported resolution or hardware limitation" -ForegroundColor Red
                $errorCount++
            }
        }
    }
    catch {
        Write-Host "[ERROR] $($_.Exception.Message)" -ForegroundColor Red
        $errorCount++
    }

    Write-Host ""
}

# ========================================
# Result Summary
# ========================================
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Resolution Configuration Results" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
if ($successCount -gt 0) {
    Write-Host "Success: $successCount items" -ForegroundColor Green
}
if ($skipCount -gt 0) {
    Write-Host "Skipped: $skipCount items (already set)" -ForegroundColor Gray
}
if ($errorCount -gt 0) {
    Write-Host "Failed:  $errorCount items" -ForegroundColor Red
}
Write-Host ""

# Return ModuleResult
$overallStatus = if ($errorCount -eq 0 -and $successCount -gt 0) { "Success" }
    elseif ($errorCount -eq 0 -and $skipCount -gt 0 -and $successCount -eq 0) { "Skipped" }
    elseif ($successCount -gt 0 -and $errorCount -gt 0) { "Partial" }
    elseif ($errorCount -gt 0) { "Error" }
    else { "Success" }
return (New-ModuleResult -Status $overallStatus -Message "Success: $successCount, Skip: $skipCount, Fail: $errorCount")
