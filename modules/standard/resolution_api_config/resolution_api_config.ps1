# ========================================
# Resolution Configuration Script (Live)
# ========================================
# Changes display resolution immediately using
# Win32 ChangeDisplaySettings API.
# No restart required.
# ========================================

Write-Host ""
Show-Separator
Write-Host "Resolution Configuration (Live)" -ForegroundColor Cyan
Show-Separator
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

$enabledItems = Import-ModuleCsv -Path $csvPath -FilterEnabled -RequiredColumns @("Enabled", "Width", "Height")
if ($null -eq $enabledItems) { return (New-ModuleResult -Status "Error" -Message "Failed to load resolution_list.csv") }
if ($enabledItems.Count -eq 0) { return (New-ModuleResult -Status "Skipped" -Message "No enabled entries") }

# ========================================
# Show Current Resolution
# ========================================
$current = [ResolutionHandler]::GetCurrentResolution()
$currentW = $current[0]
$currentH = $current[1]

Show-Info "Current resolution: $currentW x $currentH"
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
$failCount = 0

$hasChanges = $false

foreach ($item in $enabledItems) {
    $targetW = [int]$item.Width
    $targetH = [int]$item.Height
    $desc    = if ($item.Description) { $item.Description } else { "" }

    if ($targetW -eq $currentW -and $targetH -eq $currentH) {
        Show-Skip "$targetW x $targetH  $desc (already set)"
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
    Show-Skip "All resolutions already match current settings"
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "All resolutions already match")
}

# ========================================
# Confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Apply the above resolution settings?"
if ($null -ne $cancelResult) { return $cancelResult }

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
        Show-Skip "$targetW x $targetH$desc - already set"
        $skipCount++
        continue
    }

    Show-Info "Changing resolution to $targetW x $targetH$desc..."

    try {
        $result = [ResolutionHandler]::ChangeRes($targetW, $targetH)

        switch ($result) {
            ([ResolutionHandler]::DISP_CHANGE_SUCCESSFUL) {
                Show-Success "Resolution changed to $targetW x $targetH"
                $successCount++

                # Update current resolution for subsequent checks
                $currentW = $targetW
                $currentH = $targetH
            }
            ([ResolutionHandler]::DISP_CHANGE_RESTART) {
                Show-Warning "Resolution set but restart required to take effect"
                $successCount++
            }
            default {
                Show-Error "Failed to change resolution - unsupported resolution or hardware limitation"
                $failCount++
            }
        }
    }
    catch {
        Show-Error "$($_.Exception.Message)"
        $failCount++
    }

    Write-Host ""
}

# ========================================
# Result Summary
# ========================================
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount -Title "Resolution Configuration Results")
