# ========================================
# DPI Scaling Configuration Script (Live)
# ========================================
# Uses undocumented Windows APIs to set DPI
# scaling directly without restart.
# Based on SetDPI logic.
# ========================================

Write-Host ""
Show-Separator
Write-Host "DPI Scaling Configuration (Live)" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# ========================================
# C# Native Class Definition
# ========================================
$dpiSource = @"
using System;
using System.Runtime.InteropServices;

public class NativeDpiHelper {

    // --- Constants & Enums ---
    public const int S_OK = 0;
    public const int ERROR_SUCCESS = 0;

    public enum DISPLAYCONFIG_DEVICE_INFO_TYPE : int {
        GET_DPI_SCALE = -3,
        SET_DPI_SCALE = -4
    }

    // --- Structures ---
    [StructLayout(LayoutKind.Sequential)]
    public struct LUID {
        public uint LowPart;
        public int HighPart;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct DISPLAYCONFIG_PATH_SOURCE_INFO {
        public LUID adapterId;
        public uint id;
        public uint modeInfoIdx;
        public uint statusFlags;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct DISPLAYCONFIG_PATH_TARGET_INFO {
        public LUID adapterId;
        public uint id;
        public uint modeInfoIdx;
        public uint outputTechnology;
        public uint rotation;
        public uint scaling;
        public uint refreshRateNumerator;
        public uint refreshRateDenominator;
        public uint statusFlags;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct DISPLAYCONFIG_PATH_INFO {
        public DISPLAYCONFIG_PATH_SOURCE_INFO sourceInfo;
        public DISPLAYCONFIG_PATH_TARGET_INFO targetInfo;
        public uint flags;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct DISPLAYCONFIG_MODE_INFO {
        public uint infoType;
        public uint id;
        public LUID adapterId;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 64)]
        public byte[] modeInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct DISPLAYCONFIG_DEVICE_INFO_HEADER {
        public int type;
        public uint size;
        public LUID adapterId;
        public uint id;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct DISPLAYCONFIG_SOURCE_DPI_SCALE_GET {
        public DISPLAYCONFIG_DEVICE_INFO_HEADER header;
        public int minScaleRel;
        public int curScaleRel;
        public int maxScaleRel;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct DISPLAYCONFIG_SOURCE_DPI_SCALE_SET {
        public DISPLAYCONFIG_DEVICE_INFO_HEADER header;
        public int scaleRel;
    }

    // --- P/Invoke ---
    [DllImport("user32.dll")]
    public static extern int GetDisplayConfigBufferSizes(uint flags, out uint numPathArrayElements, out uint numModeInfoArrayElements);

    [DllImport("user32.dll")]
    public static extern int QueryDisplayConfig(uint flags, ref uint numPathArrayElements, [Out] DISPLAYCONFIG_PATH_INFO[] pathArray, ref uint numModeInfoArrayElements, [Out] DISPLAYCONFIG_MODE_INFO[] modeInfoArray, IntPtr currentTopologyId);

    [DllImport("user32.dll")]
    public static extern int DisplayConfigGetDeviceInfo(ref DISPLAYCONFIG_SOURCE_DPI_SCALE_GET requestPacket);

    [DllImport("user32.dll")]
    public static extern int DisplayConfigSetDeviceInfo(ref DISPLAYCONFIG_SOURCE_DPI_SCALE_SET requestPacket);

    // --- Logic ---

    private static readonly int[] DpiVals = { 100, 125, 150, 175, 200, 225, 250, 300, 350, 400, 450, 500 };

    private static bool GetPaths(out DISPLAYCONFIG_PATH_INFO[] paths, out uint numPaths) {
        uint numModes = 0;
        numPaths = 0;
        paths = null;
        uint QDC_ONLY_ACTIVE_PATHS = 0x00000002;

        int status = GetDisplayConfigBufferSizes(QDC_ONLY_ACTIVE_PATHS, out numPaths, out numModes);
        if (status != ERROR_SUCCESS) return false;

        paths = new DISPLAYCONFIG_PATH_INFO[numPaths];
        var modes = new DISPLAYCONFIG_MODE_INFO[numModes];

        status = QueryDisplayConfig(QDC_ONLY_ACTIVE_PATHS, ref numPaths, paths, ref numModes, modes, IntPtr.Zero);
        return status == ERROR_SUCCESS;
    }

    public static int GetCurrentDpi(int monitorIndex) {
        DISPLAYCONFIG_PATH_INFO[] paths;
        uint numPaths;
        if (!GetPaths(out paths, out numPaths)) return -1;
        if (monitorIndex < 0 || monitorIndex >= numPaths) return -1;

        var getPacket = new DISPLAYCONFIG_SOURCE_DPI_SCALE_GET();
        getPacket.header.type = (int)DISPLAYCONFIG_DEVICE_INFO_TYPE.GET_DPI_SCALE;
        getPacket.header.size = (uint)Marshal.SizeOf(typeof(DISPLAYCONFIG_SOURCE_DPI_SCALE_GET));
        getPacket.header.adapterId = paths[monitorIndex].sourceInfo.adapterId;
        getPacket.header.id = paths[monitorIndex].sourceInfo.id;

        int status = DisplayConfigGetDeviceInfo(ref getPacket);
        if (status != ERROR_SUCCESS) return -1;

        int recommendedIdx = Math.Abs(getPacket.minScaleRel);
        int currentIdx = recommendedIdx + getPacket.curScaleRel;

        if (currentIdx >= 0 && currentIdx < DpiVals.Length) {
            return DpiVals[currentIdx];
        }
        return -1;
    }

    public static int GetMonitorCount() {
        DISPLAYCONFIG_PATH_INFO[] paths;
        uint numPaths;
        if (!GetPaths(out paths, out numPaths)) return 0;
        return (int)numPaths;
    }

    public static string SetDpi(int monitorIndex, int scalePercent) {
        DISPLAYCONFIG_PATH_INFO[] paths;
        uint numPaths;
        if (!GetPaths(out paths, out numPaths))
            return "Error: Failed to query display configuration";

        if (monitorIndex < 0 || monitorIndex >= numPaths) {
            return "Error: Invalid Monitor Index. Available: " + numPaths;
        }

        LUID adapterId = paths[monitorIndex].sourceInfo.adapterId;
        uint sourceId = paths[monitorIndex].sourceInfo.id;

        var getPacket = new DISPLAYCONFIG_SOURCE_DPI_SCALE_GET();
        getPacket.header.type = (int)DISPLAYCONFIG_DEVICE_INFO_TYPE.GET_DPI_SCALE;
        getPacket.header.size = (uint)Marshal.SizeOf(typeof(DISPLAYCONFIG_SOURCE_DPI_SCALE_GET));
        getPacket.header.adapterId = adapterId;
        getPacket.header.id = sourceId;

        int status = DisplayConfigGetDeviceInfo(ref getPacket);
        if (status != ERROR_SUCCESS) return "Error: DisplayConfigGetDeviceInfo failed: " + status;

        int minAbs = Math.Abs(getPacket.minScaleRel);
        int recommendedIdx = minAbs;

        int targetIdx = -1;
        for (int i = 0; i < DpiVals.Length; i++) {
            if (DpiVals[i] == scalePercent) {
                targetIdx = i;
                break;
            }
        }

        if (targetIdx == -1) return "Error: Unsupported Scale Percent: " + scalePercent;

        int stepRel = targetIdx - recommendedIdx;

        if (stepRel < getPacket.minScaleRel || stepRel > getPacket.maxScaleRel) {
            return "Error: Scale " + scalePercent + "% is out of supported bounds for this monitor.";
        }

        var setPacket = new DISPLAYCONFIG_SOURCE_DPI_SCALE_SET();
        setPacket.header.type = (int)DISPLAYCONFIG_DEVICE_INFO_TYPE.SET_DPI_SCALE;
        setPacket.header.size = (uint)Marshal.SizeOf(typeof(DISPLAYCONFIG_SOURCE_DPI_SCALE_SET));
        setPacket.header.adapterId = adapterId;
        setPacket.header.id = sourceId;
        setPacket.scaleRel = stepRel;

        status = DisplayConfigSetDeviceInfo(ref setPacket);
        if (status == ERROR_SUCCESS) {
            return "Success";
        } else {
            return "Error: DisplayConfigSetDeviceInfo failed: " + status;
        }
    }
}
"@

try {
    Add-Type -TypeDefinition $dpiSource -Language CSharp -ErrorAction SilentlyContinue
}
catch {
    Show-Error "Failed to compile NativeDpiHelper: $($_.Exception.Message)"
    return (New-ModuleResult -Status "Error" -Message "C# compilation failed")
}

# ========================================
# Load dpi_list.csv
# ========================================
$csvPath = Join-Path $PSScriptRoot "dpi_list.csv"

$enabledItems = Import-ModuleCsv -Path $csvPath -FilterEnabled -RequiredColumns @("Enabled", "MonitorIndex", "ScalePercent")
if ($null -eq $enabledItems) { return (New-ModuleResult -Status "Error" -Message "Failed to load dpi_list.csv") }
if ($enabledItems.Count -eq 0) { return (New-ModuleResult -Status "Skipped" -Message "No enabled entries") }

# ========================================
# Show Current DPI & Monitor Info
# ========================================
$monitorCount = [NativeDpiHelper]::GetMonitorCount()
Show-Info "Active monitors: $monitorCount"

for ($i = 0; $i -lt $monitorCount; $i++) {
    $currentDpi = [NativeDpiHelper]::GetCurrentDpi($i)
    $dpiStr = if ($currentDpi -gt 0) { "${currentDpi}%" } else { "Unknown" }
    Show-Info "Monitor[$i] current DPI: $dpiStr"
}

Write-Host ""

# ========================================
# Show Target Settings
# ========================================
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "Target DPI Settings" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

$successCount = 0
$skipCount = 0
$failCount = 0

$hasChanges = $false

foreach ($item in $enabledItems) {
    $idx   = [int]$item.MonitorIndex
    $scale = [int]$item.ScalePercent
    $desc  = if ($item.Description) { $item.Description } else { "" }

    $currentDpi = [NativeDpiHelper]::GetCurrentDpi($idx)

    if ($currentDpi -eq $scale) {
        Show-Skip "Monitor[$idx] -> ${scale}%  $desc (already set)"
    }
    else {
        $currentStr = if ($currentDpi -gt 0) { "${currentDpi}%" } else { "Unknown" }
        Write-Host "  [CHANGE] Monitor[$idx] $currentStr -> ${scale}%  $desc" -ForegroundColor White
        $hasChanges = $true
    }
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

if (-not $hasChanges) {
    Show-Skip "All DPI settings already match current values"
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "All DPI settings already match")
}

# ========================================
# Confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Apply the above DPI settings?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""

# ========================================
# Apply DPI Changes
# ========================================
Write-Host "--- Applying DPI Settings ---" -ForegroundColor Cyan
Write-Host ""

foreach ($item in $enabledItems) {
    $idx   = [int]$item.MonitorIndex
    $scale = [int]$item.ScalePercent
    $desc  = if ($item.Description) { " ($($item.Description))" } else { "" }

    $currentDpi = [NativeDpiHelper]::GetCurrentDpi($idx)

    if ($currentDpi -eq $scale) {
        Show-Skip "Monitor[$idx] -> ${scale}%$desc - already set"
        $skipCount++
        continue
    }

    Show-Info "Setting Monitor[$idx] -> ${scale}%$desc..."

    try {
        $result = [NativeDpiHelper]::SetDpi($idx, $scale)

        if ($result -eq "Success") {
            Show-Success "Monitor[$idx] DPI changed to ${scale}%"
            $successCount++
        }
        else {
            Show-Error "$result"
            $failCount++
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
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount -Title "DPI Scaling Results")
