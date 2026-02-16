# =============================================================================
# Native DPI Scaling Configuration Module (Fixed for compatibility)
# =============================================================================
# Uses undocumented Windows APIs to set DPI scaling directly from PowerShell
# without external executables.
# Based on SetDPI logic.
# =============================================================================

# Check Administrator Privileges
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "Administrator privileges are required."
    return (New-ModuleResult -Status "Error" -Message "Administrator privileges required")
}

# ========================================
# C# Native Class Definition
# ========================================
$dpiSource = @"
using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Linq;

public class NativeDpiHelper {
    
    // --- Constants & Enums ---
    public const int S_OK = 0;
    public const int ERROR_SUCCESS = 0;
    
    // Undocumented constants for DPI scaling
    public enum DISPLAYCONFIG_DEVICE_INFO_TYPE : int {
        GET_DPI_SCALE = -3,
        SET_DPI_SCALE = -4
    }

    public enum DISPLAYCONFIG_MODE_INFO_TYPE : uint {
        SOURCE = 1,
        TARGET = 2,
        DESKTOP_IMAGE = 3
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
        // Union part simplified for size matching (mode info is large)
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
    
    // Standard DPI values used by Windows
    private static readonly int[] DpiVals = { 100, 125, 150, 175, 200, 225, 250, 300, 350, 400, 450, 500 };

    public static string SetDpi(int monitorIndex, int scalePercent) {
        // 1. Get Paths
        uint numPaths = 0, numModes = 0;
        uint QDC_ONLY_ACTIVE_PATHS = 0x00000002;
        
        int status = GetDisplayConfigBufferSizes(QDC_ONLY_ACTIVE_PATHS, out numPaths, out numModes);
        if (status != ERROR_SUCCESS) return "Error: GetDisplayConfigBufferSizes failed: " + status;

        var paths = new DISPLAYCONFIG_PATH_INFO[numPaths];
        var modes = new DISPLAYCONFIG_MODE_INFO[numModes];
        
        status = QueryDisplayConfig(QDC_ONLY_ACTIVE_PATHS, ref numPaths, paths, ref numModes, modes, IntPtr.Zero);
        if (status != ERROR_SUCCESS) return "Error: QueryDisplayConfig failed: " + status;

        if (monitorIndex < 0 || monitorIndex >= numPaths) {
            return "Error: Invalid Monitor Index. Available: " + numPaths;
        }

        // Target Adapter & ID
        LUID adapterId = paths[monitorIndex].sourceInfo.adapterId;
        uint sourceId = paths[monitorIndex].sourceInfo.id;

        // 2. Get Current/Recommended Info to calculate relative step
        var getPacket = new DISPLAYCONFIG_SOURCE_DPI_SCALE_GET();
        getPacket.header.type = (int)DISPLAYCONFIG_DEVICE_INFO_TYPE.GET_DPI_SCALE;
        getPacket.header.size = (uint)Marshal.SizeOf(typeof(DISPLAYCONFIG_SOURCE_DPI_SCALE_GET));
        getPacket.header.adapterId = adapterId;
        getPacket.header.id = sourceId;

        status = DisplayConfigGetDeviceInfo(ref getPacket);
        if (status != ERROR_SUCCESS) return "Error: DisplayConfigGetDeviceInfo failed: " + status;

        // Determine min absolute index in DpiVals
        // minScaleRel is usually negative (steps down from recommended)
        // recommended is at index [minAbs]
        int minAbs = Math.Abs(getPacket.minScaleRel);
        int recommendedIdx = minAbs;
        
        // Find target index
        int targetIdx = -1;
        for (int i = 0; i < DpiVals.Length; i++) {
            if (DpiVals[i] == scalePercent) {
                targetIdx = i;
                break;
            }
        }

        if (targetIdx == -1) return "Error: Unsupported Scale Percent: " + scalePercent;

        // Calculate relative value (Target - Recommended)
        int stepRel = targetIdx - recommendedIdx;

        // Check bounds
        if (stepRel < getPacket.minScaleRel || stepRel > getPacket.maxScaleRel) {
            // FIXED: Using string concatenation instead of interpolation ($) for compatibility
            return "Error: Scale " + scalePercent + "% is out of supported bounds for this monitor.";
        }
        
        // 3. Set DPI
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

# Compile C# Code
try {
    Add-Type -TypeDefinition $dpiSource -Language CSharp
}
catch {
    Write-Host "[ERROR] Failed to compile NativeDpiHelper type." -ForegroundColor Red
    Write-Host $_.Exception.Message
    return (New-ModuleResult -Status "Error" -Message "C# Compilation Failed")
}

# ========================================
# Main Logic
# ========================================

# Check if running as a module or standalone script
$scriptDir = $PSScriptRoot
if (-not $scriptDir) { $scriptDir = $PWD }
$csvPath = Join-Path $scriptDir "dpi_list.csv"

if (-not (Test-Path $csvPath)) {
    Write-Warning "dpi_list.csv not found at $csvPath"
    return (New-ModuleResult -Status "Error" -Message "dpi_list.csv not found")
}

$items = Import-Csv -Path $csvPath -Encoding Default | Where-Object { $_.Enabled -eq "1" }
if ($items.Count -eq 0) {
    return (New-ModuleResult -Status "Skipped" -Message "No enabled entries")
}

Write-Host "--- Native DPI Configuration ---" -ForegroundColor Cyan

$successCount = 0
$failCount = 0

foreach ($item in $items) {
    $idx = [int]$item.MonitorIndex
    $scale = [int]$item.ScalePercent
    $desc = $item.Description

    Write-Host "Applying: Monitor[$idx] -> $scale% ($desc)" -NoNewline

    try {
        # Call the C# static method
        $result = [NativeDpiHelper]::SetDpi($idx, $scale)

        if ($result -eq "Success") {
            Write-Host " [OK]" -ForegroundColor Green
            $successCount++
        }
        else {
            Write-Host " [FAILED]" -ForegroundColor Red
            Write-Host "  Reason: $result" -ForegroundColor Red
            $failCount++
        }
    }
    catch {
        Write-Host " [ERROR]" -ForegroundColor Red
        Write-Host "  Exception: $_" -ForegroundColor Red
        $failCount++
    }
}

Write-Host ""
$status = if ($failCount -eq 0) { "Success" } else { "Error" }
return (New-ModuleResult -Status $status -Message "Success: $successCount, Fail: $failCount")