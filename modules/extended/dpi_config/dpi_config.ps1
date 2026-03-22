# ========================================
# Display DPI Scaling Configuration Script
# ========================================
# Modifies per-monitor DPI scaling via registry
# (HKCU + Default User Hive) based on CSV
# configuration (dpi_list.csv).
# Sign-out or restart required for changes.
# ========================================

# Check Administrator Privileges
if (-not (Test-AdminPrivilege)) {
    Show-Error "This script requires administrator privileges."
    return (New-ModuleResult -Status "Error" -Message "Administrator privileges required")
}

# Resolve logged-on user's HKCU target
$hkcuInfo = Resolve-HkcuRoot

Write-Host ""
Show-Separator
Write-Host "Display DPI Scaling Configuration" -ForegroundColor Cyan
if ($hkcuInfo.Redirected) {
    Write-Host "  Target: $($hkcuInfo.Label)" -ForegroundColor Magenta
}
Show-Separator
Write-Host ""

# ========================================
# Constants
# ========================================
$script:PerMonitorBasePath = $hkcuInfo.PsDrivePath + '\Control Panel\Desktop\PerMonitorSettings'
$script:GraphicsConfigPath = 'HKLM:\SYSTEM\CurrentControlSet\Control\GraphicsDrivers\Configuration'
$HIVE_PATH = "$env:SystemDrive\Users\Default\ntuser.dat"
$HIVE_KEY = "HKEY_USERS\Hive"

# Supported scale values
$script:SupportedScales = @(100, 125, 150, 175, 200)

# ========================================
# C# DPI Scale Resolver
# ========================================
# Uses DisplayConfig API to determine each monitor's
# recommended DPI and compute correct relative DpiValue.
# Class name differs from dpi_api_config's NativeDpiHelper
# to avoid type conflicts in the same session.
# ========================================
$dpiResolverSource = @"
using System;
using System.Runtime.InteropServices;

public class DpiScaleResolver {

    public const int ERROR_SUCCESS = 0;

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

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct DISPLAYCONFIG_TARGET_DEVICE_NAME {
        public DISPLAYCONFIG_DEVICE_INFO_HEADER header;
        public uint flags;
        public uint outputTechnology;
        public ushort edidManufactureId;
        public ushort edidProductCodeId;
        public uint connectorInstance;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 64)]
        public string monitorFriendlyDeviceName;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)]
        public string monitorDevicePath;
    }

    // --- P/Invoke ---
    [DllImport("user32.dll")]
    public static extern int GetDisplayConfigBufferSizes(uint flags, out uint numPathArrayElements, out uint numModeInfoArrayElements);

    [DllImport("user32.dll")]
    public static extern int QueryDisplayConfig(uint flags, ref uint numPathArrayElements, [Out] DISPLAYCONFIG_PATH_INFO[] pathArray, ref uint numModeInfoArrayElements, [Out] DISPLAYCONFIG_MODE_INFO[] modeInfoArray, IntPtr currentTopologyId);

    [DllImport("user32.dll")]
    public static extern int DisplayConfigGetDeviceInfo(ref DISPLAYCONFIG_SOURCE_DPI_SCALE_GET requestPacket);

    [DllImport("user32.dll")]
    public static extern int DisplayConfigGetDeviceInfo(ref DISPLAYCONFIG_TARGET_DEVICE_NAME requestPacket);

    // --- Constants ---
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

    public static int GetMonitorCount() {
        DISPLAYCONFIG_PATH_INFO[] paths;
        uint numPaths;
        if (!GetPaths(out paths, out numPaths)) return 0;
        return (int)numPaths;
    }

    public static int GetRecommendedDpiPercent(int monitorIndex) {
        DISPLAYCONFIG_PATH_INFO[] paths;
        uint numPaths;
        if (!GetPaths(out paths, out numPaths)) return -1;
        if (monitorIndex < 0 || monitorIndex >= numPaths) return -1;

        var getPacket = new DISPLAYCONFIG_SOURCE_DPI_SCALE_GET();
        getPacket.header.type = -3; // GET_DPI_SCALE
        getPacket.header.size = (uint)Marshal.SizeOf(typeof(DISPLAYCONFIG_SOURCE_DPI_SCALE_GET));
        getPacket.header.adapterId = paths[monitorIndex].sourceInfo.adapterId;
        getPacket.header.id = paths[monitorIndex].sourceInfo.id;

        int status = DisplayConfigGetDeviceInfo(ref getPacket);
        if (status != ERROR_SUCCESS) return -1;

        int recommendedIdx = Math.Abs(getPacket.minScaleRel);
        if (recommendedIdx >= 0 && recommendedIdx < DpiVals.Length) {
            return DpiVals[recommendedIdx];
        }
        return -1;
    }

    public static string GetMonitorHardwareId(int monitorIndex) {
        DISPLAYCONFIG_PATH_INFO[] paths;
        uint numPaths;
        if (!GetPaths(out paths, out numPaths)) return null;
        if (monitorIndex < 0 || monitorIndex >= numPaths) return null;

        var namePacket = new DISPLAYCONFIG_TARGET_DEVICE_NAME();
        namePacket.header.type = 2; // DISPLAYCONFIG_DEVICE_INFO_GET_TARGET_NAME
        namePacket.header.size = (uint)Marshal.SizeOf(typeof(DISPLAYCONFIG_TARGET_DEVICE_NAME));
        namePacket.header.adapterId = paths[monitorIndex].targetInfo.adapterId;
        namePacket.header.id = paths[monitorIndex].targetInfo.id;

        int status = DisplayConfigGetDeviceInfo(ref namePacket);
        if (status != ERROR_SUCCESS) return null;

        if (string.IsNullOrEmpty(namePacket.monitorDevicePath)) return null;

        // Device path format: \\?\DISPLAY#CMN14D4#5&xxx&0&UID256#{guid}
        string[] parts = namePacket.monitorDevicePath.Split(new char[] { '#' });
        if (parts.Length >= 3) return parts[1];
        return null;
    }

    public static int ComputeDpiValue(int recommendedPercent, int targetPercent) {
        int recIdx = -1;
        int targetIdx = -1;
        for (int i = 0; i < DpiVals.Length; i++) {
            if (DpiVals[i] == recommendedPercent) recIdx = i;
            if (DpiVals[i] == targetPercent) targetIdx = i;
        }
        if (recIdx < 0 || targetIdx < 0) return int.MinValue;
        return targetIdx - recIdx;
    }

    public static int ComputePercent(int recommendedPercent, int dpiValue) {
        int recIdx = -1;
        for (int i = 0; i < DpiVals.Length; i++) {
            if (DpiVals[i] == recommendedPercent) { recIdx = i; break; }
        }
        if (recIdx < 0) return -1;
        int idx = recIdx + dpiValue;
        if (idx >= 0 && idx < DpiVals.Length) return DpiVals[idx];
        return -1;
    }
}
"@

$script:DpiResolverAvailable = $false
try {
    Add-Type -TypeDefinition $dpiResolverSource -Language CSharp -ErrorAction SilentlyContinue
    $script:DpiResolverAvailable = $true
}
catch {
    # Type may already be loaded from a previous run
    try {
        $null = [DpiScaleResolver]::GetMonitorCount()
        $script:DpiResolverAvailable = $true
    }
    catch {
        Show-Warning "DpiScaleResolver compilation failed. Using fallback (150% recommended assumed)."
    }
}

# ========================================
# Build Monitor Recommended DPI Map
# ========================================
# Maps HardwareID prefix -> Recommended DPI percent
$script:MonitorRecommendedMap = @{}
$script:FallbackRecommended = 150

if ($script:DpiResolverAvailable) {
    try {
        $monCount = [DpiScaleResolver]::GetMonitorCount()
        for ($i = 0; $i -lt $monCount; $i++) {
            $hwId = [DpiScaleResolver]::GetMonitorHardwareId($i)
            $rec  = [DpiScaleResolver]::GetRecommendedDpiPercent($i)
            if ($hwId -and $rec -gt 0) {
                $script:MonitorRecommendedMap[$hwId] = $rec
                Show-Info "Monitor[$i] $hwId : Recommended ${rec}%"
            }
        }
    }
    catch {
        Show-Warning "Monitor enumeration failed: $($_.Exception.Message)"
    }
}

# ========================================
# Helper Functions
# ========================================

function Get-RecommendedPercent {
    param([string]$HardwareID)
    if ($script:MonitorRecommendedMap.Count -gt 0 -and $HardwareID) {
        foreach ($key in $script:MonitorRecommendedMap.Keys) {
            if ($HardwareID -like "$key*" -or $key -like "$HardwareID*") {
                return $script:MonitorRecommendedMap[$key]
            }
        }
    }
    return $script:FallbackRecommended
}

function Get-DpiValueForScale {
    param(
        [int]$ScalePercent,
        [string]$HardwareID = ""
    )
    $rec = Get-RecommendedPercent -HardwareID $HardwareID
    if ($script:DpiResolverAvailable) {
        $val = [DpiScaleResolver]::ComputeDpiValue($rec, $ScalePercent)
        if ($val -ne [int]::MinValue) { return $val }
    }
    # Fallback: manual calculation
    $dpiVals = @(100, 125, 150, 175, 200, 225, 250, 300, 350, 400, 450, 500)
    $recIdx = [Array]::IndexOf($dpiVals, $rec)
    $targetIdx = [Array]::IndexOf($dpiVals, $ScalePercent)
    if ($recIdx -ge 0 -and $targetIdx -ge 0) { return ($targetIdx - $recIdx) }
    return $null
}

function Convert-DpiToPercent {
    param(
        [int]$DpiValue,
        [string]$HardwareID = ""
    )
    $rec = Get-RecommendedPercent -HardwareID $HardwareID
    if ($script:DpiResolverAvailable) {
        $percent = [DpiScaleResolver]::ComputePercent($rec, $DpiValue)
        if ($percent -gt 0) { return "${percent}%" }
    }
    # Fallback: manual calculation
    $dpiVals = @(100, 125, 150, 175, 200, 225, 250, 300, 350, 400, 450, 500)
    $recIdx = [Array]::IndexOf($dpiVals, $rec)
    if ($recIdx -ge 0) {
        $idx = $recIdx + $DpiValue
        if ($idx -ge 0 -and $idx -lt $dpiVals.Length) { return "$($dpiVals[$idx])%" }
    }
    return "Unknown($DpiValue)"
}

function Get-CurrentDpiValue {
    param([string]$KeyPath)
    try {
        $props = Get-ItemProperty $KeyPath -Name 'DpiValue' -ErrorAction Stop
        return [int]$props.DpiValue
    }
    catch {
        return $null
    }
}

function Find-PerMonitorKeys {
    param(
        [string]$HardwareID,
        [string]$BasePath = $script:PerMonitorBasePath
    )
    if (-not (Test-Path $BasePath)) { return @() }
    $allKeys = Get-ChildItem $BasePath -ErrorAction SilentlyContinue
    if (-not $allKeys) { return @() }

    if ([string]::IsNullOrWhiteSpace($HardwareID)) {
        return @($allKeys)
    }
    return @($allKeys | Where-Object { $_.PSChildName -like "$HardwareID*" })
}

function Find-DisplayKeyNames {
    param([string]$HardwareID)
    if (-not (Test-Path $script:GraphicsConfigPath)) { return @() }
    $allKeys = Get-ChildItem $script:GraphicsConfigPath -ErrorAction SilentlyContinue
    if (-not $allKeys) { return @() }

    if ([string]::IsNullOrWhiteSpace($HardwareID)) {
        return @($allKeys | ForEach-Object { $_.PSChildName })
    }
    return @($allKeys | Where-Object { $_.PSChildName -like "$HardwareID*" } |
             ForEach-Object { $_.PSChildName })
}

function Select-DisplayInteractive {
    param([string]$Description)

    # Collect displays from PerMonitorSettings + GraphicsDrivers
    $pmKeys = Find-PerMonitorKeys -HardwareID ""
    $gdKeyNames = Find-DisplayKeyNames -HardwareID ""

    $allDisplays = @()
    $seenNames = @{}

    foreach ($k in $pmKeys) {
        $path = Join-Path $script:PerMonitorBasePath $k.PSChildName
        $val = Get-CurrentDpiValue -KeyPath $path
        $currentStr = if ($null -ne $val) { Convert-DpiToPercent -DpiValue $val -HardwareID $k.PSChildName } else { "Not set" }
        $allDisplays += [PSCustomObject]@{
            Name       = $k.PSChildName
            Source     = "PerMonitorSettings"
            CurrentDpi = $currentStr
        }
        $seenNames[$k.PSChildName] = $true
    }

    foreach ($name in $gdKeyNames) {
        if (-not $seenNames.ContainsKey($name)) {
            $allDisplays += [PSCustomObject]@{
                Name       = $name
                Source     = "GraphicsDrivers"
                CurrentDpi = "Not set"
            }
        }
    }

    if ($allDisplays.Count -eq 0) {
        Show-Error "No display keys found in registry"
        return $null
    }

    Write-Host ""
    Show-Separator
    Write-Host "Available Displays" -ForegroundColor Cyan
    Show-Separator
    Write-Host ""

    $idx = 0
    foreach ($d in $allDisplays) {
        $idx++
        Write-Host "[$idx] $($d.Name)" -ForegroundColor White
        Write-Host "    Current DPI: $($d.CurrentDpi) ($($d.Source))" -ForegroundColor Gray
        Write-Host ""
    }

    if (-not [string]::IsNullOrEmpty($Description)) {
        Write-Host "Description: $Description" -ForegroundColor Cyan
    }
    Write-Host ""

    $selection = Read-Host "Select display number (or 0 to skip)"
    $selNum = 0
    if (-not [int]::TryParse($selection, [ref]$selNum)) {
        Show-Info "Invalid input, skipping"
        return $null
    }

    if ($selNum -le 0 -or $selNum -gt $allDisplays.Count) {
        Show-Info "Skipped"
        return $null
    }

    return $allDisplays[$selNum - 1].Name
}

function Select-ScaleInteractive {
    param([string]$HardwareID = "")

    $recPercent = Get-RecommendedPercent -HardwareID $HardwareID

    Write-Host ""
    Write-Host "Available Scale Settings:" -ForegroundColor Cyan
    $scaleOptions = @(100, 125, 150, 175, 200)
    $idx = 0
    foreach ($s in $scaleOptions) {
        $idx++
        if ($s -eq $recPercent) {
            Write-Host "  [$idx] ${s}% (Recommended)" -ForegroundColor Green
        }
        else {
            Write-Host "  [$idx] ${s}%" -ForegroundColor White
        }
    }
    Write-Host ""

    $selection = Read-Host "Select scale (or 0 to skip)"
    $selNum = 0
    if (-not [int]::TryParse($selection, [ref]$selNum)) {
        Show-Info "Invalid input, skipping"
        return $null
    }

    if ($selNum -le 0 -or $selNum -gt $scaleOptions.Count) {
        Show-Info "Skipped"
        return $null
    }

    return $scaleOptions[$selNum - 1]
}

function Write-DpiValue {
    param(
        [string]$FullKeyName,
        [int]$DpiValue,
        [bool]$HiveLoaded
    )

    $hkcuChanged = $false
    $hiveChanged = $false
    $hasError = $false

    # --- HKCU ---
    $hkcuPath = Join-Path $script:PerMonitorBasePath $FullKeyName
    $currentHkcu = Get-CurrentDpiValue -KeyPath $hkcuPath

    if ($null -ne $currentHkcu -and $currentHkcu -eq $DpiValue) {
        Write-Host "  [HKCU] Already configured (Skip)" -ForegroundColor Gray
    }
    else {
        try {
            if (-not (Test-Path $hkcuPath)) {
                $null = New-Item -Path $hkcuPath -Force
            }
            $existing = Get-ItemProperty $hkcuPath -Name 'DpiValue' -ErrorAction SilentlyContinue
            if ($existing) {
                Set-ItemProperty -Path $hkcuPath -Name 'DpiValue' -Value $DpiValue -Type DWord -Force -ErrorAction Stop
            }
            else {
                $null = New-ItemProperty -Path $hkcuPath -Name 'DpiValue' -Value $DpiValue -PropertyType DWord -Force -ErrorAction Stop
            }
            Write-Host "  [HKCU] Configured" -ForegroundColor Green
            $hkcuChanged = $true
        }
        catch {
            Write-Host "  [HKCU] Error: $_"
            $hasError = $true
        }
    }

    # --- HIVE (Default User) ---
    if ($HiveLoaded) {
        $hivePath = "HKU:\Hive\Control Panel\Desktop\PerMonitorSettings\$FullKeyName"
        $currentHive = Get-CurrentDpiValue -KeyPath $hivePath

        if ($null -ne $currentHive -and $currentHive -eq $DpiValue) {
            Write-Host "  [HIVE] Already configured (Skip)" -ForegroundColor Gray
        }
        else {
            try {
                if (-not (Test-Path $hivePath)) {
                    $null = New-Item -Path $hivePath -Force
                }
                $existingHive = Get-ItemProperty $hivePath -Name 'DpiValue' -ErrorAction SilentlyContinue
                if ($existingHive) {
                    Set-ItemProperty -Path $hivePath -Name 'DpiValue' -Value $DpiValue -Type DWord -Force -ErrorAction Stop
                }
                else {
                    $null = New-ItemProperty -Path $hivePath -Name 'DpiValue' -Value $DpiValue -PropertyType DWord -Force -ErrorAction Stop
                }
                Write-Host "  [HIVE] Configured" -ForegroundColor Green
                $hiveChanged = $true
            }
            catch {
                Write-Host "  [HIVE] Error: $_"
                $hasError = $true
            }
        }
    }
    else {
        Write-Host "  [HIVE] Skipped (Hive not loaded)" -ForegroundColor Gray
    }

    return [PSCustomObject]@{
        HkcuChanged = $hkcuChanged
        HiveChanged = $hiveChanged
        HasError    = $hasError
    }
}

# ========================================
# Load CSV
# ========================================
$csvPath = Join-Path $PSScriptRoot "dpi_list.csv"

$items = Import-ModuleCsv -Path $csvPath -FilterEnabled
if ($null -eq $items) { return (New-ModuleResult -Status "Error" -Message "Failed to load dpi_list.csv") }
if ($items.Count -eq 0) { return (New-ModuleResult -Status "Skipped" -Message "No enabled entries") }
Write-Host ""

# ========================================
# Validate CSV Data
# ========================================
$validItems = @()
foreach ($item in $items) {
    $sp = $item.ScalePercent.Trim()

    # ScalePercent empty = interactive selection (valid)
    if ([string]::IsNullOrWhiteSpace($sp)) {
        $validItems += $item
        continue
    }

    $spNum = 0
    if (-not [int]::TryParse($sp, [ref]$spNum)) {
        Show-Warning "Invalid ScalePercent for '$($item.Description)': '$sp' — skipping"
        continue
    }
    if ($spNum -notin $script:SupportedScales) {
        Show-Warning "Unsupported ScalePercent '$spNum' for '$($item.Description)' (valid: $($script:SupportedScales -join ',')) — skipping"
        continue
    }
    $validItems += $item
}

if ($validItems.Count -eq 0) {
    Show-Error "No valid entries after validation"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "No valid entries after validation")
}

# ========================================
# Resolve Display Targets
# ========================================
$targets = @()

foreach ($item in $validItems) {
    $hwId = $item.HardwareID.Trim()
    $sp = $item.ScalePercent.Trim()

    $interactiveScale = [string]::IsNullOrWhiteSpace($sp)

    if ([string]::IsNullOrWhiteSpace($hwId)) {
        $targets += [PSCustomObject]@{
            Item               = $item
            MatchedKeyNames    = @()
            InteractiveDisplay = $true
            InteractiveScale   = $interactiveScale
        }
    }
    elseif ($hwId -eq "AUTO") {
        # Auto-detect: PerMonitorSettings first, then GraphicsDrivers fallback
        $pmKeys = Find-PerMonitorKeys -HardwareID ""
        $allKeyNames = @($pmKeys | ForEach-Object { $_.PSChildName })
        if ($allKeyNames.Count -eq 0) {
            $allKeyNames = Find-DisplayKeyNames -HardwareID ""
        }

        if ($allKeyNames.Count -eq 0) {
            Show-Warning "AUTO: No display keys found for '$($item.Description)' — falling back to Interactive mode"
            $targets += [PSCustomObject]@{
                Item               = $item
                MatchedKeyNames    = @()
                InteractiveDisplay = $true
                InteractiveScale   = $interactiveScale
            }
        }
        elseif ($allKeyNames.Count -eq 1) {
            Show-Info "AUTO: Single display detected — '$($allKeyNames[0])'"
            $targets += [PSCustomObject]@{
                Item               = $item
                MatchedKeyNames    = $allKeyNames
                InteractiveDisplay = $false
                InteractiveScale   = $interactiveScale
            }
        }
        else {
            Show-Info "AUTO: Multiple displays detected ($($allKeyNames.Count)) — falling back to Interactive mode"
            $targets += [PSCustomObject]@{
                Item               = $item
                MatchedKeyNames    = @()
                InteractiveDisplay = $true
                InteractiveScale   = $interactiveScale
            }
        }
    }
    else {
        # Search PerMonitorSettings first
        $pmMatched = Find-PerMonitorKeys -HardwareID $hwId
        $matchedNames = @($pmMatched | ForEach-Object { $_.PSChildName })

        # Fallback: search GraphicsDrivers\Configuration
        if ($matchedNames.Count -eq 0) {
            $matchedNames = Find-DisplayKeyNames -HardwareID $hwId
        }

        $targets += [PSCustomObject]@{
            Item               = $item
            MatchedKeyNames    = $matchedNames
            InteractiveDisplay = $false
            InteractiveScale   = $interactiveScale
        }
    }
}

# ========================================
# List Settings with Idempotency Check
# ========================================
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host "Target DPI Scaling Settings" -ForegroundColor Cyan
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""

$index = 0
foreach ($target in $targets) {
    $index++
    $item = $target.Item
    $sp = $item.ScalePercent.Trim()
    $scaleStr = if ($target.InteractiveScale) { "TBD (Interactive)" } else { "${sp}%" }

    if ($target.InteractiveDisplay) {
        Write-Host "[$index] (Interactive Selection) -> $scaleStr"
        Write-Host "    Display and/or scale will be selected during apply phase" -ForegroundColor Gray
        Write-Host "    $($item.Description)" -ForegroundColor Gray
        Write-Host ""
        continue
    }

    if ($target.MatchedKeyNames.Count -eq 0) {
        Write-Host "[$index] $($item.HardwareID) -> $scaleStr  [ERROR]"
        Write-Host "    No display found matching '$($item.HardwareID)'"
        Write-Host "    $($item.Description)" -ForegroundColor Gray
        Write-Host ""
        continue
    }

    foreach ($keyName in $target.MatchedKeyNames) {
        $hkcuPath = Join-Path $script:PerMonitorBasePath $keyName
        $currentVal = Get-CurrentDpiValue -KeyPath $hkcuPath
        $currentStr = if ($null -ne $currentVal) { Convert-DpiToPercent -DpiValue $currentVal -HardwareID $keyName } else { "Not set" }

        if (-not $target.InteractiveScale) {
            $targetDpi = Get-DpiValueForScale -ScalePercent ([int]$sp) -HardwareID $keyName
            if ($null -ne $currentVal -and $currentVal -eq $targetDpi) {
                $marker = "[Current]"
                $markerColor = "Gray"
            }
            else {
                $marker = "[Change]"
                $markerColor = "White"
            }
        }
        else {
            $marker = "[Interactive]"
            $markerColor = "Yellow"
        }

        Write-Host "[$index] $($item.HardwareID) -> $scaleStr  $marker" -ForegroundColor $markerColor
        Write-Host "    Matched: $keyName" -ForegroundColor Gray
        Write-Host "    Current: $currentStr | $($item.Description)" -ForegroundColor Gray
        Write-Host ""
    }
}

Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""

# ========================================
# Confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Apply the above DPI scaling settings?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""

# ========================================
# Load Default Profile Hive
# ========================================
$hiveLoaded = $false

if (Test-Path $HIVE_PATH) {
    Show-Info "Loading Default Profile Hive..."
    & reg load $HIVE_KEY $HIVE_PATH 2>&1 | Out-Null
    if ($LASTEXITCODE -eq 0) {
        Show-Success "Hive loaded: $HIVE_KEY"
        $hiveLoaded = $true
        if (-not (Get-PSDrive -Name HKU -ErrorAction SilentlyContinue)) {
            $null = New-PSDrive -Name HKU -PSProvider Registry -Root HKEY_USERS
        }
    }
    else {
        Show-Error "Failed to load Hive"
        Show-Info "Configuring HKCU only"
    }
}
else {
    Show-Error "ntuser.dat not found: $HIVE_PATH"
    Show-Info "Configuring HKCU only"
}
Write-Host ""

# ========================================
# Apply Settings
# ========================================
Show-Info "Applying DPI scaling settings..."
Write-Host ""

$successCount = 0
$skipCount = 0
$failCount = 0

$index = 0
foreach ($target in $targets) {
    $index++
    $item = $target.Item
    $sp = $item.ScalePercent.Trim()

    # Resolve display key name(s)
    $keyNamesToProcess = @()
    if ($target.InteractiveDisplay) {
        $selectedName = Select-DisplayInteractive -Description $item.Description
        if ($null -eq $selectedName) {
            Show-Skip "No display selected"
            $skipCount++
            Write-Host ""
            continue
        }
        $keyNamesToProcess = @($selectedName)
    }
    else {
        if ($target.MatchedKeyNames.Count -eq 0) {
            Write-Host "[$index] $($item.HardwareID)"
            Show-Error "No display found matching '$($item.HardwareID)'"
            $failCount++
            Write-Host ""
            continue
        }
        $keyNamesToProcess = $target.MatchedKeyNames
    }

    # Resolve scale percent
    $scalePercent = $null
    if ($target.InteractiveScale) {
        $interactiveHwId = if ($keyNamesToProcess.Count -gt 0) { $keyNamesToProcess[0] } else { "" }
        $scalePercent = Select-ScaleInteractive -HardwareID $interactiveHwId
        if ($null -eq $scalePercent) {
            Show-Skip "No scale selected"
            $skipCount++
            Write-Host ""
            continue
        }
    }
    else {
        $scalePercent = [int]$sp
    }

    foreach ($keyName in $keyNamesToProcess) {
        $targetDpi = Get-DpiValueForScale -ScalePercent $scalePercent -HardwareID $keyName
        $displayLabel = if ($target.InteractiveDisplay) { $keyName } else { $item.HardwareID }
        Write-Host "[$index] $displayLabel -> $scalePercent% (DpiValue: $targetDpi)" -ForegroundColor Cyan
        Write-Host "  Key: $keyName" -ForegroundColor Gray

        $result = Write-DpiValue -FullKeyName $keyName -DpiValue $targetDpi -HiveLoaded $hiveLoaded

        if ($result.HasError) {
            $failCount++
        }
        elseif (-not $result.HkcuChanged -and -not $result.HiveChanged) {
            $skipCount++
        }
        else {
            $successCount++
        }

        Write-Host ""
    }
}

# ========================================
# Unload Hive
# ========================================
if ($hiveLoaded) {
    Show-Info "Unloading Hive..."

    if (Get-PSDrive -Name HKU -ErrorAction SilentlyContinue) {
        Remove-PSDrive -Name HKU -Force
    }

    [gc]::Collect()
    Start-Sleep -Seconds 1

    & reg unload $HIVE_KEY 2>&1 | Out-Null
    if ($LASTEXITCODE -eq 0) {
        Show-Success "Hive unloaded"
    }
    else {
        Show-Warning "Failed to unload Hive. Retrying..."
        Start-Sleep -Seconds 2
        [gc]::Collect()
        & reg unload $HIVE_KEY 2>&1 | Out-Null
        if ($LASTEXITCODE -eq 0) {
            Show-Success "Hive unloaded (Retry success)"
        }
        else {
            Show-Error "Failed to unload Hive. Please unload manually."
        }
    }
    Write-Host ""
}

# ========================================
# Result Summary
# ========================================
if ($successCount -gt 0) {
    Show-Warning "Sign-out or restart may be required for changes to take effect."
    Write-Host ""
}
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount -Title "Execution Results")
