# ========================================
# SPI Configuration Script (SystemParametersInfo)
# ========================================
# Windows SystemParametersInfo API to change settings that
# cannot be safely controlled via direct registry writes.
# Covers visual effects (UserPreferencesMask bits), mouse
# speed, keyboard settings, etc. driven by CSV.
#
# [NOTES]
# - Applies immediately to the current user session via SPI
# - Registers Active Setup (HKLM) so new users receive
#   the same SPI settings at first logon (admin required)
# ========================================

$ENABLE_STARTUP_BATCH = $true  # Set to $false to disable Startup batch deployment

Write-Host ""
Show-Separator
Write-Host "SPI Configuration (SystemParametersInfo)" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# ========================================
# C# Native Class Definition
# ========================================
# dpi_api_config の P/Invoke パターンを参考に実装。
# 3つの ValueMode (bool, uiParam, pvParam) に対応。
# GET は SET action - 1 の慣例に基づき自動算出。
# ========================================
$spiSource = @"
using System;
using System.Runtime.InteropServices;

public class SpiHelper {

    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool SystemParametersInfo(
        uint uiAction, uint uiParam, IntPtr pvParam, uint fWinIni);

    // SPIF_UPDATEINIFILE | SPIF_SENDCHANGE
    private const uint SPIF = 0x0003;

    /// <summary>
    /// GET: Read current value via pvParam pointer.
    /// Works for all SPI types (bool returns 0/1, int returns value).
    /// GET action = SET action - 1 (standard SPI convention).
    /// </summary>
    public static int GetValue(uint getAction) {
        IntPtr ptr = Marshal.AllocHGlobal(4);
        try {
            Marshal.WriteInt32(ptr, 0);
            SystemParametersInfo(getAction, 0, ptr, 0);
            return Marshal.ReadInt32(ptr);
        } finally {
            Marshal.FreeHGlobal(ptr);
        }
    }

    /// <summary>
    /// SET (bool / pvParam mode): value passed as pvParam IntPtr.
    /// For bool: 0=OFF, 1=ON. For pvParam: arbitrary integer.
    /// </summary>
    public static bool SetPvParam(uint action, int value) {
        return SystemParametersInfo(action, 0, (IntPtr)value, SPIF);
    }

    /// <summary>
    /// SET (uiParam mode): value passed as uiParam.
    /// Used by keyboard speed, keyboard delay, etc.
    /// </summary>
    public static bool SetUiParam(uint action, uint value) {
        return SystemParametersInfo(action, value, IntPtr.Zero, SPIF);
    }
}
"@

try {
    Add-Type -TypeDefinition $spiSource -Language CSharp -ErrorAction SilentlyContinue
}
catch {
    Show-Error "Failed to compile SpiHelper: $($_.Exception.Message)"
    return (New-ModuleResult -Status "Error" -Message "C# compilation failed")
}


# ========================================
# Step 1: CSV 読み込み
# ========================================
$csvPath = Join-Path $PSScriptRoot "spi_list.csv"

$enabledItems = Import-ModuleCsv -Path $csvPath -FilterEnabled `
    -RequiredColumns @("Enabled", "SpiAction", "ValueMode", "Value", "Description")

if ($null -eq $enabledItems) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load spi_list.csv")
}
if ($enabledItems.Count -eq 0) {
    return (New-ModuleResult -Status "Skipped" -Message "No enabled entries")
}


# ========================================
# Step 2: 前提条件チェック
# ========================================
# SPI は管理者権限不要（HKCU 操作のため）。
# 特別な前提条件なし。


# ========================================
# Step 3: 実行前の確認表示
# ========================================
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "SPI Settings" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

$hasChanges = $false

foreach ($item in $enabledItems) {
    $displayName = $item.Description
    $setAction = [Convert]::ToUInt32($item.SpiAction, 16)
    $getAction = $setAction - 1
    $targetValue = [int]$item.Value
    $valueMode = $item.ValueMode

    # GET current value for idempotency display
    $currentValue = $null
    try {
        $currentValue = [SpiHelper]::GetValue($getAction)
    }
    catch {
        $currentValue = $null
    }

    if ($null -ne $currentValue -and $currentValue -eq $targetValue) {
        $marker = "[SKIP]"
        $markerColor = "Gray"
    }
    else {
        $marker = "[APPLY]"
        $markerColor = "Yellow"
        $hasChanges = $true
    }

    # Display value in readable format
    $targetDisplay = switch ($valueMode) {
        'bool' { if ($targetValue -eq 0) { "OFF" } else { "ON" } }
        default { $targetValue }
    }
    $currentDisplay = if ($null -eq $currentValue) { "Unknown" }
        elseif ($valueMode -eq 'bool') { if ($currentValue -eq 0) { "OFF" } else { "ON" } }
        else { $currentValue }

    Write-Host "  $marker $displayName" -ForegroundColor $markerColor
    Write-Host "    Action: $($item.SpiAction) ($valueMode)  Current: $currentDisplay -> Target: $targetDisplay" -ForegroundColor DarkGray
    Write-Host ""
}

Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

if (-not $hasChanges) {
    Show-Skip "All SPI settings already match target values"
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "All SPI settings already match")
}


# ========================================
# Step 4: 実行確認
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Apply the above SPI settings?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""


# ========================================
# Step 5: 設定適用ループ
# ========================================
$successCount = 0
$skipCount    = 0
$failCount    = 0

foreach ($item in $enabledItems) {
    $displayName = $item.Description
    $setAction = [Convert]::ToUInt32($item.SpiAction, 16)
    $getAction = $setAction - 1
    $targetValue = [int]$item.Value
    $valueMode = $item.ValueMode

    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "Processing: $displayName" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor White

    # ----------------------------------------
    # 冪等性チェック（Skip 判定）
    # ----------------------------------------
    $currentValue = $null
    try {
        $currentValue = [SpiHelper]::GetValue($getAction)
    }
    catch { }

    if ($null -ne $currentValue -and $currentValue -eq $targetValue) {
        $currentDisplay = if ($valueMode -eq 'bool') {
            if ($currentValue -eq 0) { "OFF" } else { "ON" }
        } else { $currentValue }
        Show-Skip "Already set ($currentDisplay)"
        $skipCount++
        Write-Host ""
        continue
    }

    # ----------------------------------------
    # メイン処理
    # ----------------------------------------
    try {
        $result = switch ($valueMode) {
            'bool'    { [SpiHelper]::SetPvParam($setAction, $targetValue) }
            'pvParam' { [SpiHelper]::SetPvParam($setAction, $targetValue) }
            'uiParam' { [SpiHelper]::SetUiParam($setAction, [uint32]$targetValue) }
            default {
                Show-Error "Unknown ValueMode: $valueMode"
                $null
            }
        }

        if ($null -eq $result) {
            $failCount++
        }
        elseif ($result) {
            $targetDisplay = if ($valueMode -eq 'bool') {
                if ($targetValue -eq 0) { "OFF" } else { "ON" }
            } else { $targetValue }
            Show-Success "Set to $targetDisplay"
            $successCount++
        }
        else {
            Show-Error "SystemParametersInfo returned false (Action: $($item.SpiAction))"
            $failCount++
        }
    }
    catch {
        Show-Error "Failed: $displayName : $_"
        $failCount++
    }

    Write-Host ""
}


# ========================================
# Step 6: Active Setup Registration (for new users)
# ========================================
# Register an Active Setup entry in HKLM so that new users
# receive the same SPI settings at first logon.
# The StubPath script is generated dynamically from spi_list.csv
# (bool items only — pvParam/uiParam items are excluded).
# Requires admin privileges for HKLM write.
# ========================================
Write-Host ""
Show-Separator
Write-Host "Active Setup Registration (for new users)" -ForegroundColor Cyan
Show-Separator
Write-Host ""

if (-not (Test-AdminPrivilege)) {
    Show-Warning "Admin privileges required for Active Setup registration"
    Show-Info "Skipping Active Setup — SPI applied to current user only"
    Write-Host ""
}
else {
    # Build script content from CSV bool items
    $boolItems = @($enabledItems | Where-Object { $_.ValueMode -eq 'bool' })

    if ($boolItems.Count -eq 0) {
        Show-Skip "No bool items to register for Active Setup"
        Write-Host ""
    }
    else {
        # Generate apply_spi.ps1 script lines (SPI calls only)
        $scriptLines = @()
        $scriptLines += "Add-Type 'using System;using System.Runtime.InteropServices;public class S{[DllImport(""user32.dll"")]public static extern bool SystemParametersInfo(uint a,uint b,IntPtr c,uint d);}'"
        foreach ($bi in $boolItems) {
            $scriptLines += "[S]::SystemParametersInfo($($bi.SpiAction),0,[IntPtr]$($bi.Value),3)"
        }

        $asResult = Register-FabriqActiveSetup `
            -GUID "{fabriq-spi-config}" `
            -Description "Fabriq SPI Configuration" `
            -ScriptName "apply_spi.ps1" `
            -ScriptLines $scriptLines

        if (-not $asResult) {
            $failCount++
        }

        Write-Host ""
    }
}

# ========================================
# Startup Batch Deployment (for new users)
# ========================================
# Deploy a Startup-triggered batch that applies SPI settings
# after Explorer has fully started, then restarts Explorer.
# This supplements Active Setup to handle timing issues where
# Explorer initialization overwrites settings.
# ========================================
if ($ENABLE_STARTUP_BATCH) {
    Write-Host ""
    Show-Separator
    Write-Host "Startup Batch Deployment (for new users)" -ForegroundColor Cyan
    Show-Separator
    Write-Host ""

    # Build script lines from all enabled items (not just bool)
    $allSpiLines = @()
    $allSpiLines += "Add-Type 'using System;using System.Runtime.InteropServices;public class S{[DllImport(""user32.dll"")]public static extern bool SystemParametersInfo(uint a,uint b,IntPtr c,uint d);}'"

    foreach ($item in $enabledItems) {
        switch ($item.ValueMode) {
            'bool' {
                $allSpiLines += "[S]::SystemParametersInfo($($item.SpiAction),0,[IntPtr]$($item.Value),3)"
            }
            'pvParam' {
                $allSpiLines += "`$p=[System.Runtime.InteropServices.Marshal]::AllocHGlobal(4);[System.Runtime.InteropServices.Marshal]::WriteInt32(`$p,$($item.Value));[S]::SystemParametersInfo($($item.SpiAction),0,`$p,3);[System.Runtime.InteropServices.Marshal]::FreeHGlobal(`$p)"
            }
            'uiParam' {
                $allSpiLines += "[S]::SystemParametersInfo($($item.SpiAction),$($item.Value),[IntPtr]::Zero,3)"
            }
        }
    }

    if ($allSpiLines.Count -le 1) {
        Show-Skip "No items to deploy for Startup Batch"
    }
    else {
        # Deploy apply_spi.ps1
        $scriptDir = "C:\ProgramData\fabriq"
        if (-not (Test-Path $scriptDir)) {
            New-Item -Path $scriptDir -ItemType Directory -Force | Out-Null
        }
        $allSpiLines | Set-Content -Path (Join-Path $scriptDir "apply_spi.ps1") -Force -Encoding UTF8
        Show-Success "Script deployed: $scriptDir\apply_spi.ps1"

        # Deploy launcher and Startup trigger
        $launcherResult = Deploy-FabriqUserSetupLauncher
        $triggerResult  = Deploy-FabriqStartupTrigger

        if (-not $launcherResult -or -not $triggerResult) {
            Show-Warning "Startup Batch deployment incomplete - Active Setup still applies"
        }
    }
    Write-Host ""
}

# ========================================
# Step 7: Results
# ========================================
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount `
    -Title "SPI Configuration Results")
