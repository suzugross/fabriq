# ========================================
# Wallpaper Configuration Script (Live)
# ========================================
# Sets the desktop wallpaper immediately using
# the Win32 SystemParametersInfo API (SPI_SETDESKWALLPAPER).
# No Explorer restart required.
# ========================================

Write-Host ""
Show-Separator
Write-Host "Wallpaper Configuration (Live)" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# ========================================
# Load C# Wallpaper Handler
# ========================================
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

public class WallpaperHandler {
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);

    public const int SPI_SETDESKWALLPAPER = 0x0014;
    public const int SPIF_UPDATEINIFILE   = 0x01;
    public const int SPIF_SENDCHANGE      = 0x02;
}
'@ -ErrorAction SilentlyContinue

# ========================================
# Style -> Registry value mapping
# ========================================
$styleMap = @{
    "Fill"    = @{ WallpaperStyle = "10"; TileWallpaper = "0" }
    "Fit"     = @{ WallpaperStyle = "6";  TileWallpaper = "0" }
    "Stretch" = @{ WallpaperStyle = "2";  TileWallpaper = "0" }
    "Tile"    = @{ WallpaperStyle = "0";  TileWallpaper = "1" }
    "Center"  = @{ WallpaperStyle = "0";  TileWallpaper = "0" }
    "Span"    = @{ WallpaperStyle = "22"; TileWallpaper = "0" }
}

$validExtensions = @(".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tif", ".tiff")

# ========================================
# Load wallpaper_list.csv
# ========================================
$csvPath = Join-Path $PSScriptRoot "wallpaper_list.csv"

$enabledItems = Import-ModuleCsv -Path $csvPath -FilterEnabled -RequiredColumns @("Enabled", "FileName", "Style")
if ($null -eq $enabledItems) { return (New-ModuleResult -Status "Error" -Message "Failed to load wallpaper_list.csv") }
if ($enabledItems.Count -eq 0) { return (New-ModuleResult -Status "Skipped" -Message "No enabled entries") }

# ========================================
# Validate wallpaper/ directory
# Only required when relative-path entries exist
# ========================================
$wallpaperDir = Join-Path $PSScriptRoot "wallpaper"

$hasRelativePaths = @($enabledItems | Where-Object {
    -not [System.IO.Path]::IsPathRooted($_.FileName)
}).Count -gt 0

if ($hasRelativePaths -and -not (Test-Path $wallpaperDir)) {
    Show-Error "'wallpaper' directory not found: $wallpaperDir"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "'wallpaper' directory not found")
}

# ========================================
# Show Target List with pre-flight checks
# ========================================
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "Target Wallpapers" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

$invalidCount = 0

foreach ($item in $enabledItems) {
    $imagePath = if ([System.IO.Path]::IsPathRooted($item.FileName)) {
        $item.FileName
    } else {
        Join-Path $wallpaperDir $item.FileName
    }
    $ext       = [System.IO.Path]::GetExtension($item.FileName).ToLower()
    $desc      = if ($item.Description) { "  ($($item.Description))" } else { "" }
    $style     = if ($item.Style) { $item.Style } else { "Fill" }

    if (-not (Test-Path $imagePath)) {
        Write-Host "  [NOT FOUND] $($item.FileName)$desc" -ForegroundColor Red
        Write-Host "    Path: $imagePath" -ForegroundColor DarkGray
        $invalidCount++
    }
    elseif ($ext -notin $validExtensions) {
        Write-Host "  [INVALID EXT] $($item.FileName)$desc" -ForegroundColor Red
        Write-Host "    Supported: $($validExtensions -join ', ')" -ForegroundColor DarkGray
        $invalidCount++
    }
    else {
        Write-Host "  [SET] $($item.FileName)$desc" -ForegroundColor Yellow
        Write-Host "    Style: $style" -ForegroundColor DarkGray
    }
    Write-Host ""
}

Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

if ($invalidCount -gt 0) {
    Show-Warning "$invalidCount item(s) will be skipped (file not found or unsupported extension)"
    Write-Host ""
}

# ========================================
# Confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Apply the above wallpaper settings?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""

# ========================================
# Apply Wallpaper Settings
# ========================================
$successCount = 0
$skipCount    = 0
$failCount    = 0

foreach ($item in $enabledItems) {
    $imagePath = if ([System.IO.Path]::IsPathRooted($item.FileName)) {
        $item.FileName
    } else {
        Join-Path $wallpaperDir $item.FileName
    }
    $ext       = [System.IO.Path]::GetExtension($item.FileName).ToLower()
    $desc      = if ($item.Description) { $item.Description } else { $item.FileName }
    $styleKey  = if ($item.Style -and $styleMap.ContainsKey($item.Style)) { $item.Style } else { "Fill" }

    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "Applying: $desc" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor White

    # File existence check
    if (-not (Test-Path $imagePath)) {
        Show-Skip "File not found: $($item.FileName)"
        Write-Host ""
        $skipCount++
        continue
    }

    # Extension check
    if ($ext -notin $validExtensions) {
        Show-Skip "Unsupported extension: $ext"
        Write-Host ""
        $skipCount++
        continue
    }

    # Style fallback warning
    if ($item.Style -and -not $styleMap.ContainsKey($item.Style)) {
        Show-Warning "Unknown Style '$($item.Style)' - falling back to Fill"
    }

    try {
        # Resolve to absolute path (required by SystemParametersInfo)
        $absolutePath = (Resolve-Path $imagePath).Path

        # Write registry style values
        $regPath = "HKCU:\Control Panel\Desktop"
        Set-ItemProperty -Path $regPath -Name "WallpaperStyle" -Value $styleMap[$styleKey].WallpaperStyle -ErrorAction Stop
        Set-ItemProperty -Path $regPath -Name "TileWallpaper"  -Value $styleMap[$styleKey].TileWallpaper  -ErrorAction Stop

        # Apply wallpaper immediately via Win32 API
        $apiResult = [WallpaperHandler]::SystemParametersInfo(
            [WallpaperHandler]::SPI_SETDESKWALLPAPER,
            0,
            $absolutePath,
            [WallpaperHandler]::SPIF_UPDATEINIFILE -bor [WallpaperHandler]::SPIF_SENDCHANGE
        )

        if ($apiResult -ne 0) {
            Show-Success "Wallpaper applied: $($item.FileName) (Style: $styleKey)"
            $successCount++
        }
        else {
            Show-Error "SystemParametersInfo returned 0 (failed) for: $($item.FileName)"
            $failCount++
        }
    }
    catch {
        Show-Error "Error applying wallpaper '$($item.FileName)': $_"
        $failCount++
    }

    Write-Host ""
}

# ========================================
# Result Summary
# ========================================
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount -Title "Wallpaper Configuration Results")
