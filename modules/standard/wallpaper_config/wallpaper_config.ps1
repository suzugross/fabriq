# ========================================
# Wallpaper Configuration Script (Live)
# ========================================
# Sets the desktop wallpaper for kitting in a robust, removable-media-safe
# way:
#   1. Image files are NORMALIZED to a local, world-readable folder
#      (C:\Windows\Web\Wallpaper\fabriq) before being referenced. The
#      registry / SystemParametersInfo never reference the source path
#      (e.g. a USB drive), so the wallpaper survives ejecting the media,
#      reboots, and drive-letter changes.
#   2. The wallpaper is reflected to the current/next-logon user immediately
#      (SPI) or as a staged registry write, AND - so that users created
#      AFTER kitting also receive it - an Active Setup entry is registered
#      that re-applies the wallpaper once per user at their first logon.
#      (Writing the value into the Default profile hive is NOT reliable on
#      Windows 10/11: the theme engine overwrites Control Panel\Desktop\
#      WallPaper at first logon. Active Setup applies AFTER that, and the
#      result remains user-changeable.)
# SolidColor entries are not copied (no image); they clear WallPaper and
# set Control Panel\Colors\Background.
# Requires administrator privileges (C:\Windows write + Active Setup HKLM).
# ========================================

# ----------------------------------------
# Resolve logged-on user's HKCU target
# ----------------------------------------
$hkcuInfo = Resolve-HkcuRoot

Write-Host ""
Show-Separator
Write-Host "Wallpaper Configuration (Live)" -ForegroundColor Cyan
if ($hkcuInfo.Redirected) {
    Write-Host "  HKCU target: $($hkcuInfo.Label)" -ForegroundColor Magenta
}
Write-Host "  Targets: current/next-logon user + new users (via Active Setup)" -ForegroundColor DarkGray
Show-Separator
Write-Host ""

# Administrator privileges are required for the C:\Windows copy, the HKLM
# Active Setup registration, and the Default-profile hive load. Fail closed.
if (-not (Test-AdminPrivilege)) {
    Show-Error "Administrator privileges are required (C:\Windows write + Active Setup + hive load)."
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Administrator privileges required")
}

# ----------------------------------------
# Local destination + Active Setup constants
# ----------------------------------------
$wallpaperLocalDir = Join-Path $env:WINDIR "Web\Wallpaper\fabriq"
$ACTIVE_SETUP_GUID   = "{FAB71C04-0000-4000-8000-000000000038}"
$ACTIVE_SETUP_NAME   = "wallpaper_default_apply.ps1"
# Must match Register-FabriqActiveSetup's hard-coded deploy dir (C:\ProgramData\fabriq)
# so Post-Apply Verification checks the file the kernel actually wrote.
$ACTIVE_SETUP_SCRIPT = "C:\ProgramData\fabriq\$ACTIVE_SETUP_NAME"

# Default-profile hive (used only to disable Desktop Spotlight for new users).
$HIVE_PATH      = "$env:SystemDrive\Users\Default\ntuser.dat"
$HIVE_KEY       = "HKEY_USERS\fabriq_wallpaper_default"   # unique mount key
$HIVE_PSPATH    = "HKU:\fabriq_wallpaper_default"
$SPOTLIGHT_SUB  = '\Software\Microsoft\Windows\CurrentVersion\DesktopSpotlight\Settings'

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
# Helper functions (module-local)
# ========================================

function Resolve-WallpaperLocalDest {
    # Maps a source image path to its canonical local destination under
    # $BaseDir, returning $null when the leaf filename is unsafe or the
    # resolved path would escape the base (Section 8 destructive-op guard).
    param(
        [string]$SourcePath,
        [string]$BaseDir
    )
    $leaf = [System.IO.Path]::GetFileName($SourcePath)
    if (-not (Test-FabriqSafePathComponent -Value $leaf)) { return $null }
    $baseFull = [System.IO.Path]::GetFullPath($BaseDir)
    $destFull = [System.IO.Path]::GetFullPath((Join-Path $baseFull $leaf))
    $baseGuard = $baseFull.TrimEnd('\') + '\'
    if (-not $destFull.StartsWith($baseGuard, [System.StringComparison]::OrdinalIgnoreCase)) { return $null }
    return $destFull
}

function Copy-WallpaperToLocal {
    # Idempotently copies $Source to $Dest. Skips the copy when an
    # identical file (by SHA256) is already in place. Throws on failure.
    param(
        [string]$Source,
        [string]$Dest
    )
    # Self-copy guard: when the source already IS the canonical local
    # destination, Copy-Item would throw ("being used by another process").
    if ([System.IO.Path]::GetFullPath($Source) -ieq $Dest) {
        Show-Info "Source already at the local destination: $Dest"
        return
    }
    if (Test-Path $Dest) {
        try {
            $sh = (Get-FileHash -Path $Source -Algorithm SHA256 -ErrorAction Stop).Hash
            $dh = (Get-FileHash -Path $Dest   -Algorithm SHA256 -ErrorAction Stop).Hash
            if ($sh -eq $dh) {
                Show-Info "Local copy already current: $Dest"
                return
            }
        } catch {
            # Hash failed - fall through and re-copy.
        }
    }
    $destDir = Split-Path -Parent $Dest
    if (-not (Test-Path $destDir)) { New-Item -Path $destDir -ItemType Directory -Force | Out-Null }
    Copy-Item -Path $Source -Destination $Dest -Force -ErrorAction Stop
    Show-Info "Copied to local: $Dest"
}

function Test-RegStringEquals {
    # Reads a REG_SZ value and compares it (string) to $Expected. Returns
    # $false on any read error. Used by Post-Apply Verification.
    param(
        [string]$Key,
        [string]$Name,
        [string]$Expected
    )
    try {
        $cur = (Get-ItemProperty -Path $Key -Name $Name -ErrorAction Stop).$Name
        return ([string]$cur -eq [string]$Expected)
    } catch {
        return $false
    }
}

function Test-RegDwordEquals {
    # Reads a REG_DWORD value and compares it (int) to $Expected. Returns
    # $false on any read error. Used by Post-Apply Verification.
    param(
        [string]$Key,
        [string]$Name,
        [int]$Expected
    )
    try {
        $cur = (Get-ItemProperty -Path $Key -Name $Name -ErrorAction Stop).$Name
        return ([int]$cur -eq $Expected)
    } catch {
        return $false
    }
}

function Set-SpotlightDisabled {
    # Disables Desktop Spotlight (auto-rotating desktop background) on a hive
    # root by setting DesktopSpotlight\Settings\EnabledState = 0 (REG_DWORD).
    # Spotlight, when enabled, takes priority over a static wallpaper.
    # $HiveRoot is e.g. 'HKCU:' / 'HKU:\<SID>' / 'HKU:\fabriq_wallpaper_default'.
    param([string]$HiveRoot)
    $key = $HiveRoot + $SPOTLIGHT_SUB
    if (-not (Test-Path $key)) { New-Item -Path $key -Force | Out-Null }
    New-ItemProperty -Path $key -Name 'EnabledState' -Value 0 -PropertyType DWord -Force | Out-Null
}

function New-WallpaperActiveSetupScript {
    # Generates the self-contained script (string[] lines) deployed by
    # Active Setup. It re-applies the final desktop state for each user at
    # their first logon via SystemParametersInfo (so it survives the theme
    # reset and stays user-changeable). Values are baked in at kit time;
    # single quotes in paths/colors are escaped ('' inside a '...' literal).
    param([hashtable]$State)

    $header = @(
        '# Fabriq default wallpaper - applied once per user at first logon (Active Setup).',
        '$ErrorActionPreference = ''SilentlyContinue''',
        'Add-Type -TypeDefinition @''',
        'using System; using System.Runtime.InteropServices;',
        'public class FabriqWp { [DllImport("user32.dll", CharSet=CharSet.Auto)] public static extern int SystemParametersInfo(int a, int u, string p, int f); }',
        '''@',
        '$desktop = ''HKCU:\Control Panel\Desktop'''
    )

    if ($State.Type -eq 'SolidColor') {
        $color = ([string]$State.Color).Replace("'", "''")
        return @($header + @(
            "Set-ItemProperty -Path `$desktop -Name 'WallPaper' -Value ''",
            "Set-ItemProperty -Path 'HKCU:\Control Panel\Colors' -Name 'Background' -Value '$color'",
            "[FabriqWp]::SystemParametersInfo(0x0014, 0, '', 0x03) | Out-Null"
        ))
    }

    $p = ([string]$State.Path).Replace("'", "''")
    return @($header + @(
        "Set-ItemProperty -Path `$desktop -Name 'WallpaperStyle' -Value '$($State.Style)'",
        "Set-ItemProperty -Path `$desktop -Name 'TileWallpaper' -Value '$($State.Tile)'",
        "[FabriqWp]::SystemParametersInfo(0x0014, 0, '$p', 0x03) | Out-Null"
    ))
}

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
    $type = if ($item.PSObject.Properties['Type'] -and $item.Type -eq 'SolidColor') { 'SolidColor' } else { 'Image' }

    if ($type -eq 'SolidColor') {
        $colorVal = if ($item.PSObject.Properties['Color'] -and $item.Color) { $item.Color } else { "0 0 0" }
        $desc     = if ($item.Description) { "  ($($item.Description))" } else { "" }
        if ($colorVal -match '^\d{1,3}\s+\d{1,3}\s+\d{1,3}$') {
            Write-Host "  [COLOR] $colorVal$desc" -ForegroundColor Yellow
        } else {
            Write-Host "  [INVALID COLOR] $colorVal$desc" -ForegroundColor Red
            $invalidCount++
        }
        Write-Host ""
        continue
    }

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
        Write-Host "    Style: $style  ->  copied to $wallpaperLocalDir" -ForegroundColor DarkGray
    }
    Write-Host ""
}

Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

if ($invalidCount -gt 0) {
    Show-Warning "$invalidCount item(s) will be skipped (file not found or unsupported extension)"
    Write-Host ""
}

# Warn on duplicate image filenames: different sources with the same leaf
# name resolve to the same local destination, so the last one wins.
$imageLeaves = @($enabledItems |
    Where-Object { -not ($_.PSObject.Properties['Type'] -and $_.Type -eq 'SolidColor') } |
    ForEach-Object { [System.IO.Path]::GetFileName($_.FileName).ToLower() } |
    Where-Object { $_ })
$dupLeaves = @($imageLeaves | Group-Object | Where-Object { $_.Count -gt 1 } | ForEach-Object { $_.Name })
if ($dupLeaves.Count -gt 0) {
    Show-Warning "Duplicate image filename(s): $($dupLeaves -join ', '). Same-name images share one local destination (last one wins); use distinct filenames."
    Write-Host ""
}

# ========================================
# Confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Apply the above wallpaper settings?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""

# ========================================
# Disable Desktop Spotlight (current user + Default profile / new users)
# Spotlight (the auto-rotating desktop background) takes priority over a
# static wallpaper, so it must be turned off BEFORE the wallpaper is applied.
# EnabledState is a plain inherited setting (NOT the theme-managed wallpaper
# path), so writing it into the Default profile hive DOES reach new users.
# ========================================
$spotlightDegraded = $false   # could not disable Spotlight somewhere
$hiveUnloadFailed   = $false

Write-Host "----------------------------------------" -ForegroundColor White
Write-Host "Disabling Desktop Spotlight" -ForegroundColor Cyan
Write-Host "----------------------------------------" -ForegroundColor White

# Current / next-logon user
try {
    Set-SpotlightDisabled -HiveRoot $hkcuInfo.PsDrivePath
    if (Test-RegDwordEquals ($hkcuInfo.PsDrivePath + $SPOTLIGHT_SUB) 'EnabledState' 0) {
        Show-Success "Desktop Spotlight disabled for $($hkcuInfo.Label)"
    } else {
        Show-Warning "Desktop Spotlight value did not verify for $($hkcuInfo.Label)"
        $spotlightDegraded = $true
    }
}
catch {
    Show-Warning "Failed to disable Desktop Spotlight for current user: $_"
    $spotlightDegraded = $true
}

# Default profile (new users) via a temporary hive mount.
$hiveLoaded = $false
& reg query $HIVE_KEY 2>&1 | Out-Null
if ($LASTEXITCODE -eq 0) {
    Show-Warning "Stale hive mount '$HIVE_KEY' found - unloading before reload"
    & reg unload $HIVE_KEY 2>&1 | Out-Null
}
if (Test-Path $HIVE_PATH) {
    & reg load $HIVE_KEY $HIVE_PATH 2>&1 | Out-Null
    if ($LASTEXITCODE -eq 0) {
        $hiveLoaded = $true
        if (-not (Get-PSDrive -Name HKU -ErrorAction SilentlyContinue)) {
            New-PSDrive -Name HKU -PSProvider Registry -Root HKEY_USERS | Out-Null
        }
        try {
            Set-SpotlightDisabled -HiveRoot $HIVE_PSPATH
            if (Test-RegDwordEquals ($HIVE_PSPATH + $SPOTLIGHT_SUB) 'EnabledState' 0) {
                Show-Success "Desktop Spotlight disabled in Default profile (new users)"
            } else {
                Show-Warning "Desktop Spotlight value did not verify in Default profile"
                $spotlightDegraded = $true
            }
        }
        catch {
            Show-Warning "Failed to disable Desktop Spotlight in Default profile: $_"
            $spotlightDegraded = $true
        }
    }
    else {
        Show-Warning "Failed to load Default profile hive - Spotlight not disabled for new users"
        $spotlightDegraded = $true
    }
}
else {
    Show-Warning "Default profile ntuser.dat not found: $HIVE_PATH - Spotlight not disabled for new users"
    $spotlightDegraded = $true
}

# Unload the Default hive (release .NET handles first, retry once).
if ($hiveLoaded) {
    if (Get-PSDrive -Name HKU -ErrorAction SilentlyContinue) {
        Remove-PSDrive -Name HKU -Force -ErrorAction SilentlyContinue
    }
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()
    Start-Sleep -Seconds 1
    & reg unload $HIVE_KEY 2>&1 | Out-Null
    if ($LASTEXITCODE -ne 0) {
        Show-Warning "Failed to unload Default hive. Retrying..."
        Start-Sleep -Seconds 2
        [gc]::Collect()
        [gc]::WaitForPendingFinalizers()
        & reg unload $HIVE_KEY 2>&1 | Out-Null
        if ($LASTEXITCODE -ne 0) {
            Show-Error "Failed to unload Default hive. ntuser.dat may remain locked."
            $hiveUnloadFailed = $true
        }
    }
    # Removing the HKU PSDrive above also dropped the drive that
    # Resolve-HkcuRoot uses for a redirected (HKU:\<SID>) target. Re-create
    # it so the wallpaper apply loop below can still write to that hive.
    if ($hkcuInfo.Redirected -and -not (Get-PSDrive -Name HKU -ErrorAction SilentlyContinue)) {
        New-PSDrive -Name HKU -PSProvider Registry -Root HKEY_USERS | Out-Null
    }
}
Write-Host ""

# ========================================
# Apply Wallpaper Settings (current/next-logon user)
# ========================================
$successCount = 0
$skipCount    = 0
$failCount    = 0
$verifyPass   = 0
$verifyFail   = 0

# Last successfully-applied item; this is what the desktop ends up showing,
# and what new users receive via Active Setup.
$finalState   = $null

foreach ($item in $enabledItems) {
    $type = if ($item.PSObject.Properties['Type'] -and $item.Type -eq 'SolidColor') { 'SolidColor' } else { 'Image' }

    # ---------------------------------- SolidColor ----------------------------------
    if ($type -eq 'SolidColor') {
        $colorVal = if ($item.PSObject.Properties['Color'] -and $item.Color) { $item.Color } else { "0 0 0" }
        $desc     = if ($item.Description) { $item.Description } else { "Solid Color: $colorVal" }

        Write-Host "----------------------------------------" -ForegroundColor White
        Write-Host "Applying: $desc" -ForegroundColor Cyan
        Write-Host "----------------------------------------" -ForegroundColor White

        if ($colorVal -notmatch '^\d{1,3}\s+\d{1,3}\s+\d{1,3}$') {
            Show-Skip "Invalid color format: $colorVal"
            $skipCount++
            Write-Host ""
            continue
        }

        $itemOk = $true
        try {
            $regDesktop = $hkcuInfo.PsDrivePath + '\Control Panel\Desktop'
            $regColors  = $hkcuInfo.PsDrivePath + '\Control Panel\Colors'

            Set-ItemProperty -Path $regDesktop -Name "WallPaper"  -Value ""        -ErrorAction Stop
            Set-ItemProperty -Path $regColors  -Name "Background" -Value $colorVal -ErrorAction Stop

            if ($hkcuInfo.Redirected) {
                Show-Success "Solid color staged for $($hkcuInfo.Label): $colorVal ($desc) (applies at next logon)"
            }
            else {
                [WallpaperHandler]::SystemParametersInfo(
                    [WallpaperHandler]::SPI_SETDESKWALLPAPER, 0, '',
                    [WallpaperHandler]::SPIF_UPDATEINIFILE -bor [WallpaperHandler]::SPIF_SENDCHANGE
                ) | Out-Null
                Show-Success "Solid color applied: $colorVal ($desc)"
            }
        }
        catch {
            Show-Error "Error applying solid color '$colorVal': $_"
            $itemOk = $false
        }
        if (-not $itemOk) { $failCount++; Write-Host ""; continue }

        # Post-Apply Verification (current/next-logon user)
        $ok = (Test-RegStringEquals $regDesktop 'WallPaper' '') -and `
              (Test-RegStringEquals $regColors  'Background' $colorVal)
        if ($ok) { $verifyPass++ } else { $verifyFail++ }

        $finalState = @{ Type = 'SolidColor'; Color = $colorVal }
        $successCount++
        Write-Host ""
        continue
    }

    # ---------------------------------- Image ----------------------------------
    $imagePath = if ([System.IO.Path]::IsPathRooted($item.FileName)) {
        $item.FileName
    } else {
        Join-Path $wallpaperDir $item.FileName
    }
    $ext      = [System.IO.Path]::GetExtension($item.FileName).ToLower()
    $desc     = if ($item.Description) { $item.Description } else { $item.FileName }
    $styleKey = if ($item.Style -and $styleMap.ContainsKey($item.Style)) { $item.Style } else { "Fill" }

    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "Applying: $desc" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor White

    if (-not (Test-Path $imagePath)) {
        Show-Skip "File not found: $($item.FileName)"
        Write-Host ""
        $skipCount++
        continue
    }
    if ($ext -notin $validExtensions) {
        Show-Skip "Unsupported extension: $ext"
        Write-Host ""
        $skipCount++
        continue
    }
    if ($item.Style -and -not $styleMap.ContainsKey($item.Style)) {
        Show-Warning "Unknown Style '$($item.Style)' - falling back to Fill"
    }

    $itemOk   = $true
    $destPath = $null
    try {
        # Normalize the image to a local, world-readable path. The registry /
        # SPI / Active Setup only ever reference this local path, never the
        # (possibly removable) source.
        $destPath = Resolve-WallpaperLocalDest -SourcePath $imagePath -BaseDir $wallpaperLocalDir
        if (-not $destPath) { throw "Unsafe destination filename for '$($item.FileName)'" }
        Copy-WallpaperToLocal -Source $imagePath -Dest $destPath

        $regPath = $hkcuInfo.PsDrivePath + '\Control Panel\Desktop'
        Set-ItemProperty -Path $regPath -Name "WallpaperStyle" -Value $styleMap[$styleKey].WallpaperStyle -ErrorAction Stop
        Set-ItemProperty -Path $regPath -Name "TileWallpaper"  -Value $styleMap[$styleKey].TileWallpaper  -ErrorAction Stop

        if ($hkcuInfo.Redirected) {
            Set-ItemProperty -Path $regPath -Name "WallPaper" -Value $destPath -ErrorAction Stop
            Show-Success "Wallpaper staged for $($hkcuInfo.Label): $($item.FileName) (Style: $styleKey, applies at next logon)"
        }
        else {
            $apiResult = [WallpaperHandler]::SystemParametersInfo(
                [WallpaperHandler]::SPI_SETDESKWALLPAPER,
                0,
                $destPath,
                [WallpaperHandler]::SPIF_UPDATEINIFILE -bor [WallpaperHandler]::SPIF_SENDCHANGE
            )
            if ($apiResult -eq 0) { throw "SystemParametersInfo returned 0 (failed)" }
            Show-Success "Wallpaper applied: $($item.FileName) (Style: $styleKey)"
        }
    }
    catch {
        Show-Error "Error applying wallpaper '$($item.FileName)': $_"
        $itemOk = $false
    }
    if (-not $itemOk) { $failCount++; Write-Host ""; continue }

    # Post-Apply Verification (current/next-logon user)
    $vDesktop = $hkcuInfo.PsDrivePath + '\Control Panel\Desktop'
    $ok = (Test-Path $destPath) -and (Test-RegStringEquals $vDesktop 'WallPaper' $destPath)
    if ($ok) { $verifyPass++ } else { $verifyFail++ }

    $finalState = @{
        Type  = 'Image'
        Path  = $destPath
        Style = $styleMap[$styleKey].WallpaperStyle
        Tile  = $styleMap[$styleKey].TileWallpaper
    }
    $successCount++
    Write-Host ""
}

# ========================================
# New-user reflection via Active Setup
# Register a per-user first-logon script that re-applies the final desktop
# state (the Default-profile hive value is unreliable - the theme engine
# overwrites it at first logon).
# ========================================
$newUserDegraded = $false

if ($null -ne $finalState) {
    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "Registering new-user reflection (Active Setup)" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor White

    $asLines = New-WallpaperActiveSetupScript -State $finalState
    $asOk = Register-FabriqActiveSetup -GUID $ACTIVE_SETUP_GUID `
                                       -Description "Fabriq default wallpaper" `
                                       -ScriptName $ACTIVE_SETUP_NAME `
                                       -ScriptLines $asLines

    $asKey = "HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\$ACTIVE_SETUP_GUID"
    if ($asOk -and (Test-Path $asKey) -and (Test-Path $ACTIVE_SETUP_SCRIPT)) {
        Show-Success "New users will receive the wallpaper at first logon (Active Setup)"
    }
    else {
        Show-Warning "Active Setup registration failed - new users may not receive the wallpaper"
        $newUserDegraded = $true
    }
    Write-Host ""
}

# ========================================
# Result Summary
# ========================================
# Verified reflects every operation that ran after confirmation: the Desktop
# Spotlight disable, the wallpaper verifications, and the Active Setup
# registration. (Spotlight always runs, so Verified is always a real bool.)
$verified = (-not $spotlightDegraded -and -not $hiveUnloadFailed -and $verifyFail -eq 0 -and -not $newUserDegraded)

if ($hiveUnloadFailed) {
    # A locked ntuser.dat can break downstream sysprep / new-profile creation.
    return (New-BatchResult -Success $successCount -Skip $skipCount -Fail ($failCount + 1) `
        -Title "Wallpaper Configuration Results" `
        -MessageSuffix "(Default hive unload FAILED - ntuser.dat may remain locked)" -Verified $false)
}

$notes = @()
if ($spotlightDegraded) { $notes += "Desktop Spotlight not fully disabled" }
if ($newUserDegraded)   { $notes += "new-user Active Setup registration failed" }
$msgSuffix = if ($notes.Count -gt 0) { "(" + ($notes -join "; ") + ")" } else { "" }

return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount `
    -Title "Wallpaper Configuration Results" -MessageSuffix $msgSuffix -Verified $verified)
