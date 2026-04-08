# ========================================
# Fabriq Launcher Build Script
# Compiles Launcher.cs into Fabriq.exe using csc.exe (.NET Framework).
# Produces the final binary at the fabriq root: ..\..\Fabriq.exe
# ========================================

[CmdletBinding()]
param(
    [switch]$Force
)

$ErrorActionPreference = 'Stop'

# Load fabriq common library for unified logging
$scriptDir   = Split-Path -Parent $MyInvocation.MyCommand.Path
$fabriqRoot  = Resolve-Path (Join-Path $scriptDir '..\..')
$commonPath  = Join-Path $fabriqRoot 'kernel\common.ps1'

if (-not (Test-Path $commonPath)) {
    Write-Host "[ERROR] kernel\common.ps1 not found at: $commonPath" -ForegroundColor Red
    exit 1
}
. $commonPath

Show-Separator
Show-CategorySeparator "Fabriq Launcher Build"
Show-Separator

# ========================================
# Step 1: Locate csc.exe
# ========================================
Show-Info "Locating csc.exe (.NET Framework compiler)..."

$cscCandidates = @(
    'C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe',
    'C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe'
)

$csc = $cscCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1

if (-not $csc) {
    Show-Error ".NET Framework 4.x csc.exe was not found."
    Show-Error "Expected one of:"
    $cscCandidates | ForEach-Object { Show-Error "  $_" }
    exit 1
}
Show-Success "csc.exe: $csc"

# ========================================
# Step 2: Ensure icon exists (generate fallback if missing)
# ========================================
$iconPath = Join-Path $scriptDir 'fabriq.ico'
if (-not (Test-Path $iconPath)) {
    Show-Info "fabriq.ico not found. Generating a multi-size fallback icon procedurally..."
    try {
        Add-Type -AssemblyName System.Drawing

        # P/Invoke DestroyIcon for cleaning up HICONs from Bitmap.GetHicon().
        $iconApiSig = @'
using System;
using System.Runtime.InteropServices;
public static class IconApi {
    [DllImport("user32.dll")]
    public static extern bool DestroyIcon(IntPtr hIcon);
}
'@
        if (-not ([System.Management.Automation.PSTypeName]'IconApi').Type) {
            Add-Type -TypeDefinition $iconApiSig
        }

        # Draw a simple "F" badge at each size. Using our own drawing code
        # (rather than extracting from shell32.dll) guarantees that every
        # size in the resulting .ico contains visually identical artwork,
        # which avoids the issue where Windows Explorer shows different
        # images depending on icon view size / cache state.
        $sizes   = @(16, 32, 48, 256)
        $perSize = @{}

        # fabriq brand-ish colors (deep indigo background, white glyph)
        $bgTop     = [System.Drawing.Color]::FromArgb(255, 30, 45, 95)
        $bgBottom  = [System.Drawing.Color]::FromArgb(255, 15, 25, 55)
        $fgColor   = [System.Drawing.Color]::White
        $edgeColor = [System.Drawing.Color]::FromArgb(255, 80, 120, 200)

        foreach ($sz in $sizes) {
            $bmp = New-Object System.Drawing.Bitmap($sz, $sz, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
            $g   = [System.Drawing.Graphics]::FromImage($bmp)
            try {
                $g.SmoothingMode     = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
                $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
                $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit

                # Rounded-rect background with vertical gradient
                $pad  = [int][math]::Max(1, $sz * 0.06)
                $rect = New-Object System.Drawing.Rectangle($pad, $pad, ($sz - $pad * 2), ($sz - $pad * 2))
                $radius = [int][math]::Max(2, $sz * 0.18)

                $path = New-Object System.Drawing.Drawing2D.GraphicsPath
                $d = $radius * 2
                $path.AddArc($rect.X, $rect.Y, $d, $d, 180, 90)
                $path.AddArc(($rect.Right - $d), $rect.Y, $d, $d, 270, 90)
                $path.AddArc(($rect.Right - $d), ($rect.Bottom - $d), $d, $d, 0, 90)
                $path.AddArc($rect.X, ($rect.Bottom - $d), $d, $d, 90, 90)
                $path.CloseFigure()

                $brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
                    $rect, $bgTop, $bgBottom,
                    [System.Drawing.Drawing2D.LinearGradientMode]::Vertical)
                $g.FillPath($brush, $path)
                $brush.Dispose()

                # Subtle edge highlight
                if ($sz -ge 24) {
                    $pen = New-Object System.Drawing.Pen($edgeColor, [single]([math]::Max(1, $sz * 0.03)))
                    $g.DrawPath($pen, $path)
                    $pen.Dispose()
                }
                $path.Dispose()

                # "F" glyph, centered
                # Size the font so the glyph fills most of the inner area
                $fontSize = [single]([math]::Max(6, $sz * 0.62))
                $font = New-Object System.Drawing.Font(
                    'Arial Black', $fontSize,
                    [System.Drawing.FontStyle]::Bold,
                    [System.Drawing.GraphicsUnit]::Pixel)

                $sf = New-Object System.Drawing.StringFormat
                $sf.Alignment     = [System.Drawing.StringAlignment]::Center
                $sf.LineAlignment = [System.Drawing.StringAlignment]::Center

                $fgBrush = New-Object System.Drawing.SolidBrush($fgColor)
                # Optical centering: "F" leans left, nudge slightly right
                $cx = $sz / 2.0 + ($sz * 0.03)
                $cy = $sz / 2.0 - ($sz * 0.02)
                $g.DrawString('F', $font, $fgBrush, [single]$cx, [single]$cy, $sf)
                $fgBrush.Dispose()
                $font.Dispose()
                $sf.Dispose()
            }
            finally {
                $g.Dispose()
            }

            # Encode the bitmap directly as PNG bytes. PNG-in-ICO is supported
            # since Windows Vista and is MUCH more reliable than BMP-in-ICO,
            # especially at 256x256 where Windows is known to reject malformed
            # BMP entries silently and fall back to the default icon.
            $ms = New-Object System.IO.MemoryStream
            try {
                $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
                $perSize[$sz] = $ms.ToArray()
            }
            finally {
                $ms.Dispose()
                $bmp.Dispose()
            }
        }

        if ($perSize.Count -eq 0) {
            throw "Icon generation produced no output"
        }
        else {
            # Assemble a multi-image ICO file with PNG-encoded entries.
            # ICO layout:
            #   [ICONDIR 6B]
            #   [ICONDIRENTRY 16B x N]
            #   [PNG data x N]
            #
            # ICONDIRENTRY fields (16 bytes):
            #   0  : width  (1B, 0 = 256)
            #   1  : height (1B, 0 = 256)
            #   2  : color count (1B, 0 = no palette)
            #   3  : reserved (1B, 0)
            #   4-5: color planes (2B, 1 for PNG)
            #   6-7: bits per pixel (2B, 32 for 32bpp RGBA)
            #   8-11: image data size (4B)
            #   12-15: image data offset from start of ICO file (4B)

            $orderedSizes = $sizes | Where-Object { $perSize.ContainsKey($_) }
            $count = $orderedSizes.Count

            $out = New-Object System.IO.MemoryStream
            $bw  = New-Object System.IO.BinaryWriter($out)
            try {
                # ICONDIR
                $bw.Write([uint16]0)       # reserved
                $bw.Write([uint16]1)       # type: 1 = icon
                $bw.Write([uint16]$count)  # image count

                # First image data starts after header + all directory entries
                $dataOffset = 6 + (16 * $count)
                $offsets = @{}
                foreach ($sz in $orderedSizes) {
                    $offsets[$sz] = $dataOffset
                    $dataOffset += $perSize[$sz].Length
                }

                # Directory entries
                foreach ($sz in $orderedSizes) {
                    $pngLen = $perSize[$sz].Length
                    $wByte = if ($sz -ge 256) { [byte]0 } else { [byte]$sz }
                    $hByte = $wByte
                    $bw.Write([byte]$wByte)          # width
                    $bw.Write([byte]$hByte)          # height
                    $bw.Write([byte]0)               # colors (0 = >= 256 colors)
                    $bw.Write([byte]0)               # reserved
                    $bw.Write([uint16]1)             # planes
                    $bw.Write([uint16]32)            # bpp
                    $bw.Write([uint32]$pngLen)       # image data size
                    $bw.Write([uint32]$offsets[$sz]) # image data offset
                }

                # Image data (PNG blobs)
                foreach ($sz in $orderedSizes) {
                    $bw.Write($perSize[$sz])
                }

                $bw.Flush()
                [System.IO.File]::WriteAllBytes($iconPath, $out.ToArray())
            }
            finally {
                $bw.Dispose()
                $out.Dispose()
            }

            # Self-test: reload the written .ico via System.Drawing.Icon to make
            # sure it's well-formed before we hand it to csc.exe.
            try {
                $verify = New-Object System.Drawing.Icon($iconPath)
                $verify.Dispose()
            }
            catch {
                throw "Generated .ico failed self-test load: $_"
            }

            $sizeList = ($orderedSizes | ForEach-Object { "${_}x${_}" }) -join ', '
            Show-Success "Multi-size PNG icon saved: $iconPath ($sizeList, procedurally drawn)"
        }
    }
    catch {
        Show-Error "Failed to generate fallback icon: $_"
        exit 1
    }
} else {
    Show-Info "Using existing icon: $iconPath"
}

# ========================================
# Step 3: Verify source files
# ========================================
$sourceCs   = Join-Path $scriptDir 'Launcher.cs'
$manifest   = Join-Path $scriptDir 'app.manifest'
$outputExe  = Join-Path $fabriqRoot 'Fabriq.exe'

foreach ($required in @($sourceCs, $manifest)) {
    if (-not (Test-Path $required)) {
        Show-Error "Required file missing: $required"
        exit 1
    }
}

# Warn if Fabriq.exe is currently running (would lock the output path)
$running = Get-Process -Name 'Fabriq' -ErrorAction SilentlyContinue
if ($running) {
    Show-Warning "Fabriq.exe is currently running. Terminate it before rebuild."
    if (-not $Force) {
        Show-Error "Build aborted. Re-run with -Force to ignore (not recommended)."
        exit 1
    }
}

# ========================================
# Step 4: Compile
# ========================================
Show-Info "Compiling Launcher.cs -> Fabriq.exe ..."

$cscArgs = @(
    '/target:winexe',
    '/platform:anycpu',
    '/optimize+',
    '/nologo',
    "/win32icon:$iconPath",
    "/win32manifest:$manifest",
    "/out:$outputExe",
    $sourceCs
)

$proc = Start-Process -FilePath $csc -ArgumentList $cscArgs `
    -NoNewWindow -Wait -PassThru `
    -WorkingDirectory $scriptDir

if ($proc.ExitCode -ne 0) {
    Show-Error "csc.exe failed with exit code $($proc.ExitCode)"
    exit $proc.ExitCode
}

if (-not (Test-Path $outputExe)) {
    Show-Error "Build reported success but output is missing: $outputExe"
    exit 1
}

# ========================================
# Step 5: Report result
# ========================================
Show-Success "Build succeeded."
$info = Get-Item $outputExe
Show-Info "Output    : $($info.FullName)"
Show-Info "Size      : $([math]::Round($info.Length / 1KB, 2)) KB"
Show-Info "Product   : $($info.VersionInfo.ProductName)"
Show-Info "Version   : $($info.VersionInfo.ProductVersion)"
Show-Info "Company   : $($info.VersionInfo.CompanyName)"

Show-Separator
Show-CategorySeparator "Done"
Show-Separator
exit 0
