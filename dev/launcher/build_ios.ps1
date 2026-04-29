# ========================================
# Fabriq IOS Launcher Build Script
# Compiles Launcher_IOS.cs into Fabriq_IOS.exe using csc.exe.
# Produces the final binary at the fabriq root: ..\..\Fabriq_IOS.exe
# Reuses fabriq.ico (no separate icon for the joke shell).
# ========================================

[CmdletBinding()]
param(
    [switch]$Force
)

$ErrorActionPreference = 'Stop'

$scriptDir   = Split-Path -Parent $MyInvocation.MyCommand.Path
$fabriqRoot  = Resolve-Path (Join-Path $scriptDir '..\..')
$commonPath  = Join-Path $fabriqRoot 'kernel\common.ps1'

if (-not (Test-Path $commonPath)) {
    Write-Host "[ERROR] kernel\common.ps1 not found at: $commonPath" -ForegroundColor Red
    exit 1
}
. $commonPath

Show-Separator
Show-CategorySeparator "Fabriq IOS Launcher Build"
Show-Separator

Show-Info "Locating csc.exe (.NET Framework compiler)..."
$cscCandidates = @(
    'C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe',
    'C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe'
)
$csc = $cscCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $csc) {
    Show-Error ".NET Framework 4.x csc.exe was not found."
    exit 1
}
Show-Success "csc.exe: $csc"

# Reuse the existing fabriq.ico (joke shell shares branding for now).
$iconPath = Join-Path $scriptDir 'fabriq.ico'
if (-not (Test-Path $iconPath)) {
    Show-Warning "fabriq.ico not found; building without an icon. Run build.ps1 first to generate one."
    $iconPath = $null
}

$sourceCs   = Join-Path $scriptDir 'Launcher_IOS.cs'
$manifest   = Join-Path $scriptDir 'app_ios.manifest'
$outputExe  = Join-Path $fabriqRoot 'Fabriq_IOS.exe'

foreach ($required in @($sourceCs, $manifest)) {
    if (-not (Test-Path $required)) {
        Show-Error "Required file missing: $required"
        exit 1
    }
}

$running = Get-Process -Name 'Fabriq_IOS' -ErrorAction SilentlyContinue
if ($running) {
    Show-Warning "Fabriq_IOS.exe is currently running. Terminate it before rebuild."
    if (-not $Force) {
        Show-Error "Build aborted. Re-run with -Force to ignore (not recommended)."
        exit 1
    }
}

Show-Info "Compiling Launcher_IOS.cs -> Fabriq_IOS.exe ..."

$cscArgs = @(
    '/target:winexe',
    '/platform:anycpu',
    '/optimize+',
    '/nologo',
    "/win32manifest:$manifest",
    "/out:$outputExe"
)
if ($iconPath) {
    $cscArgs += "/win32icon:$iconPath"
}
$cscArgs += $sourceCs

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

Show-Success "Build succeeded."
$info = Get-Item $outputExe
Show-Info "Output    : $($info.FullName)"
Show-Info "Size      : $([math]::Round($info.Length / 1KB, 2)) KB"
Show-Info "Product   : $($info.VersionInfo.ProductName)"
Show-Info "Version   : $($info.VersionInfo.ProductVersion)"

Show-Separator
Show-CategorySeparator "Done"
Show-Separator
exit 0
