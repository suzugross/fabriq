# show * command implementations.

function Get-FabriqIosVersion {
    # Returns the fabriq_ios independent SemVer (from apps/fabriq_ios/VERSION).
    # Self-loads the VERSION file if $script:FabriqIosVersion was not seeded
    # by the entry point (e.g. when called from a smoke test that dot-sources
    # only the lib files).
    if ($script:FabriqIosVersion) { return $script:FabriqIosVersion }
    $verFile = Join-Path $script:FabriqIosRoot 'VERSION'
    if (Test-Path $verFile) {
        $script:FabriqIosVersion = (Get-Content -Path $verFile -Raw).Trim()
        return $script:FabriqIosVersion
    }
    return '0.0.0'
}

function Invoke-ShowCommand {
    param(
        [string[]]$ArgList,
        [hashtable]$State
    )
    if (-not $ArgList -or $ArgList.Count -eq 0) {
        Write-Host "% Incomplete command: 'show' requires a sub-command (try 'show ?')" -ForegroundColor Red
        return
    }
    switch ($ArgList[0]) {
        'version'        { Show-FabriqIosVersion }
        'manifesto'      { Show-FabriqIosManifesto }
        'profiles'       { Show-FabriqIosProfiles }
        'modules'        { Show-FabriqIosModules }
        'evidence'       { Show-FabriqIosEvidence }
        'running-config' { Show-FabriqIosRunningConfig  -State $State }
        default {
            Write-Host ("% Unknown show subcommand: {0}" -f $ArgList[0]) -ForegroundColor Red
        }
    }
}

function Show-FabriqIosVersion {
    $bannerPath = Join-Path $script:FabriqIosRoot 'data\version_banner.txt'
    if (Test-Path $bannerPath) {
        Get-Content -Path $bannerPath -Raw -Encoding UTF8 | Write-Host
    }
    Write-Host ("Fabriq IOS version:       {0}" -f (Get-FabriqIosVersion))
    Write-Host ("Subprocess PID:           {0}" -f $PID)
    Write-Host ""
}

function Show-FabriqIosProfiles {
    $profilesDir = Join-Path $script:FabriqRoot 'profiles'
    if (-not (Test-Path $profilesDir)) {
        Write-Host "% profiles/ not found" -ForegroundColor Red
        return
    }
    $files = @(Get-ChildItem -Path $profilesDir -Filter '*.csv' -File | Sort-Object Name)
    if ($files.Count -eq 0) {
        Write-Host "% No profile CSVs found" -ForegroundColor Red
        return
    }
    Write-Host ""
    foreach ($f in $files) {
        try {
            $rows = @(Import-Csv -Path $f.FullName -Encoding UTF8)
            $enabledCount = @($rows | Where-Object { $_.Enabled -eq '1' }).Count
            Write-Host ("  {0,-32} {1,3} entries ({2} enabled)" -f $f.Name, $rows.Count, $enabledCount)
        } catch {
            Write-Host ("  {0,-32} (parse error)" -f $f.Name)
        }
    }
    Write-Host ""
}

function Show-FabriqIosModules {
    $stdDir = Join-Path $script:FabriqRoot 'modules\standard'
    $extDir = Join-Path $script:FabriqRoot 'modules\extended'

    Write-Host ""
    foreach ($pair in @(@{ Title = 'Standard'; Path = $stdDir }, @{ Title = 'Extended'; Path = $extDir })) {
        if (-not (Test-Path $pair.Path)) { continue }
        Write-Host ("  {0} modules:" -f $pair.Title)
        $dirs = @(Get-ChildItem -Path $pair.Path -Directory | Sort-Object Name)
        foreach ($d in $dirs) {
            $csv = Join-Path $d.FullName 'module.csv'
            $menu = ''
            if (Test-Path $csv) {
                try {
                    $row = Import-Csv -Path $csv -Encoding UTF8 | Select-Object -First 1
                    if ($row -and $row.MenuName) { $menu = $row.MenuName }
                } catch {}
            }
            Write-Host ("    {0,-30} {1}" -f $d.Name, $menu)
        }
        Write-Host ""
    }
}

function Show-FabriqIosEvidence {
    $evidenceDir = Join-Path $script:FabriqRoot 'evidence'
    if (-not (Test-Path $evidenceDir)) {
        Write-Host "% evidence/ does not exist yet"
        return
    }
    $dirs = @(Get-ChildItem -Path $evidenceDir -Directory | Sort-Object LastWriteTime -Descending)
    if ($dirs.Count -eq 0) {
        Write-Host "% No evidence directories yet"
        return
    }
    Write-Host ""
    Write-Host "  Recent evidence directories (newest first):"
    $top = @($dirs | Select-Object -First 10)
    foreach ($d in $top) {
        $size = 0
        try {
            $size = (Get-ChildItem -Path $d.FullName -Recurse -File -ErrorAction SilentlyContinue |
                     Measure-Object -Property Length -Sum).Sum
        } catch {}
        $sizeMb = if ($size) { [Math]::Round($size / 1MB, 2) } else { 0 }
        Write-Host ("    {0:yyyy-MM-dd HH:mm}  {1,-40} {2,7} MB" -f $d.LastWriteTime, $d.Name, $sizeMb)
    }
    if ($dirs.Count -gt 10) {
        Write-Host ("    ... and {0} more" -f ($dirs.Count - 10))
    }
    Write-Host ""
}

function Show-FabriqIosManifesto {
    # The kernel's manifesto.ps1 is a WinForms GUI viewer, unsuitable
    # for a terminal session. We emit a poetic syslog quote here and
    # point users at the operator dashboard for the full text.
    Write-Host ""
    Write-FabriqIosSyslog -Severity 7 -Mnemonic 'MANIFESTO' -Key 'quote' -Placeholders @{}
    Write-Host ""
    Write-Host "  (Full text available via the operator dashboard's"
    Write-Host "   [Manifeste du Surkitinisme] button.)"
    Write-Host ""
}

function Show-FabriqIosRunningConfig {
    param([hashtable]$State)

    # Cisco IOS shows the configuration definitions in `show
    # running-config`; on Windows the most truthful analog is
    # ipconfig /all (live interface state, MAC, lease, gateway, DNS).
    # We deliberately deviate from Cisco semantics and emit real
    # ipconfig output, wrapped in the Cisco-style preamble / footer
    # so the joke aesthetic still survives. ipconfig is invoked
    # directly (not captured into a variable) so the console
    # encoding handles the locale-dependent output without any
    # PowerShell pipeline re-interpretation that would mojibake on
    # JP Windows.
    Write-Host ""
    Write-Host "Building configuration..."
    Write-Host ""
    Write-Host "Current configuration : (live, sourced from ipconfig /all)"
    Write-Host "!"
    Write-Host "! Surkittinist artefact - Windows is not Cisco; the body is honest."
    Write-Host "!"
    Write-Host ("version {0}" -f (Get-FabriqIosVersion))
    Write-Host "!"
    Write-Host "banner motd ^C"
    Write-Host "    Surkittinism is the convulsive beauty of mass deployment, or it is nothing."
    Write-Host "^C"
    Write-Host "!"
    try {
        & ipconfig.exe /all
    } catch {
        Write-Host ("! ipconfig invocation failed: {0}" -f $_.Exception.Message)
    }
    Write-Host "!"
    Write-Host "end"
    Write-Host ""
}
