# show * command implementations.

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
        'host'           { Show-FabriqIosHost           -State $State }
        'hosts'          { Show-FabriqIosHosts          -State $State }
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
    $kernelVer = ''
    $kernelVerFile = Join-Path $script:FabriqRoot 'kernel\KERNEL_VERSION'
    if (Test-Path $kernelVerFile) {
        $kernelVer = (Get-Content -Path $kernelVerFile -Raw).Trim()
    }
    Write-Host ("Underlying Fabriq Kernel: {0}" -f $kernelVer)
    Write-Host ("Subprocess PID:           {0}" -f $PID)
    Write-Host ""
}

function Show-FabriqIosHost {
    param([hashtable]$State)

    if (-not $env:SELECTED_NEW_PCNAME) {
        Write-Host ""
        Write-Host "% No host is currently selected."
        Write-Host "  (Selected host comes from the parent fabriq session via SELECTED_*"
        Write-Host "   environment variables. Launch fabriq_ios after picking a host in"
        Write-Host "   the operator dashboard, or use 'show hosts' to view the catalogue.)"
        Write-Host ""
        return
    }

    Write-Host ""
    Write-Host "Selected Host:"
    Write-Host ("  AdminID:     {0}" -f $env:SELECTED_KANRI_NO)
    Write-Host ("  OldName:     {0}" -f $env:SELECTED_OLD_PCNAME)
    Write-Host ("  NewName:     {0}" -f $env:SELECTED_NEW_PCNAME)
    Write-Host ("  EthernetIP:  {0}/{1}" -f $env:SELECTED_ETH_IP, $env:SELECTED_ETH_SUBNET)
    Write-Host ("  Gateway:     {0}" -f $env:SELECTED_ETH_GATEWAY)
    $dns = @($env:SELECTED_DNS1, $env:SELECTED_DNS2, $env:SELECTED_DNS3, $env:SELECTED_DNS4) | Where-Object { $_ }
    Write-Host ("  DNS:         {0}" -f ($dns -join ', '))
    if ($env:SELECTED_WIFI_IP) {
        Write-Host ("  WiFiIP:      {0}/{1}" -f $env:SELECTED_WIFI_IP, $env:SELECTED_WIFI_SUBNET)
    }
    Write-Host ""
}

function Show-FabriqIosHosts {
    param([hashtable]$State)

    $csvPath = Join-Path $script:FabriqRoot 'kernel\csv\hostlist.csv'
    if (-not (Test-Path $csvPath)) {
        Write-Host ("% hostlist.csv not found at {0}" -f $csvPath) -ForegroundColor Red
        return
    }

    if (-not $State.Passphrase) {
        Write-Host ""
        Write-Host "% Passphrase not set. ENC: cells will appear as raw ciphertext."
        Write-Host "  Use 'enable' to authenticate first."
    }

    $previousPass = $global:FabriqMasterPassphrase
    if ($State.Passphrase) {
        $global:FabriqMasterPassphrase = $State.Passphrase
    }
    try {
        $rows = @(Import-ModuleCsv -Path $csvPath)
    } finally {
        $global:FabriqMasterPassphrase = $previousPass
    }

    if (-not $rows -or $rows.Count -eq 0) {
        Write-Host "% No rows in hostlist.csv" -ForegroundColor Red
        return
    }

    $displayCols = @('AdminID','OldPCName','NewPCName','EthernetIP','EthernetGateway')
    $availCols = $rows[0].PSObject.Properties.Name
    $colsShown = @($displayCols | Where-Object { $_ -in $availCols })
    if ($colsShown.Count -eq 0) {
        $colsShown = @($availCols | Select-Object -First 5)
    }

    Write-Host ""
    $header = ($colsShown | ForEach-Object { '{0,-18}' -f $_ }) -join ' '
    Write-Host ("  " + $header)
    Write-Host ("  " + ('-' * $header.Length))
    foreach ($r in $rows) {
        $vals = $colsShown | ForEach-Object {
            $v = $r.$_
            $s = if ($null -eq $v) { '' } else { "$v" }
            if ($s.Length -gt 17) { $s.Substring(0, 15) + '..' } else { $s }
        }
        Write-Host ("  " + ((@($vals) | ForEach-Object { '{0,-18}' -f $_ }) -join ' '))
    }
    Write-Host ""
    Write-Host ("  Total: {0} hosts." -f $rows.Count)
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
    $kernelVer = ''
    $kernelVerFile = Join-Path $script:FabriqRoot 'kernel\KERNEL_VERSION'
    if (Test-Path $kernelVerFile) {
        $kernelVer = (Get-Content -Path $kernelVerFile -Raw).Trim()
    }
    $worker = if ($env:FABRIQ_WORKER_NAME) { $env:FABRIQ_WORKER_NAME } else { 'anonymous' }

    Write-Host ""
    Write-Host "Building configuration..."
    Write-Host ""
    Write-Host "Current configuration : (live, sourced from ipconfig /all)"
    Write-Host "!"
    Write-Host "! Surkittinist artefact - Windows is not Cisco; the body is honest."
    Write-Host "!"
    Write-Host ("version {0}" -f $kernelVer)
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
    Write-Host ("session worker {0}" -f $worker)
    Write-Host "!"
    Write-Host "end"
    Write-Host ""
}
