# ========================================
# System Launcher
# ========================================
# Quick access to Windows settings, control panel, and system tools
# without using Windows Search (no search history left behind).
# ========================================

$tools = @(
    # --- Settings (ms-settings) ---
    [PSCustomObject]@{ Num=1;  Name="Display";            Command="ms-settings:display";          Type="uri";     Category="Settings (ms-settings)" }
    [PSCustomObject]@{ Num=2;  Name="Sound";              Command="ms-settings:sound";             Type="uri";     Category="Settings (ms-settings)" }
    [PSCustomObject]@{ Num=3;  Name="Notifications";      Command="ms-settings:notifications";     Type="uri";     Category="Settings (ms-settings)" }
    [PSCustomObject]@{ Num=4;  Name="Network";            Command="ms-settings:network";           Type="uri";     Category="Settings (ms-settings)" }
    [PSCustomObject]@{ Num=5;  Name="Wi-Fi";              Command="ms-settings:network-wifi";      Type="uri";     Category="Settings (ms-settings)" }
    [PSCustomObject]@{ Num=6;  Name="Proxy";              Command="ms-settings:network-proxy";     Type="uri";     Category="Settings (ms-settings)" }
    [PSCustomObject]@{ Num=7;  Name="Apps & Features";    Command="ms-settings:appsfeatures";      Type="uri";     Category="Settings (ms-settings)" }
    [PSCustomObject]@{ Num=8;  Name="Default Apps";       Command="ms-settings:defaultapps";       Type="uri";     Category="Settings (ms-settings)" }
    [PSCustomObject]@{ Num=9;  Name="Taskbar";            Command="ms-settings:taskbar";           Type="uri";     Category="Settings (ms-settings)" }
    [PSCustomObject]@{ Num=10; Name="Personalization";    Command="ms-settings:personalization";   Type="uri";     Category="Settings (ms-settings)" }
    [PSCustomObject]@{ Num=11; Name="Date & Time";        Command="ms-settings:dateandtime";       Type="uri";     Category="Settings (ms-settings)" }
    [PSCustomObject]@{ Num=12; Name="Region";             Command="ms-settings:regionformatting";  Type="uri";     Category="Settings (ms-settings)" }
    [PSCustomObject]@{ Num=13; Name="Windows Update";     Command="ms-settings:windowsupdate";     Type="uri";     Category="Settings (ms-settings)" }
    [PSCustomObject]@{ Num=14; Name="About";              Command="ms-settings:about";             Type="uri";     Category="Settings (ms-settings)" }

    # --- Control Panel ---
    [PSCustomObject]@{ Num=15; Name="Control Panel (All)"; Command="control.exe";                  Type="exe";     Category="Control Panel" }
    [PSCustomObject]@{ Num=16; Name="Network Connections"; Command="ncpa.cpl";                     Type="exe";     Category="Control Panel" }
    [PSCustomObject]@{ Num=17; Name="Programs & Features"; Command="appwiz.cpl";                   Type="exe";     Category="Control Panel" }
    [PSCustomObject]@{ Num=18; Name="System Properties";   Command="sysdm.cpl";                    Type="exe";     Category="Control Panel" }
    [PSCustomObject]@{ Num=19; Name="Power Options";       Command="powercfg.cpl";                 Type="exe";     Category="Control Panel" }
    [PSCustomObject]@{ Num=20; Name="Sound (Legacy)";      Command="mmsys.cpl";                    Type="exe";     Category="Control Panel" }
    [PSCustomObject]@{ Num=21; Name="Firewall";            Command="firewall.cpl";                 Type="exe";     Category="Control Panel" }
    [PSCustomObject]@{ Num=22; Name="User Accounts";       Command="netplwiz";                     Type="exe";     Category="Control Panel" }

    # --- System Tools ---
    [PSCustomObject]@{ Num=23; Name="God Mode";            Command="shell:::{ED7BA470-8E54-465E-825C-99712043E01C}"; Type="shell"; Category="System Tools" }
    [PSCustomObject]@{ Num=24; Name="Device Manager";      Command="devmgmt.msc";                  Type="exe";     Category="System Tools" }
    [PSCustomObject]@{ Num=25; Name="Computer Management"; Command="compmgmt.msc";                 Type="exe";     Category="System Tools" }
    [PSCustomObject]@{ Num=26; Name="Event Viewer";        Command="eventvwr.msc";                 Type="exe";     Category="System Tools" }
    [PSCustomObject]@{ Num=27; Name="Services";            Command="services.msc";                 Type="exe";     Category="System Tools" }
    [PSCustomObject]@{ Num=28; Name="Task Scheduler";      Command="taskschd.msc";                 Type="exe";     Category="System Tools" }
    [PSCustomObject]@{ Num=29; Name="Group Policy";        Command="gpedit.msc";                   Type="exe";     Category="System Tools" }
    [PSCustomObject]@{ Num=30; Name="Registry Editor";     Command="regedit.exe";                  Type="exe";     Category="System Tools" }
    [PSCustomObject]@{ Num=31; Name="Disk Management";     Command="diskmgmt.msc";                 Type="exe";     Category="System Tools" }

    # --- Shell ---
    [PSCustomObject]@{ Num=32; Name="CMD";                 Command="cmd.exe";                      Type="exe";     Category="Shell" }
    [PSCustomObject]@{ Num=33; Name="PowerShell";          Command="powershell.exe";               Type="exe";     Category="Shell" }
    [PSCustomObject]@{ Num=34; Name="PowerShell (Admin)";  Command="powershell.exe";               Type="runas";   Category="Shell" }
)

function Show-LauncherMenu {
    Write-Host ""
    Show-Separator
    Write-Host "System Launcher" -ForegroundColor Cyan
    Show-Separator
    Write-Host ""

    $currentCategory = ""
    foreach ($t in $tools) {
        if ($t.Category -ne $currentCategory) {
            $currentCategory = $t.Category
            Write-Host "  --- $currentCategory ---" -ForegroundColor DarkCyan
        }

        $numStr = "[{0,2}]" -f $t.Num
        $nameStr = $t.Name.PadRight(22)

        if ($t.Num % 2 -eq 1) {
            Write-Host "    $numStr $nameStr" -NoNewline -ForegroundColor White
        }
        else {
            Write-Host "  $numStr $($t.Name)" -ForegroundColor White
        }
    }

    if ($tools.Count % 2 -eq 1) {
        Write-Host ""
    }

    Write-Host ""
    Write-Host "    [ 0] Return" -ForegroundColor Gray
    Write-Host ""
}

function Invoke-Tool {
    param([PSCustomObject]$Tool)

    switch ($Tool.Type) {
        "uri" {
            Start-Process $Tool.Command -ErrorAction SilentlyContinue
        }
        "shell" {
            Start-Process explorer.exe $Tool.Command -ErrorAction SilentlyContinue
        }
        "runas" {
            Start-Process $Tool.Command -Verb RunAs -ErrorAction SilentlyContinue
        }
        "exe" {
            Start-Process $Tool.Command -ErrorAction SilentlyContinue
        }
    }
}

# Main loop
while ($true) {
    Show-LauncherMenu

    Write-Host -NoNewline "  Enter number: "
    $input = Read-Host

    if ($input -eq '0' -or [string]::IsNullOrWhiteSpace($input)) {
        Write-Host ""
        Show-Info "Returning to menu"
        Write-Host ""
        break
    }

    $num = 0
    if (-not [int]::TryParse($input, [ref]$num)) {
        Show-Warning "Invalid input: $input"
        continue
    }

    $selected = $tools | Where-Object { $_.Num -eq $num }
    if (-not $selected) {
        Show-Warning "Invalid number: $num"
        continue
    }

    Show-Info "Launching: $($selected.Name)..."
    try {
        Invoke-Tool -Tool $selected
        Show-Success "Launched: $($selected.Name)"
    }
    catch {
        Show-Error "Failed to launch: $($selected.Name) - $_"
    }
}
