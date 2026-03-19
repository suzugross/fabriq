# ========================================
# Evidence Collection Script
# ========================================

Show-Info "Executing evidence collection..."
Write-Host ""

# ========================================
# Directory and Path Settings
# ========================================
$pcName = if (-not [string]::IsNullOrEmpty($env:SELECTED_NEW_PCNAME)) {
    $env:SELECTED_NEW_PCNAME
} else {
    $env:COMPUTERNAME
}
$dateStr    = Get-Date -Format "yyyy_MM_dd_HHmmss"
$uid        = if ($global:FabriqUniqueId) { $global:FabriqUniqueId } else { Get-HardwareUniqueId }

if (-not [string]::IsNullOrWhiteSpace($global:FabriqEvidenceBasePath)) {
    # Unified path: flat (no date/uid/pc subfolder)
    $targetDir = Join-Path $global:FabriqEvidenceBasePath "pc_information"
}
else {
    # Fallback: legacy path with date/uid/pc subfolder
    $folderName = "${dateStr}_${uid}_${pcName}"
    $targetDir  = Join-Path $PSScriptRoot "..\..\..\evidence\pc_information\$folderName"
}

if (-not (Test-Path $targetDir)) {
    $null = New-Item -ItemType Directory -Path $targetDir -Force
}

$masterLogFile = Join-Path $targetDir "_ALL_${pcName}_Log.txt"
$currentSplitFile = $null

# ========================================
# Helper: Log Output (Console + Master + Split)
# ========================================
function Out-Log {
    param(
        [string]$Text,
        [ConsoleColor]$Color = "White"
    )
    Write-Host $Text -ForegroundColor $Color
    $Text | Out-File -FilePath $masterLogFile -Append -Encoding UTF8
    if (-not [string]::IsNullOrEmpty($currentSplitFile)) {
        $splitPath = Join-Path $targetDir $currentSplitFile
        $Text | Out-File -FilePath $splitPath -Append -Encoding UTF8
    }
}

# ========================================
# Helper: Start Section
# ========================================
function Start-Section {
    param(
        [string]$Title,
        [string]$FileName
    )
    $script:currentSplitFile = $FileName
    Out-Log ""
    Out-Log "========================================" -Color Cyan
    Out-Log "$Title" -Color Cyan
    Out-Log "========================================" -Color Cyan
}

# ========================================
# Display Settings
# ========================================
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host "Evidence Collection" -ForegroundColor Cyan
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""
Write-Host "  Target PC:     $pcName" -ForegroundColor Yellow
Write-Host "  Save Location: $targetDir" -ForegroundColor White
Write-Host ""
Write-Host "  Sections:" -ForegroundColor Cyan
Write-Host "    [1] System Basic Info" -ForegroundColor White
Write-Host "    [2] Local Administrators" -ForegroundColor White
Write-Host "    [3] Network Settings (IP/DNS)" -ForegroundColor White
Write-Host "    [4] Printers / Ports List" -ForegroundColor White
Write-Host "    [5] BitLocker Status" -ForegroundColor White
Write-Host "    [6] MAC Address List" -ForegroundColor White
Write-Host "    [7] PC Serial Number" -ForegroundColor White
Write-Host "    [8] Installed Software List (CSV)" -ForegroundColor White
Write-Host "    [9] Firewall Status (CSV)" -ForegroundColor White
Write-Host "    [10] Windows Optional Features (CSV)" -ForegroundColor White
Write-Host "    [11] Server Roles & Features (CSV) *Server only" -ForegroundColor White
Write-Host ""
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""

# ========================================
# Confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Collect evidence for '$pcName'?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""

# ========================================
# Collection Process
# ========================================
$sectionCount = 0
$failCount = 0

$now = Get-Date -Format "yyyy/MM/dd HH:mm:ss.ff"
$currentSplitFile = $null
Out-Log "==== Evidence Log ====" -Color Cyan
Out-Log "Date: $now"
Out-Log "Computer: $pcName"
Out-Log "Save Location: $targetDir"

# ----------------------------------------
# 1. Basic Info (Hostname / OS / Specs)
# ----------------------------------------
Start-Section -Title "System Basic Info" -FileName "01_SystemInfo.txt"

try {
    Out-Log "Hostname:       $env:COMPUTERNAME"

    $os = Get-CimInstance Win32_OperatingSystem
    Out-Log "OS Name:        $($os.Caption)"
    Out-Log "Version:        $($os.Version) (Build: $($os.BuildNumber))"

    $cpu = Get-CimInstance Win32_Processor
    Out-Log "CPU:            $($cpu.Name)"

    $mem = Get-CimInstance Win32_PhysicalMemory | Measure-Object -Property Capacity -Sum
    $memGB = [Math]::Round($mem.Sum / 1GB, 1)
    Out-Log "Memory:         $memGB GB"

    $sectionCount++
}
catch {
    Out-Log "[ERROR] Failed to get basic info: $_" -Color Red
    $failCount++
}

# ----------------------------------------
# 2. Local Administrators
# ----------------------------------------
Start-Section -Title "Local Administrators" -FileName "02_LocalAdmins.txt"

try {
    $admins = Get-LocalGroupMember -Group "Administrators"
    foreach ($admin in $admins) {
        Out-Log "  - $($admin.Name) ($($admin.ObjectClass))"
    }
    $sectionCount++
}
catch {
    Out-Log "[ERROR] Failed to get administrator info: $_" -Color Red
    $failCount++
}

# ----------------------------------------
# 3. IP / DNS Settings
# ----------------------------------------
Start-Section -Title "Network Settings (IP/DNS)" -FileName "03_NetworkConfig.txt"

try {
    $netConfigs = Get-NetIPConfiguration | Where-Object { $_.IPv4Address -ne $null }
    foreach ($nc in $netConfigs) {
        Out-Log "Interface: $($nc.InterfaceAlias)"
        Out-Log "  IPv4 Address:   $($nc.IPv4Address.IPAddress)"

        # Subnet Mask: PrefixLength → dotted-decimal conversion
        $ipEntry = Get-NetIPAddress -InterfaceIndex $nc.InterfaceIndex `
                   -AddressFamily IPv4 -ErrorAction SilentlyContinue |
                   Where-Object { $_.PrefixOrigin -ne "WellKnown" } |
                   Select-Object -First 1
        if ($ipEntry) {
            $prefixLen = $ipEntry.PrefixLength
            $maskInt = if ($prefixLen -gt 0) {
                [uint32]([math]::Pow(2, 32) - [math]::Pow(2, 32 - $prefixLen))
            } else { [uint32]0 }
            $subnet = "{0}.{1}.{2}.{3}" -f `
                (($maskInt -shr 24) -band 0xFF),
                (($maskInt -shr 16) -band 0xFF),
                (($maskInt -shr 8) -band 0xFF),
                ($maskInt -band 0xFF)
            Out-Log "  Subnet Mask:    $subnet"
        }

        Out-Log "  Default Gateway: $($nc.IPv4DefaultGateway.NextHop)"
        Out-Log "  DNS Servers:     $($nc.DNSServer.ServerAddresses -join ', ')"
        Out-Log ""
    }
    $sectionCount++
}
catch {
    Out-Log "[ERROR] Failed to get IP settings: $_" -Color Red
    $failCount++
}

# ----------------------------------------
# 4. Printers / Ports List
# ----------------------------------------
Start-Section -Title "Printers / Ports List" -FileName "04_Printers.txt"

try {
    $printers = Get-Printer -ErrorAction SilentlyContinue
    if ($printers) {
        foreach ($p in $printers) {
            Out-Log "Name=$($p.Name)|Driver=$($p.DriverName)|Port=$($p.PortName)"
        }
    } else {
        Out-Log "(No printers installed)"
    }
    $sectionCount++
}
catch {
    Out-Log "[ERROR] Failed to get printer info: $_" -Color Red
    $failCount++
}

# ----------------------------------------
# 5. BitLocker Status
# ----------------------------------------
Start-Section -Title "BitLocker Status" -FileName "05_BitLocker.txt"

try {
    $volumes = Get-BitLockerVolume
    foreach ($v in $volumes) {
        Out-Log "Volume $($v.MountPoint) [$($v.VolumeType)]"
        Out-Log "    Size:                 $( [Math]::Round($v.CapacityGB, 2) ) GB"
        Out-Log "    BitLocker Version:    $($v.BitLockerVersion)"
        Out-Log "    Conversion Status:    $($v.VolumeStatus)"
        Out-Log "    Encryption Percentage: $($v.EncryptionPercentage)%"
        Out-Log "    Encryption Method:    $($v.EncryptionMethod)"
        Out-Log "    Protection Status:    $($v.ProtectionStatus)"

        Out-Log "    Key Protectors:"
        foreach ($key in $v.KeyProtector) {
            Out-Log "        $($key.KeyProtectorType)"
        }
        Out-Log ""
    }
    $sectionCount++
}
catch {
    Out-Log "[ERROR] Failed to get BitLocker info: $_" -Color Red
    $failCount++
}

# ----------------------------------------
# 6. MAC Address List
# ----------------------------------------
Start-Section -Title "MAC Address List" -FileName "06_MacAddress.txt"

try {
    $adapters = Get-NetAdapter | Select-Object Name, InterfaceDescription, MacAddress, Status
    foreach ($a in $adapters) {
        Out-Log "Connection Name:      $($a.Name)"
        Out-Log "Adapter:              $($a.InterfaceDescription)"
        Out-Log "Physical Address:     $($a.MacAddress)"
        Out-Log "Status:               $($a.Status)"
        Out-Log ""
    }
    $sectionCount++
}
catch {
    Out-Log "[ERROR] Failed to get network adapter info: $_" -Color Red
    $failCount++
}

# ----------------------------------------
# 7. PC Serial Number
# ----------------------------------------
Start-Section -Title "PC Serial Number" -FileName "07_SerialNumber.txt"

try {
    $bios = Get-CimInstance -ClassName Win32_BIOS
    Out-Log $bios.SerialNumber
    $sectionCount++
}
catch {
    Out-Log "[ERROR] Failed to get serial number: $_" -Color Red
    $failCount++
}

# ----------------------------------------
# 8. Installed Software List (CSV Export)
# ----------------------------------------
Start-Section -Title "Installed Software List (CSV)" -FileName $null

try {
    # 8a. Desktop Apps (Registry)
    $desktopPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    $desktop = Get-ItemProperty $desktopPaths -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName } |
        Select-Object @{N='Name';E={$_.DisplayName}},
                      @{N='Version';E={$_.DisplayVersion}},
                      Publisher,
                      InstallDate |
        Sort-Object Name

    $outDesktop = Join-Path $targetDir "08_DesktopApps.csv"
    $desktop | Export-Csv -Path $outDesktop -NoTypeInformation -Encoding UTF8

    Out-Log "Desktop apps: $($desktop.Count) items -> 08_DesktopApps.csv"

    # 8b. Store / UWP Apps
    $store = Get-AppxPackage |
        Select-Object @{N='Name';E={$_.Name}},
                      @{N='Version';E={$_.Version}},
                      @{N='Publisher';E={$_.PublisherId}} |
        Sort-Object Name

    $outStore = Join-Path $targetDir "08_StoreApps.csv"
    $store | Export-Csv -Path $outStore -NoTypeInformation -Encoding UTF8

    Out-Log "Store apps: $($store.Count) items -> 08_StoreApps.csv"

    $sectionCount++
}
catch {
    Out-Log "[ERROR] Failed to get software list: $_" -Color Red
    $failCount++
}

# OS type detection for server-only sections
$osProductType = (Get-CimInstance Win32_OperatingSystem).ProductType
$isServer = ($osProductType -ne 1)

# ----------------------------------------
# 9. Firewall Status (CSV Export)
# ----------------------------------------
Start-Section -Title "Firewall Status (CSV)" -FileName $null

try {
    # 9a. Firewall Profiles
    $fwProfiles = Get-NetFirewallProfile -ErrorAction Stop |
        Select-Object Name, Enabled, DefaultInboundAction, DefaultOutboundAction, LogFileName

    $outFwProfiles = Join-Path $targetDir "09_FirewallProfiles.csv"
    $fwProfiles | Export-Csv -Path $outFwProfiles -NoTypeInformation -Encoding UTF8

    Out-Log "Firewall profiles: $($fwProfiles.Count) profiles -> 09_FirewallProfiles.csv"

    # 9b. Firewall Rules
    $fwRules = Get-NetFirewallRule -ErrorAction Stop |
        Select-Object DisplayName, Enabled, Direction, Action, Profile |
        Sort-Object DisplayName

    $outFwRules = Join-Path $targetDir "09_FirewallRules.csv"
    $fwRules | Export-Csv -Path $outFwRules -NoTypeInformation -Encoding UTF8

    Out-Log "Firewall rules: $($fwRules.Count) rules -> 09_FirewallRules.csv"

    $sectionCount++
}
catch {
    Out-Log "[ERROR] Failed to get firewall info: $_" -Color Red
    $failCount++
}

# ----------------------------------------
# 10. Windows Optional Features (CSV Export)
# ----------------------------------------
Start-Section -Title "Windows Optional Features (CSV)" -FileName $null

try {
    $optFeatures = Get-WindowsOptionalFeature -Online -ErrorAction Stop |
        Select-Object FeatureName, State |
        Sort-Object FeatureName

    $outOptFeatures = Join-Path $targetDir "10_OptionalFeatures.csv"
    $optFeatures | Export-Csv -Path $outOptFeatures -NoTypeInformation -Encoding UTF8

    Out-Log "Optional features: $($optFeatures.Count) features -> 10_OptionalFeatures.csv"

    $sectionCount++
}
catch {
    Out-Log "[ERROR] Failed to get optional features: $_" -Color Red
    $failCount++
}

# ----------------------------------------
# 11. Server Roles & Features (CSV Export)
# ----------------------------------------
Start-Section -Title "Server Roles & Features (CSV)" -FileName $null

if ($isServer) {
    try {
        $serverFeatures = Get-WindowsFeature -ErrorAction Stop |
            Select-Object Name, DisplayName, InstallState, FeatureType |
            Sort-Object Name

        $outServerFeatures = Join-Path $targetDir "11_ServerRolesFeatures.csv"
        $serverFeatures | Export-Csv -Path $outServerFeatures -NoTypeInformation -Encoding UTF8

        Out-Log "Server roles & features: $($serverFeatures.Count) items -> 11_ServerRolesFeatures.csv"

        $sectionCount++
    }
    catch {
        Out-Log "[ERROR] Failed to get server features: $_" -Color Red
        $failCount++
    }
}
else {
    Out-Log "Skipped: Client OS detected (Server-only section)"
    $sectionCount++
}

# ----------------------------------------
# Completion
# ----------------------------------------
$currentSplitFile = $null
Out-Log ""
Out-Log "==== Evidence Collection Completed ====" -Color Cyan

Write-Host ""
Show-Info "Evidence saved to: $targetDir"
Write-Host ""

return (New-BatchResult -Success $sectionCount -Fail $failCount -Title "Evidence Collection Results")
