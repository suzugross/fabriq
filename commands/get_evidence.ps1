# ========================================
# Kitting Information Retrieval Script (Split Directory Version)
# ========================================

# --- 1. Directory and Path Settings ---
if (-not [string]::IsNullOrWhiteSpace($global:FabriqEvidenceBasePath)) {
    # Unified path: flat (no date/pc subfolder)
    $targetDir = Join-Path $global:FabriqEvidenceBasePath "pc_information"
}
else {
    # Fallback: legacy path with date/pc subfolder
    $baseDir = ".\evidence\pc_information"
    $dateStr = Get-Date -Format "yyyyMMdd_HHmmss"
    $folderName = "$($env:COMPUTERNAME)_$dateStr"
    $targetDir = Join-Path $baseDir $folderName
}

# Create destination directory
if (-not (Test-Path $targetDir)) {
    New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
}

# Master Log File Path
$masterLogFile = Join-Path $targetDir "_ALL_$($env:COMPUTERNAME)_Log.txt"

# Variable to hold current split log filename (Initial value $null)
$currentSplitFile = $null

# --- 2. Log Output Function (Output to 3 locations) ---
function Out-Log {
    param(
        [string]$Text,
        [ConsoleColor]$Color = "White"
    )

    # A. Output to Console
    Write-Host $Text -ForegroundColor $Color

    # B. Output to Master Log
    $Text | Out-File -FilePath $masterLogFile -Append -Encoding UTF8

    # C. Output to Split Log (Only if configured)
    if (-not [string]::IsNullOrEmpty($currentSplitFile)) {
        $splitPath = Join-Path $targetDir $currentSplitFile
        $Text | Out-File -FilePath $splitPath -Append -Encoding UTF8
    }
}

# --- 3. Helper Function to Start Section ---
function Start-Section {
    param(
        [string]$Title,
        [string]$FileName
    )
    # Switch split log target
    $script:currentSplitFile = $FileName
    
    Out-Log ""
    Out-Log "========================================" -ForegroundColor Cyan
    Out-Log "$Title" -ForegroundColor Cyan
    Out-Log "========================================" -ForegroundColor Cyan
}

# ========================================
# Main Process Start
# ========================================

$now = Get-Date -Format "yyyy/MM/dd HH:mm:ss.ff"

# Header (No split file here)
$currentSplitFile = $null 
Out-Log "==== Kitting Log ====" -ForegroundColor Cyan
Out-Log "Date: $now"
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
}
catch {
    Out-Log "[ERROR] Failed to get basic info: $_" -ForegroundColor Red
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
}
catch {
    Out-Log "[ERROR] Failed to get administrator info: $_" -ForegroundColor Red
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
}
catch {
    Out-Log "[ERROR] Failed to get IP settings: $_" -ForegroundColor Red
}

# ----------------------------------------
# 4. Printers / Ports List
# ----------------------------------------
Start-Section -Title "Printers / Ports List" -FileName "04_Printers.txt"

$printers = Get-Printer -ErrorAction SilentlyContinue
if ($printers) {
    foreach ($p in $printers) {
        Out-Log "Name=$($p.Name)|Driver=$($p.DriverName)|Port=$($p.PortName)"
    }
} else {
    Out-Log "(No printers installed)"
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
}
catch {
    Out-Log "[ERROR] Failed to get BitLocker info: $_" -ForegroundColor Red
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
}
catch {
    Out-Log "[ERROR] Failed to get network adapter info: $_" -ForegroundColor Red
}

# ----------------------------------------
# 7. PC Serial Number
# ----------------------------------------
Start-Section -Title "PC Serial Number" -FileName "07_SerialNumber.txt"

try {
    $bios = Get-CimInstance -ClassName Win32_BIOS
    Out-Log $bios.SerialNumber
}
catch {
    Out-Log "[ERROR] Failed to get serial number: $_" -ForegroundColor Red
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
}
catch {
    Out-Log "[ERROR] Failed to get software list: $_" -ForegroundColor Red
}

# ----------------------------------------
# Completion Process
# ----------------------------------------
$currentSplitFile = $null # Stop split output
Out-Log ""
Out-Log "==== Log Output Completed ====" -ForegroundColor Cyan
Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host " All processes completed" -ForegroundColor Green
Write-Host " Destination Folder: $targetDir" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""