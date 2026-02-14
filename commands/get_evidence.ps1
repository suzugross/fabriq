# ========================================
# Kitting Information Retrieval Script (Split Directory Version)
# ========================================

# --- 1. Directory and Path Settings ---
$baseDir = ".\evidence\pc_information"
$dateStr = Get-Date -Format "yyyyMMdd_HHmmss"
$folderName = "$($env:COMPUTERNAME)_$dateStr"
$targetDir = Join-Path $baseDir $folderName

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

$printerList = Get-Printer | Select-Object Name, PortName | Format-Table -AutoSize | Out-String
Out-Log $printerList.TrimEnd()

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
# 8. Installed Software List
# ----------------------------------------
Start-Section -Title "Installed Software List" -FileName "08_InstalledSoftware.txt"

try {
    $uninstallKeys = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    
    $softwareList = Get-ItemProperty $uninstallKeys -ErrorAction SilentlyContinue | 
                    Where-Object { $_.DisplayName -ne $null } | 
                    Select-Object DisplayName, DisplayVersion | 
                    Sort-Object DisplayName | 
                    Format-Table -AutoSize | Out-String

    Out-Log $softwareList.TrimEnd()
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