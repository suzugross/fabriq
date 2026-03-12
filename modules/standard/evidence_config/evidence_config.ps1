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
Write-Host "    [8] Installed Software List" -ForegroundColor White
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
    $printerList = Get-Printer | Select-Object Name, PortName | Format-Table -AutoSize | Out-String
    Out-Log $printerList.TrimEnd()
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
    $sectionCount++
}
catch {
    Out-Log "[ERROR] Failed to get software list: $_" -Color Red
    $failCount++
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
