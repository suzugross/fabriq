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
Write-Host "    [1]  System Basic Info" -ForegroundColor White
Write-Host "    [2]  Local Users (CSV)" -ForegroundColor White
Write-Host "    [3]  Local Groups (CSV)" -ForegroundColor White
Write-Host "    [4]  Local Group Members (CSV)" -ForegroundColor White
Write-Host "    [5]  Domain / Azure AD Status + User Profiles" -ForegroundColor White
Write-Host "    [6]  Network Settings (CSV)" -ForegroundColor White
Write-Host "    [7]  Printers / Ports List (CSV)" -ForegroundColor White
Write-Host "    [8]  BitLocker Status" -ForegroundColor White
Write-Host "    [8b] Disk & Partition Info (CSV)" -ForegroundColor White
Write-Host "    [9]  MAC Address List (CSV)" -ForegroundColor White
Write-Host "    [10] PC Serial Number" -ForegroundColor White
Write-Host "    [11] Installed Software List (CSV)" -ForegroundColor White
Write-Host "    [12] Firewall Status (CSV)" -ForegroundColor White
Write-Host "    [13] Windows Optional Features (CSV)" -ForegroundColor White
Write-Host "    [14] Server Roles & Features (CSV) *Server only" -ForegroundColor White
Write-Host "    [15] Power Settings" -ForegroundColor White
Write-Host "    [16] WiFi Profiles" -ForegroundColor White
Write-Host "    [17] Restore Points (CSV)" -ForegroundColor White
Write-Host "    [18] Windows Defender Status" -ForegroundColor White
Write-Host "    [19] Windows Update History (CSV)" -ForegroundColor White
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

    $tz = Get-TimeZone
    Out-Log "TimeZone:       $($tz.Id) (UTC$( if ($tz.BaseUtcOffset.TotalHours -ge 0) {'+'} )$($tz.BaseUtcOffset.TotalHours))"

    $culture = Get-Culture
    Out-Log "Locale:         $($culture.Name) ($($culture.DisplayName))"

    $sectionCount++
}
catch {
    Out-Log "[ERROR] Failed to get basic info: $_" -Color Red
    $failCount++
}

# ----------------------------------------
# 2. Local Users (CSV Export)
# ----------------------------------------
Start-Section -Title "Local Users (CSV)" -FileName $null

try {
    $localUsers = Get-LocalUser | Select-Object `
        Name, Enabled, FullName, Description, SID,
        LastLogon, PasswordLastSet, PasswordRequired,
        PasswordExpires, AccountExpires, PrincipalSource |
        Sort-Object Name

    $outLocalUsers = Join-Path $targetDir "02_LocalUsers.csv"
    $localUsers | Export-Csv -Path $outLocalUsers -NoTypeInformation -Encoding UTF8

    Out-Log "Local users: $($localUsers.Count) accounts -> 02_LocalUsers.csv"
    $sectionCount++
}
catch {
    Out-Log "[ERROR] Failed to get local users: $_" -Color Red
    $failCount++
}

# ----------------------------------------
# 3. Local Groups (CSV Export)
# ----------------------------------------
Start-Section -Title "Local Groups (CSV)" -FileName $null

try {
    $localGroups = Get-LocalGroup | Select-Object Name, Description, SID |
        Sort-Object Name

    $outLocalGroups = Join-Path $targetDir "03_LocalGroups.csv"
    $localGroups | Export-Csv -Path $outLocalGroups -NoTypeInformation -Encoding UTF8

    Out-Log "Local groups: $($localGroups.Count) groups -> 03_LocalGroups.csv"
    $sectionCount++
}
catch {
    Out-Log "[ERROR] Failed to get local groups: $_" -Color Red
    $failCount++
}

# ----------------------------------------
# 4. Local Group Members (CSV Export)
# ----------------------------------------
Start-Section -Title "Local Group Members (CSV)" -FileName $null

try {
    $allMembers = @()
    $groups = Get-LocalGroup

    foreach ($group in $groups) {
        try {
            $members = Get-LocalGroupMember -Group $group.Name -ErrorAction Stop
            foreach ($member in $members) {
                $allMembers += [PSCustomObject]@{
                    GroupName       = $group.Name
                    MemberName      = $member.Name
                    ObjectClass     = $member.ObjectClass
                    PrincipalSource = $member.PrincipalSource
                }
            }
        }
        catch {
            # Fallback: net localgroup (handles error 1789 / orphaned SIDs)
            Out-Log "  [WARN] Get-LocalGroupMember failed for '$($group.Name)', using fallback..." -Color Yellow
            try {
                $netOutput = net localgroup "$($group.Name)" 2>&1
                $inMembers = $false
                foreach ($line in $netOutput) {
                    $lineStr = "$line"
                    if ($lineStr -match '^----') { $inMembers = $true; continue }
                    if ($lineStr -match '^\s*$') { continue }
                    if ($lineStr -match '^(コマンドは正常に|The command completed)') { break }
                    if ($inMembers) {
                        $allMembers += [PSCustomObject]@{
                            GroupName       = $group.Name
                            MemberName      = $lineStr.Trim()
                            ObjectClass     = "Unknown"
                            PrincipalSource = "Unknown"
                        }
                    }
                }
            }
            catch {
                Out-Log "  [WARN] Fallback also failed for '$($group.Name)': $_" -Color Yellow
            }
        }
    }

    $outGroupMembers = Join-Path $targetDir "04_LocalGroupMembers.csv"
    $allMembers | Export-Csv -Path $outGroupMembers -NoTypeInformation -Encoding UTF8

    Out-Log "Group memberships: $($allMembers.Count) entries -> 04_LocalGroupMembers.csv"
    $sectionCount++
}
catch {
    Out-Log "[ERROR] Failed to get group members: $_" -Color Red
    $failCount++
}

# ----------------------------------------
# 5. Domain / Azure AD Status
# ----------------------------------------
Start-Section -Title "Domain / Azure AD Status" -FileName "05_DomainStatus.txt"

try {
    # 5a. Domain join status
    $cs = Get-CimInstance Win32_ComputerSystem
    $domainRoleMap = @{
        0 = "Standalone Workstation"
        1 = "Member Workstation"
        2 = "Standalone Server"
        3 = "Member Server"
        4 = "Backup Domain Controller"
        5 = "Primary Domain Controller"
    }
    $roleName = $domainRoleMap[[int]$cs.DomainRole]

    Out-Log "PartOfDomain:   $($cs.PartOfDomain)"
    Out-Log "Domain:         $($cs.Domain)"
    Out-Log "DomainRole:     $($cs.DomainRole) ($roleName)"
    Out-Log ""

    # 5b. Current user identity
    $currentIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    Out-Log "Current User:   $($currentIdentity.Name)"
    Out-Log ""

    # 5c. Azure AD / Entra ID status (dsregcmd)
    Out-Log "---- dsregcmd /status ----"
    $dsregOutput = dsregcmd /status 2>&1
    foreach ($line in $dsregOutput) {
        Out-Log "  $line"
    }
    Out-Log ""

    # 5d. User Profiles on this PC (CSV)
    Out-Log "---- User Profiles ----" -Color Cyan
    try {
        $profiles = Get-CimInstance Win32_UserProfile | Where-Object { -not $_.Special } |
            Select-Object @{N='LocalPath';E={$_.LocalPath}},
                          @{N='SID';E={$_.SID}},
                          @{N='LastUseTime';E={$_.LastUseTime}},
                          @{N='Loaded';E={$_.Loaded}} |
            Sort-Object LocalPath

        $outProfiles = Join-Path $targetDir "05_UserProfiles.csv"
        $profiles | Export-Csv -Path $outProfiles -NoTypeInformation -Encoding UTF8
        Out-Log "User profiles: $($profiles.Count) profiles -> 05_UserProfiles.csv"
    }
    catch {
        Out-Log "  [WARN] Could not retrieve user profiles: $_" -Color Yellow
    }

    $sectionCount++
}
catch {
    Out-Log "[ERROR] Failed to get domain/Azure AD status: $_" -Color Red
    $failCount++
}

# ----------------------------------------
# 6. IP / DNS Settings (CSV Export)
# ----------------------------------------
Start-Section -Title "Network Settings (CSV)" -FileName $null

try {
    $netConfigs = Get-NetIPConfiguration | Where-Object { $_.IPv4Address -ne $null }
    $networkRows = @()

    foreach ($nc in $netConfigs) {
        # Filter out APIPA / link-local addresses (169.254.x.x)
        $validIPs = @($nc.IPv4Address.IPAddress | Where-Object { $_ -notmatch '^169\.254\.' })
        if ($validIPs.Count -eq 0) { continue }

        # Subnet Mask: PrefixLength -> dotted-decimal conversion (exclude link-local)
        $subnet = ""
        $ipEntry = Get-NetIPAddress -InterfaceIndex $nc.InterfaceIndex `
                   -AddressFamily IPv4 -ErrorAction SilentlyContinue |
                   Where-Object { $_.PrefixOrigin -ne "WellKnown" -and $_.IPAddress -notmatch '^169\.254\.' } |
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
        }

        $networkRows += [PSCustomObject]@{
            Interface      = $nc.InterfaceAlias
            IPv4Address    = ($validIPs -join ', ')
            SubnetMask     = $subnet
            DefaultGateway = $nc.IPv4DefaultGateway.NextHop
            DNSServers     = ($nc.DNSServer.ServerAddresses -join ', ')
        }
    }

    $outNetwork = Join-Path $targetDir "06_NetworkConfig.csv"
    $networkRows | Export-Csv -Path $outNetwork -NoTypeInformation -Encoding UTF8

    Out-Log "Network interfaces: $($networkRows.Count) entries -> 06_NetworkConfig.csv"
    $sectionCount++
}
catch {
    Out-Log "[ERROR] Failed to get IP settings: $_" -Color Red
    $failCount++
}

# ----------------------------------------
# 7. Printers / Ports List (CSV Export)
# ----------------------------------------
Start-Section -Title "Printers / Ports List (CSV)" -FileName $null

try {
    $printers = Get-Printer -ErrorAction SilentlyContinue
    if ($printers) {
        $printerRows = $printers | Select-Object Name, DriverName, PortName, Shared, PrinterStatus |
            Sort-Object Name

        $outPrinters = Join-Path $targetDir "07_Printers.csv"
        $printerRows | Export-Csv -Path $outPrinters -NoTypeInformation -Encoding UTF8

        Out-Log "Printers: $($printerRows.Count) entries -> 07_Printers.csv"
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
# 8. BitLocker Status
# ----------------------------------------
Start-Section -Title "BitLocker Status" -FileName "08_BitLocker.txt"

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
# 8b. Disk & Partition Info (CSV Export)
# ----------------------------------------
Start-Section -Title "Disk & Partition Info (CSV)" -FileName $null

try {
    # Physical disks
    $disks = Get-Disk | Select-Object Number, FriendlyName, SerialNumber,
        @{N='SizeGB';E={[Math]::Round($_.Size / 1GB, 2)}},
        PartitionStyle, HealthStatus, OperationalStatus |
        Sort-Object Number

    $outDisks = Join-Path $targetDir "08b_Disks.csv"
    $disks | Export-Csv -Path $outDisks -NoTypeInformation -Encoding UTF8
    Out-Log "Physical disks: $($disks.Count) disk(s) -> 08b_Disks.csv"

    # Partitions
    $partitions = Get-Partition | Select-Object DiskNumber, PartitionNumber, DriveLetter,
        @{N='SizeGB';E={[Math]::Round($_.Size / 1GB, 2)}},
        Type, IsSystem, IsBoot, IsActive |
        Sort-Object DiskNumber, PartitionNumber

    $outPartitions = Join-Path $targetDir "08b_Partitions.csv"
    $partitions | Export-Csv -Path $outPartitions -NoTypeInformation -Encoding UTF8
    Out-Log "Partitions: $($partitions.Count) partition(s) -> 08b_Partitions.csv"

    $sectionCount++
}
catch {
    Out-Log "[ERROR] Failed to get disk/partition info: $_" -Color Red
    $failCount++
}

# ----------------------------------------
# 9. MAC Address List (CSV Export)
# ----------------------------------------
Start-Section -Title "MAC Address List (CSV)" -FileName $null

try {
    $adapters = Get-NetAdapter | Select-Object Name, InterfaceDescription, MacAddress, Status |
        Sort-Object Name

    $outMac = Join-Path $targetDir "09_MacAddress.csv"
    $adapters | Export-Csv -Path $outMac -NoTypeInformation -Encoding UTF8

    Out-Log "Network adapters: $($adapters.Count) entries -> 09_MacAddress.csv"
    $sectionCount++
}
catch {
    Out-Log "[ERROR] Failed to get network adapter info: $_" -Color Red
    $failCount++
}

# ----------------------------------------
# 10. PC Serial Number
# ----------------------------------------
Start-Section -Title "PC Serial Number" -FileName "10_SerialNumber.txt"

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
# 11. Installed Software List (CSV Export)
# ----------------------------------------
Start-Section -Title "Installed Software List (CSV)" -FileName $null

try {
    # 11a. Desktop Apps (Registry)
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

    $outDesktop = Join-Path $targetDir "11_DesktopApps.csv"
    $desktop | Export-Csv -Path $outDesktop -NoTypeInformation -Encoding UTF8

    Out-Log "Desktop apps: $($desktop.Count) items -> 11_DesktopApps.csv"

    # 11b. Store / UWP Apps
    $store = Get-AppxPackage |
        Select-Object @{N='Name';E={$_.Name}},
                      @{N='Version';E={$_.Version}},
                      @{N='Publisher';E={$_.PublisherId}} |
        Sort-Object Name

    $outStore = Join-Path $targetDir "11_StoreApps.csv"
    $store | Export-Csv -Path $outStore -NoTypeInformation -Encoding UTF8

    Out-Log "Store apps: $($store.Count) items -> 11_StoreApps.csv"

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
# 12. Firewall Status (CSV Export)
# ----------------------------------------
Start-Section -Title "Firewall Status (CSV)" -FileName $null

try {
    # 12a. Firewall Profiles
    $fwProfiles = Get-NetFirewallProfile -ErrorAction Stop |
        Select-Object Name, Enabled, DefaultInboundAction, DefaultOutboundAction, LogFileName

    $outFwProfiles = Join-Path $targetDir "12_FirewallProfiles.csv"
    $fwProfiles | Export-Csv -Path $outFwProfiles -NoTypeInformation -Encoding UTF8

    Out-Log "Firewall profiles: $($fwProfiles.Count) profiles -> 12_FirewallProfiles.csv"

    # 12b. Firewall Rules
    $fwRules = Get-NetFirewallRule -ErrorAction Stop |
        Select-Object DisplayName, Enabled, Direction, Action, Profile |
        Sort-Object DisplayName

    $outFwRules = Join-Path $targetDir "12_FirewallRules.csv"
    $fwRules | Export-Csv -Path $outFwRules -NoTypeInformation -Encoding UTF8

    Out-Log "Firewall rules: $($fwRules.Count) rules -> 12_FirewallRules.csv"

    $sectionCount++
}
catch {
    Out-Log "[ERROR] Failed to get firewall info: $_" -Color Red
    $failCount++
}

# ----------------------------------------
# 13. Windows Optional Features (CSV Export)
# ----------------------------------------
Start-Section -Title "Windows Optional Features (CSV)" -FileName $null

try {
    $optFeatures = Get-WindowsOptionalFeature -Online -ErrorAction Stop |
        Select-Object FeatureName, State |
        Sort-Object FeatureName

    $outOptFeatures = Join-Path $targetDir "13_OptionalFeatures.csv"
    $optFeatures | Export-Csv -Path $outOptFeatures -NoTypeInformation -Encoding UTF8

    Out-Log "Optional features: $($optFeatures.Count) features -> 13_OptionalFeatures.csv"

    $sectionCount++
}
catch {
    Out-Log "[ERROR] Failed to get optional features: $_" -Color Red
    $failCount++
}

# ----------------------------------------
# 14. Server Roles & Features (CSV Export)
# ----------------------------------------
Start-Section -Title "Server Roles & Features (CSV)" -FileName $null

if ($isServer) {
    try {
        $serverFeatures = Get-WindowsFeature -ErrorAction Stop |
            Select-Object Name, DisplayName, InstallState, FeatureType |
            Sort-Object Name

        $outServerFeatures = Join-Path $targetDir "14_ServerRolesFeatures.csv"
        $serverFeatures | Export-Csv -Path $outServerFeatures -NoTypeInformation -Encoding UTF8

        Out-Log "Server roles & features: $($serverFeatures.Count) items -> 14_ServerRolesFeatures.csv"

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
# 15. Power Settings
# ----------------------------------------
Start-Section -Title "Power Settings" -FileName "15_PowerSettings.txt"

try {
    # 15a. Power plans (CIM - structured)
    $powerPlans = Get-CimInstance -Namespace root\cimv2\power -ClassName Win32_PowerPlan -ErrorAction Stop
    Out-Log "---- Power Plans ----"
    foreach ($plan in $powerPlans) {
        $active = if ($plan.IsActive) { " [ACTIVE]" } else { "" }
        Out-Log "  $($plan.ElementName)$active"
    }
    Out-Log ""

    # 15b. Active plan details (powercfg raw output)
    Out-Log "---- Active Plan Details (powercfg /query) ----"
    $powercfgOutput = powercfg /query 2>&1
    foreach ($line in $powercfgOutput) {
        Out-Log "  $line"
    }

    $sectionCount++
}
catch {
    Out-Log "[ERROR] Failed to get power settings: $_" -Color Red
    $failCount++
}

# ----------------------------------------
# 16. WiFi Profiles
# ----------------------------------------
Start-Section -Title "WiFi Profiles" -FileName "16_WiFiProfiles.txt"

try {
    $prevEncoding = [Console]::OutputEncoding
    try {
        [Console]::OutputEncoding = [System.Text.Encoding]::GetEncoding(932)
        $wlanOutput = netsh wlan show profiles 2>&1
    }
    finally {
        [Console]::OutputEncoding = $prevEncoding
    }
    foreach ($line in $wlanOutput) {
        Out-Log "  $line"
    }
    $sectionCount++
}
catch {
    Out-Log "[ERROR] Failed to get WiFi profiles: $_" -Color Red
    $failCount++
}

# ----------------------------------------
# 17. Restore Points (CSV Export)
# ----------------------------------------
Start-Section -Title "Restore Points (CSV)" -FileName $null

try {
    $restorePoints = Get-ComputerRestorePoint -ErrorAction Stop |
        Select-Object SequenceNumber, Description, RestorePointType,
            @{N='CreationTime';E={$_.ConvertToDateTime($_.CreationTime)}} |
        Sort-Object SequenceNumber

    $outRestorePoints = Join-Path $targetDir "17_RestorePoints.csv"
    $restorePoints | Export-Csv -Path $outRestorePoints -NoTypeInformation -Encoding UTF8
    Out-Log "Restore points: $($restorePoints.Count) point(s) -> 17_RestorePoints.csv"

    $sectionCount++
}
catch {
    Out-Log "[ERROR] Failed to get restore points: $_" -Color Red
    $failCount++
}

# ----------------------------------------
# 18. Windows Defender / Antivirus Status
# ----------------------------------------
Start-Section -Title "Windows Defender Status" -FileName "18_DefenderStatus.txt"

try {
    $mpStatus = Get-MpComputerStatus -ErrorAction Stop

    Out-Log "AMServiceEnabled:           $($mpStatus.AMServiceEnabled)"
    Out-Log "AntispywareEnabled:         $($mpStatus.AntispywareEnabled)"
    Out-Log "AntivirusEnabled:           $($mpStatus.AntivirusEnabled)"
    Out-Log "RealTimeProtectionEnabled:  $($mpStatus.RealTimeProtectionEnabled)"
    Out-Log "BehaviorMonitorEnabled:     $($mpStatus.BehaviorMonitorEnabled)"
    Out-Log "IoavProtectionEnabled:      $($mpStatus.IoavProtectionEnabled)"
    Out-Log "NISEnabled:                 $($mpStatus.NISEnabled)"
    Out-Log "OnAccessProtectionEnabled:  $($mpStatus.OnAccessProtectionEnabled)"
    Out-Log ""
    Out-Log "AMEngineVersion:            $($mpStatus.AMEngineVersion)"
    Out-Log "AMProductVersion:           $($mpStatus.AMProductVersion)"
    Out-Log "AntivirusSignatureVersion:  $($mpStatus.AntivirusSignatureVersion)"
    Out-Log "AntivirusSignatureLastUpdated: $($mpStatus.AntivirusSignatureLastUpdated)"
    Out-Log ""
    Out-Log "QuickScanEndTime:           $($mpStatus.QuickScanEndTime)"
    Out-Log "FullScanEndTime:            $($mpStatus.FullScanEndTime)"

    $sectionCount++
}
catch {
    Out-Log "[WARN] Could not get Defender status (may be replaced by 3rd-party AV): $_" -Color Yellow
    $sectionCount++
}

# ----------------------------------------
# 19. Windows Update History (CSV Export)
# ----------------------------------------
Start-Section -Title "Windows Update History (CSV)" -FileName $null

try {
    $hotfixes = Get-HotFix -ErrorAction Stop |
        Select-Object HotFixID, Description, InstalledBy,
            @{N='InstalledOn';E={$_.InstalledOn}} |
        Sort-Object InstalledOn -Descending

    $outHotfixes = Join-Path $targetDir "19_WindowsUpdates.csv"
    $hotfixes | Export-Csv -Path $outHotfixes -NoTypeInformation -Encoding UTF8
    Out-Log "Windows updates: $($hotfixes.Count) update(s) -> 19_WindowsUpdates.csv"

    $sectionCount++
}
catch {
    Out-Log "[ERROR] Failed to get Windows update history: $_" -Color Red
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
