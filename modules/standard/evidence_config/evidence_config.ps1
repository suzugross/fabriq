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
# Helper: Capture cscript output (locale-safe)
# ========================================
# cscript WScript.Echo output does NOT honor `chcp 65001` in parent cmd.
# On JP locales, the text is emitted as bytes in the OEM console codepage
# (CP932), so the cmd/chcp wrapper used for Win32 EXEs (e.g. netsh in
# Section 16) produces mojibake when captured as UTF-8 for a VBS host.
#
# The `cscript //U` flag is documented to emit UTF-16LE on redirected
# stdout but is unreliable in practice on modern Windows (often emits
# zero bytes). So instead:
#   1. Run cscript normally (no //U), letting it write OEM bytes
#   2. Tell .NET Process to decode stdout with the culture's OEM codepage
#      (so the bytes become correct Unicode in the .NET string)
#   3. Out-Log -> Out-File -Encoding UTF8 then writes valid UTF-8
# ----------------------------------------
function Invoke-CScriptCapture {
    param(
        [Parameter(Mandatory=$true)][string]$ScriptPath,
        [string[]]$ScriptArgs = @()
    )

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = 'cscript.exe'
    $argList = @('//Nologo', $ScriptPath) + $ScriptArgs
    $psi.Arguments = ($argList | ForEach-Object {
        if ($_ -match '\s') { '"' + $_ + '"' } else { $_ }
    }) -join ' '
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError  = $true
    $psi.UseShellExecute        = $false
    $psi.CreateNoWindow         = $true

    # Decode child stdout with the OEM codepage of the current culture
    # (932 on JP, 437/850 on EN, etc.). Falls back to [Encoding]::Default
    # (ANSI codepage) if GetEncoding throws for any reason.
    $oemCp = [System.Globalization.CultureInfo]::CurrentCulture.TextInfo.OEMCodePage
    $enc = try { [System.Text.Encoding]::GetEncoding($oemCp) } catch { [System.Text.Encoding]::Default }
    $psi.StandardOutputEncoding = $enc
    $psi.StandardErrorEncoding  = $enc

    $p = [System.Diagnostics.Process]::Start($psi)
    $outText = $p.StandardOutput.ReadToEnd()
    $errText = $p.StandardError.ReadToEnd()
    $p.WaitForExit()

    $combined = $outText
    if (-not [string]::IsNullOrWhiteSpace($errText)) {
        $combined += "`n[stderr]`n$errText"
    }
    return $combined
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
Write-Host "    [20] System TEMP Text-Log Backup (safety net)" -ForegroundColor White
Write-Host "    [21] Windows License / Activation Status" -ForegroundColor White
Write-Host "    [22] Office License / Activation Status" -ForegroundColor White
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
# 10. PC Serial Number (multi-source collection)
# ----------------------------------------
# Query every reasonable SMBIOS / registry source for the PC serial
# number independently, then pick a canonical value by priority.
# Every source is recorded (even rejected) so that post-facto audit
# can answer "which SMBIOS field held the SN on device X?".
#
# Why multi-source: OEMs populate SMBIOS Type 0 (BIOS) and Type 1
# (System) inconsistently. A 2400-unit rollout observed ~2 units
# where Win32_BIOS.SerialNumber was blank while
# Win32_ComputerSystemProduct.IdentifyingNumber held the real SN
# (and occasional inverse cases). Whitebox / VDI / un-burned units
# also frequently ship with the placeholder "Default string", which
# this section explicitly rejects.
# ----------------------------------------
Start-Section -Title "PC Serial Number" -FileName "10_SerialNumber.txt"

function Test-SerialValid {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return @{ Valid = $false; Reason = 'empty' } }
    $trimmed = $Value.Trim()

    $invalidExact = @(
        'None','N/A','INVALID',
        'To be filled by O.E.M.',
        'Default string',
        'System Serial Number','Chassis Serial Number',
        'Not Applicable','Not Specified','OEM'
    )
    foreach ($bad in $invalidExact) {
        if ($trimmed -ieq $bad) { return @{ Valid = $false; Reason = "`"$bad`"" } }
    }
    if ($trimmed -match '^0+$')        { return @{ Valid = $false; Reason = 'all-zero' } }
    if ($trimmed -match '^[\.\-\s]+$') { return @{ Valid = $false; Reason = 'dummy-chars-only' } }
    return @{ Valid = $true; Reason = $null }
}

function New-SerialSourceRow {
    param(
        [string]$Label,
        [bool]$IsCanonicalCandidate,
        [scriptblock]$Getter
    )
    $row = [PSCustomObject]@{
        Label                = $Label
        Value                = ''
        Tag                  = ''
        Valid                = $false
        IsCanonicalCandidate = $IsCanonicalCandidate
    }
    try {
        $raw  = & $Getter
        $text = if ($null -eq $raw) { '' } else { [string]$raw }
        $row.Value = $text
        $check     = Test-SerialValid -Value $text
        $row.Valid = $check.Valid
        $row.Tag   = if ($check.Valid) { 'VALID' } else { "INVALID: $($check.Reason)" }
    }
    catch {
        $reason = ($_.Exception.Message -replace '\s+', ' ').Trim()
        if ($reason.Length -gt 80) { $reason = $reason.Substring(0, 80) + '...' }
        $row.Tag = "QUERY FAILED: $reason"
    }
    return $row
}

try {
    # ----- Phase 1: collect every source independently -----
    $sources = @()

    $sources += New-SerialSourceRow -Label 'Win32_BIOS.SerialNumber' -IsCanonicalCandidate $true -Getter {
        (Get-CimInstance -ClassName Win32_BIOS -ErrorAction Stop).SerialNumber
    }
    $sources += New-SerialSourceRow -Label 'Win32_ComputerSystemProduct.IdentifyingNumber' -IsCanonicalCandidate $true -Getter {
        (Get-CimInstance -ClassName Win32_ComputerSystemProduct -ErrorAction Stop).IdentifyingNumber
    }
    $sources += New-SerialSourceRow -Label 'Win32_SystemEnclosure.SerialNumber' -IsCanonicalCandidate $true -Getter {
        $enc = Get-CimInstance -ClassName Win32_SystemEnclosure -ErrorAction Stop
        if ($enc -is [array]) {
            ($enc | ForEach-Object { $_.SerialNumber } |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                Select-Object -First 1)
        } else {
            $enc.SerialNumber
        }
    }
    $sources += New-SerialSourceRow -Label 'Win32_BaseBoard.SerialNumber' -IsCanonicalCandidate $false -Getter {
        (Get-CimInstance -ClassName Win32_BaseBoard -ErrorAction Stop).SerialNumber
    }
    $sources += New-SerialSourceRow -Label 'Registry SystemSerialNumber' -IsCanonicalCandidate $true -Getter {
        $regPath = 'HKLM:\HARDWARE\DESCRIPTION\System\BIOS'
        (Get-ItemProperty -Path $regPath -Name SystemSerialNumber -ErrorAction Stop).SystemSerialNumber
    }

    # Reference ID (UUID) - not a serial, captured for Default-string machines.
    $uuidRow = [PSCustomObject]@{
        Label = 'Win32_ComputerSystemProduct.UUID'
        Value = ''
        Tag   = ''
    }
    try {
        $uuidVal       = (Get-CimInstance -ClassName Win32_ComputerSystemProduct -ErrorAction Stop).UUID
        $uuidRow.Value = if ($null -eq $uuidVal) { '' } else { [string]$uuidVal }
        $uuidRow.Tag   = if ([string]::IsNullOrWhiteSpace($uuidRow.Value)) { 'EMPTY' } else { 'CAPTURED' }
    }
    catch {
        $reason = ($_.Exception.Message -replace '\s+', ' ').Trim()
        if ($reason.Length -gt 80) { $reason = $reason.Substring(0, 80) + '...' }
        $uuidRow.Tag = "QUERY FAILED: $reason"
    }

    # ----- Phase 2: pick canonical by priority -----
    $canonical       = $null
    $canonicalSource = $null
    foreach ($s in $sources) {
        if ($s.IsCanonicalCandidate -and $s.Valid) {
            $canonical       = $s.Value.Trim()
            $canonicalSource = $s.Label
            break
        }
    }

    # ----- Phase 3: emit evidence -----
    Out-Log "---- Canonical Serial Number ----"
    if ($null -eq $canonical) {
        Out-Log "(Unretrievable)" -Color Red
        Out-Log "(Source: none - all canonical candidates were invalid or failed)"
    } else {
        Out-Log $canonical
        Out-Log "(Source: $canonicalSource)"
    }
    Out-Log ""

    # Column widths for aligned table (include UUID row label).
    $allLabels  = @($sources | ForEach-Object { $_.Label }) + @($uuidRow.Label)
    $labelWidth = ($allLabels | ForEach-Object { $_.Length } | Measure-Object -Maximum).Maximum
    $allValues  = @($sources | ForEach-Object { if ([string]::IsNullOrEmpty($_.Value)) { '(empty)' } else { $_.Value.Trim() } }) +
                  @( if ([string]::IsNullOrEmpty($uuidRow.Value)) { '(empty)' } else { $uuidRow.Value.Trim() } )
    $valueWidth = [Math]::Max(20, (($allValues | ForEach-Object { $_.Length } | Measure-Object -Maximum).Maximum))

    Out-Log "---- All Sources ----"
    foreach ($s in $sources) {
        $displayValue = if ([string]::IsNullOrEmpty($s.Value)) { '(empty)' } else { $s.Value.Trim() }

        $tag = $s.Tag
        if ($s.Valid) {
            if ($null -ne $canonical -and $displayValue -eq $canonical) {
                $tag = 'VALID, MATCH'
            } elseif ($s.IsCanonicalCandidate) {
                $tag = 'VALID, DIFFERENT'
            } else {
                $tag = 'VALID, DIFFERENT (record-only)'
            }
        }

        $line = $s.Label.PadRight($labelWidth) + "  : " + $displayValue.PadRight($valueWidth) + "  [$tag]"
        Out-Log $line
    }
    Out-Log ""

    Out-Log "---- Reference ID ----"
    $refDisplay = if ([string]::IsNullOrEmpty($uuidRow.Value)) { '(empty)' } else { $uuidRow.Value.Trim() }
    $refLine    = $uuidRow.Label.PadRight($labelWidth) + "  : " + $refDisplay.PadRight($valueWidth) + "  [$($uuidRow.Tag)]"
    Out-Log $refLine
    Out-Log ""

    Out-Log "---- Selection Policy ----"
    Out-Log "Priority: Win32_BIOS.SerialNumber -> Win32_ComputerSystemProduct.IdentifyingNumber"
    Out-Log "       -> Win32_SystemEnclosure.SerialNumber -> Registry SystemSerialNumber"
    Out-Log "Win32_BaseBoard.SerialNumber is record-only (motherboard SN, not PC SN)."
    Out-Log "Win32_ComputerSystemProduct.UUID is reference ID only (not a serial number)."
    Out-Log 'Rejected values: empty / "Default string" / "To be filled by O.E.M." /'
    Out-Log '                 "None" / "N/A" / "INVALID" / "System Serial Number" /'
    Out-Log '                 "Chassis Serial Number" / "Not Applicable" / "Not Specified" /'
    Out-Log '                 "OEM" / all-zero / dots-hyphens-whitespace-only'

    if ($null -ne $canonical) {
        $sectionCount++
    } else {
        $failCount++
    }
}
catch {
    Out-Log "[ERROR] Failed to collect serial number sources: $_" -Color Red
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
    # Use cmd /c with chcp 65001 to get UTF-8 output from netsh
    # PowerShell 5.1 has issues decoding CP932 output when capturing to variable
    $prevEncoding = [Console]::OutputEncoding
    try {
        [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
        $wlanOutput = cmd /c "chcp 65001 >nul && netsh wlan show profiles" 2>&1
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
# 20. System TEMP Text-Log Backup (safety net)
# ----------------------------------------
# Top-level .log and .txt files from C:\Windows\Temp.
# Preserves raw forensic data (installer logs, ODT logs, driver logs)
# for post-incident investigation. Non-recursive, no size cap.
# Locked files are skipped silently since TEMP often holds files
# opened by active processes.
# ----------------------------------------
Start-Section -Title "System TEMP Text-Log Backup" -FileName "20_TempBackup.txt"

try {
    $tempSrc = "C:\Windows\Temp"

    if (-not (Test-Path $tempSrc)) {
        Out-Log "System TEMP not found: $tempSrc" -Color Yellow
        $sectionCount++
    }
    else {
        $backupDir = Join-Path $targetDir "20_TempBackup"
        $null = New-Item -ItemType Directory -Path $backupDir -Force -ErrorAction SilentlyContinue

        $files = @(Get-ChildItem -Path $tempSrc -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Extension -in '.log', '.txt' })

        $copied = 0
        foreach ($f in $files) {
            try {
                Copy-Item $f.FullName (Join-Path $backupDir $f.Name) -Force -ErrorAction Stop
                $copied++
            }
            catch {
                # Locked / access-denied file: skip silently. Common in TEMP
                # where installers and services may hold files open.
            }
        }

        Out-Log "Source:  $tempSrc"
        Out-Log "Matches: $($files.Count) .log/.txt file(s) at top level"
        Out-Log "Copied:  $copied file(s) -> 20_TempBackup\"

        $sectionCount++
    }
}
catch {
    Out-Log "[ERROR] Failed to backup System TEMP text-logs: $_" -Color Red
    $failCount++
}

# ----------------------------------------
# 21. Windows License / Activation Status
# ----------------------------------------
# Captures three perspectives on Windows activation:
#   21a. SoftwareLicensingProduct (CIM) - per-SKU licensing state
#   21b. SoftwareLicensingService (CIM) - machine-wide KMS configuration
#   21c. slmgr /dlv raw output - canonical human-readable diagnostic
# slmgr /dlv can take 5-30s because it queries SLS; this is acceptable for
# evidence collection (non-interactive snapshot).
# ----------------------------------------
Start-Section -Title "Windows License / Activation Status" -FileName "21_WindowsLicense.txt"

try {
    $licStatusMap = @{
        0 = "Unlicensed"
        1 = "Licensed"
        2 = "OOB Grace Period"
        3 = "Out of Tolerance Grace"
        4 = "Non-Genuine Grace"
        5 = "Notification Mode"
        6 = "Extended Grace Period"
    }

    # 21a. Per-product state (filter by PartialProductKey so only installed SKUs surface)
    Out-Log "---- SoftwareLicensingProduct (per-product) ----"
    $osProducts = @(Get-CimInstance -ClassName SoftwareLicensingProduct -ErrorAction SilentlyContinue |
                    Where-Object { $_.PartialProductKey -and $_.Name -like "*Windows*" })

    if ($osProducts.Count -eq 0) {
        Out-Log "(No installed Windows product key found)"
    }
    else {
        foreach ($p in $osProducts) {
            $statusCode = [int]$p.LicenseStatus
            $statusText = if ($licStatusMap.ContainsKey($statusCode)) { $licStatusMap[$statusCode] } else { "Unknown" }
            Out-Log "Name:                  $($p.Name)"
            Out-Log "  Description:         $($p.Description)"
            Out-Log "  LicenseStatus:       $statusText ($statusCode)"
            Out-Log "  PartialProductKey:   $($p.PartialProductKey)"
            Out-Log "  LicenseFamily:       $($p.LicenseFamily)"
            Out-Log "  ProductKeyChannel:   $($p.ProductKeyChannel)"
            if ($p.GracePeriodRemaining -gt 0) {
                Out-Log "  GracePeriodRemaining: $($p.GracePeriodRemaining) minutes"
            }
            if (-not [string]::IsNullOrEmpty($p.KeyManagementServiceMachine)) {
                Out-Log "  KMS (configured):    $($p.KeyManagementServiceMachine)"
            }
            if (-not [string]::IsNullOrEmpty($p.DiscoveredKeyManagementServiceMachineName)) {
                Out-Log "  KMS (discovered):    $($p.DiscoveredKeyManagementServiceMachineName):$($p.DiscoveredKeyManagementServiceMachinePort)"
            }
            Out-Log ""
        }
    }

    # 21b. Machine-wide licensing service (KMS client config, ClientMachineID)
    Out-Log "---- SoftwareLicensingService (machine-wide) ----"
    try {
        $sls = Get-CimInstance -ClassName SoftwareLicensingService -ErrorAction Stop
        Out-Log "ClientMachineID:             $($sls.ClientMachineID)"
        Out-Log "KeyManagementServiceMachine: $($sls.KeyManagementServiceMachine)"
        Out-Log "KeyManagementServicePort:    $($sls.KeyManagementServicePort)"
        Out-Log "OA3xOriginalProductKey:      $($sls.OA3xOriginalProductKey)"
        Out-Log "OA3xOriginalProductKeyDescription: $($sls.OA3xOriginalProductKeyDescription)"
        Out-Log "PolicyCacheRefreshRequired:  $($sls.PolicyCacheRefreshRequired)"
        Out-Log ""
    }
    catch {
        Out-Log "  [WARN] Could not query SoftwareLicensingService: $_" -Color Yellow
    }

    # 21c. slmgr /dlv raw (canonical diagnostic, matches what admins expect to see)
    # cscript WScript.Echo output ignores `chcp`, so use Invoke-CScriptCapture
    # helper (//U flag + UTF-16LE redirected stdout) for locale-safe capture.
    Out-Log "---- slmgr /dlv (raw) ----"
    try {
        $slmgrOutput = Invoke-CScriptCapture `
            -ScriptPath 'C:\Windows\System32\slmgr.vbs' `
            -ScriptArgs @('/dlv')
        foreach ($line in ($slmgrOutput -split "\r?\n")) {
            Out-Log "  $line"
        }
    }
    catch {
        Out-Log "  [WARN] slmgr /dlv failed: $_" -Color Yellow
    }

    $sectionCount++
}
catch {
    Out-Log "[ERROR] Failed to collect Windows license status: $_" -Color Red
    $failCount++
}

# ----------------------------------------
# 22. Office License / Activation Status
# ----------------------------------------
# Captures two perspectives on Office:
#   22a. Click-to-Run registry (ProductReleaseIds, channel, version)
#        - works even when OSPP.vbs is absent (Store Office / M365 Apps only)
#   22b. ospp.vbs /dstatus raw - per-product LICENSE NAME / STATUS / KMS info
# OSPP.vbs path detection mirrors office_license_config\office_license_auth.ps1
# so both modules stay in sync on Office install layout assumptions.
# ----------------------------------------
Start-Section -Title "Office License / Activation Status" -FileName "22_OfficeLicense.txt"

try {
    # 22a. Click-to-Run configuration registry
    Out-Log "---- Office Click-to-Run Configuration (registry) ----"
    $c2rKey = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
    if (Test-Path $c2rKey) {
        try {
            $c2r = Get-ItemProperty -Path $c2rKey -ErrorAction Stop
            Out-Log "ProductReleaseIds: $($c2r.ProductReleaseIds)"
            Out-Log "VersionToReport:   $($c2r.VersionToReport)"
            Out-Log "Platform:          $($c2r.Platform)"
            Out-Log "CDNBaseUrl:        $($c2r.CDNBaseUrl)"
            Out-Log "UpdateChannel:     $($c2r.UpdateChannel)"
            Out-Log "AudienceData:      $($c2r.AudienceData)"
            Out-Log "ClientCulture:     $($c2r.ClientCulture)"
            Out-Log ""
        }
        catch {
            Out-Log "  [WARN] Could not read C2R config: $_" -Color Yellow
        }
    }
    else {
        Out-Log "(No Click-to-Run configuration - MSI-based Office or Office not installed)"
        Out-Log ""
    }

    # 22b. OSPP.vbs path detection (same candidates as office_license_auth.ps1)
    $osppCandidates = @(
        "$env:ProgramFiles\Microsoft Office\root\Office16\OSPP.vbs"
        "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\OSPP.vbs"
        "$env:ProgramFiles\Microsoft Office\Office16\OSPP.vbs"
        "${env:ProgramFiles(x86)}\Microsoft Office\Office16\OSPP.vbs"
    )
    $osppPath = $null
    foreach ($cand in $osppCandidates) {
        if (Test-Path $cand) { $osppPath = $cand; break }
    }

    if ($null -eq $osppPath) {
        Out-Log "---- OSPP.vbs ----"
        Out-Log "(OSPP.vbs not found - Office not installed, or uses Microsoft 365 per-user licensing only)"
    }
    else {
        Out-Log "---- OSPP.vbs path ----"
        Out-Log "$osppPath"
        Out-Log ""
        Out-Log "---- cscript OSPP.vbs /dstatus (raw) ----"
        try {
            $osppOutput = Invoke-CScriptCapture `
                -ScriptPath $osppPath `
                -ScriptArgs @('/dstatus')
            foreach ($line in ($osppOutput -split "\r?\n")) {
                Out-Log "  $line"
            }
        }
        catch {
            Out-Log "  [WARN] ospp /dstatus failed: $_" -Color Yellow
        }
    }

    $sectionCount++
}
catch {
    Out-Log "[ERROR] Failed to collect Office license status: $_" -Color Red
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
