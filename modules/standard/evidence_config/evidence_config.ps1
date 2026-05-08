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
# Initializes per-section manifest tracking state (id, title, file list,
# stopwatch). The -Id parameter records the section number for the
# manifest.json output (kernel/EVIDENCE_MANIFEST.md schemaVersion=1).
function Start-Section {
    param(
        [string]$Id,
        [string]$Title,
        [string]$FileName
    )
    $script:CurrentSectionId = $Id
    $script:CurrentSectionTitle = $Title
    $script:CurrentSectionFiles = @()
    if (-not [string]::IsNullOrEmpty($FileName)) {
        $script:CurrentSectionFiles += $FileName
    }
    $script:CurrentSectionStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $script:currentSplitFile = $FileName
    Out-Log ""
    Out-Log "========================================" -Color Cyan
    Out-Log "$Title" -Color Cyan
    Out-Log "========================================" -Color Cyan
}

# ========================================
# Helper: Register additional file written by a section
# ========================================
# Sections that write files via Export-Csv (or other dynamic writes) call
# Add-SectionFile after each successful write so the manifest accurately
# records all output files.
function Add-SectionFile {
    param([string]$FileName)
    if ([string]::IsNullOrEmpty($FileName)) { return }
    if ($null -eq $script:CurrentSectionFiles) {
        $script:CurrentSectionFiles = @($FileName)
        return
    }
    if ($script:CurrentSectionFiles -notcontains $FileName) {
        $script:CurrentSectionFiles += $FileName
    }
}

# ========================================
# Helper: Close current section and append manifest entry
# ========================================
# Stops the section stopwatch and appends a Section object to
# $script:ManifestSections. Failed sections always emit empty files[]
# (potentially partial / corrupt files are not advertised to consumers).
function Close-Section {
    param(
        [Parameter(Mandatory=$true)]
        [ValidateSet('Success','Skipped','Failed','Partial')]
        [string]$Status,
        [string]$Reason = $null
    )
    if ($null -ne $script:CurrentSectionStopwatch) {
        $script:CurrentSectionStopwatch.Stop()
        $elapsed = [int]$script:CurrentSectionStopwatch.ElapsedMilliseconds
    } else {
        $elapsed = 0
    }
    $files = if ($Status -eq 'Failed') { @() } else { @($script:CurrentSectionFiles) }
    $script:ManifestSections += [PSCustomObject]@{
        id        = $script:CurrentSectionId
        title     = $script:CurrentSectionTitle
        files     = @($files)
        status    = $Status
        reason    = $Reason
        elapsedMs = $elapsed
    }
    $script:CurrentSectionId = $null
    $script:CurrentSectionTitle = $null
    $script:CurrentSectionFiles = @()
    $script:CurrentSectionStopwatch = $null
}

# ========================================
# Helper: Write evidence manifest.json
# ========================================
# Public contract documented in kernel/EVIDENCE_MANIFEST.md (schemaVersion=1).
# Called once at the end of collection. Rotates any existing manifest.json
# to manifest.json.bak (1-generation only) before writing the new one.
function Write-EvidenceManifest {
    param(
        [Parameter(Mandatory=$true)][string]$TargetDir,
        [Parameter(Mandatory=$true)][string]$EvidenceConfigVersion,
        [Parameter(Mandatory=$true)][string]$ComputerName,
        [Parameter(Mandatory=$true)][string]$HardwareUniqueId,
        [Parameter(Mandatory=$true)][string]$SelectedNewPcName,
        [Parameter(Mandatory=$true)][datetime]$CollectedAt
    )

    # Resolve kernel version (best-effort — manifest is informational)
    $kernelVersionFile = Join-Path $PSScriptRoot "..\..\..\kernel\KERNEL_VERSION"
    $kernelVersion = if (Test-Path $kernelVersionFile) {
        (Get-Content $kernelVersionFile -Raw).Trim()
    } else {
        "unknown"
    }

    $workerName = if ([string]::IsNullOrWhiteSpace($env:FABRIQ_WORKER_NAME)) {
        $null
    } else {
        $env:FABRIQ_WORKER_NAME
    }

    $sections = @($script:ManifestSections)
    $successCount = @($sections | Where-Object { $_.status -eq 'Success' }).Count
    $skippedCount = @($sections | Where-Object { $_.status -eq 'Skipped' }).Count
    $failedCount  = @($sections | Where-Object { $_.status -eq 'Failed'  }).Count
    $partialCount = @($sections | Where-Object { $_.status -eq 'Partial' }).Count

    $manifest = [ordered]@{
        schemaVersion         = 1
        manifestType          = "fabriq-evidence-manifest"
        evidenceConfigVersion = $EvidenceConfigVersion
        fabriqKernelVersion   = $kernelVersion
        collectedAt           = $CollectedAt.ToString("yyyy-MM-ddTHH:mm:sszzz")
        computerName          = $ComputerName
        hardwareUniqueId      = $HardwareUniqueId
        selectedNewPcName     = $SelectedNewPcName
        workerName            = $workerName
        sections              = $sections
        summary               = [ordered]@{
            sectionCount = $sections.Count
            successCount = $successCount
            skippedCount = $skippedCount
            failedCount  = $failedCount
            partialCount = $partialCount
        }
    }

    $manifestPath = Join-Path $TargetDir "manifest.json"
    $bakPath      = Join-Path $TargetDir "manifest.json.bak"

    # Rotate previous manifest (1-generation; .bak.bak is not preserved)
    if (Test-Path $manifestPath) {
        if (Test-Path $bakPath) {
            Remove-Item $bakPath -Force -ErrorAction SilentlyContinue
        }
        Move-Item $manifestPath $bakPath -Force -ErrorAction SilentlyContinue
    }

    $json = $manifest | ConvertTo-Json -Depth 6
    $json | Out-File -FilePath $manifestPath -Encoding UTF8 -Force

    return $manifestPath
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
# Helper: Section enable/disable lookup
# ========================================
# Loaded once at startup from evidence_list.csv (Id, Title, Enabled).
# Default-on policy: missing CSV / missing row / non-0 value -> enabled.
# This preserves existing behavior when the CSV is absent and ensures
# that adding a new section in code without updating the CSV does not
# accidentally disable it.
function Test-SectionEnabled {
    param([Parameter(Mandatory=$true)][string]$Id)
    if ($null -eq $script:sectionEnabled) { return $true }
    if (-not $script:sectionEnabled.ContainsKey($Id)) { return $true }
    return [bool]$script:sectionEnabled[$Id]
}

# ========================================
# Helper: Emit Skipped manifest entry for a disabled section
# ========================================
# Used in place of Start-Section/Close-Section when a section is disabled
# via evidence_list.csv. Writes one gray master-log line and appends a
# Skipped manifest entry with files=[] and a stable reason prefix so that
# external consumers can distinguish user-disabled skips from intrinsic
# skips (Server-only on client OS, no-battery, Defender absent, etc.).
function Write-DisabledSection {
    param(
        [Parameter(Mandatory=$true)][string]$Id,
        [Parameter(Mandatory=$true)][string]$Title
    )
    # Reset split-file routing before logging the skip notice. Close-Section
    # does NOT reset $script:currentSplitFile (only Start-Section does), so
    # without this reset the gray skip line would bleed into the previous
    # section's split log file (e.g. a "Section 02 disabled" notice ending
    # up appended to 01_SystemInfo.txt).
    $script:currentSplitFile = $null
    Out-Log ""
    Out-Log "[Section $Id] $Title : Skipped (disabled by configuration)" -Color DarkGray
    $script:ManifestSections += [PSCustomObject]@{
        id        = $Id
        title     = $Title
        files     = @()
        status    = 'Skipped'
        reason    = 'Disabled by configuration (evidence_list.csv)'
        elapsedMs = 0
    }
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
Write-Host "    [22] Office License / Activation Status (C2R + OSPP + vNext)" -ForegroundColor White
Write-Host "    [23] Security Baseline (TPM / Secure Boot / VBS / LSA / BIOS)" -ForegroundColor White
Write-Host "    [24] Group Policy Report (gpresult /h HTML)" -ForegroundColor White
Write-Host "    [25] Certificates (4 stores in single CSV)" -ForegroundColor White
Write-Host "    [26] Battery Report (laptop only)" -ForegroundColor White
Write-Host "    [27] Environment Variables (Machine + User scopes, CSV)" -ForegroundColor White
Write-Host "    [28] Startup Items (Win32_StartupCommand + logon ScheduledTask, CSV)" -ForegroundColor White
Write-Host "    [29] Memory Slots & Array Summary (CSV)" -ForegroundColor White
Write-Host "    [30] PnP Devices (full enumeration with driver version/date, CSV)" -ForegroundColor White
Write-Host "    [31] Hardware Identifiers (System / BaseBoard / Enclosure, TXT)" -ForegroundColor White
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

# Manifest state (kernel/EVIDENCE_MANIFEST.md schemaVersion=1)
$script:ManifestSections = @()
$script:CollectedAt = Get-Date

# Load section enable/disable map from evidence_list.csv (default-on policy:
# absent CSV / missing row / non-0 Enabled -> enabled, preserving prior
# behavior when the CSV is not provided).
$script:sectionEnabled = @{}
$_evListPath = Join-Path $PSScriptRoot 'evidence_list.csv'
if (Test-Path $_evListPath) {
    try {
        $_evRows = Import-ModuleCsv -Path $_evListPath
        if ($_evRows) {
            foreach ($_evRow in $_evRows) {
                $_evId = "$($_evRow.Id)".Trim()
                if ([string]::IsNullOrEmpty($_evId)) { continue }
                $script:sectionEnabled[$_evId] = ("$($_evRow.Enabled)".Trim() -ne '0')
            }
        }
    }
    catch {
        Out-Log "[WARN] Failed to load evidence_list.csv (all sections will run): $_" -Color Yellow
        $script:sectionEnabled = @{}
    }
}

$now = Get-Date -Format "yyyy/MM/dd HH:mm:ss.ff"
$currentSplitFile = $null
Out-Log "==== Evidence Log ====" -Color Cyan
Out-Log "Date: $now"
Out-Log "Computer: $pcName"
Out-Log "Save Location: $targetDir"

# ----------------------------------------
# 1. Basic Info (Hostname / OS / Specs)
# ----------------------------------------
if (Test-SectionEnabled "01") {
Start-Section -Id "01" -Title "System Basic Info" -FileName "01_SystemInfo.txt"
$sectionStatus = 'Success'
$sectionReason = $null

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
    $sectionStatus = 'Failed'
    $sectionReason = "$($_.Exception.Message)"
}
Close-Section -Status $sectionStatus -Reason $sectionReason
} else {
    Write-DisabledSection -Id "01" -Title "System Basic Info"
}

# ----------------------------------------
# 2. Local Users (CSV Export)
# ----------------------------------------
if (Test-SectionEnabled "02") {
Start-Section -Id "02" -Title "Local Users (CSV)" -FileName $null
$sectionStatus = 'Success'
$sectionReason = $null

try {
    $localUsers = Get-LocalUser | Select-Object `
        Name, Enabled, FullName, Description, SID,
        LastLogon, PasswordLastSet, PasswordRequired,
        PasswordExpires, AccountExpires, PrincipalSource |
        Sort-Object Name

    $outLocalUsers = Join-Path $targetDir "02_LocalUsers.csv"
    $localUsers | Export-Csv -Path $outLocalUsers -NoTypeInformation -Encoding UTF8
    Add-SectionFile "02_LocalUsers.csv"

    Out-Log "Local users: $($localUsers.Count) accounts -> 02_LocalUsers.csv"
    $sectionCount++
}
catch {
    Out-Log "[ERROR] Failed to get local users: $_" -Color Red
    $failCount++
    $sectionStatus = 'Failed'
    $sectionReason = "$($_.Exception.Message)"
}
Close-Section -Status $sectionStatus -Reason $sectionReason
} else {
    Write-DisabledSection -Id "02" -Title "Local Users (CSV)"
}

# ----------------------------------------
# 3. Local Groups (CSV Export)
# ----------------------------------------
if (Test-SectionEnabled "03") {
Start-Section -Id "03" -Title "Local Groups (CSV)" -FileName $null
$sectionStatus = 'Success'
$sectionReason = $null

try {
    $localGroups = Get-LocalGroup | Select-Object Name, Description, SID |
        Sort-Object Name

    $outLocalGroups = Join-Path $targetDir "03_LocalGroups.csv"
    $localGroups | Export-Csv -Path $outLocalGroups -NoTypeInformation -Encoding UTF8
    Add-SectionFile "03_LocalGroups.csv"

    Out-Log "Local groups: $($localGroups.Count) groups -> 03_LocalGroups.csv"
    $sectionCount++
}
catch {
    Out-Log "[ERROR] Failed to get local groups: $_" -Color Red
    $failCount++
    $sectionStatus = 'Failed'
    $sectionReason = "$($_.Exception.Message)"
}
Close-Section -Status $sectionStatus -Reason $sectionReason
} else {
    Write-DisabledSection -Id "03" -Title "Local Groups (CSV)"
}

# ----------------------------------------
# 4. Local Group Members (CSV Export)
# ----------------------------------------
if (Test-SectionEnabled "04") {
Start-Section -Id "04" -Title "Local Group Members (CSV)" -FileName $null
$sectionStatus = 'Success'
$sectionReason = $null

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
    Add-SectionFile "04_LocalGroupMembers.csv"

    Out-Log "Group memberships: $($allMembers.Count) entries -> 04_LocalGroupMembers.csv"
    $sectionCount++
}
catch {
    Out-Log "[ERROR] Failed to get group members: $_" -Color Red
    $failCount++
    $sectionStatus = 'Failed'
    $sectionReason = "$($_.Exception.Message)"
}
Close-Section -Status $sectionStatus -Reason $sectionReason
} else {
    Write-DisabledSection -Id "04" -Title "Local Group Members (CSV)"
}

# ----------------------------------------
# 5. Domain / Azure AD Status
# ----------------------------------------
if (Test-SectionEnabled "05") {
Start-Section -Id "05" -Title "Domain / Azure AD Status" -FileName "05_DomainStatus.txt"
$sectionStatus = 'Success'
$sectionReason = $null

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
        Add-SectionFile "05_UserProfiles.csv"
        Out-Log "User profiles: $($profiles.Count) profiles -> 05_UserProfiles.csv"
    }
    catch {
        Out-Log "  [WARN] Could not retrieve user profiles: $_" -Color Yellow
        $sectionStatus = 'Partial'
        $sectionReason = "User profile sub-collection failed: $($_.Exception.Message)"
    }

    $sectionCount++
}
catch {
    Out-Log "[ERROR] Failed to get domain/Azure AD status: $_" -Color Red
    $failCount++
    $sectionStatus = 'Failed'
    $sectionReason = "$($_.Exception.Message)"
}
Close-Section -Status $sectionStatus -Reason $sectionReason
} else {
    Write-DisabledSection -Id "05" -Title "Domain / Azure AD Status"
}

# ----------------------------------------
# 6. IP / DNS Settings (CSV Export)
# ----------------------------------------
if (Test-SectionEnabled "06") {
Start-Section -Id "06" -Title "Network Settings (CSV)" -FileName $null
$sectionStatus = 'Success'
$sectionReason = $null

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
    Add-SectionFile "06_NetworkConfig.csv"

    Out-Log "Network interfaces: $($networkRows.Count) entries -> 06_NetworkConfig.csv"
    $sectionCount++
}
catch {
    Out-Log "[ERROR] Failed to get IP settings: $_" -Color Red
    $failCount++
    $sectionStatus = 'Failed'
    $sectionReason = "$($_.Exception.Message)"
}
Close-Section -Status $sectionStatus -Reason $sectionReason
} else {
    Write-DisabledSection -Id "06" -Title "Network Settings (CSV)"
}

# ----------------------------------------
# 7. Printers / Ports List (CSV Export)
# ----------------------------------------
if (Test-SectionEnabled "07") {
Start-Section -Id "07" -Title "Printers / Ports List (CSV)" -FileName $null
$sectionStatus = 'Success'
$sectionReason = $null

try {
    $printers = Get-Printer -ErrorAction SilentlyContinue
    if ($printers) {
        $printerRows = $printers | Select-Object Name, DriverName, PortName, Shared, PrinterStatus |
            Sort-Object Name

        $outPrinters = Join-Path $targetDir "07_Printers.csv"
        $printerRows | Export-Csv -Path $outPrinters -NoTypeInformation -Encoding UTF8
        Add-SectionFile "07_Printers.csv"

        Out-Log "Printers: $($printerRows.Count) entries -> 07_Printers.csv"
    } else {
        Out-Log "(No printers installed)"
        $sectionStatus = 'Skipped'
        $sectionReason = 'No printers installed'
    }
    $sectionCount++
}
catch {
    Out-Log "[ERROR] Failed to get printer info: $_" -Color Red
    $failCount++
    $sectionStatus = 'Failed'
    $sectionReason = "$($_.Exception.Message)"
}
Close-Section -Status $sectionStatus -Reason $sectionReason
} else {
    Write-DisabledSection -Id "07" -Title "Printers / Ports List (CSV)"
}

# ----------------------------------------
# 8. BitLocker Status
# ----------------------------------------
if (Test-SectionEnabled "08") {
Start-Section -Id "08" -Title "BitLocker Status" -FileName "08_BitLocker.txt"
$sectionStatus = 'Success'
$sectionReason = $null

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
    $sectionStatus = 'Failed'
    $sectionReason = "$($_.Exception.Message)"
}
Close-Section -Status $sectionStatus -Reason $sectionReason
} else {
    Write-DisabledSection -Id "08" -Title "BitLocker Status"
}

# ----------------------------------------
# 8b. Disk & Partition Info (CSV Export)
# ----------------------------------------
if (Test-SectionEnabled "8b") {
Start-Section -Id "8b" -Title "Disk & Partition Info (CSV)" -FileName $null
$sectionStatus = 'Success'
$sectionReason = $null

try {
    # Physical disks
    $disks = Get-Disk | Select-Object Number, FriendlyName, SerialNumber,
        @{N='SizeGB';E={[Math]::Round($_.Size / 1GB, 2)}},
        PartitionStyle, HealthStatus, OperationalStatus |
        Sort-Object Number

    $outDisks = Join-Path $targetDir "08b_Disks.csv"
    $disks | Export-Csv -Path $outDisks -NoTypeInformation -Encoding UTF8
    Add-SectionFile "08b_Disks.csv"
    Out-Log "Physical disks: $($disks.Count) disk(s) -> 08b_Disks.csv"

    # Partitions
    $partitions = Get-Partition | Select-Object DiskNumber, PartitionNumber, DriveLetter,
        @{N='SizeGB';E={[Math]::Round($_.Size / 1GB, 2)}},
        Type, IsSystem, IsBoot, IsActive |
        Sort-Object DiskNumber, PartitionNumber

    $outPartitions = Join-Path $targetDir "08b_Partitions.csv"
    $partitions | Export-Csv -Path $outPartitions -NoTypeInformation -Encoding UTF8
    Add-SectionFile "08b_Partitions.csv"
    Out-Log "Partitions: $($partitions.Count) partition(s) -> 08b_Partitions.csv"

    $sectionCount++
}
catch {
    Out-Log "[ERROR] Failed to get disk/partition info: $_" -Color Red
    $failCount++
    $sectionStatus = 'Failed'
    $sectionReason = "$($_.Exception.Message)"
}
Close-Section -Status $sectionStatus -Reason $sectionReason
} else {
    Write-DisabledSection -Id "8b" -Title "Disk & Partition Info (CSV)"
}

# ----------------------------------------
# 9. MAC Address List (CSV Export)
# ----------------------------------------
if (Test-SectionEnabled "09") {
Start-Section -Id "09" -Title "MAC Address List (CSV)" -FileName $null
$sectionStatus = 'Success'
$sectionReason = $null

try {
    $adapters = Get-NetAdapter | Select-Object Name, InterfaceDescription, MacAddress, Status |
        Sort-Object Name

    $outMac = Join-Path $targetDir "09_MacAddress.csv"
    $adapters | Export-Csv -Path $outMac -NoTypeInformation -Encoding UTF8
    Add-SectionFile "09_MacAddress.csv"

    Out-Log "Network adapters: $($adapters.Count) entries -> 09_MacAddress.csv"
    $sectionCount++
}
catch {
    Out-Log "[ERROR] Failed to get network adapter info: $_" -Color Red
    $failCount++
    $sectionStatus = 'Failed'
    $sectionReason = "$($_.Exception.Message)"
}
Close-Section -Status $sectionStatus -Reason $sectionReason
} else {
    Write-DisabledSection -Id "09" -Title "MAC Address List (CSV)"
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
if (Test-SectionEnabled "10") {
Start-Section -Id "10" -Title "PC Serial Number" -FileName "10_SerialNumber.txt"
$sectionStatus = 'Success'
$sectionReason = $null

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
        $sectionStatus = 'Failed'
        $sectionReason = 'No valid serial number source found (all canonical candidates rejected)'
    }
}
catch {
    Out-Log "[ERROR] Failed to collect serial number sources: $_" -Color Red
    $failCount++
    $sectionStatus = 'Failed'
    $sectionReason = "$($_.Exception.Message)"
}
Close-Section -Status $sectionStatus -Reason $sectionReason
} else {
    Write-DisabledSection -Id "10" -Title "PC Serial Number"
}

# ----------------------------------------
# 11. Installed Software List (CSV Export)
# ----------------------------------------
if (Test-SectionEnabled "11") {
Start-Section -Id "11" -Title "Installed Software List (CSV)" -FileName $null
$sectionStatus = 'Success'
$sectionReason = $null

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
    Add-SectionFile "11_DesktopApps.csv"

    Out-Log "Desktop apps: $($desktop.Count) items -> 11_DesktopApps.csv"

    # 11b. Store / UWP Apps
    $store = Get-AppxPackage |
        Select-Object @{N='Name';E={$_.Name}},
                      @{N='Version';E={$_.Version}},
                      @{N='Publisher';E={$_.PublisherId}} |
        Sort-Object Name

    $outStore = Join-Path $targetDir "11_StoreApps.csv"
    $store | Export-Csv -Path $outStore -NoTypeInformation -Encoding UTF8
    Add-SectionFile "11_StoreApps.csv"

    Out-Log "Store apps: $($store.Count) items -> 11_StoreApps.csv"

    $sectionCount++
}
catch {
    Out-Log "[ERROR] Failed to get software list: $_" -Color Red
    $failCount++
    $sectionStatus = 'Failed'
    $sectionReason = "$($_.Exception.Message)"
}
Close-Section -Status $sectionStatus -Reason $sectionReason
} else {
    Write-DisabledSection -Id "11" -Title "Installed Software List (CSV)"
}

# OS type detection for server-only sections
$osProductType = (Get-CimInstance Win32_OperatingSystem).ProductType
$isServer = ($osProductType -ne 1)

# ----------------------------------------
# 12. Firewall Status (CSV Export)
# ----------------------------------------
if (Test-SectionEnabled "12") {
Start-Section -Id "12" -Title "Firewall Status (CSV)" -FileName $null
$sectionStatus = 'Success'
$sectionReason = $null

try {
    # 12a. Firewall Profiles
    $fwProfiles = Get-NetFirewallProfile -ErrorAction Stop |
        Select-Object Name, Enabled, DefaultInboundAction, DefaultOutboundAction, LogFileName

    $outFwProfiles = Join-Path $targetDir "12_FirewallProfiles.csv"
    $fwProfiles | Export-Csv -Path $outFwProfiles -NoTypeInformation -Encoding UTF8
    Add-SectionFile "12_FirewallProfiles.csv"

    Out-Log "Firewall profiles: $($fwProfiles.Count) profiles -> 12_FirewallProfiles.csv"

    # 12b. Firewall Rules
    $fwRules = Get-NetFirewallRule -ErrorAction Stop |
        Select-Object DisplayName, Enabled, Direction, Action, Profile |
        Sort-Object DisplayName

    $outFwRules = Join-Path $targetDir "12_FirewallRules.csv"
    $fwRules | Export-Csv -Path $outFwRules -NoTypeInformation -Encoding UTF8
    Add-SectionFile "12_FirewallRules.csv"

    Out-Log "Firewall rules: $($fwRules.Count) rules -> 12_FirewallRules.csv"

    $sectionCount++
}
catch {
    Out-Log "[ERROR] Failed to get firewall info: $_" -Color Red
    $failCount++
    $sectionStatus = 'Failed'
    $sectionReason = "$($_.Exception.Message)"
}
Close-Section -Status $sectionStatus -Reason $sectionReason
} else {
    Write-DisabledSection -Id "12" -Title "Firewall Status (CSV)"
}

# ----------------------------------------
# 13. Windows Optional Features (CSV Export)
# ----------------------------------------
if (Test-SectionEnabled "13") {
Start-Section -Id "13" -Title "Windows Optional Features (CSV)" -FileName $null
$sectionStatus = 'Success'
$sectionReason = $null

try {
    $optFeatures = Get-WindowsOptionalFeature -Online -ErrorAction Stop |
        Select-Object FeatureName, State |
        Sort-Object FeatureName

    $outOptFeatures = Join-Path $targetDir "13_OptionalFeatures.csv"
    $optFeatures | Export-Csv -Path $outOptFeatures -NoTypeInformation -Encoding UTF8
    Add-SectionFile "13_OptionalFeatures.csv"

    Out-Log "Optional features: $($optFeatures.Count) features -> 13_OptionalFeatures.csv"

    $sectionCount++
}
catch {
    Out-Log "[ERROR] Failed to get optional features: $_" -Color Red
    $failCount++
    $sectionStatus = 'Failed'
    $sectionReason = "$($_.Exception.Message)"
}
Close-Section -Status $sectionStatus -Reason $sectionReason
} else {
    Write-DisabledSection -Id "13" -Title "Windows Optional Features (CSV)"
}

# ----------------------------------------
# 14. Server Roles & Features (CSV Export)
# ----------------------------------------
if (Test-SectionEnabled "14") {
Start-Section -Id "14" -Title "Server Roles & Features (CSV)" -FileName $null
$sectionStatus = 'Success'
$sectionReason = $null

if ($isServer) {
    try {
        $serverFeatures = Get-WindowsFeature -ErrorAction Stop |
            Select-Object Name, DisplayName, InstallState, FeatureType |
            Sort-Object Name

        $outServerFeatures = Join-Path $targetDir "14_ServerRolesFeatures.csv"
        $serverFeatures | Export-Csv -Path $outServerFeatures -NoTypeInformation -Encoding UTF8
        Add-SectionFile "14_ServerRolesFeatures.csv"

        Out-Log "Server roles & features: $($serverFeatures.Count) items -> 14_ServerRolesFeatures.csv"

        $sectionCount++
    }
    catch {
        Out-Log "[ERROR] Failed to get server features: $_" -Color Red
        $failCount++
        $sectionStatus = 'Failed'
        $sectionReason = "$($_.Exception.Message)"
    }
}
else {
    Out-Log "Skipped: Client OS detected (Server-only section)"
    $sectionCount++
    $sectionStatus = 'Skipped'
    $sectionReason = 'Client OS detected (Server-only section)'
}
Close-Section -Status $sectionStatus -Reason $sectionReason
} else {
    Write-DisabledSection -Id "14" -Title "Server Roles & Features (CSV)"
}

# ----------------------------------------
# 15. Power Settings
# ----------------------------------------
if (Test-SectionEnabled "15") {
Start-Section -Id "15" -Title "Power Settings" -FileName "15_PowerSettings.txt"
$sectionStatus = 'Success'
$sectionReason = $null

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
    $sectionStatus = 'Failed'
    $sectionReason = "$($_.Exception.Message)"
}
Close-Section -Status $sectionStatus -Reason $sectionReason
} else {
    Write-DisabledSection -Id "15" -Title "Power Settings"
}

# ----------------------------------------
# 16. WiFi Profiles
# ----------------------------------------
if (Test-SectionEnabled "16") {
Start-Section -Id "16" -Title "WiFi Profiles" -FileName "16_WiFiProfiles.txt"
$sectionStatus = 'Success'
$sectionReason = $null

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
    $sectionStatus = 'Failed'
    $sectionReason = "$($_.Exception.Message)"
}
Close-Section -Status $sectionStatus -Reason $sectionReason
} else {
    Write-DisabledSection -Id "16" -Title "WiFi Profiles"
}

# ----------------------------------------
# 17. Restore Points (CSV Export)
# ----------------------------------------
if (Test-SectionEnabled "17") {
Start-Section -Id "17" -Title "Restore Points (CSV)" -FileName $null
$sectionStatus = 'Success'
$sectionReason = $null

try {
    $restorePoints = Get-ComputerRestorePoint -ErrorAction Stop |
        Select-Object SequenceNumber, Description, RestorePointType,
            @{N='CreationTime';E={$_.ConvertToDateTime($_.CreationTime)}} |
        Sort-Object SequenceNumber

    $outRestorePoints = Join-Path $targetDir "17_RestorePoints.csv"
    $restorePoints | Export-Csv -Path $outRestorePoints -NoTypeInformation -Encoding UTF8
    Add-SectionFile "17_RestorePoints.csv"
    Out-Log "Restore points: $($restorePoints.Count) point(s) -> 17_RestorePoints.csv"

    $sectionCount++
}
catch {
    Out-Log "[ERROR] Failed to get restore points: $_" -Color Red
    $failCount++
    $sectionStatus = 'Failed'
    $sectionReason = "$($_.Exception.Message)"
}
Close-Section -Status $sectionStatus -Reason $sectionReason
} else {
    Write-DisabledSection -Id "17" -Title "Restore Points (CSV)"
}

# ----------------------------------------
# 18. Windows Defender / Antivirus Status
# ----------------------------------------
if (Test-SectionEnabled "18") {
Start-Section -Id "18" -Title "Windows Defender Status" -FileName "18_DefenderStatus.txt"
$sectionStatus = 'Success'
$sectionReason = $null

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
    $sectionStatus = 'Skipped'
    $sectionReason = "Defender unavailable (may be replaced by 3rd-party AV): $($_.Exception.Message)"
}
Close-Section -Status $sectionStatus -Reason $sectionReason
} else {
    Write-DisabledSection -Id "18" -Title "Windows Defender Status"
}

# ----------------------------------------
# 19. Windows Update History (CSV Export)
# ----------------------------------------
if (Test-SectionEnabled "19") {
Start-Section -Id "19" -Title "Windows Update History (CSV)" -FileName $null
$sectionStatus = 'Success'
$sectionReason = $null

try {
    $hotfixes = Get-HotFix -ErrorAction Stop |
        Select-Object HotFixID, Description, InstalledBy,
            @{N='InstalledOn';E={$_.InstalledOn}} |
        Sort-Object InstalledOn -Descending

    $outHotfixes = Join-Path $targetDir "19_WindowsUpdates.csv"
    $hotfixes | Export-Csv -Path $outHotfixes -NoTypeInformation -Encoding UTF8
    Add-SectionFile "19_WindowsUpdates.csv"
    Out-Log "Windows updates: $($hotfixes.Count) update(s) -> 19_WindowsUpdates.csv"

    $sectionCount++
}
catch {
    Out-Log "[ERROR] Failed to get Windows update history: $_" -Color Red
    $failCount++
    $sectionStatus = 'Failed'
    $sectionReason = "$($_.Exception.Message)"
}
Close-Section -Status $sectionStatus -Reason $sectionReason
} else {
    Write-DisabledSection -Id "19" -Title "Windows Update History (CSV)"
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
if (Test-SectionEnabled "20") {
Start-Section -Id "20" -Title "System TEMP Text-Log Backup" -FileName "20_TempBackup.txt"
$sectionStatus = 'Success'
$sectionReason = $null

try {
    $tempSrc = "C:\Windows\Temp"

    if (-not (Test-Path $tempSrc)) {
        Out-Log "System TEMP not found: $tempSrc" -Color Yellow
        $sectionCount++
        $sectionStatus = 'Skipped'
        $sectionReason = "System TEMP directory not found: $tempSrc"
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

        # Register the dump directory (trailing slash = directory per
        # EVIDENCE_MANIFEST.md §6)
        if ($copied -gt 0) {
            Add-SectionFile "20_TempBackup/"
        }

        $sectionCount++
    }
}
catch {
    Out-Log "[ERROR] Failed to backup System TEMP text-logs: $_" -Color Red
    $failCount++
    $sectionStatus = 'Failed'
    $sectionReason = "$($_.Exception.Message)"
}
Close-Section -Status $sectionStatus -Reason $sectionReason
} else {
    Write-DisabledSection -Id "20" -Title "System TEMP Text-Log Backup"
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
if (Test-SectionEnabled "21") {
Start-Section -Id "21" -Title "Windows License / Activation Status" -FileName "21_WindowsLicense.txt"
$sectionStatus = 'Success'
$sectionReason = $null

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
    $sectionStatus = 'Failed'
    $sectionReason = "$($_.Exception.Message)"
}
Close-Section -Status $sectionStatus -Reason $sectionReason
} else {
    Write-DisabledSection -Id "21" -Title "Windows License / Activation Status"
}

# ----------------------------------------
# 22. Office License / Activation Status
# ----------------------------------------
# Office has TWO parallel licensing mechanisms which fabriq must observe
# both to produce an accurate audit verdict:
#
#   - OSPP (legacy)  : Volume License / buy-once Retail keys. Tracked by
#                      Software Protection Service. Surface: OSPP.vbs.
#   - vNext (current): Microsoft 365 subscriptions. Per-user license files
#                      under %LOCALAPPDATA%\Microsoft\Office\Licenses\
#                      <Category>\<NumericFilename>. UTF-16LE JSON wrapping
#                      a Base64-encoded inner license JSON, signed by the
#                      Office Licensing Service.
#
# For M365 subscriptions, the installer drops a decoy Retail key on OSPP
# which goes to Grace state and never refreshes once the user signs in.
# This produces a misleading "NOTIFICATIONS / 0xC004F009 (Grace expired)"
# in OSPP /dstatus. The actual license is the vNext subscription token.
#
# Sub-sections:
#   22a. Click-to-Run registry  (existing) - product / channel / version
#   22b. OSPP.vbs /dstatus raw  (existing) - VL / buy-once authoritative
#   22c. vNext per-user license scan       - subscription authoritative
#   22d. Interpretation                    - cross-source verdict
#
# Section status logic:
#   - subscription detected + Provisioned vNext present -> Success
#   - subscription detected + no Provisioned vNext      -> Partial
#     (typical at kitting time when end-user has not signed in)
#   - VL/buy-once + OSPP Grace/Notifications            -> Failed
#   - VL/buy-once + OSPP Licensed                       -> Success
#   - Office not installed                              -> Success (text only)
# ----------------------------------------
if (Test-SectionEnabled "22") {
Start-Section -Id "22" -Title "Office License / Activation Status" -FileName "22_OfficeLicense.txt"
$sectionStatus = 'Success'
$sectionReason = $null

# Detect whether ProductReleaseIds matches a known M365 subscription SKU
# pattern. O365*Retail / M365*Retail / O365EduCloudRetail / OneNoteFreeRetail
# are subscription. Volume / one-time Retail (e.g. ProPlus2021Volume,
# ProPlus2021Retail) return false here.
function Test-OfficeSubscriptionSku {
    param([string]$ProductReleaseIds)
    if ([string]::IsNullOrWhiteSpace($ProductReleaseIds)) { return $false }
    if ($ProductReleaseIds -match '(?i)(O365|M365).*Retail') { return $true }
    if ($ProductReleaseIds -match '(?i)^(O365EduCloudRetail|OneNoteFreeRetail)$') { return $true }
    return $false
}

# Decode a vNext license file. Files are UTF-16LE without BOM, containing
# JSON { License, Certificate, Signature } where License is Base64 of an
# inner JSON. Returns @{ ParseStatus; Inner }; ParseStatus 'OK' on success.
function Read-VnextLicenseFile {
    param([string]$Path)
    try {
        $bytes = [IO.File]::ReadAllBytes($Path)
        if ($bytes.Length -lt 8) {
            return @{ ParseStatus = 'TooSmall'; Inner = $null }
        }
        $text = [System.Text.Encoding]::Unicode.GetString($bytes)
        $outer = $text | ConvertFrom-Json -ErrorAction Stop
        if (-not $outer.License) {
            return @{ ParseStatus = 'OuterMissingLicense'; Inner = $null }
        }
        $licenseBytes = [Convert]::FromBase64String($outer.License)
        $licenseJson  = [System.Text.Encoding]::UTF8.GetString($licenseBytes)
        $inner = $licenseJson | ConvertFrom-Json -ErrorAction Stop
        return @{ ParseStatus = 'OK'; Inner = $inner }
    }
    catch {
        $reason = ($_.Exception.Message -replace '\s+', ' ').Trim()
        if ($reason.Length -gt 80) { $reason = $reason.Substring(0, 80) + '...' }
        return @{ ParseStatus = "ParseFailed: $reason"; Inner = $null }
    }
}

# Cross-section state tracked for the interpretation step
$productReleaseIds = $null
$isSubscription    = $false
$osppShowsGrace    = $false

try {
    # 22a. Click-to-Run configuration registry
    Out-Log "---- Office Click-to-Run Configuration (registry) ----"
    $c2rKey = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
    if (Test-Path $c2rKey) {
        try {
            $c2r = Get-ItemProperty -Path $c2rKey -ErrorAction Stop
            $productReleaseIds = $c2r.ProductReleaseIds
            $isSubscription    = Test-OfficeSubscriptionSku -ProductReleaseIds $productReleaseIds
            Out-Log "ProductReleaseIds:     $productReleaseIds"
            Out-Log "VersionToReport:       $($c2r.VersionToReport)"
            Out-Log "Platform:              $($c2r.Platform)"
            Out-Log "CDNBaseUrl:            $($c2r.CDNBaseUrl)"
            Out-Log "UpdateChannel:         $($c2r.UpdateChannel)"
            Out-Log "AudienceData:          $($c2r.AudienceData)"
            Out-Log "ClientCulture:         $($c2r.ClientCulture)"
            Out-Log "DetectedAsSubscription:$isSubscription"
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
        Out-Log ""
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
            # Detect "NOTIFICATIONS" / "0xC004F009" pattern - Grace expired.
            # This is EXPECTED for M365 subscriptions but indicates real
            # activation failure for VL/buy-once.
            if ($osppOutput -match '(?i)NOTIFICATIONS' -or $osppOutput -match '0xC004F009') {
                $osppShowsGrace = $true
            }
        }
        catch {
            Out-Log "  [WARN] ospp /dstatus failed: $_" -Color Yellow
        }
        Out-Log ""
    }

    # 22c. vNext per-user license file scan
    # Walks every C:\Users\<profile>\AppData\Local\Microsoft\Office\Licenses\
    # subtree. License files are categorized by parent folder (e.g. "5" =
    # commercial subscription). UserGUID is NOT a folder layer; ProductReleaseId
    # / TenantId / UserId live inside the decoded license JSON.
    Out-Log "---- vNext Per-User License Files ----"
    Out-Log "Scanning all user profiles under C:\Users ..."
    Out-Log ""

    $userDirs = @(Get-ChildItem 'C:\Users' -Directory -ErrorAction SilentlyContinue)
    $vnextRows = @()
    $vnextFileCount = 0

    foreach ($u in $userDirs) {
        $licDir = Join-Path $u.FullName 'AppData\Local\Microsoft\Office\Licenses'
        if (-not (Test-Path $licDir)) { continue }
        $categoryDirs = @(Get-ChildItem $licDir -Directory -ErrorAction SilentlyContinue)
        foreach ($cat in $categoryDirs) {
            $licFiles = @(Get-ChildItem $cat.FullName -File -ErrorAction SilentlyContinue)
            foreach ($f in $licFiles) {
                $vnextFileCount++
                $result = Read-VnextLicenseFile -Path $f.FullName
                if ($result.ParseStatus -eq 'OK') {
                    $inner = $result.Inner
                    $tenantId   = ''
                    $userId     = ''
                    $hardwareId = ''
                    $notBefore  = ''
                    $notAfter   = ''
                    if ($inner.Metadata) {
                        try { $tenantId   = "$($inner.Metadata.TenantId)" }   catch { }
                        try { $userId     = "$($inner.Metadata.UserId)" }     catch { }
                        try { $hardwareId = if ($inner.Metadata.HardwareId) { '(present)' } else { '' } } catch { }
                        try { $notBefore  = "$($inner.Metadata.NotBefore)" }  catch { }
                        try { $notAfter   = "$($inner.Metadata.NotAfter)" }   catch { }
                    }
                    $vnextRows += [PSCustomObject]@{
                        UserProfile      = $u.Name
                        Category         = $cat.Name
                        LicenseFile      = $f.Name
                        LicenseType      = "$($inner.LicenseType)"
                        ProductReleaseId = "$($inner.ProductReleaseId)"
                        Status           = "$($inner.Status)"
                        IsTrial          = "$($inner.IsTrial)"
                        Beneficiary      = "$($inner.Beneficiary)"
                        LicenseId        = "$($inner.LicenseId)"
                        Acid             = "$($inner.Acid)"
                        TenantId         = $tenantId
                        UserId           = $userId
                        HardwareIdBound  = $hardwareId
                        NotBefore        = $notBefore
                        NotAfter         = $notAfter
                        ParseStatus      = 'OK'
                    }
                    Out-Log ("  [$($u.Name)] cat=$($cat.Name) Status=$($inner.Status) Type=$($inner.LicenseType) Product=$($inner.ProductReleaseId)")
                }
                else {
                    $vnextRows += [PSCustomObject]@{
                        UserProfile      = $u.Name
                        Category         = $cat.Name
                        LicenseFile      = $f.Name
                        LicenseType      = ''
                        ProductReleaseId = ''
                        Status           = ''
                        IsTrial          = ''
                        Beneficiary      = ''
                        LicenseId        = ''
                        Acid             = ''
                        TenantId         = ''
                        UserId           = ''
                        HardwareIdBound  = ''
                        NotBefore        = ''
                        NotAfter         = ''
                        ParseStatus      = $result.ParseStatus
                    }
                    Out-Log ("  [$($u.Name)] cat=$($cat.Name) [$($result.ParseStatus)]") -Color Yellow
                }
            }
        }
    }

    Out-Log ""
    Out-Log "vNext files found: $vnextFileCount across $($userDirs.Count) user profile(s) scanned"

    # Always emit CSV (even with 0 rows for parser consistency)
    $outVnext = Join-Path $targetDir "22_OfficeVnextLicenses.csv"
    $vnextRows | Export-Csv -Path $outVnext -NoTypeInformation -Encoding UTF8
    Add-SectionFile "22_OfficeVnextLicenses.csv"
    Out-Log ("  -> 22_OfficeVnextLicenses.csv (" + $vnextRows.Count + " row(s))")
    Out-Log ""

    # 22d. Interpretation - cross-source verdict
    Out-Log "---- INTERPRETATION ----"
    Out-Log "ProductReleaseIds (C2R):     $(if ($productReleaseIds) { $productReleaseIds } else { '(not present)' })"
    Out-Log "Detected as subscription:    $isSubscription"
    Out-Log "OSPP shows Grace/Notify:     $osppShowsGrace"

    $provisionedCount = @($vnextRows | Where-Object { $_.Status -eq 'Provisioned' }).Count
    Out-Log "vNext licenses (any status): $($vnextRows.Count)"
    Out-Log "vNext Provisioned:           $provisionedCount"
    Out-Log ""

    if ($isSubscription) {
        if ($osppShowsGrace) {
            Out-Log "  NOTE: OSPP NOTIFICATIONS / Grace expired is EXPECTED on M365 subscription."
            Out-Log "        OSPP is the legacy path and is not authoritative for subscription editions."
            Out-Log "        Authoritative source: per-user vNext license file."
        }
        if ($provisionedCount -gt 0) {
            Out-Log "  CONCLUSION: LICENSED (M365 subscription, $provisionedCount Provisioned vNext license(s))."
        }
        else {
            Out-Log "  CONCLUSION: M365 subscription installed but no Provisioned vNext license found."
            Out-Log "              End-user may not have signed in yet (typical at kitting time)."
            $sectionStatus = 'Partial'
            $sectionReason = "M365 subscription installed but no Provisioned vNext license found (end-user sign-in pending)"
        }
    }
    elseif ($null -ne $productReleaseIds) {
        # VL or one-time Retail (non-subscription)
        if ($osppShowsGrace) {
            Out-Log "  WARNING: OSPP shows Grace/Notifications and this is NOT a subscription SKU."
            Out-Log "           This indicates a real activation failure for VL/buy-once Office."
            $sectionStatus = 'Failed'
            $sectionReason = "OSPP reports Grace/Notifications for non-subscription Office SKU ($productReleaseIds)"
        }
        else {
            Out-Log "  CONCLUSION: VL/buy-once Office, OSPP appears licensed."
        }
    }
    else {
        Out-Log "  CONCLUSION: No Office detected (no C2R registry)."
        # Status stays Success - no Office is a valid state, file is still written
    }

    $sectionCount++
}
catch {
    Out-Log "[ERROR] Failed to collect Office license status: $_" -Color Red
    $failCount++
    $sectionStatus = 'Failed'
    $sectionReason = "$($_.Exception.Message)"
}
Close-Section -Status $sectionStatus -Reason $sectionReason
} else {
    Write-DisabledSection -Id "22" -Title "Office License / Activation Status"
}

# ----------------------------------------
# 23. Security Baseline (TPM / Secure Boot / VBS / LSA / BIOS)
# ----------------------------------------
# Each probe is wrapped in its own try/catch so a single missing capability
# (legacy BIOS without Secure Boot, no TPM hardware, DeviceGuard CIM absent
# on Home SKUs, etc.) does not invalidate the whole section. The outer
# section status stays Success as long as the dispatcher itself does not
# throw — partial probe data is still useful evidence.
# ----------------------------------------
if (Test-SectionEnabled "23") {
Start-Section -Id "23" -Title "Security Baseline" -FileName "23_SecurityBaseline.txt"
$sectionStatus = 'Success'
$sectionReason = $null

try {
    # 23a. TPM
    Out-Log "---- TPM ----"
    try {
        $tpm = Get-Tpm -ErrorAction Stop
        Out-Log "TpmPresent:           $($tpm.TpmPresent)"
        Out-Log "TpmReady:             $($tpm.TpmReady)"
        Out-Log "TpmEnabled:           $($tpm.TpmEnabled)"
        Out-Log "TpmActivated:         $($tpm.TpmActivated)"
        Out-Log "TpmOwned:             $($tpm.TpmOwned)"
        Out-Log "ManufacturerId:       $($tpm.ManufacturerId)"
        Out-Log "ManufacturerVersion:  $($tpm.ManufacturerVersion)"
        Out-Log "ManagedAuthLevel:     $($tpm.ManagedAuthLevel)"
        Out-Log "OwnerAuth:            $(if ($tpm.OwnerAuth) { '(present)' } else { '(absent)' })"
        Out-Log "LockedOut:            $($tpm.LockedOut)"
        Out-Log "LockoutCount:         $($tpm.LockoutCount)"
        Out-Log "LockoutMax:           $($tpm.LockoutMax)"
    }
    catch {
        Out-Log "(Get-Tpm not available: $_)" -Color Yellow
    }
    Out-Log ""

    # 23b. Secure Boot
    Out-Log "---- Secure Boot ----"
    try {
        $sb = Confirm-SecureBootUEFI -ErrorAction Stop
        Out-Log "SecureBootEnabled:    $sb"
    }
    catch {
        Out-Log "(Secure Boot status unavailable - legacy BIOS / unsupported: $_)" -Color Yellow
    }
    Out-Log ""

    # 23c. Virtualization-Based Security / HVCI / Credential Guard
    Out-Log "---- Virtualization-Based Security (Win32_DeviceGuard) ----"
    try {
        $dg = Get-CimInstance -Namespace 'root\Microsoft\Windows\DeviceGuard' -ClassName Win32_DeviceGuard -ErrorAction Stop
        # VirtualizationBasedSecurityStatus: 0=Not running, 1=Configured but not running, 2=Running
        $vbsMap = @{ 0 = 'Not running'; 1 = 'Configured but not running'; 2 = 'Running' }
        $vbsCode = [int]$dg.VirtualizationBasedSecurityStatus
        $vbsText = if ($vbsMap.ContainsKey($vbsCode)) { $vbsMap[$vbsCode] } else { 'Unknown' }
        Out-Log "VirtualizationBasedSecurityStatus:            $vbsText ($vbsCode)"

        # SecurityServicesRunning / Configured: 0=None, 1=Credential Guard, 2=HVCI,
        # 3=System Guard Secure Launch, 4=SMM Firmware Measurement
        $sscMap = @{ 0='None'; 1='Credential Guard'; 2='HVCI'; 3='System Guard Secure Launch'; 4='SMM Firmware Measurement' }
        $running = @($dg.SecurityServicesRunning | ForEach-Object {
            $code = [int]$_
            if ($sscMap.ContainsKey($code)) { $sscMap[$code] } else { "Unknown($code)" }
        })
        $configured = @($dg.SecurityServicesConfigured | ForEach-Object {
            $code = [int]$_
            if ($sscMap.ContainsKey($code)) { $sscMap[$code] } else { "Unknown($code)" }
        })
        Out-Log "SecurityServicesRunning:                      $($running -join ', ')"
        Out-Log "SecurityServicesConfigured:                   $($configured -join ', ')"

        # CodeIntegrityPolicyEnforcementStatus: 0=Off, 1=Audit, 2=Enforced
        $ciMap = @{ 0='Off'; 1='Audit'; 2='Enforced' }
        $ciCode = [int]$dg.CodeIntegrityPolicyEnforcementStatus
        $ciText = if ($ciMap.ContainsKey($ciCode)) { $ciMap[$ciCode] } else { 'Unknown' }
        Out-Log "CodeIntegrityPolicyEnforcementStatus:         $ciText ($ciCode)"
        $umciCode = [int]$dg.UsermodeCodeIntegrityPolicyEnforcementStatus
        $umciText = if ($ciMap.ContainsKey($umciCode)) { $ciMap[$umciCode] } else { 'Unknown' }
        Out-Log "UsermodeCodeIntegrityPolicyEnforcementStatus: $umciText ($umciCode)"

        Out-Log "AvailableSecurityProperties:                  $($dg.AvailableSecurityProperties -join ', ')"
        Out-Log "RequiredSecurityProperties:                   $($dg.RequiredSecurityProperties -join ', ')"
    }
    catch {
        Out-Log "(Win32_DeviceGuard query failed: $_)" -Color Yellow
    }
    Out-Log ""

    # 23d. LSA Protection (RunAsPPL)
    Out-Log "---- LSA Protection (RunAsPPL) ----"
    try {
        $lsaPath = 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa'
        $runAsPPL = (Get-ItemProperty -Path $lsaPath -Name RunAsPPL -ErrorAction SilentlyContinue).RunAsPPL
        $runAsPPLBoot = (Get-ItemProperty -Path $lsaPath -Name RunAsPPLBoot -ErrorAction SilentlyContinue).RunAsPPLBoot
        # 0 / absent = off, 1 = PPL on, 2 = PPL with UEFI lock
        $pplMap = @{ 0 = 'Off'; 1 = 'On (PPL)'; 2 = 'On + UEFI Lock (PPL)' }
        $pplText = if ($null -eq $runAsPPL) {
            'Off (registry value absent)'
        }
        elseif ($pplMap.ContainsKey([int]$runAsPPL)) {
            $pplMap[[int]$runAsPPL]
        }
        else {
            "Unknown ($runAsPPL)"
        }
        Out-Log "RunAsPPL:             $pplText"
        if ($null -ne $runAsPPLBoot) {
            Out-Log "RunAsPPLBoot:         $runAsPPLBoot"
        }
    }
    catch {
        Out-Log "(LSA Protection registry query failed: $_)" -Color Yellow
    }
    Out-Log ""

    # 23e. BIOS / Firmware Info (SerialNumber is collected separately in §10)
    Out-Log "---- BIOS / Firmware ----"
    try {
        $bios = Get-CimInstance -ClassName Win32_BIOS -ErrorAction Stop
        Out-Log "Manufacturer:         $($bios.Manufacturer)"
        Out-Log "SMBIOSBIOSVersion:    $($bios.SMBIOSBIOSVersion)"
        Out-Log "ReleaseDate:          $($bios.ReleaseDate)"
        Out-Log "BIOSVersion:          $(($bios.BIOSVersion -join ' / '))"
        Out-Log "SystemBIOSMajor:      $($bios.SystemBiosMajorVersion)"
        Out-Log "SystemBIOSMinor:      $($bios.SystemBiosMinorVersion)"
        Out-Log "SMBIOSMajor:          $($bios.SMBIOSMajorVersion)"
        Out-Log "SMBIOSMinor:          $($bios.SMBIOSMinorVersion)"
    }
    catch {
        Out-Log "(Win32_BIOS query failed: $_)" -Color Yellow
    }

    $sectionCount++
}
catch {
    Out-Log "[ERROR] Failed to collect security baseline: $_" -Color Red
    $failCount++
    $sectionStatus = 'Failed'
    $sectionReason = "$($_.Exception.Message)"
}
Close-Section -Status $sectionStatus -Reason $sectionReason
} else {
    Write-DisabledSection -Id "23" -Title "Security Baseline"
}

# ----------------------------------------
# 24. Group Policy Report (gpresult /h)
# ----------------------------------------
# Captures Resultant Set of Policy as HTML. The /h output includes both
# computer-side and user-side RSoP. Note: user-side reflects the user
# running this command (typically the kitting profile user, e.g. admin01),
# NOT the eventual end-user. Computer-side reflects the actual machine
# GPO state which is the audit-relevant portion. This caveat is documented
# in Guide.txt.
# ----------------------------------------
if (Test-SectionEnabled "24") {
Start-Section -Id "24" -Title "Group Policy Report" -FileName "24_GroupPolicySummary.txt"
$sectionStatus = 'Success'
$sectionReason = $null

try {
    $gpHtml = Join-Path $targetDir "24_GroupPolicy.html"

    # gpresult /h has a hard 127-char limit on the path argument
    # (cmdline option-value length restriction inherited from cmd.exe).
    # Long evidence dir names (timestamp + PC name + UUID) easily exceed
    # this. Workaround: emit to a short temp path, then move to the
    # actual evidence dir.
    $rand = [System.Guid]::NewGuid().ToString('N').Substring(0, 8)
    $tempHtml = Join-Path $env:TEMP ("fabriq_gp_${rand}.html")
    if ($tempHtml.Length -gt 127) {
        # Fallback: $env:TEMP itself is unusually long. Use system-wide temp.
        $tempHtml = Join-Path 'C:\Windows\Temp' ("fabriq_gp_${rand}.html")
    }

    Out-Log "Running gpresult /h ..."
    Out-Log "Temp HTML:            $tempHtml"
    Out-Log "Target HTML:          24_GroupPolicy.html"
    Out-Log ""

    try {
        # /h <path> /f forces overwrite. stderr captured into stream for diagnostic logging.
        $gpStdout = & gpresult /h $tempHtml /f 2>&1
        $gpExit = $LASTEXITCODE

        foreach ($line in $gpStdout) {
            Out-Log "  $line"
        }
        Out-Log ""

        Out-Log "gpresult exit code:   $gpExit"

        if ($gpExit -ne 0 -or -not (Test-Path $tempHtml)) {
            throw "gpresult failed (exit=$gpExit, temp html present=$(Test-Path $tempHtml))"
        }

        # Move to final evidence location
        Move-Item -Path $tempHtml -Destination $gpHtml -Force -ErrorAction Stop
    }
    finally {
        # Cleanup temp file on any failure path
        if (Test-Path $tempHtml) {
            Remove-Item $tempHtml -Force -ErrorAction SilentlyContinue
        }
    }

    if (-not (Test-Path $gpHtml)) {
        throw "Move to target evidence dir failed: $gpHtml"
    }

    $htmlSize = (Get-Item $gpHtml).Length
    Out-Log "HTML file size:       $htmlSize bytes"

    # Quick sanity context to summarize alongside the HTML
    try {
        $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
        Out-Log "PartOfDomain:         $($cs.PartOfDomain)"
        Out-Log "Domain:               $($cs.Domain)"
    }
    catch {
        Out-Log "(Could not query domain status: $_)" -Color Yellow
    }
    Out-Log "ExecutingUser:        $env:USERNAME"
    Out-Log "(NOTE: user-side RSoP reflects the executing user, not the end-user.)"

    Add-SectionFile "24_GroupPolicy.html"
    $sectionCount++
}
catch {
    Out-Log "[ERROR] Failed to generate Group Policy report: $_" -Color Red
    $failCount++
    $sectionStatus = 'Failed'
    $sectionReason = "$($_.Exception.Message)"
}
Close-Section -Status $sectionStatus -Reason $sectionReason
} else {
    Write-DisabledSection -Id "24" -Title "Group Policy Report"
}

# ----------------------------------------
# 25. Certificates (LocalMachine\My + \Root + \CA + CurrentUser\My)
# ----------------------------------------
# Single CSV with a Store column for unified parsing. Private keys are
# never exported — only the HasPrivateKey boolean flag and standard
# metadata (Subject / Issuer / Thumbprint / NotBefore / NotAfter /
# EnhancedKeyUsageList / FriendlyName / SerialNumber).
# Per-store enumeration failures are logged as warnings and do not fail
# the section as long as Export-Csv itself succeeds.
# ----------------------------------------
if (Test-SectionEnabled "25") {
Start-Section -Id "25" -Title "Certificates (CSV)" -FileName $null
$sectionStatus = 'Success'
$sectionReason = $null

try {
    $stores = @(
        'Cert:\LocalMachine\My',
        'Cert:\LocalMachine\Root',
        'Cert:\LocalMachine\CA',
        'Cert:\CurrentUser\My'
    )

    $rows = @()
    foreach ($s in $stores) {
        $storeLabel = $s -replace '^Cert:\\', ''
        try {
            $certs = @(Get-ChildItem -Path $s -ErrorAction Stop)
            foreach ($c in $certs) {
                $eku = ''
                try {
                    if ($c.EnhancedKeyUsageList) {
                        $eku = ($c.EnhancedKeyUsageList | ForEach-Object {
                            if ($_.FriendlyName) { $_.FriendlyName } else { $_.ObjectId }
                        }) -join '; '
                    }
                }
                catch {
                    $eku = ''
                }
                $rows += [PSCustomObject]@{
                    Store                = $storeLabel
                    Subject              = $c.Subject
                    Issuer               = $c.Issuer
                    Thumbprint           = $c.Thumbprint
                    NotBefore            = $c.NotBefore
                    NotAfter             = $c.NotAfter
                    HasPrivateKey        = $c.HasPrivateKey
                    EnhancedKeyUsageList = $eku
                    FriendlyName         = $c.FriendlyName
                    SerialNumber         = $c.SerialNumber
                }
            }
            Out-Log ("  $storeLabel : " + $certs.Count + " certificate(s)")
        }
        catch {
            Out-Log "  [WARN] Could not enumerate $s : $_" -Color Yellow
        }
    }

    $outCerts = Join-Path $targetDir "25_Certificates.csv"
    $rows | Export-Csv -Path $outCerts -NoTypeInformation -Encoding UTF8
    Add-SectionFile "25_Certificates.csv"

    Out-Log ("Total certificates: " + $rows.Count + " -> 25_Certificates.csv")

    $sectionCount++
}
catch {
    Out-Log "[ERROR] Failed to enumerate certificates: $_" -Color Red
    $failCount++
    $sectionStatus = 'Failed'
    $sectionReason = "$($_.Exception.Message)"
}
Close-Section -Status $sectionStatus -Reason $sectionReason
} else {
    Write-DisabledSection -Id "25" -Title "Certificates (CSV)"
}

# ----------------------------------------
# 26. Battery Report (laptop only)
# ----------------------------------------
# powercfg /batteryreport produces an HTML showing design vs full charge
# capacity, recent usage, and lifetime estimates. Critical for laptop
# acceptance inspection (e.g. contract clause "battery initial capacity
# >= 95% of design"). Skipped when no battery is present (desktop PC).
# ----------------------------------------
if (Test-SectionEnabled "26") {
Start-Section -Id "26" -Title "Battery Report" -FileName $null
$sectionStatus = 'Success'
$sectionReason = $null

try {
    $batteries = @(Get-CimInstance Win32_Battery -ErrorAction SilentlyContinue)
    if ($batteries.Count -eq 0) {
        Out-Log "No battery detected (Win32_Battery returned 0 instances). Skipping."
        $sectionCount++
        $sectionStatus = 'Skipped'
        $sectionReason = 'No battery present (desktop PC or battery removed)'
    }
    else {
        Out-Log "Battery detected ($($batteries.Count) instance(s)). Generating report..."
        $reportHtml = Join-Path $targetDir "26_BatteryReport.html"

        $pcOut = & powercfg /batteryreport /output $reportHtml 2>&1
        $pcExit = $LASTEXITCODE

        foreach ($line in $pcOut) {
            Out-Log "  $line"
        }

        if ($pcExit -eq 0 -and (Test-Path $reportHtml)) {
            $reportSize = (Get-Item $reportHtml).Length
            Out-Log "Report generated:     26_BatteryReport.html ($reportSize bytes)"
            Add-SectionFile "26_BatteryReport.html"
            $sectionCount++
        }
        else {
            throw "powercfg /batteryreport failed (exit=$pcExit, html present=$(Test-Path $reportHtml))"
        }
    }
}
catch {
    Out-Log "[ERROR] Failed to generate battery report: $_" -Color Red
    $failCount++
    $sectionStatus = 'Failed'
    $sectionReason = "$($_.Exception.Message)"
}
Close-Section -Status $sectionStatus -Reason $sectionReason
} else {
    Write-DisabledSection -Id "26" -Title "Battery Report"
}

# ----------------------------------------
# 27. Environment Variables (Machine + User scopes)
# ----------------------------------------
# Machine and User scope only. Process scope is volatile (depends on the
# running shell context, not the system state) so it is excluded as
# evidentiary noise. Values are recorded raw — masking is intentionally
# avoided because evidence must capture what was actually configured.
# ----------------------------------------
if (Test-SectionEnabled "27") {
Start-Section -Id "27" -Title "Environment Variables (CSV)" -FileName $null
$sectionStatus = 'Success'
$sectionReason = $null

try {
    $envRows = @()
    foreach ($scope in 'Machine','User') {
        try {
            $vars = [System.Environment]::GetEnvironmentVariables($scope)
            foreach ($key in $vars.Keys) {
                $envRows += [PSCustomObject]@{
                    Scope = $scope
                    Name  = $key
                    Value = $vars[$key]
                }
            }
            Out-Log "  $scope scope: $($vars.Count) variables"
        }
        catch {
            Out-Log "  [WARN] Could not enumerate $scope scope: $_" -Color Yellow
        }
    }

    $envRows = $envRows | Sort-Object Scope, Name
    $outEnv = Join-Path $targetDir "27_EnvironmentVariables.csv"
    $envRows | Export-Csv -Path $outEnv -NoTypeInformation -Encoding UTF8
    Add-SectionFile "27_EnvironmentVariables.csv"

    Out-Log "Environment variables: $($envRows.Count) entries -> 27_EnvironmentVariables.csv"
    $sectionCount++
}
catch {
    Out-Log "[ERROR] Failed to enumerate environment variables: $_" -Color Red
    $failCount++
    $sectionStatus = 'Failed'
    $sectionReason = "$($_.Exception.Message)"
}
Close-Section -Status $sectionStatus -Reason $sectionReason
} else {
    Write-DisabledSection -Id "27" -Title "Environment Variables (CSV)"
}

# ----------------------------------------
# 28. Startup Items (Win32_StartupCommand + logon-triggered ScheduledTask)
# ----------------------------------------
# Two sources combined into a single CSV with a Source column:
#   - Win32_StartupCommand : Run / RunOnce registry keys + Startup folders
#     (legacy comprehensive view, PCView-compatible)
#   - Get-ScheduledTask    : logon-triggered tasks that are not Disabled,
#     excluding Microsoft\Windows\* OS internals (audit noise reduction)
# Disabled tasks are dropped to keep the CSV evidence-relevant.
# ----------------------------------------
if (Test-SectionEnabled "28") {
Start-Section -Id "28" -Title "Startup Items (CSV)" -FileName $null
$sectionStatus = 'Success'
$sectionReason = $null

try {
    $startupRows = @()

    # 28a. Win32_StartupCommand (Run / RunOnce / Startup folder)
    try {
        $cmds = @(Get-CimInstance Win32_StartupCommand -ErrorAction Stop)
        foreach ($c in $cmds) {
            $startupRows += [PSCustomObject]@{
                Source   = 'Win32_StartupCommand'
                Name     = $c.Name
                User     = $c.User
                Command  = $c.Command
                Location = $c.Location
                Enabled  = $true
            }
        }
        Out-Log "  Win32_StartupCommand: $($cmds.Count) entries"
    }
    catch {
        Out-Log "  [WARN] Win32_StartupCommand failed: $_" -Color Yellow
    }

    # 28b. Logon-triggered ScheduledTask (non-Disabled, exclude OS internals)
    try {
        $tasks = @(Get-ScheduledTask -ErrorAction Stop | Where-Object {
            $_.State -ne 'Disabled' -and
            $_.TaskPath -notlike '\Microsoft\Windows\*' -and
            ($_.Triggers | Where-Object { $_.CimClass.CimClassName -eq 'MSFT_TaskLogonTrigger' })
        })
        foreach ($t in $tasks) {
            $action = $t.Actions | Select-Object -First 1
            $cmdText = if ($action) {
                $exe = $action.Execute
                $args = $action.Arguments
                if ([string]::IsNullOrWhiteSpace($args)) { $exe } else { "$exe $args" }
            } else {
                ''
            }
            $startupRows += [PSCustomObject]@{
                Source   = 'ScheduledTask'
                Name     = $t.TaskName
                User     = ($t.Principal.UserId)
                Command  = $cmdText
                Location = $t.TaskPath
                Enabled  = ($t.State -ne 'Disabled')
            }
        }
        Out-Log "  ScheduledTask (logon-trigger, non-Disabled, non-MS): $($tasks.Count) entries"
    }
    catch {
        Out-Log "  [WARN] Get-ScheduledTask failed: $_" -Color Yellow
    }

    $outStartup = Join-Path $targetDir "28_StartupItems.csv"
    $startupRows | Export-Csv -Path $outStartup -NoTypeInformation -Encoding UTF8
    Add-SectionFile "28_StartupItems.csv"

    Out-Log "Startup items: $($startupRows.Count) entries -> 28_StartupItems.csv"
    $sectionCount++
}
catch {
    Out-Log "[ERROR] Failed to enumerate startup items: $_" -Color Red
    $failCount++
    $sectionStatus = 'Failed'
    $sectionReason = "$($_.Exception.Message)"
}
Close-Section -Status $sectionStatus -Reason $sectionReason
} else {
    Write-DisabledSection -Id "28" -Title "Startup Items (CSV)"
}

# ----------------------------------------
# 29. Memory Slots (per-slot detail) + 29b. Memory Array Summary
# ----------------------------------------
# Two CSVs (mirrors the §8b Disks/Partitions split convention):
#   - 29_MemorySlots.csv         : per-slot detail (Win32_PhysicalMemory)
#   - 29b_MemoryArraySummary.csv : array-level (Win32_PhysicalMemoryArray)
# FormFactor and SMBIOSMemoryType numeric codes are translated to strings
# per the WMI schema (FormFactor 8=DIMM / 12=SODIMM, SMBIOSMemoryType
# 24=DDR3 / 26=DDR4 / 30=LPDDR4 / 34=DDR5 / 35=LPDDR5).
# ----------------------------------------
if (Test-SectionEnabled "29") {
Start-Section -Id "29" -Title "Memory Slots (CSV)" -FileName $null
$sectionStatus = 'Success'
$sectionReason = $null

try {
    # Translation tables (WMI schema constants)
    $formFactorMap = @{
        0 = 'Unknown'; 1 = 'Other'; 2 = 'SIP'; 3 = 'DIP'; 4 = 'ZIP'; 5 = 'SOJ'
        6 = 'Proprietary'; 7 = 'SIMM'; 8 = 'DIMM'; 9 = 'TSOP'; 10 = 'PGA'
        11 = 'RIMM'; 12 = 'SODIMM'; 13 = 'SRIMM'; 14 = 'SMD'; 15 = 'SSMP'
        16 = 'QFP'; 17 = 'TQFP'; 18 = 'SOIC'; 19 = 'LCC'; 20 = 'PLCC'
        21 = 'BGA'; 22 = 'FPBGA'; 23 = 'LGA'; 24 = 'FB-DIMM'
    }
    $smbiosMemTypeMap = @{
        0 = 'Unknown'; 1 = 'Other'; 2 = 'DRAM'; 3 = 'Synchronous DRAM'
        4 = 'Cache DRAM'; 5 = 'EDO'; 6 = 'EDRAM'; 7 = 'VRAM'; 8 = 'SRAM'
        9 = 'RAM'; 10 = 'ROM'; 11 = 'Flash'; 12 = 'EEPROM'; 13 = 'FEPROM'
        14 = 'EPROM'; 15 = 'CDRAM'; 16 = '3DRAM'; 17 = 'SDRAM'; 18 = 'SGRAM'
        19 = 'RDRAM'; 20 = 'DDR'; 21 = 'DDR2'; 22 = 'DDR2 FB-DIMM'
        24 = 'DDR3'; 25 = 'FBD2'; 26 = 'DDR4'; 27 = 'LPDDR'; 28 = 'LPDDR2'
        29 = 'LPDDR3'; 30 = 'LPDDR4'; 31 = 'Logical non-volatile device'
        32 = 'HBM'; 33 = 'HBM2'; 34 = 'DDR5'; 35 = 'LPDDR5'
    }

    # 29. Per-slot detail
    $slotRows = @()
    $modules = @(Get-CimInstance Win32_PhysicalMemory -ErrorAction Stop)
    foreach ($m in $modules) {
        $ffCode = [int]$m.FormFactor
        $ffName = if ($formFactorMap.ContainsKey($ffCode)) { $formFactorMap[$ffCode] } else { "Code $ffCode" }
        $mtCode = if ($null -ne $m.SMBIOSMemoryType) { [int]$m.SMBIOSMemoryType } else { 0 }
        $mtName = if ($smbiosMemTypeMap.ContainsKey($mtCode)) { $smbiosMemTypeMap[$mtCode] } else { "Code $mtCode" }

        $capGB = if ($m.Capacity) { [math]::Round([double]$m.Capacity / 1GB, 2) } else { 0 }

        $slotRows += [PSCustomObject]@{
            BankLabel              = $m.BankLabel
            DeviceLocator          = $m.DeviceLocator
            Capacity_GB            = $capGB
            Speed_MHz              = $m.Speed
            ConfiguredClockSpeed_MHz = $m.ConfiguredClockSpeed
            ConfiguredVoltage_mV   = $m.ConfiguredVoltage
            Manufacturer           = $m.Manufacturer
            PartNumber             = ($m.PartNumber).Trim()
            SerialNumber           = ($m.SerialNumber).Trim()
            FormFactor             = $ffName
            SMBIOSMemoryType       = $mtName
            DataWidth_bit          = $m.DataWidth
            TotalWidth_bit         = $m.TotalWidth
        }
    }
    $outSlots = Join-Path $targetDir "29_MemorySlots.csv"
    $slotRows | Export-Csv -Path $outSlots -NoTypeInformation -Encoding UTF8
    Add-SectionFile "29_MemorySlots.csv"
    Out-Log "Memory slots: $($slotRows.Count) populated module(s) -> 29_MemorySlots.csv"

    # 29b. Array summary
    $arrayRows = @()
    try {
        $arrays = @(Get-CimInstance Win32_PhysicalMemoryArray -ErrorAction Stop)
        foreach ($a in $arrays) {
            # MaxCapacity is in KB; MaxCapacityEx (UInt64) is preferred when present
            $maxKB = if ($a.MaxCapacityEx -and $a.MaxCapacityEx -gt 0) { [double]$a.MaxCapacityEx } else { [double]$a.MaxCapacity }
            $maxGB = if ($maxKB) { [math]::Round($maxKB / 1MB, 2) } else { 0 }
            $arrayRows += [PSCustomObject]@{
                Tag             = $a.Tag
                Location        = $a.Location
                Use             = $a.Use
                MemoryErrorCorrection = $a.MemoryErrorCorrection
                MaxCapacity_GB  = $maxGB
                MemoryDevices   = $a.MemoryDevices
            }
        }
        $outArray = Join-Path $targetDir "29b_MemoryArraySummary.csv"
        $arrayRows | Export-Csv -Path $outArray -NoTypeInformation -Encoding UTF8
        Add-SectionFile "29b_MemoryArraySummary.csv"
        Out-Log "Memory array summary: $($arrayRows.Count) array(s) -> 29b_MemoryArraySummary.csv"
    }
    catch {
        Out-Log "  [WARN] Win32_PhysicalMemoryArray failed: $_" -Color Yellow
        $sectionStatus = 'Partial'
        $sectionReason = "Array summary sub-collection failed: $($_.Exception.Message)"
    }

    $sectionCount++
}
catch {
    Out-Log "[ERROR] Failed to enumerate memory slots: $_" -Color Red
    $failCount++
    $sectionStatus = 'Failed'
    $sectionReason = "$($_.Exception.Message)"
}
Close-Section -Status $sectionStatus -Reason $sectionReason
} else {
    Write-DisabledSection -Id "29" -Title "Memory Slots (CSV)"
}

# ----------------------------------------
# 30. PnP Devices (full enumeration with driver version/date)
# ----------------------------------------
# Get-PnpDevice without -PresentOnly returns past-connected devices too,
# which is intentional for audit traceability. DriverVersion / DriverDate
# are queried per-instance via Get-PnpDeviceProperty (cost: tens of seconds
# on a typical client). Per-device query failures fall back to blank cells
# without failing the section.
# ----------------------------------------
if (Test-SectionEnabled "30") {
Start-Section -Id "30" -Title "PnP Devices (CSV)" -FileName $null
$sectionStatus = 'Success'
$sectionReason = $null

try {
    $devices = @(Get-PnpDevice -ErrorAction Stop | Sort-Object Class, FriendlyName)
    $deviceRows = @()
    $queryFailures = 0

    foreach ($d in $devices) {
        $driverVersion = ''
        $driverDate    = ''
        try {
            $props = Get-PnpDeviceProperty -InstanceId $d.InstanceId `
                -KeyName 'DEVPKEY_Device_DriverVersion','DEVPKEY_Device_DriverDate' `
                -ErrorAction Stop
            foreach ($p in $props) {
                switch ($p.KeyName) {
                    'DEVPKEY_Device_DriverVersion' { $driverVersion = "$($p.Data)" }
                    'DEVPKEY_Device_DriverDate'    {
                        if ($p.Data -is [datetime]) {
                            $driverDate = $p.Data.ToString('yyyy-MM-dd')
                        } else {
                            $driverDate = "$($p.Data)"
                        }
                    }
                }
            }
        }
        catch {
            $queryFailures++
        }

        $deviceRows += [PSCustomObject]@{
            Class         = $d.Class
            FriendlyName  = $d.FriendlyName
            Status        = $d.Status
            Present       = $d.Present
            Manufacturer  = $d.Manufacturer
            Service       = $d.Service
            DriverVersion = $driverVersion
            DriverDate    = $driverDate
            InstanceId    = $d.InstanceId
        }
    }

    $outPnp = Join-Path $targetDir "30_PnpDevices.csv"
    $deviceRows | Export-Csv -Path $outPnp -NoTypeInformation -Encoding UTF8
    Add-SectionFile "30_PnpDevices.csv"

    Out-Log "PnP devices: $($deviceRows.Count) entries (driver query failures: $queryFailures) -> 30_PnpDevices.csv"
    $sectionCount++
}
catch {
    Out-Log "[ERROR] Failed to enumerate PnP devices: $_" -Color Red
    $failCount++
    $sectionStatus = 'Failed'
    $sectionReason = "$($_.Exception.Message)"
}
Close-Section -Status $sectionStatus -Reason $sectionReason
} else {
    Write-DisabledSection -Id "30" -Title "PnP Devices (CSV)"
}

# ----------------------------------------
# 31. Hardware Identifiers (System / BaseBoard / Enclosure)
# ----------------------------------------
# Aggregates four WMI classes that PCView's "WMI" tab covered but were
# previously unrecorded by fabriq evidence:
#   - Win32_ComputerSystem        (Manufacturer / Model / SystemFamily / SKU)
#   - Win32_ComputerSystemProduct (Vendor / Name / IdentifyingNumber / UUID)
#   - Win32_BaseBoard             (motherboard Manufacturer / Product / SN)
#   - Win32_SystemEnclosure       (chassis type / asset tag / SN)
# ChassisTypes numeric codes are translated to strings per the WMI schema.
# This complements §10 (PC serial number) and §23 (BIOS / TPM) without
# duplication.
# ----------------------------------------
if (Test-SectionEnabled "31") {
Start-Section -Id "31" -Title "Hardware Identifiers" -FileName "31_HardwareIdentifiers.txt"
$sectionStatus = 'Success'
$sectionReason = $null

try {
    $chassisTypeMap = @{
        1 = 'Other'; 2 = 'Unknown'; 3 = 'Desktop'; 4 = 'Low Profile Desktop'
        5 = 'Pizza Box'; 6 = 'Mini Tower'; 7 = 'Tower'; 8 = 'Portable'
        9 = 'Laptop'; 10 = 'Notebook'; 11 = 'Hand Held'; 12 = 'Docking Station'
        13 = 'All-in-One'; 14 = 'Sub-Notebook'; 15 = 'Space-Saving'
        16 = 'Lunch Box'; 17 = 'Main Server Chassis'; 18 = 'Expansion Chassis'
        19 = 'Sub-Chassis'; 20 = 'Bus Expansion Chassis'; 21 = 'Peripheral Chassis'
        22 = 'RAID Chassis'; 23 = 'Rack-Mount Chassis'; 24 = 'Sealed-Case PC'
        25 = 'Multi-System Chassis'; 26 = 'Compact PCI'; 27 = 'Advanced TCA'
        28 = 'Blade'; 29 = 'Blade Enclosure'; 30 = 'Tablet'; 31 = 'Convertible'
        32 = 'Detachable'; 33 = 'IoT Gateway'; 34 = 'Embedded PC'
        35 = 'Mini PC'; 36 = 'Stick PC'
    }

    Out-Log "---- Win32_ComputerSystem ----"
    try {
        $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
        Out-Log "Manufacturer:      $($cs.Manufacturer)"
        Out-Log "Model:             $($cs.Model)"
        Out-Log "SystemFamily:      $($cs.SystemFamily)"
        Out-Log "SystemSKUNumber:   $($cs.SystemSKUNumber)"
        Out-Log "TotalPhysicalMem:  $([math]::Round([double]$cs.TotalPhysicalMemory / 1GB, 2)) GB"
        Out-Log "NumberOfProcessors: $($cs.NumberOfProcessors)"
        Out-Log "NumberOfLogicalProcessors: $($cs.NumberOfLogicalProcessors)"
    }
    catch {
        Out-Log "  [WARN] Win32_ComputerSystem failed: $_" -Color Yellow
    }
    Out-Log ""

    Out-Log "---- Win32_ComputerSystemProduct ----"
    try {
        $csp = Get-CimInstance Win32_ComputerSystemProduct -ErrorAction Stop
        Out-Log "Vendor:            $($csp.Vendor)"
        Out-Log "Name:              $($csp.Name)"
        Out-Log "Version:           $($csp.Version)"
        Out-Log "IdentifyingNumber: $($csp.IdentifyingNumber)"
        Out-Log "UUID:              $($csp.UUID)"
        Out-Log "SKUNumber:         $($csp.SKUNumber)"
    }
    catch {
        Out-Log "  [WARN] Win32_ComputerSystemProduct failed: $_" -Color Yellow
    }
    Out-Log ""

    Out-Log "---- Win32_BaseBoard ----"
    try {
        $bb = Get-CimInstance Win32_BaseBoard -ErrorAction Stop
        Out-Log "Manufacturer:      $($bb.Manufacturer)"
        Out-Log "Product:           $($bb.Product)"
        Out-Log "Version:           $($bb.Version)"
        Out-Log "SerialNumber:      $($bb.SerialNumber)"
        Out-Log "Tag:               $($bb.Tag)"
    }
    catch {
        Out-Log "  [WARN] Win32_BaseBoard failed: $_" -Color Yellow
    }
    Out-Log ""

    Out-Log "---- Win32_SystemEnclosure ----"
    try {
        $se = Get-CimInstance Win32_SystemEnclosure -ErrorAction Stop
        $ctRaw = @($se.ChassisTypes)
        $ctNames = $ctRaw | ForEach-Object {
            $code = [int]$_
            if ($chassisTypeMap.ContainsKey($code)) {
                "$code ($($chassisTypeMap[$code]))"
            } else {
                "$code (Unmapped)"
            }
        }
        Out-Log "Manufacturer:      $($se.Manufacturer)"
        Out-Log "Model:             $($se.Model)"
        Out-Log "ChassisTypes:      $($ctNames -join ', ')"
        Out-Log "SerialNumber:      $($se.SerialNumber)"
        Out-Log "SMBIOSAssetTag:    $($se.SMBIOSAssetTag)"
        Out-Log "AssetTag:          $($se.AssetTag)"
        Out-Log "SecurityStatus:    $($se.SecurityStatus)"
        Out-Log "LockPresent:       $($se.LockPresent)"
    }
    catch {
        Out-Log "  [WARN] Win32_SystemEnclosure failed: $_" -Color Yellow
    }

    $sectionCount++
}
catch {
    Out-Log "[ERROR] Failed to collect hardware identifiers: $_" -Color Red
    $failCount++
    $sectionStatus = 'Failed'
    $sectionReason = "$($_.Exception.Message)"
}
Close-Section -Status $sectionStatus -Reason $sectionReason
} else {
    Write-DisabledSection -Id "31" -Title "Hardware Identifiers"
}

# ----------------------------------------
# Completion
# ----------------------------------------
$currentSplitFile = $null

# Write evidence manifest (kernel/EVIDENCE_MANIFEST.md schemaVersion=1)
$evidenceConfigVersionFile = Join-Path $PSScriptRoot "VERSION"
$evidenceConfigVersion = if (Test-Path $evidenceConfigVersionFile) {
    (Get-Content $evidenceConfigVersionFile -Raw).Trim()
} else {
    "unknown"
}

try {
    $manifestPath = Write-EvidenceManifest `
        -TargetDir             $targetDir `
        -EvidenceConfigVersion $evidenceConfigVersion `
        -ComputerName          $env:COMPUTERNAME `
        -HardwareUniqueId      $uid `
        -SelectedNewPcName     $pcName `
        -CollectedAt           $script:CollectedAt
    Out-Log ""
    Out-Log "Manifest written: $manifestPath" -Color Cyan
}
catch {
    Out-Log "[WARN] Failed to write manifest.json: $_" -Color Yellow
}

Out-Log ""
Out-Log "==== Evidence Collection Completed ====" -Color Cyan

Write-Host ""
Show-Info "Evidence saved to: $targetDir"
Write-Host ""

return (New-BatchResult -Success $sectionCount -Fail $failCount -Title "Evidence Collection Results")
