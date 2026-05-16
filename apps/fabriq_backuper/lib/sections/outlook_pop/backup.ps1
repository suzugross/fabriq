# ============================================================
# FabriqBackUper Section: outlook_pop / backup (Phase 2.9.0 Phase A)
#
# Enumerates Outlook (classic 2016/2019/2021/365, registry version
# 16.0; falls back to 15.0 for Outlook 2013) Mail Profiles under
#   HKCU\Software\Microsoft\Office\<ver>\Outlook\Profiles\<prof>\
#     9375CFF0413111d3B88A00104B2A6676\<NN>
# and emits a portable JSON manifest of every POP3 account it finds.
#
# Notes:
#   - The well-known GUID 9375CFF0413111d3B88A00104B2A6676 is the
#     MAPI "Internet Account" service identifier (POP / IMAP / SMTP).
#   - String values under this key are stored as REG_BINARY with a
#     UTF-16LE encoding (often null-terminated); numeric values are
#     REG_DWORD.
#   - Passwords are DPAPI-encrypted per-user/per-machine and CANNOT
#     be decrypted on another box. Microsoft's PRF format also does
#     not deploy passwords on import (Outlook 2016+ silently drops
#     them). This section therefore intentionally skips password
#     blobs entirely — the operator re-enters on first send/receive.
#   - IMAP accounts (POP3 Server value absent, IMAP Server present)
#     are detected, counted, and intentionally skipped in Phase A.
#
# SectionParams (hashtable, all optional):
#   SourceUserProfilePath : profile path of the user whose HKCU to
#                           enumerate. Matches the same parameter used
#                           by the userdata section. Resolve-HkcuRoot
#                           handles the SID lookup.
# ============================================================

param(
    [Parameter(Mandatory = $true)][string]$BackuperRoot,
    [Parameter(Mandatory = $true)][string]$FabriqRoot,
    [Parameter(Mandatory = $true)][string]$OldPcName,
    [Parameter(Mandatory = $true)][string]$AggregateBackupDir,
    [hashtable]$SectionParams = @{}
)

$sw = [System.Diagnostics.Stopwatch]::StartNew()
$warnings = @()

# ----------------------------------------------------------
# Parse SectionParams
# ----------------------------------------------------------
$sourceUserProfilePath = $null
if ($SectionParams.ContainsKey('SourceUserProfilePath') -and `
    -not [string]::IsNullOrWhiteSpace($SectionParams['SourceUserProfilePath'])) {
    $sourceUserProfilePath = "$($SectionParams['SourceUserProfilePath'])"
}

# ----------------------------------------------------------
# Section output dir
# ----------------------------------------------------------
$sectionDir = Join-Path $AggregateBackupDir 'sections\outlook_pop'
try {
    $null = New-Item -ItemType Directory -Path $sectionDir -Force -ErrorAction Stop
} catch {
    return [PSCustomObject]@{
        Status               = 'Failed'
        ElapsedMs            = [int]$sw.ElapsedMilliseconds
        Summary              = [ordered]@{}
        Warnings             = @("Failed to create section output dir: $($_.Exception.Message)")
        ExternalOutputDir    = $null
        ExternalManifestPath = $null
    }
}

Show-Info "Section output: $sectionDir"

# ----------------------------------------------------------
# Resolve HKCU root (handles admin elevation / cross-user)
# ----------------------------------------------------------
$hkcuInfo = Resolve-HkcuRoot
if ($null -eq $hkcuInfo -or [string]::IsNullOrWhiteSpace($hkcuInfo.PsDrivePath)) {
    return [PSCustomObject]@{
        Status               = 'Failed'
        ElapsedMs            = [int]$sw.ElapsedMilliseconds
        Summary              = [ordered]@{}
        Warnings             = @('Resolve-HkcuRoot returned null/empty PsDrivePath')
        ExternalOutputDir    = $null
        ExternalManifestPath = $null
    }
}
if ($hkcuInfo.Redirected) {
    Show-Info "HKCU source: $($hkcuInfo.Label) [SID=$($hkcuInfo.SID)]"
}

# ----------------------------------------------------------
# Locate the Outlook Profiles key (try 16.0 first, then 15.0)
# ----------------------------------------------------------
$outlookVersions = @('16.0', '15.0')
$profilesKeyPath = $null
$outlookVersion  = $null
foreach ($v in $outlookVersions) {
    $candidate = "$($hkcuInfo.PsDrivePath)\Software\Microsoft\Office\$v\Outlook\Profiles"
    if (Test-Path $candidate) {
        $profilesKeyPath = $candidate
        $outlookVersion  = $v
        break
    }
}
if ($null -eq $profilesKeyPath) {
    Show-Skip "No Outlook 16.0 or 15.0 mail profile registry found — skipping section"
    return [PSCustomObject]@{
        Status               = 'Skipped'
        ElapsedMs            = [int]$sw.ElapsedMilliseconds
        Summary              = [ordered]@{ note = 'no Outlook 16.0/15.0 profile registry' }
        Warnings             = @()
        ExternalOutputDir    = $null
        ExternalManifestPath = $null
    }
}
Show-Info "Outlook version: $outlookVersion"
Show-Info "Profiles root  : $profilesKeyPath"

# ----------------------------------------------------------
# Helpers
# ----------------------------------------------------------
$INTERNET_ACCOUNT_GUID = '9375CFF0413111d3B88A00104B2A6676'

function ConvertFrom-RegBinaryUtf16 {
    # REG_BINARY containing UTF-16LE text, often with a trailing null
    # pair (0x00 0x00). Strip nulls and return the resulting string.
    param($Value)
    if ($null -eq $Value) { return '' }
    $bytes = [byte[]]$Value
    if ($bytes.Length -eq 0) { return '' }
    $s = [System.Text.Encoding]::Unicode.GetString($bytes)
    return $s.TrimEnd([char]0)
}

function Get-RegValueRaw {
    # Returns the raw value (whatever type) or $null if absent.
    param(
        [Parameter(Mandatory = $true)]$RegKey,
        [Parameter(Mandatory = $true)][string]$Name
    )
    try { return $RegKey.GetValue($Name, $null) } catch { return $null }
}

function Get-RegValueString {
    # POP3-account string values are stored as REG_BINARY UTF-16LE.
    # Some installations store them as REG_SZ (rare but seen);
    # accept both.
    param(
        [Parameter(Mandatory = $true)]$RegKey,
        [Parameter(Mandatory = $true)][string]$Name
    )
    $v = Get-RegValueRaw -RegKey $RegKey -Name $Name
    if ($null -eq $v) { return $null }
    if ($v -is [byte[]]) { return (ConvertFrom-RegBinaryUtf16 -Value $v) }
    return [string]$v
}

function Get-RegValueDword {
    param(
        [Parameter(Mandatory = $true)]$RegKey,
        [Parameter(Mandatory = $true)][string]$Name
    )
    $v = Get-RegValueRaw -RegKey $RegKey -Name $Name
    if ($null -eq $v) { return $null }
    return [int]$v
}

function Get-PstPathsFromProfile {
    # Phase 2.9.2a-v2: recursive walk of the Outlook profile registry to
    # find every subkey that has a '001f6700' value — that's the PST
    # file path (REG_BINARY UTF-16LE, null-terminated) for a Personal
    # Folders store. Returns an array of file path strings.
    #
    # This is the more reliable detection path; the EntryID parser
    # below (Get-PstPathFromDeliveryStoreEntryId) only works for the
    # "wrapped PST store provider" SAMPLE format, not the production
    # Outlook PST provider which embeds "mspst.dll" plus undocumented
    # bytes before the actual path.
    param([Parameter(Mandatory = $true)][string]$ProfileKeyPath)

    $results = New-Object System.Collections.Generic.List[string]
    $stack = New-Object System.Collections.Generic.Stack[string]
    $stack.Push($ProfileKeyPath)

    while ($stack.Count -gt 0) {
        $current = $stack.Pop()
        $key = $null
        try { $key = Get-Item -LiteralPath $current -ErrorAction Stop } catch { continue }
        if ($null -eq $key) { continue }

        # Check this key for the PST path value (REG_BINARY UTF-16LE).
        try {
            $raw = $key.GetValue('001f6700', $null)
            if ($null -ne $raw -and $raw -is [byte[]] -and $raw.Length -gt 0) {
                $s = [System.Text.Encoding]::Unicode.GetString([byte[]]$raw)
                $s = $s.TrimEnd([char]0)
                if (-not [string]::IsNullOrWhiteSpace($s) -and `
                    $s -match '^[A-Za-z]:\\' -and $s -match '\.pst$') {
                    [void]$results.Add($s)
                }
            }
        } catch { }

        # Recurse into subkeys
        try {
            foreach ($sub in (Get-ChildItem -LiteralPath $current -ErrorAction Stop)) {
                $stack.Push($sub.PSPath)
            }
        } catch { }
    }

    return @($results)
}

function Get-PstPathFromDeliveryStoreEntryId {
    # Phase 2.9.2a: parse a POP3 account's "Delivery Store EntryID"
    # REG_BINARY value to extract the bound PST file path.
    #
    # Format per Microsoft Learn (wrapped PST store provider):
    #   EIDMS  (ASCII)   = [4 flags][16 MAPIUID][1 reserved][N CHAR  szPath\0]
    #   EIDMSW (Unicode) = [4 flags][16 MAPIUID][1 reserved][1 NUL    szPath]
    #                                                       [N WCHAR wzPath\0\0]
    # Detection: if byte at offset 21 is 0x00 -> EIDMSW (Unicode follows),
    # otherwise EIDMS (ASCII path starts here).
    # Returns the extracted path string on success, or $null if the value
    # is absent / too short / fails validation (.pst extension + drive letter).
    param([byte[]]$EntryIdBytes)

    if ($null -eq $EntryIdBytes -or $EntryIdBytes.Length -lt 22) { return $null }

    $headerLen = 4 + 16 + 1   # flags + MAPIUID + reserved = 21
    $offset = $headerLen
    $path = $null

    if ($EntryIdBytes[$offset] -eq 0x00) {
        # EIDMSW: skip the empty ASCII NULL byte, then read UTF-16LE
        $unicodeStart = $offset + 1
        $end = $unicodeStart
        while (($end + 1) -lt $EntryIdBytes.Length) {
            if ($EntryIdBytes[$end] -eq 0x00 -and $EntryIdBytes[$end + 1] -eq 0x00) { break }
            $end += 2
        }
        if ($end -gt $unicodeStart) {
            $path = [System.Text.Encoding]::Unicode.GetString($EntryIdBytes, $unicodeStart, $end - $unicodeStart)
        }
    } else {
        # EIDMS: ASCII path starting at offset
        $end = $offset
        while ($end -lt $EntryIdBytes.Length -and $EntryIdBytes[$end] -ne 0x00) { $end++ }
        if ($end -gt $offset) {
            $path = [System.Text.Encoding]::ASCII.GetString($EntryIdBytes, $offset, $end - $offset)
        }
    }

    # Validation: must look like a .pst file path with a drive letter.
    # Anything else (e.g. Exchange or OST) is not what we're after; return
    # null so the caller can record a "PST mapping unavailable" warning.
    if ([string]::IsNullOrWhiteSpace($path)) { return $null }
    if ($path -notmatch '^[A-Za-z]:\\') { return $null }
    if ($path -notmatch '\.pst$') { return $null }
    return $path
}

# ----------------------------------------------------------
# Walk profiles -> internet-account subkeys -> POP3 entries
# ----------------------------------------------------------
$manifestProfiles = @()
$totalPop = 0
$totalImap = 0
$totalOther = 0

$profileKeys = @()
try {
    $profileKeys = @(Get-ChildItem -LiteralPath $profilesKeyPath -ErrorAction Stop)
} catch {
    $warnings += "Failed to enumerate profiles: $($_.Exception.Message)"
}

foreach ($profKey in $profileKeys) {
    $profileName = Split-Path -Path $profKey.PSPath -Leaf
    Show-Info "[profile] $profileName"

    # Phase 2.9.2a-v2: enumerate all PSTs in this profile up front via
    # registry walk. Used as the deterministic mapping source for each
    # POP account in the same profile.
    $profilePstPaths = @(Get-PstPathsFromProfile -ProfileKeyPath $profKey.PSPath)
    if ($profilePstPaths.Count -gt 0) {
        Show-Info "  PST stores found in profile: $($profilePstPaths.Count)"
        foreach ($p in $profilePstPaths) { Show-Info "    - $p" }
    } else {
        Show-Info "  PST stores found in profile: 0"
    }

    $accountsRoot = Join-Path $profKey.PSPath $INTERNET_ACCOUNT_GUID
    if (-not (Test-Path $accountsRoot)) {
        Show-Skip "  no internet-account subkey ($INTERNET_ACCOUNT_GUID) - skipping"
        continue
    }

    $accountEntries = @()
    $accountKeys = @()
    try {
        $accountKeys = @(Get-ChildItem -LiteralPath $accountsRoot -ErrorAction Stop)
    } catch {
        $warnings += "Failed to enumerate accounts in '$profileName': $($_.Exception.Message)"
        continue
    }

    foreach ($acctKey in $accountKeys) {
        $subKeyName = Split-Path -Path $acctKey.PSPath -Leaf
        # Each subkey is named like 00000001, 00000002, ...
        if ($subKeyName -notmatch '^[0-9A-Fa-f]{8}$') {
            $totalOther++
            Show-Skip "  [$subKeyName] non-account subkey - skipping"
            continue
        }

        # Open with .Net registry API so we can read raw byte[] for REG_BINARY.
        $rk = $null
        try {
            $rk = (Get-Item -LiteralPath $acctKey.PSPath -ErrorAction Stop)
        } catch {
            $warnings += "Failed to open account subkey ${profileName}/${subKeyName}: $($_.Exception.Message)"
            continue
        }

        $pop3Server = Get-RegValueString -RegKey $rk -Name 'POP3 Server'
        $imapServer = Get-RegValueString -RegKey $rk -Name 'IMAP Server'

        if ([string]::IsNullOrWhiteSpace($pop3Server) -and `
            -not [string]::IsNullOrWhiteSpace($imapServer)) {
            $totalImap++
            Show-Skip "  [$subKeyName] IMAP account - skipping (Phase A: POP3 only)"
            continue
        }
        if ([string]::IsNullOrWhiteSpace($pop3Server)) {
            $totalOther++
            Show-Skip "  [$subKeyName] no POP3/IMAP server value - skipping"
            continue
        }

        $entry = [ordered]@{
            subKey       = $subKeyName
            type         = 'pop3'
            accountName  = (Get-RegValueString -RegKey $rk -Name 'Account Name')
            displayName  = (Get-RegValueString -RegKey $rk -Name 'Display Name')
            email        = (Get-RegValueString -RegKey $rk -Name 'Email')
            replyEmail   = (Get-RegValueString -RegKey $rk -Name 'Reply E-mail')
            organization = (Get-RegValueString -RegKey $rk -Name 'Organization')
            pop3         = [ordered]@{
                server   = $pop3Server
                # Phase 2.9.0a: corrected value names verified against an
                # actual registry dump. Outlook stores these without the
                # "Name" suffix and uses "SMTP Use Auth" (not "Use Sicily")
                # for the outgoing-server-auth flag. "POP3 Use Sicily" is
                # correct as the SPA flag, even though SMTP doesn't use
                # the "Sicily" name for auth.
                userName = (Get-RegValueString -RegKey $rk -Name 'POP3 User')
                port     = (Get-RegValueDword  -RegKey $rk -Name 'POP3 Port')
                useSSL   = (Get-RegValueDword  -RegKey $rk -Name 'POP3 Use SSL')
                useSPA   = (Get-RegValueDword  -RegKey $rk -Name 'POP3 Use Sicily')
            }
            smtp         = [ordered]@{
                server     = (Get-RegValueString -RegKey $rk -Name 'SMTP Server')
                userName   = (Get-RegValueString -RegKey $rk -Name 'SMTP User')
                port       = (Get-RegValueDword  -RegKey $rk -Name 'SMTP Port')
                useSSL     = (Get-RegValueDword  -RegKey $rk -Name 'SMTP Use SSL')
                useAuth    = (Get-RegValueDword  -RegKey $rk -Name 'SMTP Use Auth')
                authMethod = (Get-RegValueDword  -RegKey $rk -Name 'SMTP Auth Method')
            }
            # Diagnostic: whether a password blob is stored for this
            # account. We never read the actual encrypted bytes (DPAPI,
            # not portable). Just record presence as a hint to the
            # operator that a re-prompt will be needed at restore time.
            passwordStored = [ordered]@{
                pop3 = ($null -ne (Get-RegValueRaw -RegKey $rk -Name 'POP3 Password'))
                smtp = ($null -ne (Get-RegValueRaw -RegKey $rk -Name 'SMTP Password'))
            }
        }

        # Phase 2.9.2a-v2: PST detection. Two-pronged strategy:
        #   1. Primary: profile subkey walk (enumerate '001f6700' values).
        #      If exactly 1 PST in this profile, map it to this account.
        #   2. Secondary fallback: parse the 'Delivery Store EntryID'
        #      REG_BINARY via the documented EIDMSW format. Only the
        #      sample wrapped PST provider uses that format, so this
        #      typically returns $null on production Outlook — but kept
        #      as a defense-in-depth path.
        $entryIdBytes = Get-RegValueRaw -RegKey $rk -Name 'Delivery Store EntryID'
        $pstPath = $null
        $pstStatus = 'unavailable'
        $pstReason = $null
        $detectionMethod = $null
        $pstCandidates = @($profilePstPaths)

        if ($pstCandidates.Count -eq 1) {
            $pstPath = $pstCandidates[0]
            $detectionMethod = 'profile-subkey-walk'
            $pstStatus = if (Test-Path -LiteralPath $pstPath) { 'present' } else { 'path-only' }
        }
        elseif ($pstCandidates.Count -gt 1) {
            # Multiple PSTs in profile — try the EntryID parser as a
            # disambiguator (may still fail on production Outlook).
            $parsed = if ($null -ne $entryIdBytes) {
                Get-PstPathFromDeliveryStoreEntryId -EntryIdBytes ([byte[]]$entryIdBytes)
            } else { $null }
            if ($null -ne $parsed -and $pstCandidates -contains $parsed) {
                $pstPath = $parsed
                $detectionMethod = 'entryid-parse-confirmed-by-walk'
                $pstStatus = if (Test-Path -LiteralPath $pstPath) { 'present' } else { 'path-only' }
            } else {
                $pstReason = "$($pstCandidates.Count) PST stores in profile, EntryID parse could not disambiguate"
                $detectionMethod = 'profile-subkey-walk'
            }
        }
        else {
            # No PST stores found via walk — last try with EntryID parse.
            $parsed = if ($null -ne $entryIdBytes) {
                Get-PstPathFromDeliveryStoreEntryId -EntryIdBytes ([byte[]]$entryIdBytes)
            } else { $null }
            if ($null -ne $parsed) {
                $pstPath = $parsed
                $detectionMethod = 'entryid-parse-only'
                $pstStatus = if (Test-Path -LiteralPath $pstPath) { 'present' } else { 'path-only' }
            } else {
                $pstReason = 'no PST store subkey in profile, EntryID parse also failed'
                $detectionMethod = 'profile-subkey-walk'
            }
        }

        if ($null -eq $pstPath -and $null -ne $entryIdBytes) {
            # Diagnostic: dump up to 128 bytes of the EntryID for future
            # offline analysis of the production Outlook EntryID format.
            $maxIdx = [Math]::Min(127, $entryIdBytes.Length - 1)
            $hex = ($entryIdBytes[0..$maxIdx] |
                    ForEach-Object { '{0:X2}' -f $_ }) -join ' '
            $warnings += "PST mapping unavailable for ${profileName}/${subKeyName}: $pstReason. " +
                         "Profile PSTs found: $($pstCandidates.Count). " +
                         "EntryID head (max 128 bytes): $hex"
        }

        $entry.pst = [ordered]@{
            sourcePath       = $pstPath
            sourceFileName   = if ($pstPath) { Split-Path $pstPath -Leaf } else { $null }
            detectionMethod  = $detectionMethod
            detectionStatus  = $pstStatus
            detectionReason  = $pstReason
            profileCandidates = @($pstCandidates)
        }

        Show-Success "  [$subKeyName] $($entry.accountName)  <$($entry.email)>  pop=$($entry.pop3.server):$($entry.pop3.port)"
        if ($pstPath) {
            Show-Info "             pst=$pstPath  ($pstStatus)"
        } else {
            Show-Warning "             pst=(unavailable: $pstReason)"
        }
        $accountEntries += $entry
        $totalPop++
    }

    $manifestProfiles += [ordered]@{
        name     = $profileName
        accounts = @($accountEntries)
    }
}

# ----------------------------------------------------------
# Build manifest (fabriq-outlook-pop-backup schemaVersion=1)
# ----------------------------------------------------------
$hwUid = $null
try { $hwUid = (Get-CimInstance -ClassName Win32_ComputerSystemProduct -ErrorAction Stop |
                Select-Object -First 1).UUID } catch { }
$osArch = if ($env:PROCESSOR_ARCHITECTURE -eq 'ARM64') { 'arm64' }
          elseif ([Environment]::Is64BitOperatingSystem) { 'amd64' }
          else { 'x86' }
$osVersion = [System.Environment]::OSVersion.Version.ToString()
$kernelVersionFile = Join-Path $FabriqRoot 'kernel\KERNEL_VERSION'
$kernelVersion = if (Test-Path $kernelVersionFile) { (Get-Content $kernelVersionFile -Raw).Trim() } else { 'unknown' }
$moduleVersionFile = Join-Path $BackuperRoot 'VERSION'
$moduleVersion = if (Test-Path $moduleVersionFile) { (Get-Content $moduleVersionFile -Raw).Trim() } else { 'unknown' }

$sourceUserName = $null
if (-not [string]::IsNullOrWhiteSpace($sourceUserProfilePath)) {
    try { $sourceUserName = Split-Path $sourceUserProfilePath -Leaf } catch { }
}

$profileCountWithPop = @($manifestProfiles | Where-Object { $_.accounts.Count -gt 0 }).Count

$manifest = [ordered]@{
    schemaVersion       = 1
    manifestType        = 'fabriq-outlook-pop-backup'
    backupVersion       = $moduleVersion
    fabriqKernelVersion = $kernelVersion
    collectedAt         = (Get-Date).ToString('yyyy-MM-ddTHH:mm:sszzz')
    computerName        = $OldPcName
    hardwareUniqueId    = $hwUid
    osVersion           = $osVersion
    osArch              = $osArch
    outlookVersion      = $outlookVersion
    sourceUser          = [ordered]@{
        profilePath = $sourceUserProfilePath
        userName    = $sourceUserName
        sid         = $hkcuInfo.SID
        redirected  = [bool]$hkcuInfo.Redirected
    }
    counts              = [ordered]@{
        profile           = @($manifestProfiles).Count
        profileWithPop    = $profileCountWithPop
        popAccount        = $totalPop
        imapAccountSkipped= $totalImap
        otherSkipped      = $totalOther
    }
    items               = [ordered]@{ profiles = @($manifestProfiles) }
    warnings            = @($warnings)
    notes               = @(
        'Passwords are DPAPI-encrypted per-user/per-machine and excluded from this manifest.',
        'IMAP / Exchange accounts are intentionally skipped in Phase A.',
        'Restore is not yet implemented (Phase B). Use this manifest to manually craft a PRF and run: outlook.exe /importprf <prf>'
    )
}

$manifestPath = Join-Path $sectionDir 'manifest.json'
$manifest | ConvertTo-Json -Depth 8 | Out-File -FilePath $manifestPath -Encoding UTF8 -Force

$sw.Stop()

$status = if ($totalPop -eq 0) { 'Skipped' } else { 'Success' }
Show-Info "POP accounts captured: $totalPop  (IMAP skipped: $totalImap, other skipped: $totalOther)"

return [PSCustomObject]@{
    Status               = $status
    ElapsedMs            = [int]$sw.ElapsedMilliseconds
    Summary              = [ordered]@{
        profileCount    = @($manifestProfiles).Count
        popAccountCount = $totalPop
        imapSkipped     = $totalImap
        otherSkipped    = $totalOther
        outlookVersion  = $outlookVersion
    }
    Warnings             = $warnings
    ExternalOutputDir    = $null
    ExternalManifestPath = $null
    InternalSectionDir   = $sectionDir
    InternalManifestPath = $manifestPath
}
