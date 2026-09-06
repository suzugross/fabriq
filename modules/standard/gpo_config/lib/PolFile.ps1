# ========================================
# Registry.pol Codec - shared helper for gpo_config / gpo_backup
# ========================================
# Pure-PowerShell reader/writer for Group Policy registry policy files
# (MS-GPREG "PReg" format) plus gpt.ini version bookkeeping.
# Dot-sourced by gpo_config.ps1 and gpo_backup.ps1, and exercised
# directly by tests/modules/gpo_config/PolFile.tests.ps1.
# No external tools (LGPO.exe / PolicyFileEditor) are involved.
#
# File layout (all strings UTF-16LE, NUL-terminated):
#   "PReg" (ASCII, 4 bytes) + version 1 (int32 LE)
#   then records: [key;valueName;type;size;data]
# Special value names understood by the Registry CSE:
#   **del.<name>   delete a value (data = " " + NUL, type REG_SZ)
#   **delvals.     delete all values under the key
# ========================================

# ----------------------------------------
# Type name <-> code
# ----------------------------------------
function Get-PolTypeName {
    param([Parameter(Mandatory)][int]$Code)
    switch ($Code) {
        0  { return 'REG_NONE' }
        1  { return 'REG_SZ' }
        2  { return 'REG_EXPAND_SZ' }
        3  { return 'REG_BINARY' }
        4  { return 'REG_DWORD' }
        7  { return 'REG_MULTI_SZ' }
        11 { return 'REG_QWORD' }
        default { return "REG_TYPE_$Code" }
    }
}

function Get-PolTypeCode {
    param([Parameter(Mandatory)][string]$Name)
    switch ($Name.Trim().ToUpperInvariant()) {
        'REG_NONE'      { return 0 }
        'REG_SZ'        { return 1 }
        'REG_EXPAND_SZ' { return 2 }
        'REG_BINARY'    { return 3 }
        'REG_DWORD'     { return 4 }
        'REG_MULTI_SZ'  { return 7 }
        'REG_QWORD'     { return 11 }
        default { throw "Unsupported registry type '$Name'" }
    }
}

# ----------------------------------------
# Entry object
# ----------------------------------------
function New-PolEntry {
    param(
        [Parameter(Mandatory)][string]$Key,
        [AllowEmptyString()][string]$ValueName = '',
        [Parameter(Mandatory)][int]$Type,
        [byte[]]$Data = @()
    )
    if ($null -eq $Data) { $Data = @() }
    return [PSCustomObject]@{
        Key       = $Key
        ValueName = $ValueName
        Type      = $Type
        Data      = [byte[]]$Data
    }
}

# ----------------------------------------
# Data encoding (CSV string -> bytes) / decoding (bytes -> CSV string)
# ----------------------------------------
function ConvertTo-PolData {
    param(
        [Parameter(Mandatory)][int]$Type,
        [AllowEmptyString()][string]$Value = ''
    )
    $v = "$Value"
    switch ($Type) {
        4 {
            $t = $v.Trim()
            if ($t -match '^0[xX][0-9A-Fa-f]{1,8}$') { $u = [Convert]::ToUInt32($t.Substring(2), 16) }
            elseif ($t -match '^-\d+$')             { $u = [BitConverter]::ToUInt32([BitConverter]::GetBytes([int]$t), 0) }
            elseif ($t -match '^\d+$')              { $u = [uint32]$t }
            else { throw "'$v' is not a valid REG_DWORD value" }
            return ,([byte[]][BitConverter]::GetBytes([uint32]$u))
        }
        11 {
            $t = $v.Trim()
            if ($t -match '^0[xX][0-9A-Fa-f]{1,16}$') { $u = [Convert]::ToUInt64($t.Substring(2), 16) }
            elseif ($t -match '^-\d+$')              { $u = [BitConverter]::ToUInt64([BitConverter]::GetBytes([long]$t), 0) }
            elseif ($t -match '^\d+$')               { $u = [uint64]$t }
            else { throw "'$v' is not a valid REG_QWORD value" }
            return ,([byte[]][BitConverter]::GetBytes([uint64]$u))
        }
        { $_ -eq 1 -or $_ -eq 2 } {
            return ,([byte[]][Text.Encoding]::Unicode.GetBytes($v + [char]0))
        }
        7 {
            $parts = @()
            if ($v.Length -gt 0) { $parts = @($v -split '\|') }
            $sb = New-Object System.Text.StringBuilder
            foreach ($p in $parts) { [void]$sb.Append($p).Append([char]0) }
            [void]$sb.Append([char]0)
            return ,([byte[]][Text.Encoding]::Unicode.GetBytes($sb.ToString()))
        }
        3 {
            $hex = $v -replace '[^0-9A-Fa-f]', ''
            if ($hex.Length % 2 -ne 0) { throw "'$v' is not a valid REG_BINARY hex string" }
            $bytes = New-Object byte[] ($hex.Length / 2)
            for ($i = 0; $i -lt $bytes.Length; $i++) { $bytes[$i] = [Convert]::ToByte($hex.Substring($i * 2, 2), 16) }
            return ,([byte[]]$bytes)
        }
        0 {
            return ,([byte[]]@())
        }
        default { throw "Unsupported registry type code $Type" }
    }
}

function ConvertFrom-PolData {
    param(
        [Parameter(Mandatory)][int]$Type,
        [byte[]]$Data = @()
    )
    if ($null -eq $Data) { $Data = @() }
    switch ($Type) {
        4 {
            if ($Data.Length -ge 4) { return [string][BitConverter]::ToUInt32($Data, 0) }
        }
        11 {
            if ($Data.Length -ge 8) { return [string][BitConverter]::ToUInt64($Data, 0) }
        }
        { $_ -eq 1 -or $_ -eq 2 } {
            return ([Text.Encoding]::Unicode.GetString($Data)).TrimEnd([char]0)
        }
        7 {
            $s = ([Text.Encoding]::Unicode.GetString($Data)).TrimEnd([char]0)
            if ($s.Length -eq 0) { return '' }
            return (($s -split [char]0) -join '|')
        }
    }
    # REG_BINARY / REG_NONE / anything else: uppercase hex, no separators
    return (($Data | ForEach-Object { $_.ToString('X2') }) -join '')
}

# ----------------------------------------
# Identity / equality
# ----------------------------------------
# Identity groups a value with its own **del. marker so that "Set X" and
# "Delete X" replace each other on merge. **delvals. keeps its own identity.
function Get-PolEntryIdentity {
    param(
        [Parameter(Mandatory)][string]$Key,
        [AllowEmptyString()][string]$ValueName = ''
    )
    $base = "$ValueName"
    if ($base -match '^\*\*del\.(.*)$') { $base = $Matches[1] }
    return ($Key.Trim('\').ToLowerInvariant() + '|' + $base.ToLowerInvariant())
}

function Test-PolEntryEqual {
    param(
        [Parameter(Mandatory)]$A,
        [Parameter(Mandatory)]$B
    )
    if ($A.Key.Trim('\').ToLowerInvariant() -ne $B.Key.Trim('\').ToLowerInvariant()) { return $false }
    if ("$($A.ValueName)".ToLowerInvariant() -ne "$($B.ValueName)".ToLowerInvariant()) { return $false }
    if ([int]$A.Type -ne [int]$B.Type) { return $false }
    $da = [byte[]]$A.Data; $db = [byte[]]$B.Data
    if ($da.Length -ne $db.Length) { return $false }
    for ($i = 0; $i -lt $da.Length; $i++) { if ($da[$i] -ne $db[$i]) { return $false } }
    return $true
}

# ----------------------------------------
# CSV action -> entry
# ----------------------------------------
# Set             -> [Key;ValueName;Type;data]
# Delete          -> [Key;**del.ValueName;REG_SZ;" "]
# DeleteAllValues -> [Key;**delvals.;REG_SZ;" "]
# CreateKey       -> [Key;;REG_NONE;]  (key created with no values)
# Unmanage        -> $null (entry is removed from Registry.pol = Not Configured)
function ConvertTo-PolEntryFromAction {
    param(
        [Parameter(Mandatory)][string]$Key,
        [AllowEmptyString()][string]$ValueName = '',
        [Parameter(Mandatory)][string]$Action,
        [AllowEmptyString()][string]$Type = '',
        [AllowEmptyString()][string]$Value = ''
    )
    $marker = [byte[]][Text.Encoding]::Unicode.GetBytes(' ' + [char]0)
    switch ($Action) {
        'Set' {
            $code = Get-PolTypeCode -Name $Type
            if ($code -eq 0) { throw "REG_NONE cannot be used with Action=Set" }
            return (New-PolEntry -Key $Key -ValueName $ValueName -Type $code -Data (ConvertTo-PolData -Type $code -Value $Value))
        }
        'Delete'          { return (New-PolEntry -Key $Key -ValueName "**del.$ValueName" -Type 1 -Data $marker) }
        'DeleteAllValues' { return (New-PolEntry -Key $Key -ValueName '**delvals.' -Type 1 -Data $marker) }
        'CreateKey'       { return (New-PolEntry -Key $Key -ValueName '' -Type 0 -Data @()) }
        'Unmanage'        { return $null }
        default { throw "Unsupported Action '$Action'" }
    }
}

# Action inferred from an entry read back from Registry.pol (for backup/parse).
function Get-PolEntryAction {
    param([Parameter(Mandatory)]$Entry)
    $n = "$($Entry.ValueName)"
    if ($n -match '^\*\*del\.')            { return 'Delete' }
    if ($n -ieq '**delvals.')              { return 'DeleteAllValues' }
    if ($n -match '^\*\*')                 { return 'Unsupported' }
    if ($n -eq '' -and [int]$Entry.Type -eq 0) { return 'CreateKey' }
    return 'Set'
}

# ----------------------------------------
# File I/O
# ----------------------------------------
function Get-PolFilePath {
    param([Parameter(Mandatory)][ValidateSet('Machine','User')][string]$Scope)
    return (Join-Path $env:SystemRoot "System32\GroupPolicy\$Scope\Registry.pol")
}

function Get-GptIniPath {
    return (Join-Path $env:SystemRoot 'System32\GroupPolicy\gpt.ini')
}

function Read-PolFile {
    param([Parameter(Mandatory)][string]$Path)

    $entries = New-Object System.Collections.Generic.List[object]
    # Returns a plain array (unrolled by the pipeline): callers MUST wrap
    # the result in @() - PS 5.1 mishandles ",$genericList" (throws on an
    # empty List, and @(cmd) wraps a comma-returned array as one element).
    if (-not (Test-Path -LiteralPath $Path)) { return $entries.ToArray() }

    $b = [IO.File]::ReadAllBytes($Path)
    if ($b.Length -eq 0) { return $entries.ToArray() }
    if ($b.Length -lt 8 -or [Text.Encoding]::ASCII.GetString($b, 0, 4) -ne 'PReg') {
        throw "Not a registry policy (PReg) file: $Path"
    }
    $version = [BitConverter]::ToInt32($b, 4)
    if ($version -ne 1) { throw "Unsupported PReg version $version : $Path" }

    $pos = 8
    $u = [Text.Encoding]::Unicode

    while ($pos -lt $b.Length) {
        # '['
        if ($pos + 2 -gt $b.Length -or [BitConverter]::ToUInt16($b, $pos) -ne 0x5B) { throw "Corrupt PReg record at offset $pos (expected '['): $Path" }
        $pos += 2

        # key
        $start = $pos
        while ($true) {
            if ($pos + 2 -gt $b.Length) { throw "Truncated PReg key at offset $start : $Path" }
            if ([BitConverter]::ToUInt16($b, $pos) -eq 0) { break }
            $pos += 2
        }
        $key = $u.GetString($b, $start, $pos - $start); $pos += 2
        if ($pos + 2 -gt $b.Length -or [BitConverter]::ToUInt16($b, $pos) -ne 0x3B) { throw "Corrupt PReg record at offset $pos (expected ';' after key): $Path" }
        $pos += 2

        # value name
        $start = $pos
        while ($true) {
            if ($pos + 2 -gt $b.Length) { throw "Truncated PReg value name at offset $start : $Path" }
            if ([BitConverter]::ToUInt16($b, $pos) -eq 0) { break }
            $pos += 2
        }
        $valueName = $u.GetString($b, $start, $pos - $start); $pos += 2
        if ($pos + 2 -gt $b.Length -or [BitConverter]::ToUInt16($b, $pos) -ne 0x3B) { throw "Corrupt PReg record at offset $pos (expected ';' after value name): $Path" }
        $pos += 2

        # type ; size ;
        if ($pos + 12 -gt $b.Length) { throw "Truncated PReg record header at offset $pos : $Path" }
        $type = [BitConverter]::ToInt32($b, $pos); $pos += 4
        if ([BitConverter]::ToUInt16($b, $pos) -ne 0x3B) { throw "Corrupt PReg record at offset $pos (expected ';' after type): $Path" }
        $pos += 2
        $size = [BitConverter]::ToInt32($b, $pos); $pos += 4
        if ([BitConverter]::ToUInt16($b, $pos) -ne 0x3B) { throw "Corrupt PReg record at offset $pos (expected ';' after size): $Path" }
        $pos += 2

        # data ]
        if ($size -lt 0 -or $pos + $size + 2 -gt $b.Length) { throw "Truncated PReg data at offset $pos : $Path" }
        $data = New-Object byte[] $size
        if ($size -gt 0) { [Array]::Copy($b, $pos, $data, 0, $size) }
        $pos += $size
        if ([BitConverter]::ToUInt16($b, $pos) -ne 0x5D) { throw "Corrupt PReg record at offset $pos (expected ']'): $Path" }
        $pos += 2

        $entries.Add((New-PolEntry -Key $key -ValueName $valueName -Type $type -Data $data))
    }
    return $entries.ToArray()
}

function ConvertTo-PolFileBytes {
    param($Entries)
    $ms = New-Object System.IO.MemoryStream
    $w  = New-Object System.IO.BinaryWriter($ms)
    $u  = [Text.Encoding]::Unicode
    $w.Write([Text.Encoding]::ASCII.GetBytes('PReg'))
    $w.Write([int32]1)
    foreach ($e in @($Entries)) {
        if ($null -eq $e) { continue }
        $w.Write($u.GetBytes('['))
        $w.Write($u.GetBytes([string]$e.Key));       $w.Write([uint16]0); $w.Write($u.GetBytes(';'))
        $w.Write($u.GetBytes([string]$e.ValueName)); $w.Write([uint16]0); $w.Write($u.GetBytes(';'))
        $w.Write([int32]$e.Type);                    $w.Write($u.GetBytes(';'))
        $data = [byte[]]$e.Data
        $w.Write([int32]$data.Length);               $w.Write($u.GetBytes(';'))
        if ($data.Length -gt 0) { $w.Write($data) }
        $w.Write($u.GetBytes(']'))
    }
    $w.Flush()
    $bytes = $ms.ToArray()
    $w.Dispose(); $ms.Dispose()
    return ,$bytes
}

# Atomic replace: write a temp file next to the target, then move over it,
# so a crash mid-write never leaves a half-written Registry.pol behind.
function Write-PolFile {
    param(
        [Parameter(Mandatory)][string]$Path,
        $Entries
    )
    $dir = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $bytes = ConvertTo-PolFileBytes -Entries $Entries
    $tmp = "$Path.fabriq.tmp"
    [IO.File]::WriteAllBytes($tmp, $bytes)
    Move-Item -LiteralPath $tmp -Destination $Path -Force
}

# ----------------------------------------
# Merge (upsert) semantics
# ----------------------------------------
# Existing entries whose identity appears in -RemoveIdentities are dropped
# (order of the survivors preserved), then -Apply entries are appended.
function Merge-PolEntries {
    param(
        $Existing,
        [string[]]$RemoveIdentities = @(),
        $Apply
    )
    $drop = @{}
    foreach ($id in @($RemoveIdentities)) { if ($null -ne $id) { $drop[$id] = $true } }
    foreach ($e in @($Apply)) { if ($null -ne $e) { $drop[(Get-PolEntryIdentity -Key $e.Key -ValueName $e.ValueName)] = $true } }

    $result = New-Object System.Collections.Generic.List[object]
    foreach ($e in @($Existing)) {
        if ($null -eq $e) { continue }
        if ($drop.ContainsKey((Get-PolEntryIdentity -Key $e.Key -ValueName $e.ValueName))) { continue }
        $result.Add($e)
    }
    foreach ($e in @($Apply)) { if ($null -ne $e) { $result.Add($e) } }
    return $result.ToArray()   # plain array - callers wrap with @()
}

# ----------------------------------------
# gpt.ini version bookkeeping
# ----------------------------------------
# Version is a 32-bit counter: low 16 bits = Machine, high 16 bits = User.
# The Registry CSE and the Administrative Templates extension GUIDs must be
# listed in gPC{Machine|User}ExtensionNames or the policy is never processed.
function Update-GptIniVersion {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][ValidateSet('Machine','User')][string]$Scope
    )
    $cseGuid  = '{35378EAC-683F-11D2-A89A-00C04FBBCFA2}'
    $admxGuid = if ($Scope -eq 'Machine') { '{D02B1F72-3407-48AE-BA88-E8213C6761F1}' } else { '{D02B1F73-3407-48AE-BA88-E8213C6761F1}' }
    $extKey   = "gPC${Scope}ExtensionNames"

    $lines = @()
    if (Test-Path -LiteralPath $Path) { $lines = @(Get-Content -LiteralPath $Path) }

    $oldVersion = [uint32]0
    foreach ($l in $lines) {
        if ($l -match '^\s*Version\s*=\s*(\d+)\s*$') { $oldVersion = [uint32]$Matches[1]; break }
    }
    $low  = $oldVersion -band 0xFFFF
    $high = ($oldVersion -shr 16) -band 0xFFFF
    if ($Scope -eq 'Machine') { $low = ($low + 1) -band 0xFFFF } else { $high = ($high + 1) -band 0xFFFF }
    $newVersion = [uint32]((($high -shl 16) -bor $low) -band 0xFFFFFFFF)

    # Extension names: keep existing bracket groups, make sure the Registry CSE
    # group carries the Administrative Templates GUID.
    $extLineIndex = -1
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match "^\s*$extKey\s*=") { $extLineIndex = $i; break }
    }
    $groups = New-Object System.Collections.Generic.List[string]
    if ($extLineIndex -ge 0) {
        $existingValue = ($lines[$extLineIndex] -split '=', 2)[1]
        foreach ($m in [regex]::Matches($existingValue, '\[([^\]]*)\]')) {
            $guids = @([regex]::Matches($m.Groups[1].Value, '\{[0-9A-Fa-f-]+\}') | ForEach-Object { $_.Value.ToUpperInvariant() })
            if ($guids.Count -gt 0) { $groups.Add(($guids -join '')) }
        }
    }
    $found = $false
    for ($i = 0; $i -lt $groups.Count; $i++) {
        if ($groups[$i].StartsWith($cseGuid)) {
            $found = $true
            if ($groups[$i].IndexOf($admxGuid) -lt 0) { $groups[$i] = $groups[$i] + $admxGuid }
        }
    }
    if (-not $found) { $groups.Add($cseGuid + $admxGuid) }
    $extLine = "$extKey=" + (($groups | ForEach-Object { "[$_]" }) -join '')

    # Rebuild: [General] header first, then everything else with Version and
    # the extension line replaced (or appended when missing).
    $out = New-Object System.Collections.Generic.List[string]
    $out.Add('[General]')
    $versionWritten = $false; $extWritten = $false
    foreach ($l in $lines) {
        if ($l -match '^\s*\[') { continue }
        if ($l -match '^\s*Version\s*=') { if (-not $versionWritten) { $out.Add("Version=$newVersion"); $versionWritten = $true }; continue }
        if ($l -match "^\s*$extKey\s*=") { if (-not $extWritten) { $out.Add($extLine); $extWritten = $true }; continue }
        if ("$l".Trim().Length -gt 0) { $out.Add($l) }
    }
    if (-not $extWritten)     { $out.Add($extLine) }
    if (-not $versionWritten) { $out.Add("Version=$newVersion") }

    $dir = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    [IO.File]::WriteAllLines($Path, [string[]]$out)

    return [PSCustomObject]@{ Old = $oldVersion; New = $newVersion }
}

# ----------------------------------------
# Registry read-back (post-apply verification)
# ----------------------------------------
# Compares the live registry with what the policy entry demands.
# Returns $true when the registry already reflects the entry.
function Test-PolEntryRegistryState {
    param(
        [Parameter(Mandatory)]$Entry,
        [Parameter(Mandatory)][string]$Action,
        [Parameter(Mandatory)][ValidateSet('HKLM','HKCU')][string]$Hive
    )
    $root = if ($Hive -eq 'HKLM') { [Microsoft.Win32.Registry]::LocalMachine } else { [Microsoft.Win32.Registry]::CurrentUser }
    $sub = $null
    try {
        $sub = $root.OpenSubKey($Entry.Key.Trim('\'), $false)
        switch ($Action) {
            'CreateKey'       { return ($null -ne $sub) }
            'DeleteAllValues' { if ($null -eq $sub) { return $true }; return (@($sub.GetValueNames()).Count -eq 0) }
            'Delete' {
                if ($null -eq $sub) { return $true }
                $name = "$($Entry.ValueName)" -replace '^\*\*del\.', ''
                return (@($sub.GetValueNames()) -notcontains $name)
            }
            'Set' {
                if ($null -eq $sub) { return $false }
                $name = [string]$Entry.ValueName
                if (@($sub.GetValueNames()) -notcontains $name) { return $false }
                $kind = $sub.GetValueKind($name)
                $raw  = $sub.GetValue($name, $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
                $expected = ConvertFrom-PolData -Type $Entry.Type -Data $Entry.Data
                switch ([int]$Entry.Type) {
                    4  { if ($kind -ne 'DWord') { return $false }; return ([BitConverter]::ToUInt32([BitConverter]::GetBytes([int]$raw), 0) -eq [uint32]$expected) }
                    11 { if ($kind -ne 'QWord') { return $false }; return ([BitConverter]::ToUInt64([BitConverter]::GetBytes([long]$raw), 0) -eq [uint64]$expected) }
                    1  { if ($kind -ne 'String') { return $false }; return ([string]$raw -eq $expected) }
                    2  { if ($kind -ne 'ExpandString') { return $false }; return ([string]$raw -eq $expected) }
                    7  { if ($kind -ne 'MultiString') { return $false }; return ((@($raw) -join '|') -eq $expected) }
                    3  { if ($kind -ne 'Binary') { return $false }; return (((@($raw) | ForEach-Object { $_.ToString('X2') }) -join '') -eq $expected) }
                    default { return $false }
                }
            }
            default { return $false }
        }
    }
    catch {
        return $false
    }
    finally {
        if ($null -ne $sub) { $sub.Close() }
    }
}
