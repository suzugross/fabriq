# ========================================
# Userdata Backup
# ========================================
# [PURPOSE]
# Back up arbitrary files / directories on this PC into a portable
# backup folder that a companion restore module can replay on the
# same or different PC. Same backup-folder convention as printer_backup:
#   backup/<OldPCname>/<yyyy_MM_dd_HHmmss>/
#
# [NOTES]
# - Requires administrator privileges (robocopy /B backup mode, ACL).
# - Engine: robocopy (handles long paths, locked files, retries).
# - Each enabled entry in userdata_backup_list.csv becomes one entry/
#   subfolder in the backup, captured under entries/<NN>/data/.
# - manifest.json (fabriq-userdata-backup schemaVersion=1) is the
#   single source of truth that the restore module consumes.
# ========================================

Write-Host ""
Show-Separator
Write-Host "Userdata Backup" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# ========================================
# Step 1: Load CSV
# ========================================
$csvPath = Join-Path $PSScriptRoot "userdata_backup_list.csv"

$entries = Import-ModuleCsv -Path $csvPath -FilterEnabled `
    -RequiredColumns @(
        "Enabled",
        "SourcePath",
        "Recurse",
        "ExcludePattern",
        "OnConflict",
        "IncludeAcl"
    )

if ($null -eq $entries) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load userdata_backup_list.csv")
}
$entries = @($entries)
if ($entries.Count -eq 0) {
    return (New-ModuleResult -Status "Skipped" -Message "No enabled entries in userdata_backup_list.csv")
}


# ========================================
# Step 2: Prerequisite Checks
# ========================================
if (-not (Test-AdminPrivilege)) {
    Show-Error "Administrator privileges are required (robocopy /B backup mode)"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Administrator privileges required")
}

# Verify robocopy is available (always present on Windows since Vista)
$robocopyExe = Get-Command robocopy.exe -ErrorAction SilentlyContinue
if ($null -eq $robocopyExe) {
    Show-Error "robocopy.exe not found in PATH"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "robocopy.exe not found")
}


# ========================================
# Step 3: Resolve and Scan Entries
# ========================================
Show-Info "Scanning entries..."

# Resolve PC name (printer_backup uses the same convention)
$pcName = if (-not [string]::IsNullOrWhiteSpace($env:SELECTED_OLD_PCNAME)) {
    $env:SELECTED_OLD_PCNAME.Trim()
} else {
    $env:COMPUTERNAME
}

function Resolve-EntryPath {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $null }
    $expanded = [Environment]::ExpandEnvironmentVariables($Path.Trim())
    if ([string]::IsNullOrWhiteSpace($expanded)) { return $null }
    return $expanded
}

function Parse-ExcludePatterns {
    # Splits "foo;Cache/;Bar" into file patterns vs dir patterns.
    # Trailing slash marks dir patterns (consumed by robocopy /XD).
    param([string]$Raw)
    $result = @{ Files = @(); Dirs = @() }
    if ([string]::IsNullOrWhiteSpace($Raw)) { return $result }
    foreach ($tok in $Raw.Split(';')) {
        $t = $tok.Trim()
        if ([string]::IsNullOrWhiteSpace($t)) { continue }
        if ($t.EndsWith('/') -or $t.EndsWith('\')) {
            $result.Dirs += $t.TrimEnd('/', '\')
        } else {
            $result.Files += $t
        }
    }
    return $result
}

$planned = @()
$entryIndex = 0
foreach ($e in $entries) {
    $entryIndex++
    $resolved = Resolve-EntryPath -Path $e.SourcePath
    $existsAsDir  = $false
    $existsAsFile = $false
    if (-not [string]::IsNullOrWhiteSpace($resolved)) {
        $existsAsDir  = (Test-Path -LiteralPath $resolved -PathType Container)
        $existsAsFile = (Test-Path -LiteralPath $resolved -PathType Leaf)
    }
    $planned += [PSCustomObject]@{
        Index           = $entryIndex
        Id              = ("{0:D2}" -f $entryIndex)
        SourcePath      = $e.SourcePath
        ResolvedPath    = $resolved
        ExistsAsDir     = $existsAsDir
        ExistsAsFile    = $existsAsFile
        Recurse         = ($e.Recurse -eq "1")
        ExcludePattern  = if ($null -ne $e.ExcludePattern) { $e.ExcludePattern.Trim() } else { "" }
        OnConflict      = if ([string]::IsNullOrWhiteSpace($e.OnConflict)) { "skip" } else { $e.OnConflict.Trim().ToLower() }
        IncludeAcl      = ($e.IncludeAcl -eq "1")
        Description     = if ($null -ne $e.Description) { $e.Description } else { "" }
    }
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "Backup Plan" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "  PC name (folder):   $pcName" -ForegroundColor White
if ($pcName -ne $env:COMPUTERNAME) {
    Write-Host "  Computer name:      $env:COMPUTERNAME (current Windows hostname)" -ForegroundColor DarkGray
}
Write-Host "  Entries:            $($planned.Count)" -ForegroundColor White
Write-Host ""
foreach ($p in $planned) {
    $kind = if ($p.ExistsAsDir) { "[DIR ]" } elseif ($p.ExistsAsFile) { "[FILE]" } else { "[MISSING]" }
    $aclTag = if ($p.IncludeAcl) { " /ACL" } else { "" }
    $color = if ($kind -eq "[MISSING]") { "Red" } else { "White" }
    Write-Host ("  [{0}] {1} {2}{3}" -f $p.Id, $kind, $p.ResolvedPath, $aclTag) -ForegroundColor $color
    if ($p.Description) {
        Write-Host ("         desc: {0}" -f $p.Description) -ForegroundColor DarkGray
    }
    if ($p.ExcludePattern) {
        Write-Host ("         exclude: {0}" -f $p.ExcludePattern) -ForegroundColor DarkGray
    }
}
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

$missing = @($planned | Where-Object { -not $_.ExistsAsDir -and -not $_.ExistsAsFile })
if ($missing.Count -gt 0) {
    Show-Warning "$($missing.Count) entry(ies) have missing source paths (will be skipped)"
    Write-Host ""
}


# ========================================
# Step 4: Confirm
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Proceed with userdata backup?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""


# ========================================
# Step 5: Execute Backup
# ========================================
$timestamp = Get-Date -Format "yyyy_MM_dd_HHmmss"
$backupDir = Join-Path (Join-Path (Join-Path $PSScriptRoot "backup") $pcName) $timestamp

try {
    $null = New-Item -ItemType Directory -Path $backupDir -Force -ErrorAction Stop
    $null = New-Item -ItemType Directory -Path (Join-Path $backupDir "entries") -Force -ErrorAction Stop
}
catch {
    Show-Error "Failed to create backup directory: $_"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Failed to create backup dir")
}

Show-Info "Backup target: $backupDir"
Write-Host ""

$warnings = @()
$manifestEntries = @()
$successCount = 0
$skipCount    = 0
$failCount    = 0

foreach ($p in $planned) {
    $entryDir  = Join-Path (Join-Path $backupDir "entries") $p.Id
    $dataDir   = Join-Path $entryDir "data"
    $logFile   = Join-Path $entryDir "entry_log.txt"
    $entryStatus = "Success"
    $entryReason = $null
    $robocopyExit = $null
    $fileCount = 0
    $dirCount  = 0
    $byteCount = 0

    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "[$($p.Id)] $($p.ResolvedPath)" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor White

    # Skip if source path is missing
    if (-not $p.ExistsAsDir -and -not $p.ExistsAsFile) {
        Show-Skip "Source path does not exist: $($p.ResolvedPath)"
        $warnings += "Missing source: $($p.ResolvedPath)"
        $entryStatus = "Skipped"
        $entryReason = "Source path not found"
        $skipCount++
        $manifestEntries += [PSCustomObject]@{
            id              = $p.Id
            sourcePath      = $p.SourcePath
            resolvedPath    = $p.ResolvedPath
            isDirectory     = $false
            recurse         = $p.Recurse
            excludePattern  = $p.ExcludePattern
            onConflict      = $p.OnConflict
            includeAcl      = $p.IncludeAcl
            fileCount       = 0
            dirCount        = 0
            byteCount       = 0
            backupSubpath   = $null
            robocopyExitCode = $null
            status          = $entryStatus
            reason          = $entryReason
        }
        Write-Host ""
        continue
    }

    try {
        $null = New-Item -ItemType Directory -Path $dataDir -Force -ErrorAction Stop
    }
    catch {
        Show-Error "Failed to create entry data dir: $dataDir"
        $warnings += "Entry $($p.Id) dir creation failed: $($_.Exception.Message)"
        $entryStatus = "Failed"
        $entryReason = "Could not create entry data directory"
        $failCount++
        # Record the failure in the manifest so it stays visible to the
        # verification step and the evidence trail (a bare continue made
        # this failure vanish from both).
        $manifestEntries += [PSCustomObject]@{
            id              = $p.Id
            sourcePath      = $p.SourcePath
            resolvedPath    = $p.ResolvedPath
            isDirectory     = [bool]$p.ExistsAsDir
            recurse         = $p.Recurse
            excludePattern  = $p.ExcludePattern
            onConflict      = $p.OnConflict
            includeAcl      = $p.IncludeAcl
            fileCount       = 0
            dirCount        = 0
            byteCount       = 0
            backupSubpath   = $null
            robocopyExitCode = $null
            status          = $entryStatus
            reason          = $entryReason
        }
        Write-Host ""
        continue
    }

    # Build robocopy arguments
    $excludes = Parse-ExcludePatterns -Raw $p.ExcludePattern
    $copyFlag = if ($p.IncludeAcl) { '/COPYALL' } else { '/COPY:DAT' }

    if ($p.ExistsAsDir) {
        # Directory backup
        # /E   : copy subdirs including empty
        # /B   : backup mode (read-protected files)
        # /R:1 /W:1 : minimal retry/wait
        # /NP  : no progress percentage (cleaner log)
        # /NDL /NS /NC : suppress per-file noise
        $rcArgs = @($p.ResolvedPath, $dataDir, '/E', $copyFlag, '/B', '/R:1', '/W:1', '/NP')
        if (-not $p.Recurse) {
            # Replace /E with /LEV:1 (top-level only, no subdirs)
            $rcArgs = @($p.ResolvedPath, $dataDir, '/LEV:1', $copyFlag, '/B', '/R:1', '/W:1', '/NP')
        }
        if ($excludes.Files.Count -gt 0) {
            $rcArgs += '/XF'
            $rcArgs += $excludes.Files
        }
        if ($excludes.Dirs.Count -gt 0) {
            $rcArgs += '/XD'
            $rcArgs += $excludes.Dirs
        }
    }
    else {
        # Single-file backup: invoke robocopy on parent dir with single-file filter
        $parentDir = Split-Path -Path $p.ResolvedPath -Parent
        $fileName  = Split-Path -Path $p.ResolvedPath -Leaf
        $rcArgs = @($parentDir, $dataDir, $fileName, $copyFlag, '/B', '/R:1', '/W:1', '/NP', '/NDL', '/NS', '/NC')
    }

    Show-Info "  robocopy $($rcArgs -join ' ')"
    & robocopy.exe @rcArgs > $logFile 2>&1
    $robocopyExit = $LASTEXITCODE

    # robocopy exit codes (bitwise):
    #   0  = no files copied (already in sync)
    #   1  = files copied successfully
    #   2  = extra files in dest (n/a for fresh dst)
    #   4  = mismatched files/dirs
    #   8  = copy failures
    #   16 = serious error
    if ($robocopyExit -ge 8) {
        Show-Error "  robocopy reported failures (exit code $robocopyExit) - see $logFile"
        $warnings += "robocopy exit $robocopyExit for entry $($p.Id) ($($p.ResolvedPath))"
        $entryStatus = "Failed"
        $entryReason = "robocopy exit $robocopyExit"
        $failCount++
    }
    elseif ($robocopyExit -ge 4) {
        Show-Warning "  robocopy reported mismatches (exit code $robocopyExit) - see $logFile"
        $warnings += "robocopy mismatch (exit $robocopyExit) for entry $($p.Id)"
        $entryStatus = "Partial"
        $entryReason = "robocopy mismatch (exit $robocopyExit)"
        $successCount++
    }
    else {
        Show-Success "  copied (robocopy exit $robocopyExit)"
        $successCount++
    }

    # Tally
    $copied = @(Get-ChildItem -Path $dataDir -Recurse -ErrorAction SilentlyContinue)
    $fileCount = @($copied | Where-Object { -not $_.PSIsContainer }).Count
    $dirCount  = @($copied | Where-Object { $_.PSIsContainer }).Count
    $byteSum   = ($copied | Where-Object { -not $_.PSIsContainer } |
                  Measure-Object -Property Length -Sum).Sum
    if ($null -ne $byteSum) { $byteCount = [long]$byteSum }

    Write-Host ("  -> files={0} dirs={1} bytes={2:N0}" -f $fileCount, $dirCount, $byteCount) -ForegroundColor DarkGray

    $manifestEntries += [PSCustomObject]@{
        id              = $p.Id
        sourcePath      = $p.SourcePath
        resolvedPath    = $p.ResolvedPath
        isDirectory     = [bool]$p.ExistsAsDir
        recurse         = $p.Recurse
        excludePattern  = $p.ExcludePattern
        onConflict      = $p.OnConflict
        includeAcl      = $p.IncludeAcl
        fileCount       = $fileCount
        dirCount        = $dirCount
        byteCount       = $byteCount
        backupSubpath   = "entries/$($p.Id)/data"
        robocopyExitCode = $robocopyExit
        status          = $entryStatus
        reason          = $entryReason
    }

    Write-Host ""
}


# ========================================
# Step 5.5: Build manifest.json
# ========================================
Show-Info "Writing manifest.json..."

$hwUid = $null
try {
    $hwUid = (Get-CimInstance -ClassName Win32_ComputerSystemProduct -ErrorAction Stop |
              Select-Object -First 1).UUID
}
catch { }

$osArch = if ($env:PROCESSOR_ARCHITECTURE -eq 'ARM64') {
    'arm64'
} elseif ([Environment]::Is64BitOperatingSystem) {
    'amd64'
} else {
    'x86'
}

$osVersion = [System.Environment]::OSVersion.Version.ToString()

$kernelVersionFile = Join-Path $PSScriptRoot "..\..\..\kernel\KERNEL_VERSION"
$kernelVersion = if (Test-Path $kernelVersionFile) {
    (Get-Content $kernelVersionFile -Raw).Trim()
} else {
    "unknown"
}

$moduleVersionFile = Join-Path $PSScriptRoot "VERSION"
$moduleVersion = if (Test-Path $moduleVersionFile) {
    (Get-Content $moduleVersionFile -Raw).Trim()
} else {
    "unknown"
}

$totalBytes = ($manifestEntries | Measure-Object -Property byteCount -Sum).Sum
if ($null -eq $totalBytes) { $totalBytes = 0 }
$totalFiles = ($manifestEntries | Measure-Object -Property fileCount -Sum).Sum
if ($null -eq $totalFiles) { $totalFiles = 0 }
$totalDirs  = ($manifestEntries | Measure-Object -Property dirCount -Sum).Sum
if ($null -eq $totalDirs)  { $totalDirs  = 0 }
$missingCount = @($manifestEntries | Where-Object { $_.status -eq 'Skipped' -and $_.reason -eq 'Source path not found' }).Count

$manifest = [ordered]@{
    schemaVersion        = 1
    manifestType         = "fabriq-userdata-backup"
    backupVersion        = $moduleVersion
    fabriqKernelVersion  = $kernelVersion
    collectedAt          = (Get-Date).ToString("yyyy-MM-ddTHH:mm:sszzz")
    computerName         = $pcName
    hardwareUniqueId     = $hwUid
    osVersion            = $osVersion
    osArch               = $osArch
    counts               = [ordered]@{
        entry          = $manifestEntries.Count
        file           = [long]$totalFiles
        dir            = [long]$totalDirs
        missingSource  = $missingCount
    }
    sizes                = [ordered]@{
        totalBytes = [long]$totalBytes
    }
    items                = [ordered]@{
        entries = @($manifestEntries)
    }
    warnings             = @($warnings)
}

$manifestPath = Join-Path $backupDir "manifest.json"
try {
    $manifest | ConvertTo-Json -Depth 8 | Out-File -FilePath $manifestPath -Encoding UTF8 -Force
    Show-Success "manifest.json written"
}
catch {
    Show-Error "Failed to write manifest.json: $_"
    $failCount++
}

# _restore_notes.txt
$notes = @"
Userdata Backup - Restore Notes
================================
Backup taken    : $((Get-Date).ToString("yyyy-MM-dd HH:mm:ss"))
Source PC       : $pcName$(if ($pcName -ne $env:COMPUTERNAME) { " (from hostlist OldPCname)" } else { "" })
Computer name   : $env:COMPUTERNAME$(if ($pcName -ne $env:COMPUTERNAME) { " (actual Windows hostname at backup time)" } else { "" })
OS / Arch       : $osVersion / $osArch
Total entries   : $($manifestEntries.Count)
Total files     : $totalFiles
Total bytes     : $totalBytes

The companion userdata_restore.ps1 reads manifest.json from this folder.

Warnings:
$(if ($warnings.Count -gt 0) { ($warnings | ForEach-Object { "  - $_" }) -join "`n" } else { '  (none)' })
"@
$notes | Out-File -FilePath (Join-Path $backupDir "_restore_notes.txt") -Encoding UTF8 -Force


# ========================================
# Step 5.6: Post-Apply Verification
# ========================================
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Post-Apply Verification" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

$verifyFail = 0
$requiredFiles = @("manifest.json", "_restore_notes.txt")
foreach ($rf in $requiredFiles) {
    $path = Join-Path $backupDir $rf
    if (-not (Test-Path $path)) {
        Write-Host "  [VERIFY FAILED] $rf - missing" -ForegroundColor Red
        $verifyFail++
        continue
    }
    $sz = (Get-Item $path).Length
    if ($sz -eq 0) {
        Write-Host "  [VERIFY FAILED] $rf - zero bytes" -ForegroundColor Red
        $verifyFail++
        continue
    }
    Write-Host "  [VERIFIED] $rf ($sz bytes)" -ForegroundColor Green
}

# Re-read manifest and check schemaVersion / manifestType
try {
    $reread = Get-Content -Path $manifestPath -Raw | ConvertFrom-Json
    if ($reread.manifestType -eq "fabriq-userdata-backup" -and $reread.schemaVersion -eq 1) {
        Write-Host "  [VERIFIED] manifest.json schemaVersion=1, manifestType=fabriq-userdata-backup" -ForegroundColor Green
    } else {
        Write-Host "  [VERIFY FAILED] manifest.json - unexpected type/schemaVersion" -ForegroundColor Red
        $verifyFail++
    }
}
catch {
    Write-Host "  [VERIFY FAILED] manifest.json - parse error: $_" -ForegroundColor Red
    $verifyFail++
}

# Per-entry data folder presence (for non-skipped entries). A Failed
# entry must never print [VERIFIED]: its data dir was pre-created by
# this module, so bare Test-Path would pass even when robocopy copied
# nothing (exit >= 8).
foreach ($me in $manifestEntries) {
    if ($me.status -eq 'Skipped') { continue }
    if ($me.status -eq 'Failed') {
        Write-Host "  [VERIFY FAILED] entry $($me.id) backup failed ($($me.reason))" -ForegroundColor Red
        $verifyFail++
        continue
    }
    if ([string]::IsNullOrWhiteSpace($me.backupSubpath)) { continue }
    $dataPath = Join-Path $backupDir $me.backupSubpath
    if (-not (Test-Path $dataPath)) {
        Write-Host "  [VERIFY FAILED] entry $($me.id) data folder missing" -ForegroundColor Red
        $verifyFail++
    } else {
        Write-Host "  [VERIFIED] entry $($me.id) ($($me.fileCount) file(s), $('{0:N0}' -f $me.byteCount) bytes)" -ForegroundColor Green
    }
}

# A backup that captured nothing must not verify as PASS - an operator
# reading PASS may wipe the source PC. All-sources-missing (wrong user
# profile, unmounted drive) lands here. Partial missing stays PASS by
# design: it is routine (e.g. a user without a Pictures folder) and is
# already visible via warnings / manifest missingSource / Skip counts.
$backedUpEntries = @($manifestEntries | Where-Object { $_.status -ne 'Skipped' })
if ($backedUpEntries.Count -eq 0) {
    Write-Host "  [VERIFY FAILED] no data was backed up (all $($manifestEntries.Count) source path(s) missing or skipped)" -ForegroundColor Red
    $verifyFail++
}

Write-Host ""
# Any hard failure (robocopy exit >= 8, entry dir creation failure) must
# force Verified=false: an operator reading PASS may wipe the source PC.
# Sibling userdata_restore uses the same formula.
$verified = ($verifyFail -eq 0 -and $failCount -eq 0)


# ========================================
# Step 6: Result Summary
# ========================================
$totalMB = [math]::Round($totalBytes / 1MB, 1)

Show-Separator
Write-Host "Userdata Backup Results" -ForegroundColor Cyan
Show-Separator
Write-Host "  Location:  $backupDir" -ForegroundColor White
Write-Host "  Entries:   $($manifestEntries.Count) ($successCount success / $skipCount skip / $failCount fail)" -ForegroundColor White
Write-Host "  Files:     $totalFiles" -ForegroundColor White
Write-Host "  Size:      $totalMB MB" -ForegroundColor White
if ($warnings.Count -gt 0) {
    Write-Host "  Warnings:  $($warnings.Count) (see manifest.json / _restore_notes.txt)" -ForegroundColor Yellow
}
if ($verified) {
    Write-Host "  Verified:  PASS" -ForegroundColor Green
} else {
    Write-Host "  Verified:  FAIL" -ForegroundColor Red
}
Show-Separator
Write-Host ""

return (New-BatchResult `
    -Success $successCount `
    -Skip $skipCount `
    -Fail $failCount `
    -Title "Userdata Backup Results" `
    -MessageSuffix " ($totalFiles files, $totalMB MB)" `
    -Verified $verified)
