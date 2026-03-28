# ========================================
# ACL Hybrid Backup Script
# ========================================
# Backs up directory ACLs using a hybrid approach:
# 1. Full recursive icacls /save /T for the entire tree
# 2. Individual icacls /save /T for subdirectories with inheritance disabled
#
# [NOTES]
# - Administrator privileges required
# - Symbolic links and junctions are excluded from scanning
# - Backup files are stored under backup/{Id}_{SafeName}/
# ========================================

Write-Host ""
Show-Separator
Write-Host "ACL Hybrid Backup" -ForegroundColor Cyan
Show-Separator
Write-Host ""


# ========================================
# Local helper functions
# ========================================
function Get-PathHash {
    param([string]$RelativePath)
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($RelativePath.ToLower())
    $sha = [System.Security.Cryptography.SHA256]::Create()
    $hash = $sha.ComputeHash($bytes)
    return ($hash[0..7] | ForEach-Object { $_.ToString("x2") }) -join ""
}

function Get-SafeDirectoryName {
    param([string]$Path)
    $leaf = Split-Path $Path -Leaf
    return ($leaf -replace '[\\/:*?"<>|]', '_')
}


# ========================================
# Step 1: CSV reading
# ========================================
$csvPath = Join-Path $PSScriptRoot "acl_list.csv"

$enabledItems = Import-ModuleCsv -Path $csvPath -FilterEnabled `
    -RequiredColumns @("Enabled", "Id", "TargetPath")

if ($null -eq $enabledItems) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load acl_list.csv")
}
if ($enabledItems.Count -eq 0) {
    return (New-ModuleResult -Status "Skipped" -Message "No enabled entries")
}

# Expand environment variables in TargetPath
foreach ($item in $enabledItems) {
    $item.TargetPath = Expand-UserEnvironmentVariables $item.TargetPath
}


# ========================================
# Step 2: Pre-flight check
# ========================================
if (-not (Test-AdminPrivilege)) {
    Show-Error "Administrator privileges are required."
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Administrator privileges required")
}

$icaclsCmd = Get-Command "icacls.exe" -ErrorAction SilentlyContinue
if (-not $icaclsCmd) {
    Show-Error "icacls.exe not found on this system."
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "icacls.exe not found")
}


# ========================================
# Step 3: Pre-execution display (dry-run)
# ========================================
$backupBaseDir = Join-Path $PSScriptRoot "backup"

Write-Host "========================================" -ForegroundColor Yellow
Write-Host "ACL Backup Targets" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

# Pre-scan: collect non-inherited directory info for each target
$scanResults = @{}

foreach ($item in $enabledItems) {
    $displayName = if ($item.Description) { $item.Description } else { $item.TargetPath }
    $safeName = Get-SafeDirectoryName $item.TargetPath
    $backupDir = Join-Path $backupBaseDir "$($item.Id)_$safeName"

    # Target path existence check
    if (-not (Test-Path $item.TargetPath)) {
        Write-Host "  [NOT FOUND] $displayName" -ForegroundColor Red
        Write-Host "    Path: $($item.TargetPath)" -ForegroundColor DarkGray
        Write-Host ""
        continue
    }

    # Scan for non-inherited subdirectories
    $totalDirs = 0
    $nonInheritedDirs = @()
    $scanWarnings = 0

    try {
        $subdirs = @(Get-ChildItem -Path $item.TargetPath -Directory -Recurse -Attributes !ReparsePoint -ErrorAction SilentlyContinue)
        $totalDirs = $subdirs.Count

        foreach ($subdir in $subdirs) {
            try {
                $acl = Get-Acl -Path $subdir.FullName -ErrorAction Stop
                if ($acl.AreAccessRulesProtected) {
                    $nonInheritedDirs += $subdir
                }
            }
            catch {
                $scanWarnings++
            }
        }
    }
    catch {
        Show-Warning "Failed to scan subdirectories: $_"
    }

    $scanResults[$item.Id] = $nonInheritedDirs

    # Display status
    $backupExists = Test-Path $backupDir
    $statusLabel = if ($backupExists) { "[OVERWRITE]" } else { "[NEW]" }
    $statusColor = if ($backupExists) { "Yellow" } else { "Green" }

    Write-Host "  $statusLabel $displayName" -ForegroundColor $statusColor
    Write-Host "    Path: $($item.TargetPath)" -ForegroundColor DarkGray
    Write-Host "    Subdirectories: $totalDirs total, $($nonInheritedDirs.Count) non-inherited" -ForegroundColor DarkGray
    Write-Host "    Backup: $backupDir" -ForegroundColor DarkGray
    if ($scanWarnings -gt 0) {
        Write-Host "    Scan warnings: $scanWarnings (access denied)" -ForegroundColor DarkYellow
    }
    Write-Host ""
}

Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""


# ========================================
# Step 4: Confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Execute ACL backup for the above targets?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""


# ========================================
# Step 5: Execution loop
# ========================================
$successCount = 0
$skipCount    = 0
$failCount    = 0

foreach ($item in $enabledItems) {
    $displayName = if ($item.Description) { $item.Description } else { $item.TargetPath }
    $safeName = Get-SafeDirectoryName $item.TargetPath

    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "Backing up: $displayName" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor White

    # ----------------------------------------
    # Pre-check: target path existence
    # ----------------------------------------
    if (-not (Test-Path $item.TargetPath)) {
        Show-Skip "Path not found: $($item.TargetPath)"
        Write-Host ""
        $skipCount++
        continue
    }

    # ----------------------------------------
    # Main processing
    # ----------------------------------------
    try {
        # Phase A: Prepare backup directory
        $backupDir = Join-Path $backupBaseDir "$($item.Id)_$safeName"
        $individualDir = Join-Path $backupDir "individual"

        if (Test-Path $backupDir) {
            Remove-Item -Path $backupDir -Recurse -Force
        }
        New-Item -ItemType Directory -Path $individualDir -Force | Out-Null

        # Phase B: Full tree backup
        $fullBackupFile = Join-Path $backupDir "_full_acl.txt"
        Show-Info "Phase 1: Full tree backup..."
        Show-Info "  icacls `"$($item.TargetPath)`" /save `"$fullBackupFile`" /T /C"

        $proc = Start-Process -FilePath "icacls.exe" `
            -ArgumentList "`"$($item.TargetPath)`" /save `"$fullBackupFile`" /T /C" `
            -Wait -NoNewWindow -PassThru -ErrorAction Stop

        if ($proc.ExitCode -ne 0) {
            Show-Error "Full backup failed (ExitCode=$($proc.ExitCode))"
            $failCount++
            Write-Host ""
            continue
        }
        Show-Success "Full tree backup completed"

        # Phase C: Individual backups for non-inherited subdirectories
        $nonInheritedDirs = $scanResults[$item.Id]
        $manifestEntries = @()

        if ($nonInheritedDirs.Count -eq 0) {
            Show-Info "No non-inherited subdirectories found (full backup is sufficient)"
        }
        else {
            Show-Info "Phase 2: Individual backups for $($nonInheritedDirs.Count) non-inherited subdirectories..."
            $current = 0
            $indivSuccess = 0
            $indivFail = 0

            foreach ($subdir in $nonInheritedDirs) {
                $current++
                $relativePath = $subdir.FullName.Substring($item.TargetPath.Length).TrimStart('\')
                $hash = Get-PathHash $relativePath
                $backupFileName = "${hash}_acl.txt"
                $indivBackupFile = Join-Path $individualDir $backupFileName

                # Calculate depth for restore ordering
                $depth = ($relativePath.Split('\') | Where-Object { $_ -ne '' }).Count

                Write-Host "  [$current/$($nonInheritedDirs.Count)] $relativePath" -ForegroundColor DarkGray

                $indivProc = Start-Process -FilePath "icacls.exe" `
                    -ArgumentList "`"$($subdir.FullName)`" /save `"$indivBackupFile`" /T /C" `
                    -Wait -NoNewWindow -PassThru -ErrorAction Stop

                if ($indivProc.ExitCode -ne 0) {
                    Show-Warning "  Individual backup failed (ExitCode=$($indivProc.ExitCode)): $relativePath"
                    $indivFail++
                }
                else {
                    $indivSuccess++
                }

                $manifestEntries += [PSCustomObject]@{
                    RelativePath = $relativePath
                    Hash         = $hash
                    BackupFile   = $backupFileName
                    Depth        = $depth
                }
            }

            Show-Info "Individual backups: $indivSuccess succeeded, $indivFail failed"
        }

        # Phase D: Write manifest CSV
        $manifestPath = Join-Path $backupDir "_manifest.csv"
        if ($manifestEntries.Count -gt 0) {
            $manifestEntries | Export-Csv -Path $manifestPath -NoTypeInformation -Encoding UTF8 -Force
        }
        else {
            # Write empty manifest with headers
            [PSCustomObject]@{
                RelativePath = ""
                Hash         = ""
                BackupFile   = ""
                Depth        = ""
            } | Export-Csv -Path $manifestPath -NoTypeInformation -Encoding UTF8 -Force
            # Overwrite with header-only content
            "RelativePath,Hash,BackupFile,Depth" | Set-Content -Path $manifestPath -Encoding UTF8
        }

        Show-Success "Completed: $displayName (full + $($nonInheritedDirs.Count) individual)"
        $successCount++
    }
    catch {
        Show-Error "Failed: $displayName : $_"
        $failCount++
    }

    Write-Host ""
}


# ========================================
# Step 6: Result
# ========================================
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount `
    -Title "ACL Backup Results")
