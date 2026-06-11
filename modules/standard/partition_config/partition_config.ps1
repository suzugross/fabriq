# ========================================
# Partition Config Script
# ========================================
# Shrinks an existing partition and creates new partitions
# using PowerShell Storage cmdlets (Resize-Partition, New-Partition, Format-Volume).
# Supports multiple partition splits (e.g., C -> D -> E).
#
# [NOTES]
# - Administrator privileges required
# - Source drive must have enough unallocated or shrinkable space
# ========================================

Write-Host ""
Show-Separator
Write-Host "Partition Config" -ForegroundColor Cyan
Show-Separator
Write-Host ""


# ========================================
# Local Helper Functions
# ========================================

function Test-DriveLetterInUse {
    param([string]$DriveLetter)
    # Check both partitions and logical disks (CD-ROM, etc.)
    $partition = Get-Partition -DriveLetter $DriveLetter -ErrorAction SilentlyContinue
    if ($null -ne $partition) { return $true }
    $logicalDisk = Get-WmiObject Win32_LogicalDisk -Filter "DeviceID='${DriveLetter}:'" -ErrorAction SilentlyContinue
    return ($null -ne $logicalDisk)
}

function Test-DriveLetterConflict {
    param([string]$DriveLetter)
    # Returns $true if drive letter is used by a non-partition device (e.g., CD-ROM)
    $partition = Get-Partition -DriveLetter $DriveLetter -ErrorAction SilentlyContinue
    if ($null -ne $partition) { return $false }
    $logicalDisk = Get-WmiObject Win32_LogicalDisk -Filter "DeviceID='${DriveLetter}:'" -ErrorAction SilentlyContinue
    return ($null -ne $logicalDisk)
}

function Move-ConflictingDriveLetter {
    param([string]$DriveLetter)
    # Relocate a conflicting device (e.g., CD-ROM) to a free drive letter
    $cdVolume = Get-WmiObject Win32_Volume -Filter "DriveLetter='${DriveLetter}:'" -ErrorAction SilentlyContinue
    if ($null -eq $cdVolume) { return $false }

    $freeLetter = [char[]]('Z','Y','X','W','V','U','T','S','R','Q') | Where-Object {
        -not (Get-WmiObject Win32_LogicalDisk -Filter "DeviceID='$($_):'" -ErrorAction SilentlyContinue)
    } | Select-Object -First 1

    if (-not $freeLetter) {
        Show-Error "No free drive letter available to relocate $($DriveLetter):"
        return $false
    }

    $cdVolume.DriveLetter = "${freeLetter}:"
    $cdVolume.Put() | Out-Null
    Show-Info "Relocated conflicting drive $($DriveLetter): -> $($freeLetter):"
    return $true
}

function Get-PartitionSizeMB {
    param([string]$DriveLetter)
    $partition = Get-Partition -DriveLetter $DriveLetter -ErrorAction SilentlyContinue
    if ($null -eq $partition) { return 0 }
    return [math]::Round($partition.Size / 1MB, 0)
}


# ========================================
# Step 1: CSV Read
# ========================================
$csvPath = Join-Path $PSScriptRoot "partition_list.csv"

$enabledItems = Import-ModuleCsv -Path $csvPath -FilterEnabled `
    -RequiredColumns @("DiskNumber", "SourceDriveLetter", "SourceSizeMB", "NewDriveLetter", "NewSizeMB", "FileSystem")

if ($null -eq $enabledItems) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load partition_list.csv")
}
if ($enabledItems.Count -eq 0) {
    return (New-ModuleResult -Status "Skipped" -Message "No enabled entries")
}


# ========================================
# Step 2: Validation
# ========================================

# Ensure NewSizeMB=0 (use max) appears at most once and is the last entry
$useMaxEntries = @($enabledItems | Where-Object { [int]$_.NewSizeMB -eq 0 })
if ($useMaxEntries.Count -gt 1) {
    Show-Error "Multiple entries with NewSizeMB=0 found. Only one 'use all remaining space' entry is allowed."
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Multiple NewSizeMB=0 entries")
}
if ($useMaxEntries.Count -eq 1) {
    $lastItem = $enabledItems[-1]
    if ([int]$lastItem.NewSizeMB -ne 0) {
        Show-Error "NewSizeMB=0 entry must be the last row. Fixed-size partitions must come first."
        Write-Host ""
        return (New-ModuleResult -Status "Error" -Message "NewSizeMB=0 is not the last entry")
    }
}

# Check for duplicate NewDriveLetter
$driveLetters = $enabledItems | ForEach-Object { $_.NewDriveLetter.ToUpper() }
$duplicates = $driveLetters | Group-Object | Where-Object { $_.Count -gt 1 }
if ($duplicates) {
    Show-Error "Duplicate NewDriveLetter found: $($duplicates.Name -join ', ')"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Duplicate NewDriveLetter: $($duplicates.Name -join ', ')")
}


# ========================================
# Step 3: Dry-run Display
# ========================================
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "Partition Changes" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

# Show source drive shrink info, deduplicated by letter+size - the SAME
# key Phase A and the verification use. Deduplicating by letter alone
# hid second shrink rows from this preview while Phase A executed them.
$shrinkSeen = @{}
$shrinkTargets = @(foreach ($it in $enabledItems) {
    $shrinkKey = "$($it.SourceDriveLetter.ToUpper())-$([int]$it.SourceSizeMB)"
    if (-not $shrinkSeen.ContainsKey($shrinkKey)) {
        $shrinkSeen[$shrinkKey] = $true
        [PSCustomObject]@{ Letter = $it.SourceDriveLetter.ToUpper(); TargetMB = [int]$it.SourceSizeMB }
    }
})

foreach ($src in $shrinkTargets) {
    $currentMB = Get-PartitionSizeMB -DriveLetter $src.Letter
    if ($currentMB -le $src.TargetMB) {
        Write-Host "  [SKIP] $($src.Letter): shrink - already $($currentMB) MB (<= target $($src.TargetMB) MB)" -ForegroundColor DarkGray
    } else {
        Write-Host "  [APPLY] $($src.Letter): shrink $($currentMB) MB -> $($src.TargetMB) MB" -ForegroundColor Yellow
    }
    Write-Host ""
}

# Show new partition info
foreach ($item in $enabledItems) {
    $displayName = if ($item.Description) { $item.Description } else { "$($item.NewDriveLetter): partition" }
    $letter = $item.NewDriveLetter.ToUpper()

    if (Test-DriveLetterInUse -DriveLetter $letter) {
        if (Test-DriveLetterConflict -DriveLetter $letter) {
            $sizeLabel = if ([int]$item.NewSizeMB -eq 0) { "remaining space" } else { "$($item.NewSizeMB) MB" }
            Write-Host "  [RELOCATE] $($letter): is used by another device (CD-ROM etc.) - will be relocated" -ForegroundColor Yellow
            Write-Host "  [APPLY] $displayName" -ForegroundColor Yellow
            Write-Host "    Drive: $($letter):  Size: $sizeLabel  FS: $($item.FileSystem)  Label: $($item.VolumeLabel)" -ForegroundColor DarkGray
        } else {
            $existingMB = Get-PartitionSizeMB -DriveLetter $letter
            Write-Host "  [SKIP] $($letter): already exists ($($existingMB) MB)" -ForegroundColor DarkGray
        }
    } else {
        $sizeLabel = if ([int]$item.NewSizeMB -eq 0) { "remaining space" } else { "$($item.NewSizeMB) MB" }
        Write-Host "  [APPLY] $displayName" -ForegroundColor Yellow
        Write-Host "    Drive: $($letter):  Size: $sizeLabel  FS: $($item.FileSystem)  Label: $($item.VolumeLabel)" -ForegroundColor DarkGray
    }
    Write-Host ""
}

Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""


# ========================================
# Step 4: Confirm Execution
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Apply partition changes?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""


# ========================================
# Step 5: Apply Changes
# ========================================
$successCount = 0
$skipCount    = 0
$failCount    = 0

# ----------------------------------------
# Phase A: Shrink source partition (deduplicated)
# ----------------------------------------
$shrinkDone = @{}

foreach ($item in $enabledItems) {
    $srcLetter = $item.SourceDriveLetter.ToUpper()
    $targetMB  = [int]$item.SourceSizeMB
    $shrinkKey = "$srcLetter-$targetMB"

    # Skip if already processed
    if ($shrinkDone.ContainsKey($shrinkKey)) { continue }
    $shrinkDone[$shrinkKey] = $true

    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "Shrinking: $($srcLetter): -> $($targetMB) MB" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor White

    $currentMB = Get-PartitionSizeMB -DriveLetter $srcLetter
    if ($currentMB -le $targetMB) {
        Show-Skip "$($srcLetter): is already $($currentMB) MB (<= target $($targetMB) MB)"
        Write-Host ""
        continue
    }

    try {
        $targetBytes = [int64]$targetMB * 1MB
        $partition = Get-Partition -DriveLetter $srcLetter

        # Check minimum supported size
        $supportedSize = Get-PartitionSupportedSize -DriveLetter $srcLetter
        if ($targetBytes -lt $supportedSize.SizeMin) {
            Show-Error "Target size $($targetMB) MB is below minimum supported size $([math]::Round($supportedSize.SizeMin / 1MB, 0)) MB"
            $failCount++
            Write-Host ""
            continue
        }

        Resize-Partition -DriveLetter $srcLetter -Size $targetBytes -ErrorAction Stop
        Show-Success "Shrunk $($srcLetter): to $($targetMB) MB"
        $successCount++
    }
    catch {
        Show-Error "Failed to shrink $($srcLetter): $_"
        $failCount++
    }

    Write-Host ""
}

# ----------------------------------------
# Phase B: Create new partitions (in CSV order)
# ----------------------------------------
foreach ($item in $enabledItems) {
    $letter    = $item.NewDriveLetter.ToUpper()
    $diskNum   = [int]$item.DiskNumber
    $sizeMB    = [int]$item.NewSizeMB
    $fs        = $item.FileSystem
    $label     = $item.VolumeLabel
    $displayName = if ($item.Description) { $item.Description } else { "$($letter): partition" }

    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "Creating: $displayName" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor White

    # Idempotency check (partition already exists on this drive letter)
    if (Test-DriveLetterInUse -DriveLetter $letter) {
        if (Test-DriveLetterConflict -DriveLetter $letter) {
            # Drive letter is used by CD-ROM etc. - relocate it
            $relocated = Move-ConflictingDriveLetter -DriveLetter $letter
            if (-not $relocated) {
                Show-Error "Could not relocate conflicting device on $($letter):"
                $failCount++
                Write-Host ""
                continue
            }
        } else {
            # Existing partition: distinguish a healthy volume (Skip - a
            # reformat would destroy data even when the FS differs from
            # the CSV; verification flags the mismatch instead) from the
            # half-created RAW state a failed Format-Volume leaves behind
            # (New-Partition succeeded, format did not). RAW carries no
            # filesystem - and therefore no data - so retrying the format
            # is the safe self-heal that used to be a permanent Skip.
            $existingVol = Get-Volume -DriveLetter $letter -ErrorAction SilentlyContinue
            $isRaw = ($null -eq $existingVol) -or
                     ("$($existingVol.FileSystemType)" -eq 'Unknown') -or
                     ([string]::IsNullOrEmpty("$($existingVol.FileSystem)"))
            if (-not $isRaw) {
                Show-Skip "$($letter): already exists"
                Write-Host ""
                $skipCount++
                continue
            }

            Show-Warning "$($letter): exists but is unformatted (RAW) - retrying format"
            try {
                Format-Volume -DriveLetter $letter -FileSystem $fs -NewFileSystemLabel $label -Confirm:$false -ErrorAction Stop | Out-Null
                $actualMB = Get-PartitionSizeMB -DriveLetter $letter
                Show-Success "Formatted $($letter): ($($actualMB) MB, $($fs), Label: $($label))"
                $successCount++
            }
            catch {
                Show-Error "Failed to format $($letter): $_"
                $failCount++
            }
            Write-Host ""
            continue
        }
    }

    try {
        # Create partition
        if ($sizeMB -eq 0) {
            $newPartition = New-Partition -DiskNumber $diskNum -UseMaximumSize -DriveLetter $letter -ErrorAction Stop
        } else {
            $sizeBytes = [int64]$sizeMB * 1MB
            $newPartition = New-Partition -DiskNumber $diskNum -Size $sizeBytes -DriveLetter $letter -ErrorAction Stop
        }

        # Format volume
        Format-Volume -DriveLetter $letter -FileSystem $fs -NewFileSystemLabel $label -Confirm:$false -ErrorAction Stop | Out-Null

        $actualMB = Get-PartitionSizeMB -DriveLetter $letter
        Show-Success "Created $($letter): ($($actualMB) MB, $($fs), Label: $($label))"
        $successCount++
    }
    catch {
        Show-Error "Failed to create $($letter):: $_"
        $failCount++
    }

    Write-Host ""
}


# ========================================
# Step 5.5: Post-Apply Verification
# ========================================
$verifyPass = 0
$verifyFail = 0

Write-Host "----------------------------------------" -ForegroundColor White
Write-Host "Verification" -ForegroundColor Cyan
Write-Host "----------------------------------------" -ForegroundColor White

# Verify source partition size
$verifiedSources = @{}
foreach ($item in $enabledItems) {
    $srcLetter = $item.SourceDriveLetter.ToUpper()
    $targetMB  = [int]$item.SourceSizeMB
    $srcKey = "$srcLetter-$targetMB"

    if ($verifiedSources.ContainsKey($srcKey)) { continue }
    $verifiedSources[$srcKey] = $true

    $actualMB = Get-PartitionSizeMB -DriveLetter $srcLetter
    $tolerance = $targetMB * 0.05
    if ([math]::Abs($actualMB - $targetMB) -le $tolerance) {
        Write-Host "  [VERIFIED] $($srcLetter): size $($actualMB) MB (target: $($targetMB) MB)" -ForegroundColor Green
        $verifyPass++
    } else {
        Write-Host "  [VERIFY FAILED] $($srcLetter): size $($actualMB) MB (expected: ~$($targetMB) MB)" -ForegroundColor Red
        $verifyFail++
    }
}

# Verify new partitions
foreach ($item in $enabledItems) {
    $letter = $item.NewDriveLetter.ToUpper()
    $displayName = if ($item.Description) { $item.Description } else { "$($letter):" }

    if (-not (Test-DriveLetterInUse -DriveLetter $letter)) {
        Write-Host "  [VERIFY FAILED] $($letter): partition not found" -ForegroundColor Red
        $verifyFail++
        continue
    }

    $volume = Get-Volume -DriveLetter $letter -ErrorAction SilentlyContinue
    if ($null -eq $volume) {
        Write-Host "  [VERIFY FAILED] $($letter): volume not found" -ForegroundColor Red
        $verifyFail++
        continue
    }

    # Check file system
    $expectedFS = $item.FileSystem.ToUpper()
    $actualFS   = $volume.FileSystemType.ToString().ToUpper()
    if ($actualFS -ne $expectedFS) {
        Write-Host "  [VERIFY FAILED] $($letter): FileSystem is $actualFS (expected: $expectedFS)" -ForegroundColor Red
        $verifyFail++
        continue
    }

    # Check size (only for fixed-size entries)
    $sizeMB = [int]$item.NewSizeMB
    if ($sizeMB -gt 0) {
        $actualMB  = Get-PartitionSizeMB -DriveLetter $letter
        $tolerance = $sizeMB * 0.05
        if ([math]::Abs($actualMB - $sizeMB) -gt $tolerance) {
            Write-Host "  [VERIFY FAILED] $($letter): size $($actualMB) MB (expected: ~$($sizeMB) MB)" -ForegroundColor Red
            $verifyFail++
            continue
        }
    }

    Write-Host "  [VERIFIED] $($letter): $($actualFS), $( Get-PartitionSizeMB -DriveLetter $letter ) MB" -ForegroundColor Green
    $verifyPass++
}

Write-Host ""

$verified = ($verifyFail -eq 0)


# ========================================
# Step 6: Return Result
# ========================================
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount `
    -Title "Partition Config Results" -Verified $verified)
