# ========================================
# Local Group Policy Configuration Script
# ========================================
# [PURPOSE]
# Applies registry-based local Group Policy (Administrative Templates)
# from gpo_list*.csv by editing %SystemRoot%\System32\GroupPolicy\
# {Machine,User}\Registry.pol directly, bumping gpt.ini, then running
# gpupdate. LGPO.exe is NOT used (not redistributable).
#
# [NOTES]
# - Requires administrator privileges and a 64-bit PowerShell process
#   (a 32-bit process is redirected to SysWOW64\GroupPolicy).
# - Merge semantics: CSV rows are upserted into the existing Registry.pol;
#   entries not listed in the CSV are left untouched. Action=Unmanage
#   removes an entry (= "Not Configured"); Group Policy then cleans up the
#   registry value on the next refresh.
# - Machine rows are verified twice (Registry.pol + HKLM read-back).
#   User rows are verified against Registry.pol only (user policy is
#   applied per user at logon); the HKCU state is shown for information.
# - Domain GPOs take precedence over local policy on domain-joined PCs.
# ========================================

Write-Host ""
Show-Separator
Write-Host "Local Group Policy Configuration" -ForegroundColor Cyan
Show-Separator
Write-Host ""

. (Join-Path $PSScriptRoot "lib\PolFile.ps1")

# ========================================
# Step 1: Load CSV (all files matching gpo_list*.csv)
# ========================================
$csvFiles = @(Get-ChildItem -Path $PSScriptRoot -Filter "gpo_list*.csv" -File | Sort-Object Name)

if ($csvFiles.Count -eq 0) {
    Show-Error "No files matching gpo_list*.csv found"
    return (New-ModuleResult -Status "Error" -Message "No files matching gpo_list*.csv found")
}

$requiredColumns = @("Enabled", "AdminID", "SettingTitle", "Scope", "KeyPath", "ValueName", "Action", "Type", "Value")
$allItems = @()
$loadedFileCount = 0

foreach ($csvFile in $csvFiles) {
    $items = Import-ModuleCsv -Path $csvFile.FullName -RequiredColumns $requiredColumns
    if ($null -ne $items) {
        $allItems += $items
        Show-Info "Loaded $($csvFile.Name) ($($items.Count) items)"
        $loadedFileCount++
    }
}

if ($loadedFileCount -eq 0) {
    Show-Error "Failed to load any CSV files"
    return (New-ModuleResult -Status "Error" -Message "Failed to load any CSV files")
}

$gpoItems = @($allItems | Where-Object { $_.'Enabled' -eq '1' })
$disabledCount = $allItems.Count - $gpoItems.Count

Write-Host ""
$skipMsg = if ($disabledCount -gt 0) { " ($disabledCount disabled)" } else { "" }
Show-Info "Total: $($gpoItems.Count) enabled$skipMsg"
Write-Host ""

if ($gpoItems.Count -eq 0) {
    Show-Info "No enabled policy entries found"
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "No enabled policy entries found")
}

# ========================================
# Step 2: Prerequisite checks (early return)
# ========================================
if (-not (Test-AdminPrivilege)) {
    Show-Error "Administrator privileges are required"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Administrator privileges required")
}

if ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess) {
    Show-Error "32-bit PowerShell process detected: System32\GroupPolicy would be redirected to SysWOW64. Run from 64-bit PowerShell."
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "32-bit process not supported")
}

$partOfDomain = $false
try { $partOfDomain = [bool](Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop).PartOfDomain } catch { }
if ($partOfDomain) {
    Show-Warning "This PC is domain-joined: domain GPOs override local policy, so verification may report mismatches"
    Write-Host ""
}

# ========================================
# Step 2.5: Row validation (fail-closed: nothing is written on any error)
# ========================================
$plan = New-Object System.Collections.Generic.List[object]
$validationErrors = New-Object System.Collections.Generic.List[string]
$seenIdentities = @{}

foreach ($item in $gpoItems) {
    $label = "[$($item.'AdminID')] $($item.'SettingTitle')"

    $scope = switch -Regex ("$($item.'Scope')".Trim()) {
        '^(?i)machine$' { 'Machine' }
        '^(?i)user$'    { 'User' }
        default         { $null }
    }
    if (-not $scope) { $validationErrors.Add("$label : Scope must be Machine or User (got '$($item.'Scope')')"); continue }

    $action = switch -Regex ("$($item.'Action')".Trim()) {
        '^(?i)set$'             { 'Set' }
        '^(?i)delete$'          { 'Delete' }
        '^(?i)deleteallvalues$' { 'DeleteAllValues' }
        '^(?i)createkey$'       { 'CreateKey' }
        '^(?i)unmanage$'        { 'Unmanage' }
        default                 { $null }
    }
    if (-not $action) { $validationErrors.Add("$label : Action must be Set / Delete / DeleteAllValues / CreateKey / Unmanage (got '$($item.'Action')')"); continue }

    $keyPath = "$($item.'KeyPath')".Trim()
    $stripped = $keyPath -replace '^(HKEY_LOCAL_MACHINE|HKLM|HKEY_CURRENT_USER|HKCU):?\\', ''
    if ($stripped -ne $keyPath) {
        Show-Warning "$label : hive prefix removed from KeyPath (policy key paths are hive-less)"
        $keyPath = $stripped
    }
    $keyPath = $keyPath.Trim('\')
    if ([string]::IsNullOrWhiteSpace($keyPath)) { $validationErrors.Add("$label : KeyPath is empty"); continue }

    $valueName = "$($item.'ValueName')".Trim()
    if ($valueName.StartsWith('**')) { $validationErrors.Add("$label : ValueName must not start with '**' (express deletions via the Action column)"); continue }

    # Unmanage with an empty ValueName targets the key-only (CreateKey) entry.
    if ($action -in @('Set', 'Delete')) {
        if ([string]::IsNullOrWhiteSpace($valueName)) { $validationErrors.Add("$label : ValueName is required for Action=$action"); continue }
    }
    elseif ($action -ne 'Unmanage') {
        if (-not [string]::IsNullOrWhiteSpace($valueName)) {
            Show-Warning "$label : ValueName '$valueName' is ignored for Action=$action"
            $valueName = ''
        }
    }

    $typeName = "$($item.'Type')".Trim().ToUpperInvariant()
    $entry = $null
    try {
        $entry = ConvertTo-PolEntryFromAction -Key $keyPath -ValueName $valueName -Action $action -Type $typeName -Value "$($item.'Value')"
    }
    catch {
        $validationErrors.Add("$label : $($_.Exception.Message)")
        continue
    }

    $identityName = switch ($action) {
        'DeleteAllValues' { '**delvals.' }
        'CreateKey'       { '' }
        default           { $valueName }
    }
    $identity = Get-PolEntryIdentity -Key $keyPath -ValueName $identityName
    $scopedIdentity = "$scope|$identity"
    if ($seenIdentities.ContainsKey($scopedIdentity)) {
        $validationErrors.Add("$label : duplicate target ($scope $keyPath \ $identityName) - already defined by $($seenIdentities[$scopedIdentity])")
        continue
    }
    $seenIdentities[$scopedIdentity] = $label

    $plan.Add([PSCustomObject]@{
        Label     = $label
        Scope     = $scope
        Action    = $action
        KeyPath   = $keyPath
        ValueName = $valueName
        TypeName  = $typeName
        Value     = "$($item.'Value')"
        Entry     = $entry
        Identity  = $identity
        State     = ''
        Result    = ''
    })
}

if ($validationErrors.Count -gt 0) {
    foreach ($e in $validationErrors) { Show-Error $e }
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "CSV validation failed ($($validationErrors.Count) errors) - nothing was changed")
}

# ========================================
# Step 2.7: Read the existing Registry.pol per scope
# ========================================
$scopes = [ordered]@{}
foreach ($scopeName in @('Machine', 'User')) {
    $rows = @($plan | Where-Object { $_.Scope -eq $scopeName })
    if ($rows.Count -eq 0) { continue }
    $polPath = Get-PolFilePath -Scope $scopeName
    try {
        $existing = @(Read-PolFile -Path $polPath)
    }
    catch {
        Show-Error "Existing Registry.pol ($scopeName) is unreadable: $($_.Exception.Message)"
        Write-Host ""
        return (New-ModuleResult -Status "Error" -Message "Existing Registry.pol ($scopeName) unreadable - nothing was changed")
    }
    $index = @{}
    foreach ($e in @($existing)) { $index[(Get-PolEntryIdentity -Key $e.Key -ValueName $e.ValueName)] = $e }
    $scopes[$scopeName] = [PSCustomObject]@{
        Name       = $scopeName
        PolPath    = $polPath
        Existing   = $existing
        Index      = $index
        Rows       = $rows
        GpupdateOk = $null
    }
    Show-Info "$scopeName Registry.pol: $(@($existing).Count) existing entries ($polPath)"
}
Write-Host ""

# ========================================
# Step 3: Dry-run summary (idempotency probe against Registry.pol)
# ========================================
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "The following local group policy changes will be applied" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

foreach ($row in $plan) {
    $cur = $scopes[$row.Scope].Index[$row.Identity]
    if ($row.Action -eq 'Unmanage') {
        $isCurrent = ($null -eq $cur)
    }
    else {
        $isCurrent = ($null -ne $cur) -and (Test-PolEntryEqual -A $cur -B $row.Entry)
    }
    $row.State = if ($isCurrent) { 'Current' } else { 'Change' }

    $marker = if ($isCurrent) { "[Current]" } else { "[Change]" }
    $markerColor = if ($isCurrent) { "Gray" } else { "White" }

    Write-Host "$($row.Label)  $marker" -ForegroundColor $markerColor
    Write-Host "  Scope:  $($row.Scope)"
    Write-Host "  Path:   $($row.KeyPath)"
    switch ($row.Action) {
        'Set'             { Write-Host "  Action: Set  $($row.ValueName) = $($row.Value) ($($row.TypeName))" }
        'Delete'          { Write-Host "  Action: Delete value  $($row.ValueName)" }
        'DeleteAllValues' { Write-Host "  Action: Delete all values under the key" }
        'CreateKey'       { Write-Host "  Action: Create key (no values)" }
        'Unmanage'        { $target = if ($row.ValueName) { $row.ValueName } else { "(key-only entry)" }; Write-Host "  Action: Unmanage  $target  (remove from policy = Not Configured)" }
    }
    Write-Host ""
}

foreach ($scope in $scopes.Values) {
    $changes = @($scope.Rows | Where-Object { $_.State -eq 'Change' }).Count
    Write-Host "  $($scope.Name): $changes change(s) / $($scope.Rows.Count) row(s)" -ForegroundColor Cyan
}
Write-Host ""
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

# ========================================
# Step 4: User confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Apply the above local group policy changes?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""
Show-Info "Updating local group policy..."
Write-Host ""

# ========================================
# Step 5: Apply - merge into Registry.pol per scope, bump gpt.ini, gpupdate
# ========================================
$successCount = 0
$skipCount    = 0
$failCount    = 0
$timestamp    = Get-Date -Format "yyyyMMdd_HHmmss"
$backupDir    = Join-Path (Join-Path $PSScriptRoot "backup") $timestamp

foreach ($scope in $scopes.Values) {
    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "Processing: $($scope.Name) policy" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor White

    $changeRows  = @($scope.Rows | Where-Object { $_.State -eq 'Change' })
    $currentRows = @($scope.Rows | Where-Object { $_.State -eq 'Current' })

    foreach ($row in $currentRows) { $row.Result = 'Skip' }
    $skipCount += $currentRows.Count

    if ($changeRows.Count -eq 0) {
        Show-Skip "Registry.pol already contains all $($currentRows.Count) entries"
        Write-Host ""
        continue
    }

    try {
        # Backup the files we are about to touch (write-only; nothing is deleted)
        if (-not (Test-Path -LiteralPath $backupDir)) { New-Item -ItemType Directory -Path $backupDir -Force | Out-Null }
        if (Test-Path -LiteralPath $scope.PolPath) {
            Copy-Item -LiteralPath $scope.PolPath -Destination (Join-Path $backupDir "$($scope.Name)_Registry.pol") -Force
        }
        $gptIni = Get-GptIniPath
        if ((Test-Path -LiteralPath $gptIni) -and -not (Test-Path -LiteralPath (Join-Path $backupDir "gpt.ini"))) {
            Copy-Item -LiteralPath $gptIni -Destination (Join-Path $backupDir "gpt.ini") -Force
        }

        $removeIds    = @($changeRows | ForEach-Object { $_.Identity })
        $applyEntries = @($changeRows | Where-Object { $null -ne $_.Entry } | ForEach-Object { $_.Entry })
        $merged = Merge-PolEntries -Existing $scope.Existing -RemoveIdentities $removeIds -Apply $applyEntries

        Write-Host "  -> Writing Registry.pol ($(@($merged).Count) entries)" -ForegroundColor Gray
        Write-PolFile -Path $scope.PolPath -Entries $merged

        $ver = Update-GptIniVersion -Path $gptIni -Scope $scope.Name
        Show-Success "Registry.pol updated ($($changeRows.Count) change(s)), gpt.ini version $($ver.Old) -> $($ver.New)"

        foreach ($row in $changeRows) { $row.Result = 'Success' }
        $successCount += $changeRows.Count
    }
    catch {
        Show-Error "Failed to update $($scope.Name) Registry.pol: $($_.Exception.Message)"
        foreach ($row in $changeRows) { $row.Result = 'Fail' }
        $failCount += $changeRows.Count
    }
    Write-Host ""
}

# gpupdate per scope (always, so a manually tampered registry is re-aligned too)
foreach ($scope in $scopes.Values) {
    $target = if ($scope.Name -eq 'Machine') { 'computer' } else { 'user' }
    Write-Host "  -> gpupdate /target:$target /force" -ForegroundColor Gray
    $exitCode = -1
    try {
        $proc = Start-Process -FilePath "gpupdate.exe" -ArgumentList "/target:$target /force" -Wait -PassThru -NoNewWindow
        $exitCode = $proc.ExitCode
    }
    catch {
        Show-Error "gpupdate ($target) could not be started: $($_.Exception.Message)"
    }

    if ($exitCode -eq 0) {
        Show-Success "gpupdate ($target) completed"
        $scope.GpupdateOk = $true
    }
    else {
        Show-Error "gpupdate ($target) exit code: $exitCode - policy is written to Registry.pol but not applied yet (re-run this module or gpupdate)"
        $scope.GpupdateOk = $false
        foreach ($row in @($scope.Rows | Where-Object { $_.Result -eq 'Success' })) {
            $row.Result = 'Fail'
            $successCount--
            $failCount++
        }
    }
    Write-Host ""
}

# ========================================
# Step 5.5: Post-Apply Verification
# ========================================
Show-Info "Verifying applied settings..."
Write-Host ""

$verifyPass = 0
$verifyFail = 0

foreach ($scope in $scopes.Values) {
    $reread = @()
    $rereadOk = $true
    try { $reread = @(Read-PolFile -Path $scope.PolPath) } catch { $rereadOk = $false; Show-Error "Cannot re-read $($scope.Name) Registry.pol: $($_.Exception.Message)" }
    $index = @{}
    foreach ($e in $reread) { $index[(Get-PolEntryIdentity -Key $e.Key -ValueName $e.ValueName)] = $e }

    foreach ($row in $scope.Rows) {
        $cur = $index[$row.Identity]
        if ($row.Action -eq 'Unmanage') {
            $polOk = $rereadOk -and ($null -eq $cur)
        }
        else {
            $polOk = $rereadOk -and ($null -ne $cur) -and (Test-PolEntryEqual -A $cur -B $row.Entry)
        }

        $regOk = $null
        $hkcuNote = ''
        if ($row.Action -ne 'Unmanage') {
            if ($scope.Name -eq 'Machine') {
                $regOk = Test-PolEntryRegistryState -Entry $row.Entry -Action $row.Action -Hive 'HKLM'
            }
            else {
                $hkcuState = Test-PolEntryRegistryState -Entry $row.Entry -Action $row.Action -Hive 'HKCU'
                $hkcuNote = if ($hkcuState) { " (HKCU of current user: applied)" } else { " (HKCU of current user: not yet applied - user policy lands per user at logon)" }
            }
        }

        $pass = $polOk -and ($regOk -ne $false)
        if ($pass) {
            Write-Host "  [VERIFIED] $($row.Label)$hkcuNote" -ForegroundColor Green
            $verifyPass++
        }
        else {
            $detail = "Registry.pol: " + $(if ($polOk) { "ok" } else { "mismatch" })
            if ($null -ne $regOk) { $detail += ", HKLM: " + $(if ($regOk) { "ok" } else { "mismatch" }) }
            Write-Host "  [VERIFY FAILED] $($row.Label) ($detail)" -ForegroundColor Red
            $verifyFail++
        }
    }
}

Write-Host ""
$verified = ($verifyFail -eq 0)

# ========================================
# Step 6: Aggregate and return result
# ========================================
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount `
    -Title "Local Group Policy Results" -Verified $verified)
