# ========================================
# Credential Manager Config
# ========================================
# Registers credentials into the CURRENT USER's Windows Credential
# Manager (vault) via native cmdkey.exe, driven by credential_list.csv.
# Supports Generic, Domain Password, and RDP (TERMSRV/) credential types.
#
# [SCOPE - read before use]
# cmdkey writes ONLY to the vault of the account that RUNS this module
# (the kitting account). The credentials do NOT appear for any other user
# who logs in later, because each vault is DPAPI-encrypted per user. This
# module is for kitting-session credential staging, NOT for provisioning
# the delivered end-user's vault. See Guide.txt.
#
# [NOTES]
# - No admin rights are required (writes the current user's own vault).
# - Running as SYSTEM (S-1-5-18) is rejected: SYSTEM has no interactive
#   user vault, so the write would be useless.
# - The password is passed to cmdkey on its command line, so it is visible
#   to process-creation auditing (Event 4688 / Sysmon EID 1). This is an
#   exposure the registry-based autologon module does NOT have. See
#   Guide.txt before using on audited fleets.
# - Post-Apply Verification is intentionally NOT implemented: cmdkey cannot
#   read back a stored password (cmdkey /list shows target + user but never
#   the password), so presence in the vault is no proof the password is
#   correct. Asserting -Verified here would be a false PASS, so -Verified is
#   left $null (this module is on the verification-exclusion list).
# ========================================

Write-Host ""
Show-Separator
Write-Host "Credential Manager Config" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# ========================================
# Helper Functions
# ========================================

function Get-MaskedSecret {
    param([string]$Value)
    if ([string]::IsNullOrEmpty($Value)) { return "**" }
    if ($Value.Length -gt 2) {
        return $Value[0] + ("*" * ($Value.Length - 2)) + $Value[-1]
    }
    return "**"
}

function Resolve-CredentialPlan {
    # Normalizes one CSV row into an apply plan, or returns an invalid marker
    # with a human-readable reason. Computes the effective cmdkey target
    # (TERMSRV/ prefix for RDP) and the cmdkey verb up front.
    param($Item)

    $credType    = ("" + $Item.CredType).Trim()
    $target      = ("" + $Item.TargetName).Trim()
    $user        = ("" + $Item.UserName).Trim()
    $password    = "" + $Item.Password
    $displayName = if ($Item.Description) { $Item.Description } else { $target }
    $typeKey     = $credType.ToUpperInvariant()

    if ($typeKey -notin @("GENERIC", "DOMAIN", "RDP")) {
        return @{ Valid = $false; DisplayName = $displayName; Reason = "invalid CredType '$credType' (expected Generic/Domain/RDP)" }
    }
    if ([string]::IsNullOrWhiteSpace($target)) {
        return @{ Valid = $false; DisplayName = $displayName; Reason = "blank TargetName" }
    }
    if ([string]::IsNullOrWhiteSpace($user)) {
        return @{ Valid = $false; DisplayName = $displayName; Reason = "blank UserName" }
    }
    if ([string]::IsNullOrEmpty($password)) {
        return @{ Valid = $false; DisplayName = $displayName; Reason = "blank Password" }
    }

    switch ($typeKey) {
        "GENERIC" {
            $effectiveTarget = $target
            $verb = "/generic:$effectiveTarget"
            $typeLabel = "Generic"
        }
        "DOMAIN" {
            $effectiveTarget = $target
            $verb = "/add:$effectiveTarget"
            $typeLabel = "Domain Password"
        }
        "RDP" {
            $effectiveTarget = if ($target -like "TERMSRV/*") { $target } else { "TERMSRV/$target" }
            $verb = "/generic:$effectiveTarget"
            $typeLabel = "RDP (Generic, TERMSRV)"
        }
    }

    if ([string]::IsNullOrWhiteSpace($effectiveTarget)) {
        return @{ Valid = $false; DisplayName = $displayName; Reason = "empty effective target after normalization" }
    }

    return @{
        Valid           = $true
        DisplayName     = $displayName
        TypeKey         = $typeKey
        TypeLabel       = $typeLabel
        Target          = $target
        EffectiveTarget = $effectiveTarget
        Verb            = $verb
        User            = $user
        Password        = $password
    }
}

# ========================================
# Step 1: Load CSV
# ========================================
# Load WITHOUT -FilterEnabled and filter Enabled manually (autologon_config
# pattern). Import-ModuleCsv -FilterEnabled returns @() when no rows are
# enabled, which PowerShell unwraps to $null at the call site - that would
# be indistinguishable from a genuine load failure. Loading all rows lets us
# report a real load failure as Error and an all-disabled CSV as Skipped.
$csvPath = Join-Path $PSScriptRoot "credential_list.csv"
$allItems = Import-ModuleCsv -Path $csvPath `
    -RequiredColumns @("Enabled", "CredType", "TargetName", "UserName", "Password")

if ($null -eq $allItems) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load credential_list.csv")
}

$enabledItems = @($allItems | Where-Object { $_.Enabled -eq "1" })
if ($enabledItems.Count -eq 0) {
    Show-Skip "No enabled entries in credential_list.csv"
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "No enabled entries")
}

# ========================================
# Step 2: Validate rows (fail before any side effect)
# ========================================
# cmdkey returns exit 0 even for a malformed/blank target while doing
# nothing, so the only safe success oracle is "valid args + cmdkey exit 0".
# Validate the COMPUTED effective target here, not just the raw input, and
# abort the whole run on any invalid row so a typo never registers silently.
$plans = @()
$invalid = @()
foreach ($item in $enabledItems) {
    $plan = Resolve-CredentialPlan -Item $item
    if ($plan.Valid) {
        $plans += $plan
    }
    else {
        $invalid += "$($plan.DisplayName) : $($plan.Reason)"
    }
}

if ($invalid.Count -gt 0) {
    Show-Error "Invalid credential rows found ($($invalid.Count)). Fix the CSV before applying:"
    foreach ($reason in $invalid) {
        Show-Error "  - $reason"
    }
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Validation failed for $($invalid.Count) row(s)")
}

# ========================================
# Step 2.5: Vault-scope guard (current-user vault)
# ========================================
# cmdkey writes to the CURRENT USER vault only. SYSTEM has no interactive
# vault, so fail loud there instead of writing a dead vault. For any normal
# account, state plainly WHOSE vault is being written (operator visibility).
$currentIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$currentName = $currentIdentity.Name
$currentSid  = $currentIdentity.User.Value

if ($currentSid -eq "S-1-5-18") {
    Show-Error "Running as LocalSystem (SYSTEM). cmdkey cannot populate an interactive user's vault."
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Cannot register credentials under SYSTEM context")
}

# ========================================
# Step 3: Dry-run summary before execution
# ========================================
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "Credential Registration Targets" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""
Show-Info "Target vault (current user): $currentName"
Show-Warning "Credentials are stored in THIS account's vault only; other users will not see them."
Write-Host ""

foreach ($plan in $plans) {
    Write-Host "  [REGISTER] $($plan.DisplayName)" -ForegroundColor White
    Write-Host "      Type:   $($plan.TypeLabel)" -ForegroundColor DarkGray
    Write-Host "      Target: $($plan.EffectiveTarget)" -ForegroundColor DarkGray
    Write-Host "      User:   $($plan.User)" -ForegroundColor DarkGray
    Write-Host "      Pass:   $(Get-MaskedSecret $plan.Password)" -ForegroundColor DarkGray
    Write-Host ""
}

Write-Host "  Existing entries for the same target are overwritten (last-write-wins)." -ForegroundColor DarkGray
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

# ========================================
# Step 4: User confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Register the above credentials into the current user's Credential Manager?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""

# ========================================
# Step 5: Apply-settings loop
# ========================================
# cmdkey /add and /generic silently overwrite an existing target
# (last-write-wins), so registration is idempotent and we always (re)apply
# rather than skip - a skip could leave a stale password that cmdkey can
# never confirm. Success oracle = cmdkey exit code 0 (Step 2 already
# guaranteed non-blank args).
$successCount = 0
$skipCount    = 0
$failCount    = 0
$total        = $plans.Count
$current      = 0

foreach ($plan in $plans) {
    $current++
    Write-Host "[$current/$total] $($plan.DisplayName)" -ForegroundColor Cyan

    # Build the argument list so the password is a single argv element,
    # never string-interpolated into a command line. cmdkey output is
    # discarded; we report our own masked result so the secret never
    # reaches the transcript or an error message.
    $cmdkeyArgs = @($plan.Verb, "/user:$($plan.User)", "/pass:$($plan.Password)")

    try {
        & cmdkey.exe @cmdkeyArgs 2>&1 | Out-Null
        $exitCode = $LASTEXITCODE

        if ($exitCode -eq 0) {
            Show-Success "Registered: $($plan.DisplayName) [$($plan.TypeLabel)] -> $($plan.EffectiveTarget) (user $($plan.User))"
            $successCount++
        }
        else {
            Show-Error "cmdkey failed for $($plan.DisplayName) (exit $exitCode)"
            $failCount++
        }
    }
    catch {
        # Deliberately omit $_ here: a native-command error record could echo
        # the argument list, including the password. Report only non-secret
        # context (display name, type, effective target).
        Show-Error "cmdkey invocation error for $($plan.DisplayName) [$($plan.TypeLabel)] -> $($plan.EffectiveTarget)"
        $failCount++
    }

    Write-Host ""
}

# ========================================
# Step 5.5: Post-Apply Verification - intentionally NOT implemented
# ========================================
# cmdkey cannot read back a stored password, and cmdkey /list returns exit 0
# even for a missing target while always echoing the queried target name. A
# presence-based check would therefore report PASS for a stale/wrong password
# (a false PASS), so this module is on the verification-exclusion list and
# returns -Verified = $null. See Guide.txt.

# ========================================
# Step 6: Aggregate and return result
# ========================================
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount `
    -Title "Credential Manager Config Results")
