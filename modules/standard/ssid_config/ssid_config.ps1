# ========================================
# SSID Config Script
# ========================================
# Registers Wi-Fi SSID profiles defined in ssid_list.csv
# using netsh wlan add profile with dynamically generated XML.
#
# [NOTES]
# - Requires administrator privileges
# - Supports WPA2PSK, WPA3SAE, and open authentication
# - Password field supports ENC: prefix for encrypted values
# - Temporary XML files are securely deleted after import
# - Existing profiles are skipped (idempotent)
# ========================================

Write-Host ""
Show-Separator
Write-Host "SSID Config" -ForegroundColor Cyan
Show-Separator
Write-Host ""


# ========================================
# Step 1: CSV reading
# ========================================
$csvPath = Join-Path $PSScriptRoot "ssid_list.csv"

$enabledItems = Import-ModuleCsv -Path $csvPath -FilterEnabled `
    -RequiredColumns @("Enabled", "SSID", "Authentication", "Encryption", "Password", "AutoConnect")

if ($null -eq $enabledItems) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load ssid_list.csv")
}
if ($enabledItems.Count -eq 0) {
    return (New-ModuleResult -Status "Skipped" -Message "No enabled entries")
}


# ========================================
# Step 2: Pre-flight check
# ========================================

# Verify WLAN service is available by attempting to list profiles
$profileOutput = netsh wlan show profiles 2>&1
if ($LASTEXITCODE -ne 0) {
    Show-Error "WLAN service is not available on this system."
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "WLAN service not available")
}

# Extract existing profile names (locale-independent parsing)
# Format: "    All User Profile     : SSID_NAME" (EN)
#         "    すべてのユーザー プロファイル : SSID_NAME" (JA)
# Match lines with " : " pattern followed by content
$existingProfiles = @()
foreach ($line in $profileOutput) {
    if ("$line" -match "\s+:\s+(.+)$") {
        $name = $Matches[1].Trim()
        if ($name -ne "") {
            $existingProfiles += $name
        }
    }
}

Show-Info "Existing WLAN profiles: $($existingProfiles.Count) found"


# ========================================
# Step 3: Pre-execution display
# ========================================
Write-Host ""
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "Target SSIDs" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

foreach ($item in $enabledItems) {
    $displayName = if ($item.Description) { $item.Description } else { $item.SSID }

    # Case-insensitive check against existing profiles
    $alreadyExists = $existingProfiles -icontains $item.SSID

    if ($alreadyExists) {
        Write-Host "  [SKIP] $displayName" -ForegroundColor DarkGray
    }
    else {
        Write-Host "  [ADD] $displayName" -ForegroundColor Yellow
    }

    $connMode = if ($item.AutoConnect -eq "1") { "Auto" } else { "Manual" }
    $authInfo = "$($item.Authentication)/$($item.Encryption)"
    Write-Host "    SSID: $($item.SSID)  Auth: $authInfo  Connect: $connMode" -ForegroundColor DarkGray
    Write-Host ""
}

Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""


# ========================================
# Step 4: Confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Register the above SSID profiles?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""


# ========================================
# Step 5: Execution loop
# ========================================
$successCount = 0
$skipCount    = 0
$failCount    = 0

foreach ($item in $enabledItems) {
    $displayName = if ($item.Description) { $item.Description } else { $item.SSID }

    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "Processing: $displayName" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor White

    # ----------------------------------------
    # Idempotency check (outside try)
    # ----------------------------------------
    $alreadyExists = $existingProfiles -icontains $item.SSID
    if ($alreadyExists) {
        Show-Skip "Profile already exists: $($item.SSID)"
        Write-Host ""
        $skipCount++
        continue
    }

    # ----------------------------------------
    # Main processing
    # ----------------------------------------
    $tempXml = Join-Path $env:TEMP "fabriq_wlan_$([guid]::NewGuid().ToString('N')).xml"

    try {
        # ----------------------------------------
        # 5a: Generate WLAN Profile XML
        # ----------------------------------------
        $ssidEscaped = [System.Security.SecurityElement]::Escape($item.SSID)
        $connMode    = if ($item.AutoConnect -eq "1") { "auto" } else { "manual" }

        if ($item.Authentication -ieq "open") {
            # Open network — no sharedKey block
            $xmlContent = @"
<?xml version="1.0"?>
<WLANProfile xmlns="http://www.microsoft.com/networking/WLAN/profile/v1">
    <name>$ssidEscaped</name>
    <SSIDConfig>
        <SSID><name>$ssidEscaped</name></SSID>
    </SSIDConfig>
    <connectionType>ESS</connectionType>
    <connectionMode>$connMode</connectionMode>
    <MSM>
        <security>
            <authEncryption>
                <authentication>open</authentication>
                <encryption>none</encryption>
                <useOneX>false</useOneX>
            </authEncryption>
        </security>
    </MSM>
</WLANProfile>
"@
        }
        else {
            # PSK network (WPA2PSK, WPA3SAE, etc.) — include sharedKey block
            $passwordEscaped = [System.Security.SecurityElement]::Escape($item.Password)

            $xmlContent = @"
<?xml version="1.0"?>
<WLANProfile xmlns="http://www.microsoft.com/networking/WLAN/profile/v1">
    <name>$ssidEscaped</name>
    <SSIDConfig>
        <SSID><name>$ssidEscaped</name></SSID>
    </SSIDConfig>
    <connectionType>ESS</connectionType>
    <connectionMode>$connMode</connectionMode>
    <MSM>
        <security>
            <authEncryption>
                <authentication>$($item.Authentication)</authentication>
                <encryption>$($item.Encryption)</encryption>
                <useOneX>false</useOneX>
            </authEncryption>
            <sharedKey>
                <keyType>passPhrase</keyType>
                <protected>false</protected>
                <keyMaterial>$passwordEscaped</keyMaterial>
            </sharedKey>
        </security>
    </MSM>
</WLANProfile>
"@
        }

        # ----------------------------------------
        # 5b: Write XML to temp file
        # ----------------------------------------
        $xmlContent | Out-File -FilePath $tempXml -Encoding UTF8 -ErrorAction Stop
        Show-Info "Temp XML created: $tempXml"

        # ----------------------------------------
        # 5c: Import profile via netsh
        # ----------------------------------------
        Show-Info "Importing profile: $($item.SSID)"
        $netshOutput = netsh wlan add profile filename="$tempXml" 2>&1 | Out-String
        $netshExitCode = $LASTEXITCODE

        foreach ($line in ($netshOutput.Trim() -split "\r?\n")) {
            Write-Host "  $line" -ForegroundColor DarkGray
        }

        if ($netshExitCode -ne 0) {
            Show-Error "netsh wlan add profile failed (ExitCode=$netshExitCode): $($item.SSID)"
            $failCount++
            continue
        }

        Show-Success "Registered: $($item.SSID)"
        $successCount++
    }
    catch {
        Show-Error "Failed: $displayName : $_"
        $failCount++
    }
    finally {
        # ----------------------------------------
        # 5d: Secure cleanup — always delete temp XML (contains password)
        # ----------------------------------------
        if (Test-Path $tempXml) {
            Remove-Item $tempXml -Force -ErrorAction SilentlyContinue
        }
    }

    Write-Host ""
}


# ========================================
# Step 6: Result
# ========================================
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount `
    -Title "SSID Config Results")
