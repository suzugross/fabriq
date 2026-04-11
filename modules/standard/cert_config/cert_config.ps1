# ========================================
# Certificate Config - Certificate Import Tool
# ========================================
# Imports certificates (.pfx/.p12/.cer/.crt) into
# Windows certificate stores based on cert_list.csv.
# Supports AutoRoute for multi-cert PFX files and
# environment variable + wildcard file name matching.
# ========================================

Write-Host ""
Show-Separator
Write-Host "Certificate Config" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# ========================================
# Helper Functions
# ========================================

function Resolve-CertFileName {
    param(
        [string]$FileName,
        [string]$CertsDir
    )

    # Expand environment variables (%SELECTED_NEW_PCNAME%, etc.)
    $expanded = Expand-UserEnvironmentVariables $FileName

    if ($expanded -match '\*') {
        # Wildcard match: find first matching file
        $matched = Get-ChildItem -Path $CertsDir -Filter $expanded -File -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($null -eq $matched) {
            return @{ Resolved = $null; Original = $FileName; Expanded = $expanded }
        }
        return @{ Resolved = $matched.Name; Original = $FileName; Expanded = $expanded }
    }
    else {
        return @{ Resolved = $expanded; Original = $FileName; Expanded = $expanded }
    }
}

function Get-CertificateThumbprint {
    param(
        [string]$Path,
        [string]$Password
    )

    try {
        $ext = [System.IO.Path]::GetExtension($Path).ToLower()
        if ($ext -in @('.pfx', '.p12')) {
            $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2(
                $Path, $Password,
                [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::DefaultKeySet
            )
            $thumbprint = $cert.Thumbprint
            $cert.Dispose()
            return $thumbprint
        }
        else {
            $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($Path)
            $thumbprint = $cert.Thumbprint
            $cert.Dispose()
            return $thumbprint
        }
    }
    catch {
        return $null
    }
}

function Test-CertificateInStore {
    param(
        [string]$Thumbprint,
        [string]$StoreScope,
        [string]$StoreName
    )

    try {
        $scope = [System.Security.Cryptography.X509Certificates.StoreLocation]$StoreScope
        $store = New-Object System.Security.Cryptography.X509Certificates.X509Store($StoreName, $scope)
        $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadOnly)
        $found = $store.Certificates | Where-Object { $_.Thumbprint -eq $Thumbprint }
        $store.Close()
        return ($null -ne $found)
    }
    catch {
        return $false
    }
}

# ========================================
# Step 1: Load CSV
# ========================================
$csvPath = Join-Path $PSScriptRoot "cert_list.csv"
$enabledItems = Import-ModuleCsv -Path $csvPath -FilterEnabled `
    -RequiredColumns @("Enabled", "FileName")

if ($null -eq $enabledItems) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load cert_list.csv")
}
if ($enabledItems.Count -eq 0) {
    return (New-ModuleResult -Status "Skipped" -Message "No enabled entries")
}

# ========================================
# Step 2: Validate certs directory
# ========================================
$certsDir = Join-Path $PSScriptRoot "certs"
if (-not (Test-Path $certsDir)) {
    Show-Error "certs/ directory not found: $certsDir"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "certs/ directory not found")
}

# ========================================
# Step 3: Display import targets with status
# ========================================

# Pre-resolve file names and build display data
$resolvedItems = @()
foreach ($item in $enabledItems) {
    $fileInfo = Resolve-CertFileName -FileName $item.FileName -CertsDir $certsDir
    $isAutoRoute = ($item.AutoRoute -eq "1")
    $isOverwrite = ($item.Overwrite -eq "1")
    $displayName = if ($item.Description) { $item.Description } else { $item.FileName }

    $ext = ""
    $certPath = ""
    $isPfx = $false

    if ($null -ne $fileInfo.Resolved) {
        $certPath = Join-Path $certsDir $fileInfo.Resolved
        $ext = [System.IO.Path]::GetExtension($fileInfo.Resolved).ToLower()
        $isPfx = ($ext -in @('.pfx', '.p12'))
    }

    $resolvedItems += @{
        Item        = $item
        FileInfo    = $fileInfo
        CertPath    = $certPath
        IsPfx       = $isPfx
        IsAutoRoute = $isAutoRoute
        IsOverwrite = $isOverwrite
        DisplayName = $displayName
    }
}

Show-Info "Import targets: $($enabledItems.Count) items"
Write-Host ""

Write-Host "========================================" -ForegroundColor Yellow
Write-Host "Certificate Import Targets" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

$index = 0
foreach ($ri in $resolvedItems) {
    $index++
    $item = $ri.Item
    $fileInfo = $ri.FileInfo
    $certPath = $ri.CertPath
    $isPfx = $ri.IsPfx
    $isAutoRoute = $ri.IsAutoRoute
    $isOverwrite = $ri.IsOverwrite
    $displayName = $ri.DisplayName

    # File not found
    if ($null -eq $fileInfo.Resolved -or -not (Test-Path $certPath)) {
        Write-Host "  [$index] $displayName  [NOT FOUND]" -ForegroundColor Red
        if ($fileInfo.Original -ne $fileInfo.Expanded) {
            Write-Host "      Pattern: $($fileInfo.Original) -> $($fileInfo.Expanded)" -ForegroundColor DarkGray
        }
        else {
            Write-Host "      File: $($fileInfo.Original)" -ForegroundColor DarkGray
        }
        Write-Host ""
        continue
    }

    # PFX without password
    if ($isPfx -and [string]::IsNullOrEmpty($item.Password)) {
        Write-Host "  [$index] $displayName  [NO PASSWORD]" -ForegroundColor Red
        Write-Host "      File: $($fileInfo.Resolved)" -ForegroundColor DarkGray
        Write-Host ""
        continue
    }

    # AutoRoute mode: expand PFX and show contained certificates
    if ($isAutoRoute -and $isPfx) {
        Write-Host "  [$index] $displayName" -ForegroundColor Cyan
        if ($fileInfo.Original -ne $fileInfo.Resolved) {
            Write-Host "      File: $($fileInfo.Resolved) (matched: $($fileInfo.Original))" -ForegroundColor DarkGray
        }
        else {
            Write-Host "      File: $($fileInfo.Resolved)" -ForegroundColor DarkGray
        }
        Write-Host "      AutoRoute: ON" -ForegroundColor DarkGray

        try {
            $collection = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2Collection
            $collection.Import(
                $certPath, $item.Password,
                [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::DefaultKeySet
            )

            Write-Host "      Contains $($collection.Count) certificate(s):" -ForegroundColor DarkGray
            foreach ($cert in $collection) {
                # Detect CA certificates using X.509 Basic Constraints extension (OID 2.5.29.19)
                # This correctly identifies intermediate CAs (Subject != Issuer but CA=true)
                $basicConstraints = $cert.Extensions | Where-Object { $_.Oid.Value -eq "2.5.29.19" }
                $isCa = ($null -ne $basicConstraints -and $basicConstraints.CertificateAuthority)
                $targetScope = "LocalMachine"
                $targetStore = if ($isCa) { "Root" } else { "My" }
                $certType = if ($isCa) { "CA" } else { "Client" }
                $existsInStore = Test-CertificateInStore -Thumbprint $cert.Thumbprint -StoreScope $targetScope -StoreName $targetStore

                if ($existsInStore -and -not $isOverwrite) {
                    $marker = "[SKIP]"
                    $markerColor = "Gray"
                    $detail = "(already exists)"
                }
                elseif ($existsInStore -and $isOverwrite) {
                    $marker = "[REPLACE]"
                    $markerColor = "Yellow"
                    $detail = ""
                }
                else {
                    $marker = "[IMPORT]"
                    $markerColor = "White"
                    $detail = ""
                }

                Write-Host "        $marker $($cert.Subject) ($certType) -> ${targetScope}\${targetStore} $detail" -ForegroundColor $markerColor
                $cert.Dispose()
            }
        }
        catch {
            Write-Host "      [READ ERROR] Failed to read PFX: $_" -ForegroundColor Red
        }
        Write-Host ""
        continue
    }

    # Explicit mode: single certificate
    if (-not $isAutoRoute -and [string]::IsNullOrEmpty($item.StoreScope)) {
        Write-Host "  [$index] $displayName  [NO STORE]" -ForegroundColor Red
        Write-Host "      File: $($fileInfo.Resolved)" -ForegroundColor DarkGray
        Write-Host "      StoreScope/StoreName not specified (required when AutoRoute=0)" -ForegroundColor DarkGray
        Write-Host ""
        continue
    }

    $thumbprint = Get-CertificateThumbprint -Path $certPath -Password $item.Password
    if ($null -eq $thumbprint) {
        Write-Host "  [$index] $displayName  [READ ERROR]" -ForegroundColor Red
        Write-Host "      File: $($fileInfo.Resolved)" -ForegroundColor DarkGray
        Write-Host ""
        continue
    }

    $existsInStore = Test-CertificateInStore -Thumbprint $thumbprint -StoreScope $item.StoreScope -StoreName $item.StoreName

    if ($existsInStore -and -not $isOverwrite) {
        $marker = "[SKIP]"
        $markerColor = "Gray"
    }
    elseif ($existsInStore -and $isOverwrite) {
        $marker = "[REPLACE]"
        $markerColor = "Yellow"
    }
    else {
        $marker = "[IMPORT]"
        $markerColor = "White"
    }

    Write-Host "  [$index] $displayName  $marker" -ForegroundColor $markerColor
    if ($fileInfo.Original -ne $fileInfo.Resolved) {
        Write-Host "      File: $($fileInfo.Resolved) (matched: $($fileInfo.Original))" -ForegroundColor DarkGray
    }
    else {
        Write-Host "      File: $($fileInfo.Resolved) -> $($item.StoreScope)\$($item.StoreName)" -ForegroundColor DarkGray
    }
    Write-Host "      Thumbprint: $thumbprint" -ForegroundColor DarkGray
    if ($item.FriendlyName) {
        Write-Host "      FriendlyName: $($item.FriendlyName)" -ForegroundColor DarkGray
    }
    Write-Host ""
}

Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

# ========================================
# Step 4: Confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Import the above certificates?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""

# ========================================
# Step 5: Import certificates
# ========================================
$successCount = 0
$skipCount = 0
$failCount = 0
$total = $resolvedItems.Count
$current = 0

# Track imported certs for verification (array of hashtables: Thumbprint, StoreScope, StoreName, DisplayName)
$verifyTargets = @()

foreach ($ri in $resolvedItems) {
    $current++
    $item = $ri.Item
    $fileInfo = $ri.FileInfo
    $certPath = $ri.CertPath
    $isPfx = $ri.IsPfx
    $isAutoRoute = $ri.IsAutoRoute
    $isOverwrite = $ri.IsOverwrite
    $displayName = $ri.DisplayName

    Write-Host "[$current/$total] $displayName" -ForegroundColor Cyan

    # File existence check
    if ($null -eq $fileInfo.Resolved -or -not (Test-Path $certPath)) {
        Show-Error "Certificate file not found: $($fileInfo.Expanded)"
        $failCount++
        Write-Host ""
        continue
    }

    # PFX password check
    if ($isPfx -and [string]::IsNullOrEmpty($item.Password)) {
        Show-Error "Password required for PFX file: $($fileInfo.Resolved)"
        $failCount++
        Write-Host ""
        continue
    }

    # ----------------------------------------
    # AutoRoute mode
    # ----------------------------------------
    if ($isAutoRoute -and $isPfx) {
        try {
            $collection = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2Collection
            $flags = [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags](
                "Exportable, PersistKeySet, MachineKeySet"
            )
            $collection.Import($certPath, $item.Password, $flags)

            $subSuccess = 0
            $subSkip = 0
            $subFail = 0

            foreach ($cert in $collection) {
                # Detect CA certificates using X.509 Basic Constraints extension (OID 2.5.29.19)
                # This correctly identifies intermediate CAs (Subject != Issuer but CA=true)
                $basicConstraints = $cert.Extensions | Where-Object { $_.Oid.Value -eq "2.5.29.19" }
                $isCa = ($null -ne $basicConstraints -and $basicConstraints.CertificateAuthority)
                $targetScope = "LocalMachine"
                $targetStore = if ($isCa) { "Root" } else { "My" }
                $certType = if ($isCa) { "CA" } else { "Client" }
                $certDisplay = "$($cert.Subject) ($certType)"

                try {
                    $existsInStore = Test-CertificateInStore -Thumbprint $cert.Thumbprint -StoreScope $targetScope -StoreName $targetStore

                    if ($existsInStore -and -not $isOverwrite) {
                        Show-Skip "Already exists: $certDisplay"
                        $subSkip++
                        $verifyTargets += @{
                            Thumbprint  = $cert.Thumbprint
                            StoreScope  = $targetScope
                            StoreName   = $targetStore
                            DisplayName = $certDisplay
                        }
                        continue
                    }

                    $scope = [System.Security.Cryptography.X509Certificates.StoreLocation]$targetScope
                    $store = New-Object System.Security.Cryptography.X509Certificates.X509Store($targetStore, $scope)
                    $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)

                    # Remove existing if overwrite
                    if ($existsInStore -and $isOverwrite) {
                        $existing = $store.Certificates | Where-Object { $_.Thumbprint -eq $cert.Thumbprint }
                        if ($null -ne $existing) {
                            $store.Remove($existing)
                            Show-Info "Removed existing: $certDisplay"
                        }
                    }

                    # Set FriendlyName for client certs
                    if (-not $isCa -and $item.FriendlyName) {
                        $cert.FriendlyName = $item.FriendlyName
                    }

                    $store.Add($cert)
                    $store.Close()

                    Show-Success "Imported: $certDisplay -> ${targetScope}\${targetStore}"
                    $subSuccess++
                    $verifyTargets += @{
                        Thumbprint  = $cert.Thumbprint
                        StoreScope  = $targetScope
                        StoreName   = $targetStore
                        DisplayName = $certDisplay
                    }
                }
                catch {
                    Show-Error "Failed: $certDisplay : $_"
                    $subFail++
                }
                finally {
                    $cert.Dispose()
                }
            }

            $successCount += $subSuccess
            $skipCount += $subSkip
            $failCount += $subFail
        }
        catch {
            Show-Error "Failed to read PFX: $($fileInfo.Resolved) : $_"
            $failCount++
        }

        Write-Host ""
        continue
    }

    # ----------------------------------------
    # Explicit mode
    # ----------------------------------------

    # StoreScope/StoreName check
    if ([string]::IsNullOrEmpty($item.StoreScope) -or [string]::IsNullOrEmpty($item.StoreName)) {
        Show-Error "StoreScope/StoreName not specified for: $($fileInfo.Resolved)"
        $failCount++
        Write-Host ""
        continue
    }

    # Get thumbprint
    $thumbprint = Get-CertificateThumbprint -Path $certPath -Password $item.Password
    if ($null -eq $thumbprint) {
        Show-Error "Failed to read certificate: $($fileInfo.Resolved)"
        $failCount++
        Write-Host ""
        continue
    }

    # Idempotency check
    $existsInStore = Test-CertificateInStore -Thumbprint $thumbprint -StoreScope $item.StoreScope -StoreName $item.StoreName
    if ($existsInStore -and -not $isOverwrite) {
        Show-Skip "Already exists in $($item.StoreScope)\$($item.StoreName)"
        $skipCount++
        $verifyTargets += @{
            Thumbprint  = $thumbprint
            StoreScope  = $item.StoreScope
            StoreName   = $item.StoreName
            DisplayName = $displayName
        }
        Write-Host ""
        continue
    }

    try {
        $storeLocation = "Cert:\$($item.StoreScope)\$($item.StoreName)"

        # Remove existing if overwrite
        if ($existsInStore -and $isOverwrite) {
            $scope = [System.Security.Cryptography.X509Certificates.StoreLocation]($item.StoreScope)
            $store = New-Object System.Security.Cryptography.X509Certificates.X509Store($item.StoreName, $scope)
            $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
            $existing = $store.Certificates | Where-Object { $_.Thumbprint -eq $thumbprint }
            if ($null -ne $existing) {
                $store.Remove($existing)
                Show-Info "Removed existing certificate"
            }
            $store.Close()
        }

        # Import
        if ($isPfx) {
            $securePassword = ConvertTo-SecureString $item.Password -AsPlainText -Force
            $imported = Import-PfxCertificate -FilePath $certPath -CertStoreLocation $storeLocation `
                -Password $securePassword -Exportable -ErrorAction Stop

            # Set FriendlyName if specified
            if ($item.FriendlyName) {
                $imported.FriendlyName = $item.FriendlyName
            }
        }
        else {
            $imported = Import-Certificate -FilePath $certPath -CertStoreLocation $storeLocation -ErrorAction Stop
        }

        Show-Success "Imported: $displayName -> $($item.StoreScope)\$($item.StoreName)"
        $successCount++
        $verifyTargets += @{
            Thumbprint  = $thumbprint
            StoreScope  = $item.StoreScope
            StoreName   = $item.StoreName
            DisplayName = $displayName
        }
    }
    catch {
        Show-Error "Failed: $displayName : $_"
        $failCount++
    }

    Write-Host ""
}

# ========================================
# Step 5.5: Post-Apply Verification
# ========================================
if ($verifyTargets.Count -gt 0) {
    Write-Host ""
    Show-Info "Verifying imported certificates..."
    Write-Host ""

    $verifyPass = 0
    $verifyFail = 0

    foreach ($vt in $verifyTargets) {
        $found = Test-CertificateInStore -Thumbprint $vt.Thumbprint -StoreScope $vt.StoreScope -StoreName $vt.StoreName

        if ($found) {
            Write-Host "  [VERIFIED] $($vt.DisplayName) -> $($vt.StoreScope)\$($vt.StoreName)" -ForegroundColor Green
            $verifyPass++
        }
        else {
            Write-Host "  [VERIFY FAILED] $($vt.DisplayName) -> $($vt.StoreScope)\$($vt.StoreName)" -ForegroundColor Red
            $verifyFail++
        }
    }

    Write-Host ""
    $verified = ($verifyFail -eq 0)
}
else {
    $verified = $null
}

# ========================================
# Step 6: Result
# ========================================
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount `
    -Title "Certificate Config Results" -Verified $verified)
