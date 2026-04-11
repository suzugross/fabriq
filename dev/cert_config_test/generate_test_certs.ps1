# Generate test certificates for cert_config module verification
# Run this script as Administrator in PowerShell
# Output: test certs are generated in this directory (dev/cert_config_test/)
# To use: copy desired certs to modules/standard/cert_config/certs/

$certsDir = $PSScriptRoot

# Step 1: Root CA (self-signed, BasicConstraints CA=true)
Write-Host "Creating Root CA..." -ForegroundColor Cyan
$rootCa = New-SelfSignedCertificate `
    -Subject "CN=Fabriq Test Root CA" `
    -CertStoreLocation "Cert:\CurrentUser\My" `
    -KeyUsage CertSign, CRLSign `
    -KeyExportPolicy Exportable `
    -NotAfter (Get-Date).AddYears(10) `
    -TextExtension @("2.5.29.19={critical}{text}ca=TRUE")
Write-Host "  Thumbprint: $($rootCa.Thumbprint)"

# Step 2: Intermediate CA (signed by Root, BasicConstraints CA=true)
Write-Host "Creating Intermediate CA..." -ForegroundColor Cyan
$intermediateCa = New-SelfSignedCertificate `
    -Subject "CN=Fabriq Test Intermediate CA" `
    -CertStoreLocation "Cert:\CurrentUser\My" `
    -Signer $rootCa `
    -KeyUsage CertSign, CRLSign `
    -KeyExportPolicy Exportable `
    -NotAfter (Get-Date).AddYears(5) `
    -TextExtension @("2.5.29.19={critical}{text}ca=TRUE")
Write-Host "  Thumbprint: $($intermediateCa.Thumbprint)"

# Step 3: Client cert signed by Intermediate CA (for 3-tier chain)
Write-Host "Creating Client cert (3-tier chain)..." -ForegroundColor Cyan
$clientCert3Tier = New-SelfSignedCertificate `
    -Subject "CN=Fabriq Test Client 3Tier" `
    -CertStoreLocation "Cert:\CurrentUser\My" `
    -Signer $intermediateCa `
    -KeyUsage DigitalSignature, KeyEncipherment `
    -KeyExportPolicy Exportable `
    -NotAfter (Get-Date).AddYears(2)
Write-Host "  Thumbprint: $($clientCert3Tier.Thumbprint)"

# Step 4: Client cert signed by Root CA directly (for 2-tier chain)
Write-Host "Creating Client cert (2-tier chain)..." -ForegroundColor Cyan
$clientCert2Tier = New-SelfSignedCertificate `
    -Subject "CN=Fabriq Test Client 2Tier" `
    -CertStoreLocation "Cert:\CurrentUser\My" `
    -Signer $rootCa `
    -KeyUsage DigitalSignature, KeyEncipherment `
    -KeyExportPolicy Exportable `
    -NotAfter (Get-Date).AddYears(2)
Write-Host "  Thumbprint: $($clientCert2Tier.Thumbprint)"

$password = ConvertTo-SecureString "test1234" -AsPlainText -Force

# Export 1: Root CA public key (.cer)
Write-Host "`nExporting test_root_ca.cer..." -ForegroundColor Yellow
Export-Certificate -Cert $rootCa -FilePath "$certsDir\test_root_ca.cer" | Out-Null

# Export 2: Client cert only (.pfx)
Write-Host "Exporting test_client.pfx..." -ForegroundColor Yellow
Export-PfxCertificate -Cert $clientCert2Tier -FilePath "$certsDir\test_client.pfx" -Password $password | Out-Null

# Export 3: 2-tier PFX (Root CA + Client)
Write-Host "Exporting test_autoroute_2tier.pfx..." -ForegroundColor Yellow
$col2 = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2Collection
$col2.Add($rootCa) | Out-Null
$col2.Add($clientCert2Tier) | Out-Null
$pfxBytes2 = $col2.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Pfx, "test1234")
[System.IO.File]::WriteAllBytes("$certsDir\test_autoroute_2tier.pfx", $pfxBytes2)

# Export 4: 3-tier PFX (Root CA + Intermediate CA + Client)
Write-Host "Exporting test_autoroute_3tier.pfx..." -ForegroundColor Yellow
$col3 = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2Collection
$col3.Add($rootCa) | Out-Null
$col3.Add($intermediateCa) | Out-Null
$col3.Add($clientCert3Tier) | Out-Null
$pfxBytes3 = $col3.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Pfx, "test1234")
[System.IO.File]::WriteAllBytes("$certsDir\test_autoroute_3tier.pfx", $pfxBytes3)

# Export 5: PC name match PFX (same as 3-tier, named with COMPUTERNAME)
$pcFileName = "${env:COMPUTERNAME}_TEST.pfx"
Write-Host "Exporting $pcFileName..." -ForegroundColor Yellow
[System.IO.File]::WriteAllBytes("$certsDir\$pcFileName", $pfxBytes3)

# Cleanup: remove temp certs from CurrentUser\My
Write-Host "`nCleaning up temporary certificates..." -ForegroundColor Gray
Remove-Item "Cert:\CurrentUser\My\$($rootCa.Thumbprint)" -Force
Remove-Item "Cert:\CurrentUser\My\$($intermediateCa.Thumbprint)" -Force
Remove-Item "Cert:\CurrentUser\My\$($clientCert3Tier.Thumbprint)" -Force
Remove-Item "Cert:\CurrentUser\My\$($clientCert2Tier.Thumbprint)" -Force

Write-Host "`nDone! Generated files:" -ForegroundColor Green
Get-ChildItem $certsDir -File | ForEach-Object {
    Write-Host "  $($_.Name) ($($_.Length) bytes)"
}
