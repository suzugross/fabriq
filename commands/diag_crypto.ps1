# ========================================
# Diagnostic: Crypto / Passphrase Check
# ========================================

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Crypto Diagnostic" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# 1. Check passphrase
$pp = $global:FabriqMasterPassphrase
if ([string]::IsNullOrWhiteSpace($pp)) {
    Write-Host "[1] FabriqMasterPassphrase: NOT SET (null/empty)" -ForegroundColor Red
} else {
    Write-Host "[1] FabriqMasterPassphrase: SET (length=$($pp.Length))" -ForegroundColor Green
}

# 2. Check function availability
Write-Host ""
$funcs = @("Unprotect-FabriqValue", "Import-ModuleCsv", "Import-CsvSafe")
foreach ($fn in $funcs) {
    if (Get-Command $fn -ErrorAction SilentlyContinue) {
        Write-Host "[2] Function '$fn': Available" -ForegroundColor Green
    } else {
        Write-Host "[2] Function '$fn': NOT FOUND" -ForegroundColor Red
    }
}

# 3. Find CSV files containing ENC: values
Write-Host ""
Write-Host "[3] Scanning for CSV files with ENC: values..." -ForegroundColor White
$fabriqRoot = (Resolve-Path ".").Path
$csvFiles = Get-ChildItem -Path $fabriqRoot -Filter "*.csv" -Recurse -ErrorAction SilentlyContinue
$encFound = $false
foreach ($csv in $csvFiles) {
    try {
        $raw = Get-Content $csv.FullName -Raw -ErrorAction SilentlyContinue
        if ($raw -match "ENC:") {
            $encFound = $true
            $relativePath = $csv.FullName.Replace($fabriqRoot, ".")
            Write-Host "    Found: $relativePath" -ForegroundColor Yellow

            # Try loading with Import-ModuleCsv and check if decryption happened
            if (-not [string]::IsNullOrWhiteSpace($pp)) {
                $data = @(Import-Csv -Path $csv.FullName -Encoding Default)
                $stillEncrypted = $false
                foreach ($row in $data) {
                    foreach ($prop in $row.PSObject.Properties) {
                        if ($prop.Value -is [string] -and $prop.Value.StartsWith("ENC:")) {
                            $stillEncrypted = $true
                            Write-Host "      Column '$($prop.Name)' = $($prop.Value.Substring(0, [Math]::Min(30, $prop.Value.Length)))..." -ForegroundColor DarkGray

                            # Direct decryption test
                            try {
                                $decrypted = Unprotect-FabriqValue -EncryptedValue $prop.Value -Passphrase $pp
                                Write-Host "      -> Decrypt OK (length=$($decrypted.Length))" -ForegroundColor Green
                            }
                            catch {
                                Write-Host "      -> Decrypt FAILED: $_" -ForegroundColor Red
                            }
                        }
                    }
                }
                if (-not $stillEncrypted) {
                    Write-Host "      (No ENC: values in data rows - possibly in header or comment)" -ForegroundColor DarkGray
                }
            }
        }
    }
    catch { }
}

if (-not $encFound) {
    Write-Host "    No CSV files with ENC: values found" -ForegroundColor DarkGray
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Diagnostic Complete" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
