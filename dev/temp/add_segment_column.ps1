# ========================================
# Segment Column Batch Adder
# Adds ,Segment to header and trailing comma to data lines
# Preserves encoding (BOM/ANSI), line endings, quoting style
# ========================================

$targetFiles = @(
    "modules\extended\builtin_admin_config\builtin_admin.csv"
    "modules\extended\directory_cleaner\clean_list.csv"
    "modules\extended\display_config\display_list.csv"
    "modules\extended\dpi_config\dpi_list.csv"
    "modules\extended\edge_config\restore_dest.csv"
    "modules\extended\group_config\group_list.csv"
    "modules\extended\history_destroyer\ssid_list.csv"
    "modules\extended\ipv6_config\ipv6_list.csv"
    "modules\extended\manual_kitting_assistant\step_list.csv"
    "modules\extended\reg_template\reg_list.csv"
    "modules\extended\script_looper\looper_list.csv"
    "modules\standard\app_config\app_list.csv"
    "modules\standard\autologon_config\autologon_list.csv"
    "modules\standard\bitlocker_config\bitlocker_list.csv"
    "modules\standard\bloatware_remove\bloatware_list.csv"
    "modules\standard\brightness_config\brightness_list.csv"
    "modules\standard\domain_join\domain.csv"
    "modules\standard\dpi_api_config\dpi_list.csv"
    "modules\standard\fabriq_app_launcher\target_apps.csv"
    "modules\standard\firewall_config\firewall_list.csv"
    "modules\standard\generic_batch_runner\batch_list.csv"
    "modules\standard\generic_process_runner\process_list.csv"
    "modules\standard\local_user_config\local_user_list.csv"
    "modules\standard\odt_config\odt_list.csv"
    "modules\standard\power_config\power_list.csv"
    "modules\standard\printer_delete\printer_delete.csv"
    "modules\standard\process_killer\process_list.csv"
    "modules\standard\profile_delete\profile_list.csv"
    "modules\standard\reg_hkcu_config\reg_hkcu_list.csv"
    "modules\standard\reg_hklm_config\reg_hklm_list.csv"
    "modules\standard\resolution_api_config\resolution_list.csv"
    "modules\standard\storeapp_config\storeapp_list.csv"
    "modules\standard\wallpaper_config\wallpaper_list.csv"
    "modules\standard\windows_license_config\license_key.csv"
    "modules\standard\winget_install\app_list.csv"
)

$baseDir = Split-Path $PSScriptRoot -Parent | Split-Path -Parent
$successCount = 0
$failCount = 0

foreach ($relPath in $targetFiles) {
    $fullPath = Join-Path $baseDir $relPath
    if (-not (Test-Path $fullPath)) {
        Write-Host "[SKIP] Not found: $relPath" -ForegroundColor Yellow
        continue
    }

    try {
        # Read raw bytes to detect BOM and preserve encoding
        $rawBytes = [System.IO.File]::ReadAllBytes($fullPath)

        # Detect BOM (EF BB BF = UTF-8 BOM)
        $hasBom = ($rawBytes.Length -ge 3 -and $rawBytes[0] -eq 0xEF -and $rawBytes[1] -eq 0xBB -and $rawBytes[2] -eq 0xBF)
        $encoding = if ($hasBom) {
            New-Object System.Text.UTF8Encoding($true)
        } else {
            [System.Text.Encoding]::GetEncoding(932)  # Shift-JIS / ANSI
        }

        # Read content preserving encoding
        $content = $encoding.GetString($rawBytes)
        # Remove BOM char if present (U+FEFF)
        if ($content.Length -gt 0 -and $content[0] -eq [char]0xFEFF) {
            $content = $content.Substring(1)
        }

        # Detect line ending style
        $lineEnding = if ($content.Contains("`r`n")) { "`r`n" } else { "`n" }

        # Split into lines
        $textLines = @($content -split "\r?\n")

        # Remove trailing empty lines
        while ($textLines.Count -gt 0 -and [string]::IsNullOrWhiteSpace($textLines[$textLines.Count - 1])) {
            $textLines = $textLines[0..($textLines.Count - 2)]
        }

        if ($textLines.Count -eq 0) {
            Write-Host "[SKIP] Empty file: $relPath" -ForegroundColor Yellow
            continue
        }

        # --- Modify header ---
        $header = $textLines[0]

        # Detect quoted header style (e.g. "Enabled","LocalGroup",...)
        $isQuoted = ($header -match '^"[^"]*"')

        if ($isQuoted) {
            $textLines[0] = $header + ',"Segment"'
        } else {
            $textLines[0] = $header + ",Segment"
        }

        # --- Add trailing comma to data lines ---
        for ($i = 1; $i -lt $textLines.Count; $i++) {
            if (-not [string]::IsNullOrWhiteSpace($textLines[$i])) {
                $textLines[$i] = $textLines[$i] + ","
            }
        }

        # Rebuild content with original line ending + trailing newline
        $newContent = ($textLines -join $lineEnding) + $lineEnding

        # Write back with original encoding (BOM preserved via GetPreamble)
        $preamble = $encoding.GetPreamble()
        $bodyBytes = $encoding.GetBytes($newContent)
        $newBytes = New-Object byte[] ($preamble.Length + $bodyBytes.Length)
        [Array]::Copy($preamble, 0, $newBytes, 0, $preamble.Length)
        [Array]::Copy($bodyBytes, 0, $newBytes, $preamble.Length, $bodyBytes.Length)
        [System.IO.File]::WriteAllBytes($fullPath, $newBytes)

        $successCount++
        Write-Host "[OK] $relPath" -ForegroundColor Green
    }
    catch {
        $failCount++
        Write-Host "[ERROR] $relPath : $_" -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "Complete: Success=$successCount, Fail=$failCount" -ForegroundColor Cyan
