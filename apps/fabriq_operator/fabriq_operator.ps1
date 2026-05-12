# ========================================
# Fabriq Operator - GUI Entry Point
# ========================================
# Provides GUI forms for fabriq session setup
# and main dashboard navigation.
# Loaded by kernel/main.ps1 via dot-source.
# ========================================

try {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    [System.Windows.Forms.Application]::EnableVisualStyles()

    # Load GUI library files
    $operatorLibDir = Join-Path $PSScriptRoot "lib"
    . (Join-Path $operatorLibDir "theme.ps1")
    . (Join-Path $operatorLibDir "session_form.ps1")
    . (Join-Path $operatorLibDir "apps_dialog.ps1")
    . (Join-Path $operatorLibDir "quickactions_dialog.ps1")
    . (Join-Path $operatorLibDir "dashboard_form.ps1")
    . (Join-Path $operatorLibDir "flex_dashboard.ps1")
    . (Join-Path $operatorLibDir "execution_toolbar.ps1")

    $script:UseGui = $true
}
catch {
    Write-Host "[ERROR] Failed to load fabriq_operator GUI: $_" -ForegroundColor Red
    $script:UseGui = $false
}
