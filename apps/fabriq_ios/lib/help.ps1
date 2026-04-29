# '?' / 'help' command renderer.
# Reads data/help_text.csv (cached) and prints the per-mode
# vocabulary returned by Get-FabriqIosCommandVocabulary.

$script:FabriqIosHelpTable = $null

function Show-FabriqIosHelp {
    param([string]$Mode)
    if ($null -eq $script:FabriqIosHelpTable) {
        $csvPath = Join-Path $PSScriptRoot '..\data\help_text.csv'
        if (Test-Path $csvPath) {
            $script:FabriqIosHelpTable = @(Import-Csv -Path $csvPath -Encoding UTF8)
        } else {
            $script:FabriqIosHelpTable = @()
        }
    }
    $vocab = Get-FabriqIosCommandVocabulary -Mode $Mode
    Write-Host ""
    Write-Host ("  Available commands ({0}):" -f $Mode)
    foreach ($cmd in $vocab) {
        $row = $script:FabriqIosHelpTable | Where-Object {
            $_.Mode -eq $Mode -and $_.Command -eq $cmd
        } | Select-Object -First 1
        $help = if ($row) {
            $row.HelpText
        } elseif ($cmd -eq '?' -or $cmd -eq 'help') {
            'Display this help'
        } else {
            '(no description)'
        }
        Write-Host ("    {0,-12} {1}" -f $cmd, $help)
    }
    Write-Host ""
}
