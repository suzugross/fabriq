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
    # `do` is intercepted at the REPL on an exact match and is not part
    # of Get-FabriqIosCommandVocabulary (so it never participates in
    # abbreviation), so render it explicitly for the configuration modes
    # where it applies. Its description is still sourced from help_text.csv.
    if ($Mode -in @('GlobalConfig', 'InterfaceConfig', 'ModuleConfig')) {
        $doRow = $script:FabriqIosHelpTable | Where-Object {
            $_.Mode -eq $Mode -and $_.Command -eq 'do'
        } | Select-Object -First 1
        $doHelp = if ($doRow) { $doRow.HelpText } else { 'Run a privileged EXEC command without leaving config mode' }
        Write-Host ("    {0,-12} {1}" -f 'do', $doHelp)
    }
    Write-Host ""
}
