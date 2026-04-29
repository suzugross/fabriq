# Cisco IOS-style syslog formatter with surrealist message bodies.
# Templates live in data/syslog_messages.csv keyed by (Mnemonic, Severity, Key).

$script:FabriqIosSyslogTable = $null

function Get-FabriqIosSyslogTimestamp {
    # Returns a Cisco-style timestamp: '*Apr 29 14:23:01.234'.
    # Month abbreviation is forced to English (locale-independent).
    $months = @('Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec')
    $now = Get-Date
    return ('*{0} {1,2} {2:HH:mm:ss}.{3:000}' -f $months[$now.Month - 1], $now.Day, $now, $now.Millisecond)
}

function Get-FabriqIosSyslogTemplate {
    param(
        [string]$Mnemonic,
        [int]$Severity,
        [string]$Key
    )
    if ($null -eq $script:FabriqIosSyslogTable) {
        $csvPath = Join-Path $PSScriptRoot '..\data\syslog_messages.csv'
        if (-not (Test-Path $csvPath)) {
            throw "syslog_messages.csv not found at: $csvPath"
        }
        $script:FabriqIosSyslogTable = @(Import-Csv -Path $csvPath -Encoding UTF8)
    }
    $row = $script:FabriqIosSyslogTable | Where-Object {
        $_.Mnemonic -eq $Mnemonic -and
        [int]$_.Severity -eq $Severity -and
        $_.Key -eq $Key
    } | Select-Object -First 1
    if (-not $row) { return $null }
    return $row.Template
}

function Format-FabriqIosSyslogLine {
    param(
        [int]$Severity,
        [string]$Mnemonic,
        [string]$Key,
        [hashtable]$Placeholders
    )
    # Pure function returning the formatted line. No console I/O so
    # tests can assert directly against the returned string.
    $template = Get-FabriqIosSyslogTemplate -Mnemonic $Mnemonic -Severity $Severity -Key $Key
    if ($null -eq $template) {
        $template = "(missing template: %FABRIQ-$Severity-$Mnemonic key=$Key)"
    }
    if ($Placeholders) {
        foreach ($k in $Placeholders.Keys) {
            $template = $template.Replace("{$k}", "$($Placeholders[$k])")
        }
    }
    $ts = Get-FabriqIosSyslogTimestamp
    return ('{0}: %FABRIQ-{1}-{2}: {3}' -f $ts, $Severity, $Mnemonic, $template)
}

function Write-FabriqIosSyslog {
    param(
        [int]$Severity,
        [string]$Mnemonic,
        [string]$Key,
        [hashtable]$Placeholders
    )
    $line = Format-FabriqIosSyslogLine -Severity $Severity -Mnemonic $Mnemonic `
                                       -Key $Key -Placeholders $Placeholders
    $color = switch ($Severity) {
        { $_ -le 3 } { 'Red' }
        4            { 'Yellow' }
        5            { 'Cyan' }
        6            { 'Gray' }
        7            { 'DarkGray' }
        default      { 'White' }
    }
    Write-Host $line -ForegroundColor $color
}
