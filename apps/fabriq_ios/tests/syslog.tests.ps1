# Pester tests for the syslog formatter.
# Run: Invoke-Pester apps\fabriq_ios\tests\syslog.tests.ps1

BeforeAll {
    . "$PSScriptRoot\..\lib\syslog.ps1"
}

Describe 'Get-FabriqIosSyslogTimestamp' {
    It 'matches the Cisco timestamp format' {
        $ts = Get-FabriqIosSyslogTimestamp
        $ts | Should -Match '^\*[A-Z][a-z]{2}\s+\d{1,2}\s+\d{2}:\d{2}:\d{2}\.\d{3}$'
    }
}

Describe 'Get-FabriqIosSyslogTemplate' {
    It 'finds an existing row' {
        $t = Get-FabriqIosSyslogTemplate -Mnemonic 'HOSTNAME' -Severity 5 -Key 'success'
        $t | Should -Not -BeNullOrEmpty
        $t | Should -BeLike '*{NewName}*'
    }

    It 'returns null for missing key' {
        Get-FabriqIosSyslogTemplate -Mnemonic 'HOSTNAME' -Severity 5 -Key 'nonexistent' `
            | Should -BeNullOrEmpty
    }
}

Describe 'Format-FabriqIosSyslogLine' {
    It 'substitutes placeholders' {
        $line = Format-FabriqIosSyslogLine -Severity 5 -Mnemonic 'HOSTNAME' -Key 'success' `
                                           -Placeholders @{ NewName = 'NEW-PC-01' }
        $line | Should -BeLike '*NEW-PC-01*'
        $line | Should -Not -BeLike '*{NewName}*'
    }

    It 'includes the FABRIQ facility prefix' {
        $line = Format-FabriqIosSyslogLine -Severity 5 -Mnemonic 'HOSTNAME' -Key 'success' `
                                           -Placeholders @{ NewName = 'X' }
        $line | Should -BeLike '*%FABRIQ-5-HOSTNAME*'
    }

    It 'falls back gracefully on missing template' {
        $line = Format-FabriqIosSyslogLine -Severity 5 -Mnemonic 'HOSTNAME' -Key 'nonexistent' `
                                           -Placeholders @{}
        $line | Should -BeLike '*missing template*'
    }
}
