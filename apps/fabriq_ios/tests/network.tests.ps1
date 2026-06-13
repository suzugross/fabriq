# Pester tests for the pure formatting/mapping helpers of ping/traceroute.
# Run: Invoke-Pester apps\fabriq_ios\tests\network.tests.ps1
#
# Real ICMP send / Esc-abort polling depend on the network and a live
# console, so they are verified manually. Only the deterministic, side-
# effect-free helpers are unit-tested here.

BeforeAll {
    . "$PSScriptRoot\..\lib\commands\network.ps1"
}

Describe 'Get-FabriqIosPingChar' {
    It 'maps Success to !' {
        Get-FabriqIosPingChar -StatusName 'Success' | Should -Be '!'
    }
    It 'maps TimedOut to .' {
        Get-FabriqIosPingChar -StatusName 'TimedOut' | Should -Be '.'
    }
    It 'maps host-unreachable to U' {
        Get-FabriqIosPingChar -StatusName 'DestinationHostUnreachable' | Should -Be 'U'
    }
    It 'maps network-unreachable to U' {
        Get-FabriqIosPingChar -StatusName 'DestinationNetworkUnreachable' | Should -Be 'U'
    }
    It 'maps unknown statuses to .' {
        Get-FabriqIosPingChar -StatusName 'PacketTooBig' | Should -Be '.'
    }
}

Describe 'Format-FabriqIosPingSummary' {
    It 'reports 100 percent with round-trip stats' {
        $r = Format-FabriqIosPingSummary -Sent 5 -Received 5 -Rtts @(1, 2, 3)
        $r | Should -Be 'Success rate is 100 percent (5/5), round-trip min/avg/max = 1/2/3 ms'
    }
    It 'reports a partial rate and rounds the average' {
        $r = Format-FabriqIosPingSummary -Sent 5 -Received 3 -Rtts @(2, 4, 9)
        $r | Should -Be 'Success rate is 60 percent (3/5), round-trip min/avg/max = 2/5/9 ms'
    }
    It 'omits round-trip stats at 0 percent' {
        $r = Format-FabriqIosPingSummary -Sent 5 -Received 0 -Rtts @()
        $r | Should -Be 'Success rate is 0 percent (0/5)'
    }
    It 'does not divide by zero when nothing was sent' {
        $r = Format-FabriqIosPingSummary -Sent 0 -Received 0 -Rtts @()
        $r | Should -Be 'Success rate is 0 percent (0/0)'
    }
}

Describe 'Format-FabriqIosTracerouteHop' {
    It 'renders a responding hop with three times' {
        $r = Format-FabriqIosTracerouteHop -Hop 1 -Address '192.168.1.1' -Times @(4, 2, 1)
        $r | Should -Be '  1 192.168.1.1 4 msec 2 msec 1 msec'
    }
    It 'renders timeouts as asterisks within a hop' {
        $r = Format-FabriqIosTracerouteHop -Hop 2 -Address '10.0.0.1' -Times @(8, $null, 9)
        $r | Should -Be '  2 10.0.0.1 8 msec * 9 msec'
    }
    It 'renders a fully timed-out hop with no address' {
        $r = Format-FabriqIosTracerouteHop -Hop 3 -Address '' -Times @($null, $null, $null)
        $r | Should -Match '^\s+3 \* \* \*$'
    }
}
