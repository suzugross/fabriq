# Cisco IOS-style `ping` and `traceroute` over ICMP.
#
# These deliberately reproduce IOS defaults and output rather than the
# Windows ping.exe / tracert.exe behaviour:
#   ping       : 5 echoes, 2s timeout, 100-byte datagram, "!.U" result
#                string + "Success rate is N percent (r/s), round-trip
#                min/avg/max = a/b/c ms".
#   traceroute : 3 probes per hop, 3s timeout, max TTL 30, per-hop times
#                in "N msec" with "*" for no reply.
# Real IOS traceroute uses UDP on routers; here ICMP is sufficient and is
# implemented via System.Net.NetworkInformation.Ping with PingOptions.Ttl.
#
# Abort: instead of Ctrl+C (which would tear down the shell), each probe is
# issued asynchronously and the foreground polls the keyboard; pressing Esc
# stops the run cleanly. This mirrors IOS "Type escape sequence to abort".

# ---- Pure helpers (no console / network I/O; unit-tested) ----

function Get-FabriqIosPingChar {
    # Map an IPStatus name to the Cisco ping result character.
    param([string]$StatusName)
    switch -Wildcard ($StatusName) {
        'Success'       { return '!' }
        '*Unreachable*' { return 'U' }
        default         { return '.' }
    }
}

function Format-FabriqIosPingSummary {
    # Build the Cisco "Success rate ..." line. The round-trip clause is
    # only appended when at least one echo was received (matching IOS,
    # which omits min/avg/max at 0 percent).
    param(
        [int]$Sent,
        [int]$Received,
        [int[]]$Rtts
    )
    $pct = if ($Sent -gt 0) { [int][math]::Floor(($Received * 100) / $Sent) } else { 0 }
    $base = "Success rate is $pct percent ($Received/$Sent)"
    if ($Received -gt 0 -and $Rtts -and $Rtts.Count -gt 0) {
        $min = ($Rtts | Measure-Object -Minimum).Minimum
        $max = ($Rtts | Measure-Object -Maximum).Maximum
        $avg = [int][math]::Round((($Rtts | Measure-Object -Average).Average), [System.MidpointRounding]::AwayFromZero)
        return ('{0}, round-trip min/avg/max = {1}/{2}/{3} ms' -f $base, $min, $avg, $max)
    }
    return $base
}

function Format-FabriqIosTracerouteHop {
    # Render one traceroute hop line. $Times is an array whose elements are
    # either an integer (msec) or $null (no reply -> "*").
    param(
        [int]$Hop,
        [string]$Address,
        $Times
    )
    $cells = @(foreach ($t in $Times) {
        if ($null -eq $t) { '*' } else { ('{0} msec' -f [int]$t) }
    })
    if ([string]::IsNullOrEmpty($Address)) {
        return ('  {0} {1}' -f $Hop, ($cells -join ' '))
    }
    return ('  {0} {1} {2}' -f $Hop, $Address, ($cells -join ' '))
}

# ---- Network / console helpers ----

function Resolve-FabriqIosPingTarget {
    # Resolve a host/IP string to the first IPv4 address, or $null on
    # failure. An IP literal resolves to itself.
    param([string]$Target)
    try {
        $addrs = [System.Net.Dns]::GetHostAddresses($Target)
    } catch {
        return $null
    }
    return ($addrs | Where-Object { $_.AddressFamily -eq 'InterNetwork' } | Select-Object -First 1)
}

function Invoke-FabriqIosIcmpProbe {
    # Send a single ICMP echo asynchronously while polling for an Esc abort.
    # Returns @{ Reply = <PingReply|$null>; Aborted = $bool }. A faulted task
    # (e.g. transient ICMP error) yields a $null reply, which the callers
    # render as a failed probe.
    param(
        [System.Net.NetworkInformation.Ping]$Ping,
        $Target,
        [int]$TimeoutMs,
        [byte[]]$Buffer,
        [System.Net.NetworkInformation.PingOptions]$Options
    )
    $task = if ($Options) {
        $Ping.SendPingAsync($Target, $TimeoutMs, $Buffer, $Options)
    } else {
        $Ping.SendPingAsync($Target, $TimeoutMs, $Buffer)
    }
    while (-not $task.IsCompleted) {
        if ([Console]::KeyAvailable) {
            $k = [Console]::ReadKey($true)
            if ($k.Key -eq [ConsoleKey]::Escape) {
                return @{ Reply = $null; Aborted = $true }
            }
        }
        Start-Sleep -Milliseconds 50
    }
    if ($task.IsFaulted) {
        return @{ Reply = $null; Aborted = $false }
    }
    return @{ Reply = $task.Result; Aborted = $false }
}

# ---- Commands ----

function Invoke-FabriqIosPing {
    # Cisco-style ping: 5 echoes, 2s timeout, 100-byte datagram.
    param(
        [hashtable]$State,
        [string[]]$ArgList
    )
    if (-not $ArgList -or $ArgList.Count -lt 1) {
        Write-Host "% Incomplete command. Usage: 'ping <host>'" -ForegroundColor Red
        return
    }
    $ip = Resolve-FabriqIosPingTarget -Target $ArgList[0]
    if (-not $ip) {
        Write-Host '% Unrecognized host or address.' -ForegroundColor Red
        return
    }

    $count = 5; $timeoutMs = 2000; $size = 100
    Write-Host ''
    Write-Host 'Type escape sequence to abort.'
    Write-Host ('Sending {0}, {1}-byte ICMP Echos to {2}, timeout is {3} seconds:' -f `
        $count, $size, $ip, ($timeoutMs / 1000))

    $buffer = New-Object byte[] $size
    $ping   = New-Object System.Net.NetworkInformation.Ping
    $sent = 0; $received = 0; $rtts = @(); $aborted = $false
    try {
        for ($i = 0; $i -lt $count; $i++) {
            $probe = Invoke-FabriqIosIcmpProbe -Ping $ping -Target $ip -TimeoutMs $timeoutMs -Buffer $buffer -Options $null
            if ($probe.Aborted) { $aborted = $true; break }
            $sent++
            $reply = $probe.Reply
            $statusName = if ($reply) { [string]$reply.Status } else { 'TimedOut' }
            Write-Host (Get-FabriqIosPingChar -StatusName $statusName) -NoNewline
            if ($reply -and $reply.Status -eq [System.Net.NetworkInformation.IPStatus]::Success) {
                $received++
                $rtts += [int]$reply.RoundtripTime
            }
        }
    } finally {
        $ping.Dispose()
    }
    Write-Host ''
    Write-Host (Format-FabriqIosPingSummary -Sent $sent -Received $received -Rtts $rtts)
}

function Invoke-FabriqIosTraceroute {
    # Cisco-style traceroute: 3 probes per hop, 3s timeout, max TTL 30.
    param(
        [hashtable]$State,
        [string[]]$ArgList
    )
    if (-not $ArgList -or $ArgList.Count -lt 1) {
        Write-Host "% Incomplete command. Usage: 'traceroute <host>'" -ForegroundColor Red
        return
    }
    $ip = Resolve-FabriqIosPingTarget -Target $ArgList[0]
    if (-not $ip) {
        Write-Host '% Unrecognized host or address.' -ForegroundColor Red
        return
    }

    $maxTtl = 30; $probes = 3; $timeoutMs = 3000; $size = 100
    Write-Host ''
    Write-Host 'Type escape sequence to abort.'
    Write-Host ('Tracing the route to {0}' -f $ip)
    Write-Host ''

    $buffer  = New-Object byte[] $size
    $ping    = New-Object System.Net.NetworkInformation.Ping
    $aborted = $false; $reached = $false
    try {
        for ($ttl = 1; $ttl -le $maxTtl -and -not $reached -and -not $aborted; $ttl++) {
            $opts    = New-Object System.Net.NetworkInformation.PingOptions($ttl, $false)
            $hopAddr = $null
            $times   = @()
            for ($p = 0; $p -lt $probes; $p++) {
                $probe = Invoke-FabriqIosIcmpProbe -Ping $ping -Target $ip -TimeoutMs $timeoutMs -Buffer $buffer -Options $opts
                if ($probe.Aborted) { $aborted = $true; break }
                $reply = $probe.Reply
                if ($reply -and (
                        $reply.Status -eq [System.Net.NetworkInformation.IPStatus]::Success -or
                        $reply.Status -eq [System.Net.NetworkInformation.IPStatus]::TtlExpired)) {
                    if (-not $hopAddr) { $hopAddr = [string]$reply.Address }
                    $times += [int]$reply.RoundtripTime
                    if ($reply.Status -eq [System.Net.NetworkInformation.IPStatus]::Success) { $reached = $true }
                } else {
                    $times += $null
                }
            }
            if ($aborted) { break }
            Write-Host (Format-FabriqIosTracerouteHop -Hop $ttl -Address $hopAddr -Times $times)
        }
    } finally {
        $ping.Dispose()
    }
}
