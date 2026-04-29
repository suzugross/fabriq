# Command line parser: tokenization, abbreviation expansion,
# mode-scoped command resolution.

function Get-FabriqIosCommandVocabulary {
    param([string]$Mode)
    switch ($Mode) {
        'UserExec'        { return @('enable','show','help','exit','?') }
        'PrivilegedExec'  { return @('show','configure','reload','disable','exit','help','?') }
        'GlobalConfig'    { return @('hostname','interface','module','cleanup','copy','install','script','exit','end','help','?') }
        'InterfaceConfig' { return @('ip','exit','end','help','?') }
        'ModuleConfig'    { return @('set','add','show','exit','end','help','?') }
        default           { return @() }
    }
}

function Get-FabriqIosSubVocabulary {
    param(
        [string]$Parent,
        [string]$Mode
    )
    $key = "$Mode.$Parent"
    switch ($key) {
        'UserExec.show'            { return @('version','manifesto') }
        'PrivilegedExec.show'      { return @('version','running-config','profiles','modules','evidence','manifesto') }
        'PrivilegedExec.configure' { return @('terminal') }
        'InterfaceConfig.ip'       { return @('address') }
        default                    { return @() }
    }
}

function ConvertTo-FabriqIosTokens {
    param([string]$Line)
    if ([string]::IsNullOrWhiteSpace($Line)) {
        return ,@()
    }
    $tokens  = New-Object System.Collections.Generic.List[string]
    $current = New-Object System.Text.StringBuilder
    $inQuote = $false
    foreach ($ch in $Line.ToCharArray()) {
        if ($ch -eq '"') {
            $inQuote = -not $inQuote
            continue
        }
        if (-not $inQuote -and ($ch -eq ' ' -or $ch -eq "`t")) {
            if ($current.Length -gt 0) {
                [void]$tokens.Add($current.ToString())
                [void]$current.Clear()
            }
            continue
        }
        [void]$current.Append($ch)
    }
    if ($current.Length -gt 0) {
        [void]$tokens.Add($current.ToString())
    }
    return ,$tokens.ToArray()
}

function Resolve-Token {
    param(
        [string[]]$Vocabulary,
        [string]$Token
    )
    # Cisco-style prefix resolution with exact-match precedence
    # so 'host' resolves to 'host' rather than ambiguous with 'hosts'.
    $exact = @($Vocabulary | Where-Object { $_ -ieq $Token })
    if ($exact.Count -eq 1) {
        return @{ Match = $exact[0]; Ambiguous = $false; Candidates = $exact }
    }
    $prefix = @($Vocabulary | Where-Object {
        $_.StartsWith($Token, [System.StringComparison]::OrdinalIgnoreCase)
    })
    if ($prefix.Count -eq 0) {
        return @{ Match = $null; Ambiguous = $false; Candidates = @() }
    }
    if ($prefix.Count -eq 1) {
        return @{ Match = $prefix[0]; Ambiguous = $false; Candidates = $prefix }
    }
    return @{ Match = $null; Ambiguous = $true; Candidates = $prefix }
}

function Expand-FabriqIosAbbreviation {
    param(
        [string[]]$Tokens,
        [string]$Mode
    )
    if ($null -eq $Tokens -or $Tokens.Count -eq 0) {
        return ,@()
    }

    $vocab = Get-FabriqIosCommandVocabulary -Mode $Mode
    $first = Resolve-Token -Vocabulary $vocab -Token $Tokens[0]

    if ($first.Ambiguous) {
        throw "Ambiguous command: '$($Tokens[0])' matches: $($first.Candidates -join ', ')"
    }
    if (-not $first.Match) {
        # Unknown - return verbatim so Resolve can report it.
        return ,$Tokens
    }

    $expanded = New-Object System.Collections.Generic.List[string]
    [void]$expanded.Add($first.Match)

    if ($Tokens.Count -ge 2) {
        $subVocab = Get-FabriqIosSubVocabulary -Parent $first.Match -Mode $Mode
        if ($subVocab.Count -gt 0) {
            $second = Resolve-Token -Vocabulary $subVocab -Token $Tokens[1]
            if ($second.Ambiguous) {
                throw "Ambiguous sub-command: '$($Tokens[1])' matches: $($second.Candidates -join ', ')"
            }
            if ($second.Match) {
                [void]$expanded.Add($second.Match)
            } else {
                [void]$expanded.Add($Tokens[1])
            }
        } else {
            [void]$expanded.Add($Tokens[1])
        }
        # Remaining tokens are dynamic args (hostname value, ip,
        # interface alias, etc.). Pass through verbatim.
        if ($Tokens.Count -gt 2) {
            for ($i = 2; $i -lt $Tokens.Count; $i++) {
                [void]$expanded.Add($Tokens[$i])
            }
        }
    }

    return ,$expanded.ToArray()
}

function Resolve-FabriqIosCommand {
    param(
        [string[]]$Tokens,
        [string]$Mode
    )
    if ($null -eq $Tokens -or $Tokens.Count -eq 0) {
        return $null
    }

    try {
        $expanded = Expand-FabriqIosAbbreviation -Tokens $Tokens -Mode $Mode
    } catch {
        return @{
            Command = $null
            Args    = @()
            Error   = $_.Exception.Message
        }
    }

    $vocab = Get-FabriqIosCommandVocabulary -Mode $Mode
    if ($expanded[0] -notin $vocab) {
        return @{
            Command = $null
            Args    = @()
            Error   = "Unknown command: $($Tokens[0])"
        }
    }

    $resolvedArgs = @()
    if ($expanded.Count -gt 1) {
        $resolvedArgs = @($expanded[1..($expanded.Count - 1)])
    }

    return @{
        Command = $expanded[0]
        Args    = $resolvedArgs
        Error   = $null
    }
}
