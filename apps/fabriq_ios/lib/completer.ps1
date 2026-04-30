# PSReadLine integration plus a pure completion engine.
# Subprocess-isolated: handlers vanish with the child process so the
# parent shell never sees them.

$global:_FabriqIosShellState = $null

function Get-FabriqIosCompletion {
    # Pure function. Returns the candidate string[] for the cursor
    # position, filtered by the prefix being typed. Testable without
    # PSReadLine.
    param(
        [string]$Line,
        [int]$Position,
        [string]$Mode,
        [hashtable]$State
    )

    if ($null -eq $Line) { $Line = '' }
    if ($Position -lt 0) { $Position = 0 }
    if ($Position -gt $Line.Length) { $Position = $Line.Length }
    $upToCursor = $Line.Substring(0, $Position)

    $tokens = ConvertTo-FabriqIosTokens $upToCursor
    $endsWithSpace = ($upToCursor.Length -eq 0) -or ($upToCursor[$upToCursor.Length - 1] -match '\s')

    if ($tokens.Count -eq 0) {
        $current = ''
        $tokensBefore = @()
    } elseif ($endsWithSpace) {
        $current = ''
        $tokensBefore = $tokens
    } else {
        $current = $tokens[-1]
        # NOTE: if-as-RHS unwraps single-element arrays in PS 5.1, so
        # split into a direct assignment to preserve array type.
        $tokensBefore = @()
        if ($tokens.Count -gt 1) {
            $tokensBefore = @($tokens[0..($tokens.Count - 2)])
        }
    }

    $candidates = @()
    if ($tokensBefore.Count -eq 0) {
        $candidates = @(Get-FabriqIosCommandVocabulary -Mode $Mode)
    } else {
        $vocab = Get-FabriqIosCommandVocabulary -Mode $Mode
        $first = Resolve-Token -Vocabulary $vocab -Token $tokensBefore[0]
        if (-not $first.Match) {
            return ,@()
        }
        if ($tokensBefore.Count -eq 1) {
            $candidates = @(Get-DynamicCompletion -Source ($first.Match) -Mode $Mode -State $State)
        } elseif ($first.Match -in @('set','add')) {
            # set/add are N-deep alternating col/val sequences inside
            # ModuleConfig, not 2-token sub-vocabulary verbs. Position
            # parity decides the candidate kind: even args after verb
            # -> next is a column name; odd -> next is an enum value
            # for the column at $tokensBefore[-1].
            $candidates = @(Get-SetAddPositionalCompletion -TokensBefore $tokensBefore -State $State)
        } else {
            $subVocab = @(Get-FabriqIosSubVocabulary -Parent $first.Match -Mode $Mode)
            $second = $null
            if ($subVocab.Count -gt 0) {
                $second = Resolve-Token -Vocabulary $subVocab -Token $tokensBefore[1]
            }
            if ($second -and $second.Match) {
                $key = '{0}.{1}' -f $first.Match, $second.Match
                $candidates = @(Get-DynamicCompletion -Source $key -Mode $Mode -State $State)
            } else {
                $candidates = @()
            }
        }
    }

    if (-not $candidates) { return ,@() }

    if ($current) {
        $candidates = @($candidates | Where-Object {
            $_ -and $_.StartsWith($current, [System.StringComparison]::OrdinalIgnoreCase)
        })
    }

    return ,@($candidates | Where-Object { $_ } | Sort-Object -Unique)
}

function Get-SetAddPositionalCompletion {
    # Pure helper for ModuleConfig `set`/`add` completion at any
    # position past the first column. Returns column names when the
    # next token should be a column, or enum values for the previous
    # column when the next token should be a value. Already-named
    # columns (case-insensitive) are filtered out so `set Foo 1 ?`
    # does not re-suggest Foo.
    param(
        [string[]]$TokensBefore,
        [hashtable]$State
    )
    if (-not ($State -and $State.ConfigModuleSchema)) { return @() }
    $argsAfterVerb = $TokensBefore.Count - 1
    if (($argsAfterVerb % 2) -eq 0) {
        # Next position is a column name. Drop columns already used.
        $used = New-Object System.Collections.Generic.HashSet[string] ([System.StringComparer]::OrdinalIgnoreCase)
        for ($i = 1; $i -lt $TokensBefore.Count; $i += 2) {
            [void]$used.Add($TokensBefore[$i])
        }
        return @($State.ConfigModuleSchema.Columns | Where-Object { -not $used.Contains($_) })
    } else {
        # Next position is a value for the column at TokensBefore[-1].
        $col = $TokensBefore[-1]
        $enums = $State.ConfigModuleSchema.Enums
        if ($enums -and $enums.ContainsKey($col)) {
            return @($enums[$col])
        }
        return @()
    }
}

function Get-DynamicCompletion {
    param(
        [string]$Source,
        [string]$Mode,
        [hashtable]$State
    )
    switch ($Source) {
        'show'       { return @(Get-FabriqIosSubVocabulary -Parent 'show'       -Mode $Mode) }
        'configure'  { return @(Get-FabriqIosSubVocabulary -Parent 'configure'  -Mode $Mode) }
        'ip'         { return @(Get-FabriqIosSubVocabulary -Parent 'ip'         -Mode $Mode) }
        'interface'  { return @(Get-InterfaceCompletionFromAdapters) }
        # Phase 9: each category verb uses its own JSON-defined module
        # list (apps/fabriq_ios/data/module_categories.json).
        'module'     { return @(Get-CategoryModuleCompletion -CategoryId 'settings') }
        'cleanup'    { return @(Get-CategoryModuleCompletion -CategoryId 'cleanup') }
        'copy'       { return @(Get-CategoryModuleCompletion -CategoryId 'copy') }
        'install'    { return @(Get-CategoryModuleCompletion -CategoryId 'install') }
        'script'     { return @(Get-CategoryModuleCompletion -CategoryId 'scripting') }
        'set' {
            # Inside (config-mod)# the first arg of `set` is a column
            # name from the bound module's schema.
            $s = $global:_FabriqIosShellState
            if ($s -and $s.ConfigModuleSchema) {
                return @($s.ConfigModuleSchema.Columns)
            }
            return @()
        }
        'add' {
            $s = $global:_FabriqIosShellState
            if ($s -and $s.ConfigModuleSchema) {
                return @($s.ConfigModuleSchema.Columns)
            }
            return @()
        }
        default      { return @() }
    }
}

function Get-CommonPrefix {
    param([string[]]$Strings)
    if (-not $Strings -or $Strings.Count -eq 0) { return '' }
    if ($Strings.Count -eq 1) { return $Strings[0] }
    $prefix = $Strings[0]
    for ($i = 1; $i -lt $Strings.Count; $i++) {
        $s = $Strings[$i]
        while ($prefix -and -not $s.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
            $prefix = $prefix.Substring(0, $prefix.Length - 1)
        }
        if (-not $prefix) { return '' }
    }
    return $prefix
}

function Register-FabriqIosCompleter {
    param([hashtable]$State)
    $global:_FabriqIosShellState = $State

    Set-PSReadLineKeyHandler -Key Tab -BriefDescription 'FabriqIosTab' `
                             -LongDescription 'Cisco-style tab completion' -ScriptBlock {
        param($key, $arg)
        $line   = $null
        $cursor = $null
        [Microsoft.PowerShell.PSConsoleReadLine]::GetBufferState([ref]$line, [ref]$cursor)

        $state = $global:_FabriqIosShellState
        if ($null -eq $state) {
            [Microsoft.PowerShell.PSConsoleReadLine]::Insert("`t")
            return
        }

        $candidates = Get-FabriqIosCompletion -Line $line -Position $cursor -Mode $state.Mode -State $state
        if ($candidates.Count -eq 0) {
            [Microsoft.PowerShell.PSConsoleReadLine]::Ding()
            return
        }

        $upToCursor = if ($cursor -gt 0) { $line.Substring(0, $cursor) } else { '' }
        $endsWithSpace = ($upToCursor.Length -eq 0) -or ($upToCursor[$upToCursor.Length - 1] -match '\s')
        $tokens = ConvertTo-FabriqIosTokens $upToCursor
        $currentToken = if ($endsWithSpace -or $tokens.Count -eq 0) { '' } else { $tokens[-1] }
        $currentLength = $currentToken.Length

        if ($candidates.Count -eq 1) {
            [Microsoft.PowerShell.PSConsoleReadLine]::Replace(
                $cursor - $currentLength, $currentLength, $candidates[0] + ' '
            )
            return
        }

        $commonPrefix = Get-CommonPrefix -Strings $candidates
        if ($commonPrefix.Length -gt $currentLength) {
            [Microsoft.PowerShell.PSConsoleReadLine]::Replace(
                $cursor - $currentLength, $currentLength, $commonPrefix
            )
        } else {
            Write-Host ""
            foreach ($c in ($candidates | Sort-Object)) {
                Write-Host ("  " + $c) -ForegroundColor DarkCyan
            }
            [Microsoft.PowerShell.PSConsoleReadLine]::InvokePrompt()
        }
    }

    Set-PSReadLineKeyHandler -Chord '?' -BriefDescription 'FabriqIosHelp' `
                             -LongDescription 'Cisco-style ? help' -ScriptBlock {
        param($key, $arg)
        $state = $global:_FabriqIosShellState
        if ($null -eq $state) {
            [Microsoft.PowerShell.PSConsoleReadLine]::Insert('?')
            return
        }

        # Capture the buffer at the moment '?' was pressed and hand off
        # to the REPL loop. The REPL prints help on the same code path
        # as `show modules` (Write-Host then a manual prompt render at
        # the next iteration), which renders cleanly without the
        # PSReadLine InvokePrompt overlap issues.
        $line = $null; $cursor = $null
        [Microsoft.PowerShell.PSConsoleReadLine]::GetBufferState([ref]$line, [ref]$cursor)

        $global:_FabriqIosPendingHelpRequest = $true
        $global:_FabriqIosPendingHelpBuffer  = if ($null -eq $line) { '' } else { $line.Substring(0, [Math]::Min($cursor, $line.Length)) }

        # AcceptLine submits the current buffer (so it lands in
        # PSReadLine history and the user can recall it with Up
        # arrow) and returns control to the REPL loop. Buffer
        # content itself will be ignored because the REPL detects
        # the pending-help flag first.
        [Microsoft.PowerShell.PSConsoleReadLine]::AcceptLine()
    }
}
