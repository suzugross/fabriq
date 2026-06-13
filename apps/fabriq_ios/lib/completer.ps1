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

function Clear-FabriqIosPendingHelpRows {
    # Called by the REPL right after ReadLine returns (i.e. on command
    # submit). An inline `?` (non-empty buffer, handled below) writes its
    # candidate list to the rows immediately BELOW the input line and then
    # rewinds the prompt via InvokePrompt, leaving those rows on screen.
    # They are not part of PSReadLine's model, so on submit they would be
    # only partially overwritten by the (non-width-padded, often shorter)
    # command output - e.g. `module ?` shows ~44 candidates but entering
    # config-mod prints ~2 lines, leaving fragments and orphan rows.
    #
    # The post-submit cursor sits exactly where the candidate block began
    # (the row just below the input, which is also where command output is
    # about to be written), so blank the recorded row count from here and
    # restore the cursor. Robust to scrolling because it is relative to the
    # current cursor, not an absolute remembered row. No-op when nothing is
    # pending. Always consumes the tracking state so the next `?` press and
    # the next submit start clean.
    $rows = $global:_FabriqIosLastHelpRowCount
    $global:_FabriqIosLastHelpRowCount = $null
    $global:_FabriqIosLastHelpInputY   = $null
    if ($null -eq $rows -or [int]$rows -le 0) { return }
    $rows = [int]$rows

    try {
        $w = [Console]::WindowWidth
        $padWidth = if ($w -gt 1) { $w - 1 } else { 0 }
        if ($padWidth -le 0) { return }
        $blankRow = ' ' * $padWidth

        $startY    = [Console]::CursorTop
        $bufBottom = [Console]::BufferHeight - 1
        for ($i = 0; $i -lt $rows; $i++) {
            $y = $startY + $i
            if ($y -gt $bufBottom) { break }
            [Console]::SetCursorPosition(0, $y)
            [Console]::Write($blankRow)
        }
        [Console]::SetCursorPosition(0, $startY)
    } catch {
        # Console geometry can be unavailable in odd hosts; clearing is
        # cosmetic, so never let it break the REPL.
    }
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

        $line = $null; $cursor = $null
        [Microsoft.PowerShell.PSConsoleReadLine]::GetBufferState([ref]$line, [ref]$cursor)
        $buffer = if ($null -eq $line) { '' } else { $line.Substring(0, [Math]::Min($cursor, $line.Length)) }

        if ([string]::IsNullOrWhiteSpace($buffer)) {
            # Empty prompt: defer to the REPL for full Show-FabriqIosHelp
            # output. The mode-level help can run 30+ lines, which is
            # too long for InvokePrompt to redraw cleanly underneath.
            # AcceptLine returns control to the REPL, which detects the
            # pending flag and prints help on the same path as `show
            # modules` before rendering a fresh prompt.
            $global:_FabriqIosPendingHelpRequest = $true
            $global:_FabriqIosPendingHelpBuffer  = ''
            [Microsoft.PowerShell.PSConsoleReadLine]::AcceptLine()
            return
        }

        # Cisco IOS inline `?`: show only the candidate list for what
        # the user has typed so far, then let PSReadLine redraw the
        # prompt with the original buffer intact so editing continues
        # from where it left off (e.g. `host?` -> show `hostname` ->
        # cursor lands back at the end of `host`).
        #
        # Visual subtlety #1: PSReadLine's InvokePrompt re-renders
        # prompt+buffer at the original _initialY (the row where
        # ReadLine started). After Write-Host'ing N help rows starting
        # from that row, InvokePrompt overwrites row 1 of help with
        # the prompt - leaving rows 2..N visible BELOW the prompt.
        #
        # Visual subtlety #2: each repeat `?` press writes its
        # candidates to the SAME row range as the previous press
        # (since _initialY is unchanged). If the new candidate list
        # is shorter than the old one, two failure modes appear:
        #   (a) per-row: a new short candidate ("Password") landing
        #       on a row that previously held a longer candidate
        #       ("PasswordOperators"-something) leaves the trailing
        #       characters of the old visible after the new
        #       ("Passwordperators").
        #   (b) per-list: rows beyond the new list's length still
        #       hold old candidates that were never overwritten, so
        #       a tail of stale rows hangs below.
        # Fix (a) by padding every help row to full window width so
        # any leftover trailing chars are blanked. Fix (b) by tracking
        # the previous press's row count and explicitly blanking any
        # orphan tail rows. Reset the count when the input row
        # changes (user submitted a command and a fresh prompt is on
        # a new row), since the rows below now contain unrelated
        # console output we must not overwrite.
        $candidates = Get-FabriqIosCompletion -Line $buffer -Position $buffer.Length `
                                              -Mode $state.Mode -State $state

        $inputY   = [Console]::CursorTop
        $w        = [Console]::WindowWidth
        $padWidth = if ($w -gt 1) { $w - 1 } else { 0 }
        $blankRow = if ($padWidth -gt 0) { ' ' * $padWidth } else { '' }

        $prevRows = 0
        if ($global:_FabriqIosLastHelpInputY -eq $inputY -and `
            $null -ne $global:_FabriqIosLastHelpRowCount) {
            $prevRows = [int]$global:_FabriqIosLastHelpRowCount
        }

        # Sacrificial leading newline. InvokePrompt below redraws
        # prompt+buffer at PSReadLine's _initialY (the input row),
        # overwriting whatever was written there. Without this, the
        # first Write-Host writes mid-row at the input's cursor column,
        # then the redraw wipes the appended chars - the alphabetically-
        # first candidate silently vanishes (e.g. Administrators
        # missing from `set Group ?` even though Tab returned it). The
        # Tab handler uses the same trick. The sacrificial row is not
        # counted in $rowsWritten because $tailY (captured after all
        # writes) already reflects the extra row, and $extra =
        # $prevRows - $rowsWritten only depends on the relative diff
        # which both presses share.
        Write-Host ""

        $rowsWritten = 0
        if ($candidates.Count -gt 0) {
            foreach ($c in ($candidates | Sort-Object)) {
                $row = '  ' + $c
                if ($padWidth -gt 0 -and $row.Length -lt $padWidth) {
                    $row = $row.PadRight($padWidth)
                }
                Write-Host $row -ForegroundColor DarkCyan
                $rowsWritten++
            }
        } else {
            $row = '  (no candidates)'
            if ($padWidth -gt 0 -and $row.Length -lt $padWidth) {
                $row = $row.PadRight($padWidth)
            }
            Write-Host $row -ForegroundColor DarkGray
            $rowsWritten = 1
        }

        # Wipe any orphan tail rows the previous press left behind.
        if ($prevRows -gt $rowsWritten -and $padWidth -gt 0) {
            $tailY = [Console]::CursorTop
            $extra = $prevRows - $rowsWritten
            $bufBottom = [Console]::BufferHeight - 1
            for ($i = 0; $i -lt $extra; $i++) {
                $y = $tailY + $i
                if ($y -gt $bufBottom) { break }
                [Console]::SetCursorPosition(0, $y)
                [Console]::Write($blankRow)
            }
            [Console]::SetCursorPosition(0, $tailY)
        }

        $global:_FabriqIosLastHelpInputY   = $inputY
        $global:_FabriqIosLastHelpRowCount = $rowsWritten

        [Microsoft.PowerShell.PSConsoleReadLine]::InvokePrompt()
    }
}
