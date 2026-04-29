# Changelog

本プロジェクトのすべての注目すべき変更はこのファイルに記録されます。

フォーマットは [Keep a Changelog 1.1.0](https://keepachangelog.com/ja/1.1.0/) に基づき、
バージョニングは [Semantic Versioning 2.0.0](https://semver.org/lang/ja/) に従います。

カテゴリの意味:
- **Added**: 新機能
- **Changed**: 既存機能の変更（後方互換）
- **Deprecated**: 将来削除予定の機能
- **Removed**: 削除された機能
- **Fixed**: バグ修正
- **Security**: セキュリティ関連

## [Unreleased]

### Added
- apps/fabriq_ios/: Phase 7 implementation - ModuleConfig
  (config-mod)# mode for ephemeral module configuration. Cisco IOS
  semantics: `(config)# module <name>` enters (config-mod)#, and
  inside the mode `set <col1> <val1> <col2> <val2> ...` (or `add`
  alias) IMMEDIATELY runs the module with a single ephemeral row
  built from the column/value pairs. The existing <module>_list.csv
  on disk is NEVER touched - only an in-memory PSCustomObject is
  fed to the module via a path-matched override of
  Import-ModuleCsv (`Set-Item Function:Global:Import-ModuleCsv`
  with closure-captured ephemeral row + filename, restored in
  finally). Reads of unrelated CSVs (e.g. hostlist.csv) pass
  through to the real implementation.
  
  New file: lib/modes/module_config.ps1 (mode dispatcher).
  lib/commands/module.ps1 expanded with Get-ModuleCsvSchema
  (auto-discovers <prefix>_list.csv via glob - the naming convention
  drops the action suffix, e.g. reg_hklm_config -> reg_hklm_list.csv,
  app_config -> app_list.csv), Enter-ModuleConfigMode (validates
  module + schema before mode entry), Show-ModuleConfigSchema
  (Cisco-ish schema dump with preset.csv enum hints in [val|val]
  brackets), and Invoke-ModuleEphemeralRun (the override + dispatch
  + result-to-syslog mapping). lib/shell_state.ps1 gained
  ConfigModuleName / ConfigModuleSchema fields and the new
  GlobalConfig <-> ModuleConfig transition; lib/prompt.ps1 emits
  `fabriq(config-mod)#` for the new mode; lib/parser.ps1 vocabulary
  for ModuleConfig is set/add/show/exit/end/help/?;
  lib/completer.ps1 routes `set`/`add` Tab completion to the
  bound module's column names; data/help_text.csv adds 7 rows.
  
  Phase 6 immediate-run via `(config)# module <name>` is
  intentionally retired - the same verb now enters config-mod for
  ephemeral configuration. Modules without *_list.csv (e.g.
  evidence_config) cannot be ephemeral-configured and are rejected
  at mode entry. ENC: encrypted columns: not supported for set/add
  in v0.2 (would require Protect-FabriqValue public-API promotion);
  documented limitation. Default behaviour: Enabled column auto-
  defaults to '1' so most modules don't filter the ephemeral row
  out; FilterEnabled is honoured by the override; Segment is
  intentionally ignored (ephemeral intent overrides segment scoping).
  
  Internal API coupling: still just Invoke-FabriqIosModule +
  ModuleResult contract, plus the well-known PowerShell function-
  shadowing pattern (`Set-Item Function:Global:`). No KERNEL_API
  surface change. tests/_phase7_smoke.ps1 NEW (49 assertions, all
  PASS) covers schema lookup, mode transitions, vocab/prompt,
  pair-parsing reject paths, parser abbreviations (sh/se/a/ex/en),
  Tab completion delegating to schema, and the override mechanics
  (target-CSV intercepted, unrelated CSV passes through, restore
  on exit). Real module-execution leg validated manually on a
  test machine (not part of automated smoke).

  tests/_phase6_smoke.ps1 updated for new mode-entry semantics
  (40 assertions, all PASS).

### Changed
- apps/fabriq_ios/: `show running-config` body replaced with live
  `ipconfig /all` output. The previous artificial dump (constructed
  from SELECTED_* env vars at the moment of invocation) deviated
  from reality whenever the operator had not yet committed the
  selected host's settings to the OS, so its Cisco-style fidelity
  came at the cost of accuracy. New behaviour: keep the Cisco-style
  preamble (Building configuration... / Current configuration : ...
  / version / banner motd / Surkittinist comment) and footer
  (session worker / end), but emit ipconfig /all output verbatim
  between them. ipconfig is invoked directly (no PowerShell capture)
  so the console encoding handles locale-dependent text without
  mojibake on JP Windows. Code comment acknowledges the deliberate
  deviation from Cisco semantics (where running-config shows
  definitions, not runtime state) - joke aesthetic in the wrappers,
  honest body.

### Added
- apps/fabriq_operator/lib/quickactions_dialog.ps1 (NEW): "And More"
  sub-dialog (DataGridView pattern, mirrors apps_dialog.ps1 styling)
  with 4 entries that map onto existing main.ps1 switch cases:
    - Restart PC           -> Action="Restart"
    - Export History       -> Action="HistoryExport"
    - Regenerate Checklist -> Action="RegenerateChecklist"
    - FabriqApps           -> Action="AppsMode"
  Returns @{Action; AppName} which the dashboard form propagates
  upward as its own $result, so no new main.ps1 handlers are
  required for the items inside the dialog.
- apps/fabriq_operator/lib/dashboard_form.ps1: Settings tab Quick
  Actions section reorganised. Front row now hosts five high-frequency
  shortcuts plus a single "And More..." entry point:
    Row 1: [Open CSV Editor] [Windows Update]   [Fabriq IOS]
    Row 2: [Refabriq]        [System Launcher]  [And More...]
  Removed buttons (now under And More): Export History, Regenerate
  Checklist, Restart PC, FabriqApps. New direct button [Fabriq IOS]
  emits Action="LaunchApp" with AppName="fabriq_ios" (generic shortcut
  pattern, reusable for future direct-app buttons). $result hashtable
  gained AppName field. Visible-comment list of Action tokens updated.
- kernel/main.ps1: new switch case "LaunchApp" mirrors the in-process
  invocation pattern of "AppsMode" but bypasses Show-AppsDialog by
  resolving apps/<AppName>/<AppName>.ps1 directly from
  $guiSelection.AppName. Self-spawn logic inside the launched app
  (e.g. fabriq_ios.ps1's $env:FABRIQ_IOS_SUBPROCESS guard) still
  enforces subprocess isolation so the dashboard process is never
  contaminated. Internal-only change (KERNEL_API.md surface unchanged);
  PATCH-class influence on kernel/main.ps1, version-file bump deferred
  until next release instruction.
- apps/fabriq_ios/: Phase 6 implementation - `module <name>` command
  in GlobalConfig (`fabriq(config)#`) for single-module execution
  (profile execution intentionally not implemented per scope
  decision). New file lib/commands/module.ps1 hosts
  `Get-ModuleCompletionFromFilesystem` (auto-discovers
  modules/standard and modules/extended via the `<dir>/<dir>.ps1`
  convention, cached for the session), `Find-ModulePath` (standard
  -> extended fallback), and `Invoke-ModuleByName` (passphrase
  staging, dispatch via Phase 4's Invoke-FabriqIosModule, severity-
  mapped MODULE syslog). Parser vocabulary
  ('hostname','interface','module','exit','end','help','?'),
  completer dynamic source `'module'`, GlobalConfig dispatcher
  branch, help_text.csv row, and 5 new MODULE syslog templates
  (success / partial / skipped / cancelled / error) added to
  syslog_messages.csv. Excluded from listing / dispatch (6
  modules):
    - windows_update     (GUI-button-only, no module.csv)
    - test_error_module / test_harness_config (test scaffolding)
    - fabriq_app_launcher (recursive launcher)
    - hostname_config    (covered by `(config)# hostname <name>`)
    - ipaddress_config   (covered by `(config-if)# ip address ...`)
  Internal API coupling unchanged (still just Invoke-FabriqIosModule
  + ModuleResult contract). Tests/_phase6_smoke.ps1 (36 assertions,
  all PASS) covers vocab / discovery / exclusion enforcement /
  parser abbreviation (`mod` -> `module`) / dispatcher reject paths
  / completion filtering. Actual module execution validated by
  manual test on a target machine.

### Fixed
- apps/fabriq_ios/: prompt did not update on entry to Global Config
  / Interface Config modes - `fabriq(config)#` and
  `fabriq(config-if)#` never appeared, with the prompt freezing on
  the prior `#`. The expected prompt only flashed momentarily when
  the user pressed `?` (because the `?` handler ends with
  `[...PSConsoleReadLine]::InvokePrompt()` which forces a re-invoke
  of the prompt function). Root cause: PSConsoleReadLine.ReadLine()
  does NOT invoke the global `prompt` function when called
  programmatically - prompt rendering is normally the PowerShell
  Console host's responsibility, and our subprocess REPL bypasses
  the host's main loop entirely. The `Set-Item Function:\global:prompt`
  override therefore only mattered for InvokePrompt() calls inside
  PSReadLine key handlers, which is why the prompt only refreshed
  on Tab/`?` press. UserExec / PrivilegedExec branches happened to
  appear correctly because Read-Host -AsSecureString during
  `enable` re-engaged the host UI subsystem and incidentally
  rendered the next prompt; pure dispatch paths (configure
  terminal, interface, end) never triggered that side-effect.
  Fix: render the prompt explicitly with [Console]::Write inside
  the REPL loop just before each PSConsoleReadLine.ReadLine() call.
  Kept the global prompt function so Tab/`?` handlers'
  InvokePrompt() still re-renders with the correct mode.
- apps/fabriq_ios/: `?` help handler still overlapped help output
  with the redrawn prompt+buffer even after switching to
  `[Console]::WriteLine` + Out.Flush - InvokePrompt's anchor /
  buffer-redraw step proved unreliable in our custom REPL setup
  (where the prompt is rendered manually before each ReadLine,
  not by the host). Replaced the in-handler InvokePrompt approach
  entirely: the `?` handler now captures the current buffer text
  into `$global:_FabriqIosPendingHelpRequest` /
  `$global:_FabriqIosPendingHelpBuffer` and calls AcceptLine so
  ReadLine returns immediately. The REPL loop in fabriq_ios.ps1
  detects the pending-help flag, prints the help block via the
  same Write-Host path that `show modules` uses (which the user
  has confirmed renders cleanly), and `continue`s to the next loop
  iteration. The next iteration's manual `[Console]::Write` of the
  prompt then naturally lands below the help, matching Cisco IOS
  layout. Side effect: the typed buffer is cleared after `?`
  (Cisco preserves it), but the original line is submitted via
  AcceptLine so PSReadLine adds it to history - the user can recall
  with Up arrow if they want to continue editing.
- apps/fabriq_ios/: `show running-config` layout deviated from
  Cisco IOS in three places: (1) the `Current configuration : ! Surkittinist artefact`
  header concatenated the byte-count line with a comment instead of
  putting the comment on its own `!` line; (2) `ip default-gateway`
  and `ip name-server` were emitted *inside* the `interface` block
  with one space of indentation, but in real Cisco IOS those are
  global config commands and belong at the top level; (3) multiple
  DNS servers shared a single space-delimited `ip name-server` line
  rather than one server per line. Also `manifesto-banner ^C ...`
  was a fabricated banner type - real Cisco IOS only has motd /
  login / exec / incoming / slip-ppp. Refactored
  Show-FabriqIosRunningConfig to assemble the body as a list,
  compute the byte count via UTF8.GetByteCount, render
  `Current configuration : <N> bytes` followed by the body verbatim,
  switch to `banner motd ^C ... ^C` (real Cisco syntax with
  Surkittinist body preserved), and emit ip default-gateway /
  ip name-server at the top level with one server per line. Smoke
  test updated: `manifesto-banner` regex -> `banner motd`.
  inserted *all* candidates space-joined (e.g. typing `hostname `
  + Tab inserted "NEW-PC-01 NEW-PC-02 NEW-PC-03 " in one go) instead
  of extending to the common prefix Cisco-style. Cause: the Tab
  handler wrapped Get-FabriqIosCompletion's output with `@(...)`
  while the function returns `,$arr` (1-element wrapping). The
  outer `@(...)` therefore yielded a 1-element array containing the
  inner candidate array, so `$candidates.Count -eq 1` was true,
  the single-candidate branch fired, and `$candidates[0] + ' '`
  coerced the inner array to a `$OFS`-joined string before
  inserting. Removed the @() wrap on the two
  `$candidates = Get-FabriqIosCompletion ...` call sites in
  completer.ps1 (Tab and ? handlers). Common-prefix extension and
  list-display behaviour now match Cisco IOS: first Tab extends to
  the longest unique prefix; a second Tab when no extension is
  possible shows the candidate list and re-prompts.
- apps/fabriq_ios/: shared shell-state container moved from
  `$script:FabriqIosShellState` to `$global:_FabriqIosShellState`
  so the prompt scriptblock and Tab/`?` handlers observe the same
  hashtable across the dot-source boundary between fabriq_ios.ps1
  and lib/completer.ps1. (Prerequisite for the prompt fix above;
  alone this was insufficient because of the deeper PSReadLine
  invocation issue.)

### Added
- apps/fabriq_ios/: NEW joke shell (Cisco IOS-style overlay).
  Phase 1 skeleton: directory layout + function signatures + data
  CSVs (syslog_messages / help_text / version_banner). Self-spawning
  subprocess isolation (`$env:FABRIQ_IOS_SUBPROCESS` sentinel +
  `Start-Process powershell.exe -NoProfile -File ... -Wait`)
  prevents PSReadLine handler / env var / global state leakage when
  launched in-process via fabriq_operator's [FabriqApps] dialog
  (`& $appResult.AppPath` at kernel/main.ps1:1340). Auto-discovered
  by apps_dialog.ps1 -- no CSV registration required. Implementation
  bodies stubbed with `throw 'Not implemented: ...'`; Pester tests
  are `-Skip` placeholders. Future mutation phase will couple to
  KERNEL_API.md §6 internal symbols (`Invoke-BatchExecution`,
  `Initialize-ModuleSystem`, `Resolve-ProfileModules`,
  `Test-MasterPassphrase`); coupling acknowledged in app README.
  Kernel / KERNEL_API.md / KERNEL_VERSION / modules unchanged.
- apps/fabriq_ios/: Phase 5 implementation (PSReadLine completion -
  v0.1 acceptance criteria complete):
    - lib/completer.ps1: full implementation. `Get-FabriqIosCompletion`
      is a pure function (cursor position -> candidate string array)
      driven by Get-FabriqIosCommandVocabulary +
      Get-FabriqIosSubVocabulary + Get-DynamicCompletion (which
      dispatches to hostlist / NetAdapter / IP completion sources).
      `Get-CommonPrefix` enables progressive Tab completion. Tab
      handler: single candidate -> insert + space; multiple
      candidates with shared prefix -> extend to common prefix; no
      shared prefix -> list candidates and re-render prompt.
      `?` handler: empty buffer -> mode help; non-empty buffer ->
      contextual candidates list; both then re-render prompt.
    - fabriq_ios.ps1: `Initialize-FabriqIos` now imports PSReadLine
      and throws if unavailable. `Start-FabriqIosShell` switches
      from `Read-Host` (which bypasses PSReadLine in PS 5.1) to
      `[Microsoft.PowerShell.PSConsoleReadLine]::ReadLine($host.Runspace,
      $ExecutionContext)` and overrides `global:prompt` to render
      the mode-aware Cisco-style prompt. Subprocess isolation
      ensures the prompt override and PSReadLine handlers vanish
      with the child process.
    - tests/completer.tests.ps1: un-Skipped, real Pester assertions
      for command vocabulary / prefix filter / subcommand expansion
      / dynamic sources / Get-CommonPrefix.
    - tests/_phase5_smoke.ps1 (NEW): 28-assertion smoke covering
      every completion path; PSReadLine integration is verified
      manually in interactive REPL.
  SPEC v0.1 acceptance criteria checklist now fully met. Phase 4's
  `Get-*Completion*` purely-data helpers are now wired into the
  Tab and `?` keys.
- apps/fabriq_ios/: Phase 4 implementation (mutation - real
  hostname / IP address dispatch through fabriq modules):
    - lib/dispatch.ps1 (NEW): minimal local dispatcher
      `Invoke-FabriqIosModule` reproduces just enough of
      kernel/main.ps1's Invoke-KittingScript to run a single module
      script and capture its ModuleResult (we cannot dot-source
      main.ps1 because that would launch a second fabriq_operator
      GUI inside the isolated subprocess). Also hosts
      `Set-FabriqIosHostEnvironment` (mirror of main.ps1's
      Set-SelectedHostEnvironment, plaintext-only since
      Import-ModuleCsv has already decrypted ENC: cells),
      `Find-HostlistRowByNewName`, and `ConvertFrom-SubnetMaskToPrefix`.
    - lib/commands/hostname.ps1: Invoke-HostnameSelection looks up
      the hostlist row by NewPCName, sets SELECTED_* env vars,
      runs modules/standard/hostname_config/hostname_config.ps1
      via Invoke-FabriqIosModule, emits HOSTNAME/success +
      HOSTNAME/reboot_required syslog on success or HOSTNAME/refused
      on Error/missing-row. Get-HostnameCompletionFromHostlist
      returns NewPCName values (used by Phase 5 PSReadLine wiring).
    - lib/commands/ip_address.ps1: Invoke-IpAddressFromHostlist
      runs modules/standard/ipaddress_config/ipaddress_config.ps1
      against the currently bound SELECTED_* host context and emits
      the three-line whispered/gateway/dns syslog on success.
      Invoke-IpAddressManual overrides SELECTED_ETH_IP/SUBNET with
      operator-supplied values for the duration of the call (env
      restore in finally). Get-IpAddressCompletionFromHostlist
      returns the literal 'from-hostlist' shortcut plus the bound
      Ethernet/WiFi IPs.
    - lib/commands/interface.ps1: Get-InterfaceCompletionFromAdapters
      enumerates Get-NetAdapter -Physical (active) InterfaceAlias
      values; Japanese aliases are preserved.
    - lib/modes/global_config.ps1: hostname stub replaced with
      Invoke-HostnameSelection.
    - lib/modes/interface_config.ps1: ip address stub replaced with
      Invoke-IpAddressFromHostlist (2 args) / Invoke-IpAddressManual
      (3 args) routing.
    - fabriq_ios.ps1: sets `$global:AutoPilotMode = $true` and
      `$global:AutoPilotWaitSec = 0` so module-level
      Confirm-ModuleExecution / Wait-KeyPress prompts auto-yes -
      the act of typing the command in the REPL is the
      confirmation, matching Cisco IOS UX. Adds lib/dispatch.ps1
      to the dot-source list.
  Internal API coupling expanded (KERNEL_API.md §6, accepted): the
  ModuleResult contract (§5), `$global:_LastModuleResult` fallback,
  `$global:AutoPilotMode` / `$global:AutoPilotWaitSec` flags, and
  the Invoke-KittingScript dispatch pattern (locally re-implemented).
  Evidence capture and execution-history writes are intentionally
  NOT performed by fabriq_ios v0.1 - the joke shell stays out of
  the audit trail. Documented as a v0.2 candidate.
- apps/fabriq_ios/: Phase 3 implementation (REPL + read-only show
  commands, no mutating kernel calls):
    - fabriq_ios.ps1: real `Start-FabriqIosShell` REPL loop with
      mode-aware prompt, '?'/'help' fast-path, error display,
      try/catch around dispatchers. `Show-FabriqIosHelp` reads
      data/help_text.csv (cached).
    - lib/modes/{user_exec,privileged_exec,global_config,interface_config}.ps1:
      mode dispatchers wired to Invoke-Enable/Invoke-Disable,
      Invoke-ShowCommand, Set-ShellMode, and stub markers for
      Phase 4 mutation commands (hostname/ip address).
    - lib/commands/show.ps1: implemented Invoke-ShowCommand router
      and 8 show commands (version/host/hosts/profiles/modules/
      evidence/manifesto/running-config). 'show running-config'
      emits Cisco-style '!' commented dump from SELECTED_*
      environment variables; 'show manifesto' emits a poetic syslog
      and points to the operator dashboard for the GUI manifesto
      (kernel/ps1/manifesto.ps1 is a WinForms viewer, unsuitable
      for the terminal). 'show hosts' uses Import-ModuleCsv with
      $global:FabriqMasterPassphrase set/restored around the call
      so ENC: cells decrypt only when the passphrase is set.
    - lib/commands/enable_disable.ps1: Invoke-Enable validates via
      Test-MasterPassphrase against
      kernel/txt/passphrase_verify.txt; on success it stores the
      plaintext in State.Passphrase and transitions to
      PrivilegedExec with a poetic seance_begins syslog. Failure
      emits passphrase_refused and stays in UserExec.
    - lib/commands/interface.ps1: Set-FabriqIosCurrentInterface
      stores alias and transitions to InterfaceConfig with an
      INTERFACE/opened syslog.
    - data/help_text.csv: added rows for help/? in every mode (now
      23 rows covering all 4 mode vocabularies).
  Internal API usage (KERNEL_API.md §6, accepted coupling):
  Test-MasterPassphrase, Import-ModuleCsv, Show-Info / Show-Warning
  (transitively via Import-ModuleCsv).
  Phase 3 explicitly does NOT mutate any system state - hostname /
  ip address / reload commands are stubbed with '% Phase 3 stub:'
  messages or performative-only behaviour.
  Kernel / KERNEL_API.md / KERNEL_VERSION / modules unchanged.
- apps/fabriq_ios/: Phase 2 implementation (pure-logic layer, no
  kernel calls yet):
    - shell_state.ps1: Set-ShellMode with mode-transition table
      (UserExec/PrivilegedExec/GlobalConfig/InterfaceConfig).
    - parser.ps1: tokenization (quoted-arg + Japanese-alias aware),
      Cisco-style prefix-uniqueness expansion (e.g. 'conf t' ->
      'configure terminal', 'sh ru' -> 'show running-config'),
      exact-match precedence to disambiguate 'host' vs 'hosts'.
      Helpers: Get-FabriqIosCommandVocabulary,
      Get-FabriqIosSubVocabulary, Resolve-Token.
    - prompt.ps1: 4-mode prompt builder.
    - syslog.ps1: Cisco-format timestamp ('*Apr 29 14:23:01.234',
      locale-independent month abbreviation), CSV template lookup
      with caching, placeholder substitution, severity-based
      coloring. Refactored to expose `Format-FabriqIosSyslogLine`
      pure function for testability.
  Added Pester suites: shell_state.tests.ps1, prompt.tests.ps1,
  syslog.tests.ps1; un-Skipped parser.tests.ps1.
  Kernel / KERNEL_API.md / KERNEL_VERSION / modules unchanged.
- .gitignore に例外行を追加: `!/Fabriq.exe`、
  `!/modules/standard/printer_driver_config/tools/7z.exe`、
  `!/modules/standard/printer_driver_config/tools/7z.dll`。
  従来 `*.exe` / `*.dll` で全除外していた 3 つの同梱バイナリを
  リポジトリ管理対象に昇格（クローンだけで printer_driver_config が
  動作可能に、Fabriq ランチャーも GitHub 経由で配布可能に）
- Fabriq.exe をリポジトリで管理開始（自社ランチャー、ソースは
  dev/launcher/Launcher.cs。標準 System.* のみ依存で第三者
  ライブラリなし）
- modules/standard/printer_driver_config/tools/7z.exe と 7z.dll を
  リポジトリで管理開始（7-Zip 25.01 (x64) 同梱、SHA-256 は
  THIRD_PARTY_NOTICES.md に明記）
- THIRD_PARTY_NOTICES.md（リポジトリ root に新規）: サードパーティ
  同梱物のアトリビューション・ライセンス概要を一元化。現状は 7-Zip
  25.01 (x64) のみ（modules/standard/printer_driver_config/tools/
  配下の 7z.exe / 7z.dll）。同梱バイナリの SHA-256・公式 URL・
  ソースコード入手先・unRAR 制限を明記
- LICENSES/LGPL-2.1.txt（新規）: GNU LGPL v2.1 全文（FSF 公式から
  取得した verbatim コピー）。LGPL v2.1 第 1-2 条が要求する
  「ライセンスのコピーの同梱」義務に対応
- LICENSES/7-Zip-license.txt（新規）: 7-Zip 公式 license.txt の
  verbatim コピー（GNU LGPL / BSD 3-clause / BSD 2-clause / unRAR
  制限の組み合わせを明文化した一次ライセンス文書）
- modules/standard/printer_driver_config/tools/README-license.txt
  （新規）: バイナリ配置場所での attribution。tools/ を切り出して
  別ロケーションに持ち出した場合に表示が伴走するための副次表示
- README.md「サードパーティ同梱物」節（新規）: ルート
  THIRD_PARTY_NOTICES.md への入口リンクを追加
- dev/verify_comments_only.ps1: 新規追加。.ps1 ファイルの変更が
  「コメントのみ」であることを PowerShell parser の AST トークン
  比較で機械的に証明するための検証ツール。日本語コメントを英語に
  一掃する作業の安全装置として導入。Mode 1（任意 2 パス比較）と
  Mode 2（working tree vs git HEAD 比較）をサポート

### Changed
- modules/standard/printer_driver_config/Guide.txt: tools/ 配下の
  7z.exe / 7z.dll が同梱済みである旨に書き換え（従来は「ユーザーが
  配置」と記述されていたが実態と乖離）。差し替え時の更新箇所
  （README-license.txt と THIRD_PARTY_NOTICES.md のバージョン・
  SHA-256）も明示
- dev/template/_template_script.ps1: コメントを日本語から英語へ
  全面置換（41 行）。新規モジュールの種ファイルなので最優先で
  整理。dev/verify_comments_only.ps1 と sed-based code-strip diff
  の二系統で「コードトークン不変・コメントのみ変更」を確認済み
- kernel/common.ps1: コメントを日本語から英語へ全面置換（29 行の
  単行コメント + 4 ブロック分の Comment-Based Help 記述）。
  Comment-Based Help の `.SYNOPSIS` / `.DESCRIPTION` / `.PARAMETER`
  等の dot-keyword は不変保持。runtime に必要な日本語ロケール
  エラー検出文字列（`'used by another process|別のプロセス'`）は
  機能コードのため残置。AST verifier PASS（13606 トークン一致）
- kernel/main.ps1: コメントを日本語から英語へ全面置換（6 行）。
  sed-based code-strip diff: IDENTICAL（block コメントなし）
- modules/extended/manual_kitting_assistant: コメントを日本語から
  英語へ全面置換（48 行）。Phase 3 着手 1 ファイル目。
  C# here-string（@'...'@）内部の `//` コメントと UI 用文字列
  リテラル（ボタンラベル、MessageBox メッセージ等）は機能コード
  のため不変保持。sed code-strip diff: IDENTICAL
- modules/extended/history_destroyer: コメントを日本語から英語へ
  全面置換（47 行）。sed code-strip diff: IDENTICAL
- modules/standard/sysprep_config: コメントを日本語から英語へ全面
  置換（47 行）。sed code-strip diff: IDENTICAL
- modules/standard/bitlocker_config/bitlocker_await: コメントを
  日本語から英語へ全面置換（27 行）。sed code-strip diff: IDENTICAL
- modules/standard/driver_config/driver_import_config: コメントを
  日本語から英語へ全面置換（22 行）。sed code-strip diff: IDENTICAL
- modules/extended/script_looper: コメントを日本語から英語へ全面
  置換（20 行）。sed code-strip diff: IDENTICAL
- Phase 3 batch 2-4 (16 ファイル): odt_install / odt_download /
  local_user_setup / driver_export_config / restore_point /
  generic_process_runner / bitlocker_disable / taskbar_config /
  spi_config / process_killer / kernel/ps1/status_monitor /
  export_app_associations / default_app_config / browser_addon_config /
  ssid_config / log_uploader のコメントを日本語から英語へ翻訳。
  string literal 内の機能コード（UI text / 正規表現マッチ /
  Write-Log メッセージ / 技術エビデンス byte 列）は不変保持。
  Phase 3 全 22 ファイル完了
- 残存日本語ファイル 10 件はすべて機能コード扱い（不変保持）:
  ssid_config / common.ps1 / evidence_config / printer_driver_install /
  odt_install / odt_download / local_user_setup /
  manual_kitting_assistant / firewall_rule_import /
  firewall_rule_export。それぞれ regex match パターン、UI 文字列
  リテラル、Write-Log メッセージ、または UTF-8/CP932 mojibake の
  技術エビデンスとしての日本語であり、コメントのみ変更ポリシーの
  対象外
- Phase 3 全体で計 ~440 行のコメントを翻訳。Phase 1 (template) +
  Phase 2 (kernel/common + main) + Phase 3 (modules/apps) で
  日本語コメント残ゼロ達成（`feedback_scripts_english_only.md`
  ポリシー準拠）

## [3.0.0] - 2026-04-29

### Removed (BREAKING)
- kernel: 特殊マーカー 4 種を削除: `__SHUTDOWN__` / `__PAUSE__` /
  `__STOPLOG__` / `__STARTLOG__`
  - 公開 API surface の破壊的変更（KERNEL_API.md §4.2）。次期 MAJOR
    （3.0.0 候補）に向けた整理
  - 削除理由: 実運用での参照ゼロまたは唯一の使用箇所も廃止済み
    （`__SHUTDOWN__` は profiles/Master_Pre01.csv のみ使用 → 同時削除）。
    fabriq_studio のマーカーパレットでも既に除外されており、UX 上は
    事実上 deprecated だった
  - kernel/main.ps1: 4 つのハンドラブロック（約 75 行）削除
  - kernel/common.ps1: `$specialMarkers` ハッシュから 4 エントリ削除
    （6 → 2）、`Show-BatchConfirmation` の `$isSpecial` 検出を 6 フラグ
    から 2 フラグに簡素化
  - kernel/KERNEL_API.md: §4.2 表から 4 エントリ削除、§8 に
    [Unreleased] セクション追加（次期 MAJOR 候補の破壊的変更を記録）
  - profiles/Master_Pre01.csv: `__SHUTDOWN__` 行を削除（driver_export
    のみの 1 行プロファイルに変更）
  - kernel/common.ps1: `Invoke-CountdownShutdown` 関数（15 行）削除。
    `__SHUTDOWN__` の唯一の呼び出し元だったためデッドコード化。
    KERNEL_API.md §6 内部 API 一覧からも除去
  - 互換性: 削除済みマーカーを含む旧プロファイルは `$invalidPaths` 経由
    で「module not found」warning に降格、kernel はクラッシュせず他
    モジュールの実行を継続する（graceful degradation）

### Removed
- modules/extended/edge_config: 削除（Edge プロファイルの robocopy /MIR
  バックアップ/復元。プロファイル参照ゼロ・コード依存ゼロを確認の上、
  現場固有度が高すぎるためフレームワークから除去）
- modules/extended/heif_config: 削除（HEIF/HEVC AppxBundle 同梱モジュール。
  プロファイル参照ゼロ・コード依存ゼロを確認の上、使用頻度が低すぎる
  ためフレームワークから除去。約 27MB の同梱バイナリも除去）
- apps/99_old/: 削除（autokey_recipe_editor / digital_gyotaq_editor /
  profile_editor / registry_collection_app の退役 GUI アプリ群。
  target_apps.csv からの参照ゼロを確認）
- kernel/csv/categories.csv: heif_config 削除に伴い孤児化した
  `Media Codec,120` 行を除去
- README.md: Extended 一覧から edge_config / heif_config 行と
  `Media Codec` カテゴリを除去。モジュール数 73 → 71、Extended 16 → 14
  に更新

### Changed
- modules/standard/fabriq_app_launcher: Guide.txt の記載例を退役 GUI
  アプリ（autokey_recipe_editor / digital_gyotaq_editor）から現存する
  apps/ 配下のもの（winget_gui / storeapp_editor / local_user_setup）に
  差し替え
- modules/standard/evidence_config: VERSION 1.4.0 → **1.5.0**
  - **§22 Office License を 4 段構成に拡張**（22a C2R / 22b OSPP /
    22c vNext per-user / 22d 自動解釈）。M365 subscription 環境で
    OSPP が誤って "NOTIFICATIONS / 0xC004F009 Grace expired" を
    返す既知挙動に対し、vNext per-user license file を権威ある情報源
    として併走させ、納品物に正しい LICENSED 判定を出せるようにした
  - **新規ファイル: `22_OfficeVnextLicenses.csv`**（manifest files[] に
    追加）
    - `C:\Users\*\AppData\Local\Microsoft\Office\Licenses\<Category>\
      <NumericFilename>` 全プロファイル横断走査
    - UTF-16LE → 外側 JSON → Base64 内側 JSON のデコード経路を実装
      （`Read-VnextLicenseFile` ヘルパ関数）
    - 列: UserProfile / Category / LicenseFile / LicenseType /
           ProductReleaseId / Status / IsTrial / Beneficiary /
           LicenseId / Acid / TenantId / UserId / HardwareIdBound /
           NotBefore / NotAfter / ParseStatus
    - 秘密鍵 / Signature 生 / RenewalToken 生は **出力しない**
      （HardwareIdBound は (present) フラグのみ）
  - **subscription SKU 自動検出**（`Test-OfficeSubscriptionSku`）:
    `O365*Retail` / `M365*Retail` / `O365EduCloudRetail` /
    `OneNoteFreeRetail` パターン
  - **manifest §22 status の解釈ロジック更新**:
    - subscription 検出 + Provisioned vNext あり → **Success**
    - subscription 検出 + vNext 0 件 → **Partial**
      （reason="M365 subscription installed but no Provisioned vNext
      license found (end-user sign-in pending)"）
    - VL/買い切り + OSPP Grace/Notifications → **Failed**
      （reason に検出された ProductReleaseIds 含む）
    - VL/買い切り + OSPP Licensed → **Success**
    - Office 未インストール → **Success**（既存挙動、テキストのみ）
  - 既存の 22a / 22b 出力は完全に保持。subscription 環境でも OSPP raw
    出力を残すことで audit 監査人が両系統を確認できる
  - kernel 不変、KERNEL_API §10 不変、REQUIRES_KERNEL 不変
    （`[System.Text.Encoding]::Unicode` / `[Convert]::FromBase64String` /
    `ConvertFrom-Json` は .NET / PowerShell 標準）

### Added
- modules/standard/evidence_config: VERSION 1.3.0 → **1.4.0**
  - **§23 Security Baseline** 新規追加（`23_SecurityBaseline.txt`）
    - TPM (`Get-Tpm`)、Secure Boot (`Confirm-SecureBootUEFI`)、
      VBS / HVCI / Credential Guard (`Win32_DeviceGuard`)、
      LSA Protection (`RunAsPPL` / `RunAsPPLBoot`)、
      BIOS / Firmware (`Win32_BIOS`) を統合
    - 各 probe は inner try/catch で個別退避し、1 つの失敗が
      section 全体を落とさない設計（取得できたものだけ記録）
  - **§24 Group Policy Report** 新規追加（`24_GroupPolicy.html` +
    `24_GroupPolicySummary.txt`）
    - `gpresult /h` の HTML 生成（コンピュータ側 + ユーザー側 RSoP）
    - サマリ TXT に HTML サイズ・ドメイン参加状態・実行ユーザーを記録
    - 実行ユーザー側 RSoP は kitting プロファイルユーザー視点である
      ことを Guide.txt に明記
    - **`gpresult /h` の 127 文字パス制限** を回避するため、`$env:TEMP`
      に短いランダム名で出力した後 Move-Item で本来の evidence パスへ
      移動する 2 段階処理を採用（evidence ディレクトリ名 = timestamp +
      PC名 + UUID で容易に 127 文字を超えるため）
  - **§25 Certificates** 新規追加（`25_Certificates.csv`）
    - LocalMachine\My / LocalMachine\Root / LocalMachine\CA /
      CurrentUser\My の 4 ストアを Store 列で統合した単一 CSV
    - 列: Store / Subject / Issuer / Thumbprint / NotBefore / NotAfter /
      HasPrivateKey / EnhancedKeyUsageList / FriendlyName / SerialNumber
    - 秘密鍵そのものは決して出力せず、HasPrivateKey フラグのみ記録
  - **§26 Battery Report** 新規追加（`26_BatteryReport.html`）
    - `powercfg /batteryreport` の HTML 生成
    - 受入検査の初期バッテリ容量証跡用
    - バッテリ非搭載（`Win32_Battery` 0 件）時は **Skipped**
      （reason="No battery present"）
  - manifest.json の sectionCount: 23 → **27** に増加
  - kernel/EVIDENCE_MANIFEST.md schemaVersion=1 内での後方互換な
    section 追加。kernel 不変、KERNEL_API §10 不変、REQUIRES_KERNEL
    不変（2.0.0 baseline のまま）
  - Min Kernel API: 2.0.0（baseline）。新規依存は Windows / PowerShell
    標準 cmdlet のみ（Get-Tpm / Confirm-SecureBootUEFI / gpresult /
    powercfg / Get-ChildItem Cert:）

## [2.2.2] - 2026-04-25

### Added
- kernel: **§10 Evidence Manifest 公開契約**を新設
  ([kernel/EVIDENCE_MANIFEST.md](kernel/EVIDENCE_MANIFEST.md), schemaVersion=1)
  - `evidence_config` v1.3.0 以降が `pc_information/<dir>/manifest.json` を
    出力。各セクションの `{ id, title, files, status, reason, elapsedMs }` +
    `summary` を機械可読形式で記録
  - 外部 evidence consumer ツール（fabriq_evidence_manager 等）が
    前方互換にパースするための正式契約
  - status enum: Success / Skipped / Failed / Partial の 4 値
  - manifest 不在の旧 evidence は外部ツール側でファイル列挙ベースに
    フォールバックする責任を明記（後方互換）
  - 公開 API §10 として KERNEL_API.md に追加（schemaVersion 1 を真実源とし、
    破壊的変更時は schemaVersion 昇格、後方互換な追加は schemaVersion=1 内で許容）
- modules/standard/evidence_config: VERSION 1.2.0 → **1.3.0**
  - 各セクションに Id を付与し、Section オブジェクト
    `{ id, title, files, status, reason, elapsedMs }` を逐次 `$script:ManifestSections`
    に蓄積
  - 完了時に `Write-EvidenceManifest` で manifest.json を atomic 書き出し
    （既存 manifest は manifest.json.bak に 1 世代 rotate）
  - 新ヘルパー関数: `Add-SectionFile`（ファイル登録）/
    `Close-Section`（セクション完了 + manifest 追加）/ `Write-EvidenceManifest`
  - `Start-Section` に `-Id` 引数追加、Stopwatch + ファイル列追跡を内部実装
  - §7 Printers が「プリンタ未インストール」のとき Skipped として記録
  - §14 Server 機能が Client OS で実行されたとき Skipped として記録
  - §18 Defender が利用不可（3rd-party AV 等）のとき Skipped として記録
  - §10 Serial が canonical 不可のとき Failed として記録
  - 既存の console ログ・CSV・テキスト出力・`New-BatchResult` 集計には影響なし
- kernel: KERNEL_VERSION 2.2.1 → **2.2.2**（§10 公開 API 追加に伴う MINOR 昇格）

## [2.2.1] - 2026-04-25

### Changed
- kernel/main.ps1 + kernel/common.ps1: Profile 経過時間トラッキングを
  `ElapsedSeconds` 累積方式から **`ProfileStartTime` 絶対時刻方式** に
  簡素化。**`__RESTART__` を含む profile 実行で表示される elapsed time
  が実 wall clock と一致するように修正**
  - 旧方式: 各 cycle で `(now - batchStart) + PriorElapsedSeconds` を
    再帰的に累積し resume_state.json に running sum として保存。
    結果として `Invoke-CountdownRestart` カウントダウン + Windows reboot
    + autologin + fabriq 起動再初期化（合計 1 restart あたり 1〜5 分）
    が **計測対象から漏れる** → 表示 elapsed が実時間より短い
  - 新方式: 最初の `Invoke-BatchExecution` 開始時刻を `ProfileStartTime`
    として保存し、終端で `(Get-Date) - $ProfileStartTime` の単純減算で
    elapsed を求める。reboot 中の経過時間が自然に含まれる。多段 restart
    でも累積誤差なし（純粋な timestamp 算術）
  - 後方互換: `Save-ResumeState` の param `-ElapsedSeconds` を
    `-ProfileStartTime` に置換。resume_state.json schema は
    `ElapsedSeconds: <double>` から `ProfileStartTime: <ISO 8601 string>`
    に変更。旧フォーマット（`ElapsedSeconds` のみ）の resume_state.json
    を読んだ場合は warning を出して resume 後の経過のみ計測（kitting
    flow 自体は完走）。fabriq update / rollback の in-flight 遷移を
    防御
  - 影響範囲: 内部実装変更のみ、公開 API（KERNEL_API.md §1〜§5）不変。
    `Save-ResumeState` は §6 内部実装。`resume_state.json` も §6 内部
    artifact で外部 consumer 契約なし。モジュール側への波及ゼロ
  - 修正規模: 削除 ~10 行、追加 ~20 行（fallback warning 含む）。
    純粋に簡素化方向

### Added
- modules/standard/temp_ipaddress_config: **NEW** モジュール。
  顧客現場で本番 IP がまだ旧 PC で使用中の状況に、CSV プールから未使用
  IP を自動検出して NIC に付与する一時 IP 割当モジュール。
  VERSION=`0.1.0`、REQUIRES_KERNEL=`2.0.0`（Min Kernel API = 2.0.0、
  KERNEL_API §1〜§5 baseline のみ使用）
  - Order=11、Network カテゴリ
  - 11 列 CSV: Enabled / IPAddress / SubnetPrefix / Gateway / DNS1 / DNS2 /
    DNS3 / AdapterPattern（`ipv6_config` 同方式の wildcard）/ Description /
    Segment
  - GUI ベース作業者選択（probe ベースの自動選択は採用せず）:
    - **理由**: probe（ICMP/ARP）は送信元 IP が必要で、真新しい kitting
      PC（IP 未取得）では一切機能しない。また切断中 PC が後で再接続する
      collision は probe では原理的に検出不能。これらの制約を honest に
      受け入れ、作業者の状況把握 + 口頭調整 + Windows DAD を組み合わせる
      設計とした
    - **WinForms ダイアログ**: AutoPilot mode に関わらず必ず modal 表示。
      pool 一覧（IP / Prefix / Gateway / Description / Status）を表示、
      [CURRENT] / [DUPLICATE] マーカーで状況提示。作業者が選択して [Assign]
    - **Sticky**: 現 NIC IP が pool 内なら GUI で初期選択（Enter で keep
      可能、別 IP 選択で reassignment）
    - **DAD 検証**: assignment 後 500ms 待機して `Get-NetIPAddress` の
      AddressState を確認。Duplicate なら roll back + GUI を再表示して
      当該 IP を gray-out した [DUPLICATE] マーカーで除外、作業者が別 IP
      を選び直し（無限ループ可能、cancel で中止）
  - NIC 解決: 全行で同じ AdapterPattern を要求（v1）。`-like` で 1 NIC に
    一意解決できなければ Error
  - subnet 整合性チェック: NIC が pool と異なる subnet の場合、警告を
    GUI ヘッダにも表示（probe しないので致命ではないが、作業者への情報）
  - Post-Apply Verification: PrefixLength / AddressState=Preferred /
    DefaultGateway route / DNS server list / Gateway L3 ping
  - 既知の制約（Guide.txt に明記）: 切断中 PC との collision は本モジュール
    では検出不能。出荷前検査（Status Monitor 監視）で最終チェックする運用
  - 既存 `ipaddress_config`（本番 IP、SELECTED_ETH_IP env var 経由）とは
    排他。Profile で切替（一時運用フェーズ → 本番切替フェーズ）

### Added
- modules/standard/firewall_rule_make_config: **NEW** モジュール。
  CSV 定義から個別の Windows ファイアウォール rule を `New-NetFirewallRule`
  で作成。VERSION=`0.1.0`、REQUIRES_KERNEL=`2.0.0`（Min Kernel API = 2.0.0、
  KERNEL_API §1〜§5 baseline のみ使用）
  - Order=43、Security カテゴリ
  - 24 列 CSV: 必須 4（Enabled/DisplayName/Direction/Action）+ ID 1
    （Name、optional）+ 推奨 5（RuleEnabled/Profile/Protocol/LocalPort/
    RemoteAddress）+ ネットワーク詳細 6（Description/Group/RemotePort/
    LocalAddress/IcmpType/EdgeTraversalPolicy）+ 対象 4（Program/Service/
    InterfaceType/InterfaceAlias）+ セキュリティ 3 SDDL（LocalUser/
    RemoteUser/RemoteMachine）+ fabriq 標準の Segment
  - multi-value セパレータは `;`（CSV 標準 `,` との衝突回避）。
    例: `Profile=Domain;Private`、`LocalPort=80;443`
  - 冪等性: Name 指定時は Name で、未指定なら DisplayName で既存チェック
    → ある場合 SKIP（更新せず、操作者が手動削除して再実行する運用）
  - 行レベル validation: Direction / Action / Profile / Protocol / Port 形式 /
    RuleEnabled / EdgeTraversalPolicy / InterfaceType を pre-apply で
    チェック。不正行は Fail カウントに加算し他の行は継続（部分成功を許容）。
    Program パス不在は warn のみ
  - splat hash パターンで `New-NetFirewallRule` を呼出（空セルは parameter
    自体を渡さず cmdlet 既定値を採用）
  - Post-Apply Verification: 作成 rule を `Get-NetFirewallRule -Name`
    （GUID 一致）で読み戻し、Name（指定時）/ Direction / Action / Enabled /
    Profile / EdgeTraversalPolicy（指定時）を比較（Profile は集合として
    順序非依存）。InterfaceType / InterfaceAlias / *User / RemoteMachine は
    別 Filter cmdlet 経由のため v1 では verification 対象外
  - 三層責務分離: `firewall_config`（profile on/off） /
    `firewall_rule_config`（全体 backup/restore） /
    `firewall_rule_make_config`（個別 rule 作成）。Profile での順序
    （Import → Maker → firewall_config）を Guide.txt に明記

### Added
- modules/standard/firewall_rule_config: **NEW** モジュール。Windows
  ファイアウォール全体（rule + profile 状態 + logging 設定 + IPsec）を
  `.wfw` 形式で丸ごとバックアップ／復元する 2 スクリプト構成。
  `VERSION=0.1.0`、`REQUIRES_KERNEL=2.0.0`（Min Kernel API = 2.0.0、
  KERNEL_API §1〜§5 baseline のみ使用）
  - `firewall_rule_export.ps1`（Order=41）: `netsh advfirewall export` で
    `policy.wfw` を採取。透明性のため監査用サイドカー
    （`rules_show.txt` / `rules.json` / `profiles.json` / `manifest.txt`）
    を併産。保存先は module-local `backup/<yyyyMMdd_HHmmss>/`、CSV の
    `DestinationPath` を指定すれば任意のパスに変更可能
  - `firewall_rule_import.ps1`（Order=42）: `netsh advfirewall import` で
    全ポリシー復元（破壊的）。CSV の `IAcknowledgeReplace=1` 必須の
    暴発防止ゲート。`manifest.txt` 同梱時は OS 版・期待 rule 数を読んで
    現在環境と比較し、相違時に警告
  - `firewall_rule_export` が成功するたび、`firewall_rule_list.csv` の
    末尾に対応する Import 行を自動追加（`Enabled=0` /
    `IAcknowledgeReplace=0` 固定で暴発防止）。`SourcePath` は backup/
    配下なら相対形式（`<ts>/policy.wfw`）、外部 destination 指定時は
    絶対パス。`Segment` は元 Export 行から継承。CSV エンコーディング
    （UTF-8 BOM / ANSI）と末尾改行状態を既存ファイルから検出して
    保持。Excel ファイルロック等で追記失敗時は warn のみで Export
    自体は success のまま継続
  - 両スクリプトで netsh.exe の stdout 取り込み時に
    `Invoke-NetshCapture` ヘルパーを使用。.NET `ProcessStartInfo` で
    `StandardOutputEncoding=UTF-8` を設定して netsh の redirected
    stdout を直接デコード（経験的検証: Win10/11 上で netsh は redirected
    stdout に UTF-8 で書き出す。raw byte で `E8 A6 8F E5 89 87 E5 90 8D`
    = "規則名" を確認）。PS 5.1 `& exe 2>&1` 取り込みが
    `[Console]::OutputEncoding=932` で UTF-8 バイト列を CP932 として
    誤デコードし mojibake（`規則名:` → `隕丞援蜷・`）になる問題を、
    `[Console]` 状態を変更しない side-effect-free な方法で回避。
    （`evidence_config.ps1` の `Invoke-CScriptCapture` は cscript 用の
    OEM デコードで、netsh とは異なる stdout encoding を持つため別実装）
  - `firewall_rule_import` の SourcePath 解決規則:
    (1) 絶対パス（drive letter / UNC）はそのまま使用、
    (2) 相対パスは `<module>\backup\` 配下を基準に解決、
    (3) 解決後がディレクトリの場合 `policy.wfw` を自動付与。
    プレビューに CSV 記述値と (resolved) 絶対パスの両方を表示
  - Post-Apply Verification（Name-set 包含方式）:
    - Export: `policy.wfw` 存在 + サイズ > 1KB + `rule_names.txt` 存在 +
      行数 == 記録 rule 数。**現在の `Get-NetFirewallRule` count との
      比較は意図的に外している**（Windows の background 動的変動 -
      mpssvc / AppX / GPO による rule の add/remove - で誤 fail を出す
      ため）
    - Export 成果物に `rule_names.txt` を追加（1 行 1 個の rule Name、
      Name-set 検証用 sidecar）
    - Import: `rule_names.txt` がある場合、記録 Name 集合が import 後の
      `Get-NetFirewallRule` 集合に **包含** されているかで判定（after が
      expected の superset であれば PASS）。dynamic rule が追加されても
      影響を受けず、特定 rule の欠損を直接検出。欠損があれば最大 5 個まで
      Name を表示
    - Import: `rule_names.txt` 不在の旧 backup では従来の count 一致
      検証に fallback（後方互換）。verify ログに方式 `[name-set ...]` /
      `[count-only ...]` / `[weak ...]` を表示
  - 既存 `firewall_config`（profile レベル on/off）とは責務分離。Profile
    で併用する場合の順序ルール（Import → firewall_config）を Guide.txt
    に明記
  - 動機: 大規模案件で複雑な firewall ルールを base image から複製する
    必要が生じた際、`*-NetFirewallRule` cmdlet で per-rule に再構築する
    アプローチは Filter オブジェクト群の往復で取りこぼしを起こすため、
    netsh の `.wfw` 形式（rule + profile + logging + IPsec を一括復元）
    を真実源として採用

## [2.2.0] - 2026-04-23

### Added
- modules/standard/evidence_config: Section 10「PC Serial Number」を
  多重ソース収集型に拡張（`1.1.1` → `1.2.0`、MINOR / 後方互換な機能追加）
  - 動機: 2400 台規模の案件で `Win32_BIOS.SerialNumber` が空を返す機器が
    ~2 台程度出現（SMBIOS Type 0 未投入 / Type 1 のみ投入の OEM ケース）。
    また OEM 出荷前焼き込み漏れで `"Default string"` が返る機器も散見
  - 実装: Phase 1 で全ソースを独立 try/catch で取得 → Phase 2 で canonical
    を優先順位で選定 → Phase 3 で全ソースを txt に記録
  - 収集ソース:
      a. `Win32_BIOS.SerialNumber`（canonical 優先 1、従来唯一のソース）
      b. `Win32_ComputerSystemProduct.IdentifyingNumber`（canonical 優先 2）
      c. `Win32_SystemEnclosure.SerialNumber`（canonical 優先 3）
      d. `Win32_BaseBoard.SerialNumber`（記録のみ、母板 SN は canonical 候補外）
      e. `HKLM:\HARDWARE\DESCRIPTION\System\BIOS\SystemSerialNumber`
         （canonical 優先 4、最終 fallback）
      f. `Win32_ComputerSystemProduct.UUID`（参考 ID、SN ではない）
  - 無効値拒否リスト拡張: 従来 `"Default string"` 等に加え `"Not Applicable"`
    / `"Not Specified"` / `"OEM"` / `"Chassis Serial Number"` / 全桁ゼロ /
    全桁ドット・ハイフン・空白を追加
  - 出力仕様: 1 行目 = canonical SN（従来互換の grep 位置）、
    その下に全ソースの値と `[VALID, MATCH]` / `[VALID, DIFFERENT]` /
    `[INVALID: <reason>]` / `[QUERY FAILED: <reason>]` タグ、
    末尾に選定ポリシー
  - Section 成否: canonical が 1 件でも決定できれば success、全滅時のみ fail
  - `Get-HardwareUniqueId`（kernel 側・出力フォルダ名に使う UID）は今回
    未変更（既に MAC fallback があり欠番化しないため）。必要になったら
    別コミットで拡張
- dev/seed_module_versions.ps1: 全モジュールへの VERSION/REQUIRES_KERNEL
  baseline 一斉 seed スクリプト（idempotent）
- **全 71 モジュールに `VERSION=1.0.0` と `REQUIRES_KERNEL=2.0.0` を
  baseline として配備**（142 ファイル新規作成）。既存の
  `evidence_config/VERSION=1.1.1` と `odt_config/VERSION=1.0.0` は
  履歴保持のため書き換えず
  - 動機: fabriq_studio の「update from template」機能の実機検証中に、
    両側 VERSION 欠損モジュールが SKIP 扱いとなり、現実には差分がある
    のに studio が overlay を取りこぼす問題が発見された
  - 解決: 一斉 seed で全モジュールに VERSION を行き渡らせ、`KERNEL_API.md
    § 9.4` の「template あり/target 欠損 = UPDATE（lazy seed）」パターン
    を全モジュールで機能させる
  - `CLAUDE.md` ルール H / J を baseline seed 済み前提に更新。lazy seed
    方針を「個別 touched 時の打刻」から「baseline + 以降の SemVer bump」
    に切り替え（歴史的経緯セクションで lazy seed 方針からの移行を記録）

### Added
- dev/framework_overlay_rules.json: **NEW** フレームワーク更新ルールの
  マニフェスト（schemaVersion 1）。`dev/build_framework_patch.ps1` と
  外部更新ツール（fabriq_studio 等）が共通で consume する単一真実源
  - `excludeDirsTopLevel` / `excludeDirsRecursive` / `excludeFilesKernelLevel`
    / `moduleCsvWhitelist` を列挙
  - `bundles.kernel` / `bundles.module` で version 単位を定義
    - kernel bundle は `kernel/`, `apps/`, `commands/`, `dev/`, および
      トップレベル framework ファイルを包含（`VERSION` ファイルを持たない
      apps/commands/dev は kernel バンドルに束ねる判断）
    - module bundle は各モジュール独立に `modules/<type>/<name>/VERSION`
      で管理
  - `profiles/` 配下は `excludeDirsRecursive` に含まれ、**更新時に丸ごと
    保護**（`Master_*.csv` / `_test_harness*.csv` / `sysprep.csv` /
    `easy_template/` も対象）。プロファイル書式のアップデートが発生
    しても既存を絶対優先の方針
- kernel/KERNEL_API.md: **§ 9「更新・オーバーレイ契約（外部ツール向け
  公開契約）」を新設**。外部更新ツールが依存できる公開契約を明文化
  - 9.1 マニフェスト JSON の位置付け、9.2 bundle 定義、9.3 site-specific
    絶対保護、9.4 SemVer 比較セマンティクス（4 状態）、9.5
    REQUIRES_KERNEL 事前チェック、9.6 新/欠損モジュールの扱い、9.7
    更新前の安全チェック、9.8 schemaVersion 後方互換
  - 本セクション追加は **公開 API サーフェスへの追加**のため、
    次回リリース時に KERNEL_VERSION の MINOR 昇格（2.1.0 → 2.2.0）
    が必要（現時点では [Unreleased] 扱いで `KERNEL_VERSION` は据え置き）

### Changed
- dev/build_framework_patch.ps1: マニフェスト消費型にリファクタ。
  除外ルールのハードコードを廃し、`dev/framework_overlay_rules.json`
  を読み込む方式に変更（外部ツールとルール同期を保証）
  - 挙動変更: `profiles/` 全体を recursive 除外（従来は
    `Custom Plan.csv` 1 ファイルのみ除外）。`Master_*.csv` /
    `_test_harness*.csv` / `sysprep.csv` / `easy_template/` は
    framework patch に含まれなくなる（上記 9.3 の方針と整合）
  - 手順を 4 ステップから 5 ステップに再編（Step 2 = 再帰ディレクトリ
    除外、Step 3 = kernel 個別ファイル除外）
  - PATCH_README.md の「Excluded」セクションに profiles 保護方針を追記

### Changed
- apps/fabriq_operator/lib/session_form.ps1: Target Host Grid に検索
  フィルタ + 列ソートを追加（大規模案件でのターゲット選択を改善）
  - 検索対象列: `AdminID` と `NewPCName` のみ（大文字小文字無視の
    部分一致）。`OldPCName` / `EthernetIP` / `Pin` は意図的に対象外
    （Pin は機密性、他は操作者が覚えている可能性が低いため）
  - ヒット件数表示（`{visible} / {total}`）を右側ラベルに表示
  - Host Grid / Worker Grid ともに全列で `SortMode = Automatic` を明示
    （列ヘッダクリックで昇順/降順切替）
  - 検索欄 UX:
      - Esc で検索欄クリア
      - Enter で passphrase 欄へフォーカス移動
      - 文字入力中にエラーメッセージも自動クリア
  - 実装方針: DataGridView.Rows.Clear() + 再ポピュレート時に選択ホスト
    参照が壊れないよう、各行の `.Tag` プロパティにソースの PSCustomObject
    を格納。Start Session ハンドラは `$HostList[$index]` ではなく
    `$hostGrid.SelectedRows[0].Tag` から解決（sort + filter 状態に非依存）
  - 自動検出ヒント（`* Auto-detected: ...`）の挙動は従来通り維持。
    検索を空に戻した際は自動検出ホストを再選択
  - フォーム高 600 → 630 px に拡張（検索行の +30px 分）

### Added
- dev/build_framework_patch.ps1: 現行 fabriq ツリーから設定用 CSV と
  ランタイム成果物を除外した「フレームワーク全面アップデート用パッチ」
  を任意のディレクトリに生成する再利用可能スクリプト
  - パラメータ: `-OutDir`（デフォルト: Desktop）, `-PatchName`（省略時:
    `fabriq_patch_{日付}_kernel-v{KERNEL_VERSION}`）, `-Purpose`
    （PATCH_README.md への追記用自由文、省略可）
  - 除外ルール: 既存 `make_fabriq_patch.ps1`（Desktop 上の kernel-v2.1.0
    リリース用の一回限りスクリプト）の方針をそのまま踏襲
      - `modules/` 配下の全 CSV から `module.csv` / `preset.csv` のみ保持
      - `kernel/csv/hostlist.csv` / `workers.csv` / `log_destinations.csv`
      - ランタイム成果物（`kernel/json/*.json`, `art_pulse.txt`,
        `skip_request.flag`, `passphrase_verify.txt`, `silence.flag`）
      - `profiles/Custom Plan.csv`
      - `.git/`, `.claude/`, `evidence/`, `logs/`
  - PATCH_README.md を自動生成（`$kernelVersion` / `$PatchName` /
    `$Purpose` を埋め込み）
  - 用途: `powershell.exe -File .\dev\build_framework_patch.ps1` 単体で
    実行可能。Claude がリクエストを受けてそのまま実行する運用を想定
- modules/standard/evidence_config: Section 21「Windows License / Activation
  Status」および Section 22「Office License / Activation Status」を追加。
  ライセンス認証状況のエビデンス取得を強化
  - Section 21（`21_WindowsLicense.txt`）:
      a. `SoftwareLicensingProduct`（CIM）: 製品別 LicenseStatus /
         PartialProductKey / ProductKeyChannel / KMS Machine
      b. `SoftwareLicensingService`（CIM）: ClientMachineID、マシン全体の
         KMS 設定
      c. `slmgr /dlv` の raw 出力（Installation ID 含む診断情報）
  - Section 22（`22_OfficeLicense.txt`）:
      a. Click-to-Run レジストリ（`HKLM\SOFTWARE\Microsoft\Office\ClickToRun\
         Configuration` の ProductReleaseIds / VersionToReport /
         UpdateChannel / Platform など）
      b. OSPP.vbs 自動検出（`office_license_auth.ps1` と同一候補パスを
         流用）+ `cscript OSPP.vbs /dstatus` の raw 出力
      - OSPP.vbs 不在時は C2R レジストリのみ記録
  - PS 5.1 の CP932 キャプチャ問題回避のため、cscript 呼び出しは
    Section 16（WiFi Profiles）と同じ `cmd /c "chcp 65001 && ..."` パターン
    を採用
  - `modules/standard/evidence_config/VERSION` を `1.0.0` → `1.1.0` に昇格
    （MINOR / エビデンス収集項目の後方互換な追加）。`REQUIRES_KERNEL` は
    `2.0.0` のまま（新規公開 API 依存なし）
- modules/standard/evidence_config: Section 20「System TEMP Text-Log Backup」
  を追加。C:\Windows\Temp 直下の `.log` / `.txt` ファイルを `20_TempBackup\`
  に非再帰でバックアップするセーフティネット
  - 目的: ODT / ドライバ / インストーラ系の原本ログを非常時調査用に保持
  - ロック中ファイルは静的スキップ、サイズ上限なし、サブディレクトリ非対象
  - touched 初回のため `modules/standard/evidence_config/VERSION`（`1.0.0`）
    と `REQUIRES_KERNEL`（`2.0.0`）を打刻（ルール H/J）

### Fixed
- modules/standard/evidence_config: Section 21「Windows License」および
  Section 22「Office License」の raw 出力で日本語が文字化けする不具合
  を修正（`modules/standard/evidence_config/VERSION` を `1.1.0` →
  `1.1.1` に昇格、PATCH）
  - 原因: cscript は WScript.Echo 出力時に親 cmd の `chcp 65001` を
    尊重せず、JP ロケールでは OEM コードページ（CP932）のバイト列を
    吐く。これを UTF-8 として取り込んだ結果、U+FFFD 置換文字に化けて
    いた（Section 16 の netsh は chcp を尊重するため問題なし）
  - 修正方針: cscript 安全呼び出しヘルパー `Invoke-CScriptCapture` を
    追加し、`System.Diagnostics.Process` で stdout をリダイレクト、
    `StandardOutputEncoding` にカルチャの OEM コードページを指定して
    .NET 側に正しくデコードさせる。内部的には UTF-16 の .NET 文字列と
    なるため、後続の `Out-File -Encoding UTF8` で正しい UTF-8 が書かれる
  - `cscript //U` で UTF-16LE redirected stdout にする案も検討したが、
    現行 Windows 実機検証で //U が 0 バイト出力となるケースがあり採用
    せず。OEM コードページデコード方式はロケール透過で堅牢
  - slmgr /dlv と OSPP /dstatus の両方に適用
- kernel/common.ps1: `Export-HtmlChecklist` の HTML レポートで「Target PC
  (New)」「Hardware ID」メタカードと `<title>` タグ内の PC 名が空で出力
  される不具合を修正
  - 原因: `$pcName` / `$uid` の解決が `$global:FabriqEvidenceBasePath`
    未設定時の fallback ファイル名分岐内でしか行われておらず、通常運用
    （セッション初期化後は FabriqEvidenceBasePath が必ずセットされる）
    では両変数が未定義のまま HTML テンプレートに埋め込まれていた
  - 修正: `$pcName` / `$uid` の解決を if/else 分岐の外側に移動し、どちら
    の分岐でも HTML 本文で参照できるようにした
- modules/standard/odt_config: ODT セットアップログの収集パターンが現行
  setup.exe の出力形式と一致しておらず、エビデンスが常に 0 件だった
  - 収集フィルタを `SetupExe(*.log)` から `$env:COMPUTERNAME-*.log` +
    エントリ開始時刻以降の LastWriteTime に変更。ODT が実際に書き出す
    `{COMPUTERNAME}-{yyyyMMdd}-{HHmm}[a-z].log` 命名を捕捉
  - 1 エントリで複数ログが生成されるケースに対応（`-First 1` を廃止し
    該当全件をコピー）
  - 保存先を共通バケット `.\evidence\odt_log\` に統一（セッション毎の
    `FabriqEvidenceBasePath` 配下への分散配置を廃止）
  - ファイル名衝突対策としてセッションタグ（`{ts}_{PCName}_{Serial}`）
    をプレフィクスに付与。工場出荷ホスト名・クローン環境での上書き損失
    を防止
  - touched 初回のため `modules/standard/odt_config/VERSION`（`1.0.0`）と
    `REQUIRES_KERNEL`（`2.0.0`）を打刻（ルール H/J）

### Added
- compat tracking (Layer 1): モジュール touched 時の API 依存スキャン運用を
  開始。将来の中央コンパチマトリクス（Layer 3）の元データを段階的に蓄積
  - `kernel/KERNEL_API.md`: §8「API Version History」を追加（各公開 API
    の導入バージョン追跡、Min Kernel API 判定用）
  - `CLAUDE.md`: ルール I（モジュール touched 時の API 依存スキャン必須）、
    ルール J（`REQUIRES_KERNEL` ファイル lazy seed 運用）、Layer 3 の将来
    実装計画を追加
  - `dev/template/REQUIRES_KERNEL` 新規（初期値 `2.0.0`、新規モジュール用）

### Notes
- 既存 73 モジュールには `REQUIRES_KERNEL` を一斉配布しない。Claude が
  touched した時点から lazy に打刻していく（既存環境への影響ゼロを維持）
- Layer 3（`kernel/MODULE_COMPAT.md` の自動生成）は現時点では未実装。
  Layer 2 データ（`REQUIRES_KERNEL`）が貯まってから `dev/build_compat_matrix.ps1`
  として実装する想定
- 本変更は公開 API には影響しないため `KERNEL_VERSION` は据え置き

## [2.1.0] - 2026-04-23

### Added
- async execution: `__ASYNC__` マーカー以降のモジュールを子 Runspace で実行
  し、ハング時に Status Monitor の **Skip** ボタンで強制中断可能に
  - `kernel/common.ps1`: `Invoke-SafeCommandAsync`, `Get-FabriqAsyncConfig`
    を追加（既存 `Invoke-SafeCommand` は無変更）
  - `kernel/common.ps1`: `Resolve-ProfileModules` に `__ASYNC__` マーカー
    認識と `_IsAsync` フラグ付与を追加
  - `kernel/main.ps1`: `Invoke-BatchExecution` に async/sync 分岐を追加
  - `kernel/ps1/status_monitor.ps1`: Skip ボタン追加（`skip_request.flag`
    を書き出して async ポーリングループに通知）
  - `kernel/json/async_config.json` 新規（`Enabled` / `DefaultTimeoutSec`
    / `PollIntervalMs` / `SkipFlagPath`）
  - `profiles/_test_harness_async.csv` 新規（Phase 1 動作検証用）
  - `modules/standard/test_harness_config/test_harness_list.csv`:
    `hang_sim` / `async_ok` セグメント追加
  - `kernel/KERNEL_API.md`: 特殊マーカー表に `__ASYNC__` を追加
  - `README.md`: 特殊マーカー表に `__ASYNC__` を追加

### Changed
- バージョン管理をカーネル API とモジュール独立の 2 軸に再設計
  - 新規: `kernel/KERNEL_VERSION`（カーネル API 真のソース、初期値 `2.0.0`）
  - 新規: `kernel/KERNEL_API.md`（公開 API サーフェスの明文化）
  - 新規: `dev/template/VERSION`（新規モジュール用、初期値 `0.1.0`）
  - 廃止: ルート直下の `VERSION` ファイル（「全体版」概念の廃止）
  - `CLAUDE.md`: バージョン管理ルールを全面改訂（実装前宣言ルール E、
    実装サマリ報告ルール F、`KERNEL_API.md` 同期ルール G、モジュール
    `VERSION` 運用ルール H を追加）
  - `dev/check_version.ps1`: `KERNEL_VERSION` 基準に切り替え
  - `README.md` L1 / `kernel/common.ps1` L2 / `kernel/main.ps1` L3 /
    `main.ps1` 起動表示 / HTML チェックリストフッター を `2.0` に同期
- `kernel/KERNEL_VERSION`: `2.0.0` → `2.1.0`（MINOR、`__ASYNC__` マーカー
  追加による公開 API の後方互換な拡張）
- ヘッダ版表記を `2.0` → `2.1` に同期（README L1 / common.ps1 L2 /
  main.ps1 L3 / main.ps1 起動表示 / HTML チェックリストフッター）

### Notes
- `KERNEL_VERSION=2.0.0` は現行カーネルの状態を遡及的に formal SemVer
  の出発点として定義した値。過去の `VERSION=2.2.0` は「全体ディストリ版」
  という別概念だったため単純な規格合わせではなく、意味論的に新しい系列
- ランタイムでのモジュール互換性チェックは導入しない（誤判定で現場が
  止まるリスクを避けるため）。代わりに Claude が実装前後で
  `KERNEL_API.md` を参照して手動で整合性を担保する
- async 実装は既存の同期実行経路を無変更に保つ opt-in 方式。`__ASYNC__`
  マーカーを含まないプロファイルの挙動は完全に従来通り。モジュール
  全 73 件は untouched で互換性は保持される
- `async_config.json` の `Enabled: false` で Kill switch として機能し、
  `__ASYNC__` マーカーは silent ignore され全モジュールが同期実行される
- Skip / Timeout による強制中断は `$ps.Stop()` を使うため、モジュールの
  `try/finally` の実行は保証されない（レジストリ/ファイル書き込み途中の
  中断はシステム状態不整合の可能性あり）。実行履歴には警告文言が残る

## [2.2.0] - 2026-04-22

### Added
- `VERSION` ファイルをルート直下に追加（単一の真実のソース）
- `CHANGELOG.md` を新設（Keep a Changelog 形式）
- `dev/check_version.ps1` を追加（版番号整合性チェックスクリプト）
- `CLAUDE.md` に「バージョン管理ルール」セクションを追加

### Changed
- `README.md` / `kernel/common.ps1` / `kernel/main.ps1` の版番号表記を `2.2.0` に統一
  （従来 README と main.ps1 が `ver2.1`、common.ps1 が `v2.2` で不整合）

### Notes
- 本バージョンは過去コミット（Git 237 件）をまとめた初回タグ基準点。
  以降、機能追加・変更は必ず `[Unreleased]` へ追記し、リリース時に本形式で昇格する。
