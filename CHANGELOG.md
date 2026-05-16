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

### Fixed
- apps/fabriq_backuper v0.9.2 → v0.9.3 (Phase 2.9.2a-v2 / PST 検出を subkey walk に切替):
  Phase 2.9.2a で実装した EIDMSW EntryID parser が **本番 Outlook 環境では
  動作しない** ことが実機で判明 (Microsoft 公開 spec は sample wrapped PST
  provider 用、production の PST provider は `MAPIUID + 2 nulls + "mspst.dll" +
  undocumented bytes + path` という別 format)。検出戦略を二段構えに変更。
  - **新規 helper**: `Get-PstPathsFromProfile -ProfileKeyPath <path>` —
    profile registry を recursive 走査して `001f6700` (REG_BINARY UTF-16LE)
    を持つ subkey をすべて列挙 → PST file paths のリストを返す。これは
    deterministic で EntryID format に依存しない
  - **検出戦略**:
    - profile に PST 1 つ → そのまま採用 (`detectionMethod: profile-subkey-walk`)
    - profile に PST 複数 → EntryID parser で disambiguate 試行、成功なら採用
      (`detectionMethod: entryid-parse-confirmed-by-walk`)、失敗なら fail loud
    - profile に PST 0 個 → EntryID parser 単独試行
      (`detectionMethod: entryid-parse-only`)、失敗なら fail loud
  - **manifest 拡張**: `pst.profileCandidates` (profile 内 PST 一覧) 追加。
    multi-PST 環境の engineer 解析用
  - **diagnostic 強化**: EntryID hex dump を 32 → 128 bytes に拡大。
    将来 production EntryID format を offline 解析する時の参考データ用
  - operator 視点は不変: 失敗 = engineer escalate、判断要求なし

### Added
- apps/fabriq_backuper v0.9.1 → v0.9.2 (Phase 2.9.2a / outlook_pop PST mapping detection):
  Backup 時に各 POP3 アカウントが使用している **PST file path を deterministic
  に検出** して manifest に記録するように拡張。Phase B (restore 自動化) の
  「PST を `<email>.pst` に rename して path-collision attach」設計の前提を
  実装したもの。
  - **新規 helper**: `Get-PstPathFromDeliveryStoreEntryId -EntryIdBytes <byte[]>`
    - POP アカウントの `Delivery Store EntryID` (REG_BINARY) を Microsoft
      公式の wrapped PST store provider EIDMS / EIDMSW format に従って parse
    - byte 21 が 0x00 なら EIDMSW (Unicode) / それ以外なら EIDMS (ASCII) と
      判別、それぞれ対応する null-terminator (0x00 / 0x00 0x00) まで読む
    - validation: 抽出パスが drive letter で始まり `.pst` で終わる場合のみ
      採用、それ以外は `$null` を返す
  - **manifest 拡張**: 各 account に `pst` ブロック追加
    - `sourcePath` / `sourceFileName` / `detectionMethod` /
      `detectionStatus` (`present` / `path-only` / `unavailable`) /
      `detectionReason`
  - **失敗時挙動**: parse 失敗時は section 全体は `Success` のまま、当該
    account のみ `pst: { sourcePath: null }` で記録 + warnings に EntryID
    先頭 32 bytes の hex dump を残す (engineer 解析用)。restore 側は
    sourcePath 不在の account を `Failed` で扱う
  - operator 視点: 何も求められない (失敗しても section 全体は Success、
    restore 段階で初めて該当 account 単位で fail / engineer escalate)。
    `feedback_no_operator_judgment` 原則準拠

### Fixed
- apps/fabriq_backuper v0.9.0 → v0.9.1 (Phase 2.9.0a / outlook_pop value 名修正 + SectionParam 配線):
  Phase 2.9.0 Phase A 実機テスト (`E:\test\outlook_pop\manifest.json`) で
  pop3.userName / smtp.userName / smtp.useAuth が null となっていた問題を修正。
  - 修正した registry value 名 (実機 Outlook 16.0 で動作確認):
    - `POP3 User Name` → **`POP3 User`**
    - `SMTP User Name` → **`SMTP User`**
    - `SMTP Use Sicily` → **`SMTP Use Auth`**
    - (`POP3 Use Sicily` は SPA 認証フラグとして正しいので維持)
  - backup_view.ps1: `outlook_pop` セクションへ `SourceUserProfilePath`
    SectionParam を渡すよう拡張。これまで printer/userdata のみだったため
    cross-user 時に sourceUser メタデータが null だった (今回はローカル
    admin = HKCU 同一ユーザだったので実害なし)
  - Phase A 実機検証ログ:
    - 1 POP アカウント (`suzuki@e-cri.co.jp` @ sv3061.xserver.jp:995 SSL) を
      正常採取
    - 同一プロファイル下の subKey 00000001/00000003 (POP3/IMAP どちらの
      Server 値も持たない、LDAP/AddressBook 等の補助 entry) は意図通り
      `otherSkipped` として除外

### Added
- apps/fabriq_backuper v0.8.0 → v0.9.0 (Phase 2.9.0 Phase A / Outlook POP アカウント backup):
  Outlook (classic 2016/2019/2021/365、registry version 16.0) の **POP3
  アカウント設定を registry から JSON manifest にダンプする** 新セクション
  `outlook_pop` を追加。**Phase A は backup のみ** (restore は次フェーズ)。
  - **新規セクション**:
    - `data/sections.csv` に `outlook_pop` 追加 (既定 Enabled=0、明示選択方式)
    - `lib/sections/outlook_pop/backup.ps1` (新規 ~250 行) — HKCU 配下の
      `Software\Microsoft\Office\16.0\Outlook\Profiles\<prof>\
      9375CFF0413111d3B88A00104B2A6676\<NN>` を walk、各 POP3 アカウントから:
        - 文字列値 (REG_BINARY UTF-16LE) → `[Text.Encoding]::Unicode.GetString` で decode + 末尾 null 除去
        - 数値 (REG_DWORD) → そのまま整数として取得
        - Account Name / Display Name / Email / POP3 Server,User,Port,SSL,Sicily / SMTP Server,User,Port,SSL,Sicily,AuthMethod を採取
      manifest schema: `fabriq-outlook-pop-backup` v1。`items.profiles[].accounts[]` 構造
    - `lib/sections/outlook_pop/restore.ps1` (Phase B 用 stub、`Status='Skipped'`)
  - **意図的な除外**:
    - Password (DPAPI per-user/per-machine 暗号化、cross-PC 復号不可) は **manifest に含めない**。
      ただし `passwordStored` boolean (POP3/SMTP それぞれ存在フラグ) は trace 目的で記録
    - IMAP アカウントは検出 + count するが backup 対象外 (Phase A スコープ外、IMAP Server 値で判定)
    - 新 Outlook (2024〜、クラウド設定) は対象外
  - **fabriq 既存パターン踏襲**:
    - `Resolve-HkcuRoot` で admin 昇格 / cross-user の HKCU redirect に対応
    - `SourceUserProfilePath` SectionParam (userdata セクションと同じ命名) で
      target user の registry 経路を選択可能
    - Outlook 16.0 が見つからなければ 15.0 (Outlook 2013) を fallback 検索
  - **Phase A 検証手順**: backup 実行 → `sections/outlook_pop/manifest.json` を
    確認 → 手書きで Microsoft canonical 7-section PRF を組み立て → 別 PC で
    `outlook.exe /importprf <path>` テスト → 動作確認できたら Phase B (restore
    自動化、PRF 動的生成 + Stop-Process OUTLOOK + importprf + cleanup) へ
  - **PRF 仕様調査結果** (CHANGELOG memo): 元 AI 提案にあった `[ServiceEGS]` /
    `ServiceName=MSEMS` は **Exchange 専用** で POP3 では使えない。POP3 は
    `[Service List]` に `ServiceN=Internet E-mail`、本体は `[AccountN]`
    セクション。また password を PRF に書く方式は **Outlook 2016+ では機能しない**
    (Slipstick / Microsoft KB Q259957 確認済)。Phase B 設計はこの正解情報に
    沿って実装する

### Added
- apps/fabriq_backuper v0.7.10 → v0.8.0 (Phase 2.8.0 / Folder/File Browse 時の env-var 自動変換):
  Add/Edit userdata エントリダイアログで [Folder...] / [File...] からパスを
  選択した際、選択 user のプロファイル配下なら **prefix を自動的に
  `%USERPROFILE%` / `%APPDATA%` / `%LOCALAPPDATA%` に書き戻して入力欄に
  反映**するように変更。CSV に絶対パスではなく env-var 形式で保存されるので、
  cross-user 移行や admin 昇格コンテキストでも portable に維持される。
  - 例:
    - `C:\Users\foo\Desktop` → `%USERPROFILE%\Desktop`
    - `C:\Users\foo\AppData\Roaming\Microsoft\Signatures` → `%APPDATA%\Microsoft\Signatures`
    - `C:\Users\foo\AppData\Local\Google\Chrome\User Data` → `%LOCALAPPDATA%\Google\Chrome\User Data`
    - `D:\Projects\myrepo` → 絶対パスのまま (どの env var とも一致しないため)
  - **新規ヘルパー** (user_selector.ps1): `ConvertTo-EnvVarPath
    -AbsolutePath <p> -UserProfilePath <profile>` — `Expand-PathWithUser`
    の逆操作。longest-prefix-first で `%APPDATA%` / `%LOCALAPPDATA%` /
    `%USERPROFILE%` の順に判定し、いずれにも該当しなければ絶対パスを返す。
    prefix 直後の文字が separator または終端の場合のみマッチ (誤マッチ防止)
  - **呼び出し側** (userdata_edit_dialog.ps1): Folder Browse と File Browse の
    OK 時に `ConvertTo-EnvVarPath` を通してから textbox に反映。`try/catch`
    で fallback (helper が見つからなくても元の動作に戻る)
  - 既存 entry を Edit するときも同じ。手動入力した絶対パスは変換対象外

### Fixed
- apps/fabriq_backuper v0.7.9 → v0.7.10 (Phase 2.7.10 / `$script:` スコープ解決問題):
  Phase 2.7.9 で Get-Command ゲートを撤去したが、streaming/checklist は依然
  反応せず。実機で popup は動くのに streaming は動かない、というユーザ報告から
  原因を `$script:` の dynamic 解決問題と特定。
  - 原因の核心: `Add-ProgressLog` などの関数は `$script:ProgressLogBox` を
    参照しているが、PowerShell の `$script:` は **関数を実行している現在の
    script コンテキスト** によって動的に切り替わる。section script を engine
    が `& $scriptPath` で起動すると、section は独自の script スコープを持つ
    ため、そこから呼ばれた `Add-ProgressLog` 内の `$script:ProgressLogBox`
    は **section のスコープに解決され、null** となっていた。popup は
    `Show-CompletionPopup` が `Invoke-BackupStart` (fabriq script スコープ
    の関数) から呼ばれるため `$script:MainForm` が正しく解決され動作していた、
    という非対称な挙動だった
  - 修正: UI コントロール参照のうち、section から呼ばれるパス上にある 2 個 -
    `$script:ProgressLogBox` → **`$global:Fbp_ProgressLogBox`** /
    `$script:ProgressEntriesGrid` → **`$global:Fbp_ProgressEntriesGrid`**
    に昇格。`Fbp_` prefix で global namespace の汚染を最小化
  - 副作用なし: New-ProgressView / Initialize-ProgressView / Add-ProgressLog /
    Initialize-ProgressEntries / Set-EntryStatus の全箇所を一括書き換え

### Fixed
- apps/fabriq_backuper v0.7.8 → v0.7.9 (Phase 2.7.9 / streaming/checklist 無反応バグ):
  Phase 2.7.4 で入れた robocopy streaming と Phase 2.7.8 で入れた per-entry
  チェックリストが、**実機で全く反応しなかった**問題を修正。
  - 原因: section script (`userdata/backup.ps1` / `userdata/restore.ps1`) を
    engine から `& $scriptPath` で呼んだ際、`Get-Command Add-ProgressLog
    -ErrorAction SilentlyContinue` が **親スコープにある dot-source 済み関数
    を検出できない** ケースがあり、`if` ゲートが常に偽となって Add-ProgressLog /
    Initialize-ProgressEntries / Set-EntryStatus がすべて no-op になっていた。
    `Test-AdminPrivilege` のような **直接呼び出し** は dynamic scope chain で
    解決できるが、`Get-Command` 経由の lookup は別経路を通る挙動
  - 修正: section 内の全 `if (Get-Command X -EA SilentlyContinue) { X ... }`
    パターンを **`try { X ... } catch { }`** に置換。直接呼び出しに変更し、
    console-only fallback の安全性は try/catch で確保
  - 影響: backup / restore とも、Phase 2.7.4 (streaming) と Phase 2.7.8
    (checklist) が **GUI 経由でようやく実際に動作**するようになる

### Added
- apps/fabriq_backuper v0.7.7 → v0.7.8 (Phase 2.7.8 / per-entry 進捗チェックリスト):
  Progress View に **エントリ単位の状態表示 grid** を追加。
  「Documents 終わった → Chrome User Data コピー中 → Desktop pending」が
  一目でわかるようになる。アプリ設定 entry が先に終われば、その時点で
  当該アプリの動作確認を始められるユースケース対応。
  - 新規ヘルパー (progress_view.ps1):
    - `Initialize-ProgressEntries -Entries @(@{Id=...; Label=...}, ...)`
      で planned エントリ一覧をチェックリストに展開
    - `Set-EntryStatus -Id <id> -Status <Pending|InProgress|Done|Partial|Failed|Skipped>`
      で個別エントリの状態を更新。マーカー (`[v] Done` 等) + 色 (Done=green,
      Failed=red, Partial=orange, Skipped=gray, InProgress=blue) を変更
    - DoEvents pump 内蔵で同期 UI スレッドでも即時反映
    - InProgress に変わった行は自動スクロール
  - レイアウト変更: Title (Y=14) → **Entry status grid** (Y=68, h=150) →
    **Detail log** (Y=248, h=362) → Done ボタン (Y=624)。既存の streaming
    log は縮小して併設
  - 呼び出し側 (`userdata/backup.ps1` / `userdata/restore.ps1`):
    planned 配列構築直後に `Initialize-ProgressEntries`、各エントリ処理開始時に
    `Set-EntryStatus -InProgress`、完了時に `Done` / `Partial` / `Failed` /
    `Skipped` のいずれか。`Get-Command` ガードで console 実行時の後方互換維持
  - Label の表示:
    - Backup: Description (空なら SourcePath)
    - Restore: manifest の sourcePath (Description は manifest に保存していない)

### Added
- apps/fabriq_backuper v0.7.6 → v0.7.7 (Phase 2.7.7 / userdata エントリ順並べ替え):
  Backup View の userdata grid に **[Up] / [Down] ボタン**を追加。選択行を
  上下に 1 つずつ移動でき、即時 CSV 保存 (.bak rotation) で永続化される。
  - 配置: editor 行内 Delete と Save changes の間
    (Add / Edit / Delete / Up / Down / Save changes の 6 ボタン構成)
  - 共有ヘルパー `Move-BackupEntry -Delta <-1|+1>` を追加し、Up/Down は
    1 行のラッパーで呼び出し
  - 並べ替え後、選択行のハイライトは移動先に追従
  - **重要**: バックアップ実行時の処理順 = userdata_list.csv の行順 →
    manifest.items.entries[] に焼き込まれる → restore はその順で実行。
    つまり CSV を並べ替えるだけで **次回 backup と その restore 両方** の
    順序が制御される (既存 backup の manifest 順は凍結のため変更不可)

### Changed
- apps/fabriq_backuper v0.7.5 → v0.7.6 (Phase 2.7.6 / robocopy multi-thread bump):
  `userdata/backup.ps1` と `userdata/restore.ps1` の robocopy 引数に
  **`/MT:16`** を追加 (既定 `/MT:8` → 16 に倍化)。
  - 期待効果: AppData / browser cache / Outlook signatures など極小ファイル
    大量系で +10〜20% のスループット改善
  - `/MT:32` (Microsoft MIGRATE 採用値) と比較して保守的に選択:
    HDD destination での seek storm リスクを避ける + log 行錯綜を抑える
  - 1Gbps LAN のワイヤ飽和は既に robocopy 既定で達成済みのため、
    本変更の主効果は **小ファイル系のレイテンシ隠蔽**
  - 大ファイル中心の workload では効果ほぼ無し (差なし)
  - 計測した結果さらに上げたい場合は `/MT:32` への昇格を検討 (1 行修正)

### Added
- apps/fabriq_backuper v0.7.4 → v0.7.5 (Phase 2.7.5 / 完了通知ポップアップ):
  バックアップ/リストア完了時に **MessageBox モーダル**で結果を表示するように。
  Progress View を凝視していなくても完了に気付けるよう改善。
  - **新規ヘルパー** (progress_view.ps1): `Show-CompletionPopup`
    (-Title / -Body / -Status)。Status により icon を出し分け:
    `Success` → Information, `Partial` → Warning, `Failed` → Error
  - Backup 完了 popup の本文: Status + Elapsed + Data size + Files/Dirs +
    Entries + Written to (AggregateDir)
  - Restore 完了 popup の本文: Status + Elapsed + Entries (success/skip/
    fail) + Restored from (AggregateDir)
  - 副次効果: MessageBox 出現時に OS 通知音 + タスクバー点滅が自動発生
    (ユーザが他作業中でも気付ける), `$script:MainForm.Activate()` でフォーム
    自体も foreground に持ち上げ

### Added
- apps/fabriq_backuper v0.7.3 → v0.7.4 (Phase 2.7.4 / 進行可視化 + Run summary):
  バックアップ/リストアが silent に見える問題と、所要時間 / データ量がわからない
  問題をまとめて解消。
  - **robocopy 出力の line-by-line streaming**: `userdata/backup.ps1` と
    `userdata/restore.ps1` の robocopy 呼び出しを stdout/stderr 一行ずつ
    `Add-ProgressLog` に転送する形に変更。8 行毎に
    `[System.Windows.Forms.Application]::DoEvents()` を pump し、同期 UI
    スレッドでも message queue が処理されて Progress View 上で **ファイル名が
    流れる**ようになる。実行コストは無視できる範囲、Runspace 化なし
  - **Run summary block**: backup_view / restore_view の Invoke-XxxStart
    で全体 stopwatch を取り、終了後に `Run summary:` ブロックを Progress
    View 末尾に追記:
    - Elapsed: 経過時間 (`5m 23s` / `1h 12m 45s` 形式)
    - Data size: 各 section の Summary.totalBytes を集計 (Format-Bytes で
      KB/MB/GB 単位)
    - Files / Dirs / Entries: 利用可能な集計値を表示
    - restore 側は totalBytes が無いので Entries (success/skip/fail) を表示
  - **新規ヘルパー** (progress_view.ps1): `Format-Bytes`, `Format-Duration`
  - backup.ps1 の entry_log.txt は streaming 経由でも書き続ける (artifact
    保持)、restore は元から log file 書き出しなしのため UI 転送のみ

### Fixed
- apps/fabriq_backuper v0.7.2 → v0.7.3 (Phase 2.7.3 / Documents 内ジャンクション追跡問題):
  `%USERPROFILE%\Documents` をバックアップすると、Pictures / Music / Videos
  などの**兄弟フォルダの中身が Documents バックアップ内に混入**し、リストア時
  には移行先の実 Pictures / Music / Videos に書き戻されるという混乱挙動を修正。
  - 原因: Windows の XP 互換 junction (`Documents\My Pictures` →
    `Pictures` 等) を robocopy `/E` が**既定で追跡**するため、recursion が
    NTFS reparse point 越しに兄弟フォルダへ「滲み出して」いた。
    target 側でも同じ junction が存在するため、`backup\data\My Pictures\foo.jpg`
    を target の `Documents\My Pictures\foo.jpg` に書きに行くと NTFS が
    透過的に `Pictures\foo.jpg` へ再ルーティング、これがユーザ体感の
    「Documents に入っていたのに Pictures に移った」現象になっていた
  - 修正: `userdata/backup.ps1` の robocopy 引数に **`/XJD`** (eXclude
    Junction Directories) を追加。本物のサブフォルダは対象外なので副作用なし。
    Documents 配下の `My Pictures`/`My Music`/`My Videos` だけでなく、
    AppData 系の悪名高い再帰 junction (`AppData\Local\Application Data` →
    `AppData\Local` 等) も同時に回避
  - `userdata/restore.ps1` にも `/XJD` を defense-in-depth として追加。
    新規 backup には対象 junction 既に無いので no-op、旧 backup
    (Pictures content が混入したもの) に対しては誤書き込みを防止
  - **注意**: 0.7.2 以前で取得した Documents バックアップは Pictures /
    Music / Videos の内容を含んでいる可能性があるため、本パッチ適用後に
    **再バックアップ推奨**

- apps/fabriq_backuper v0.7.1 → v0.7.2 (Phase 2.7.2 / Browse mode printer list 不可視):
  Browse for backup でバックアップフォルダを選択した後、printer 一覧が grid に
  表示されない問題を 2 件の原因に分解して修正。
  - **Bug A**: `Invoke-RestoreBrowse` で
    `$script:RestoreExplicitDir = $chosen` を先にセットし、その直後に
    `$script:RestoreTimestampCombo.SelectedIndex = -1` を呼んでいたが、combo
    が以前 index 0 だった (= ローカルバックアップが存在する) 場合、
    `SelectedIndexChanged` イベントが発火して **イベントハンドラ内で
    `RestoreExplicitDir = $null` が呼ばれ、直前にセットした値が消去** され、
    `Update-RestoreSelection` が grid を Rows.Clear() で空にしてしまっていた。
    順序を入れ替えて combo を先に -1 にしてから ExplicitDir をセットする形に
    変更。Start Restore 時に "Browse mode が忘れられている" 問題も同時に解消
  - **Bug B**: `Show-RestorePrinterListFromAggregate` に `try { ... } catch { }`
    の silent catch があり、JSON parse 失敗 / `$pm.items.printers` が null /
    その他の例外がすべて握り潰され、grid が空のまま原因がわからない状態に
    なっていた。silent catch を撤廃し、エラー内容を `RestoreManifestLabel` に
    表示するよう変更。さらに manifest が空のケース / printer 件数 (shown / hidden /
    total) もラベルに出すよう改善
- apps/fabriq_backuper v0.7.0 → v0.7.1 (Phase 2.7.1 hotfix / layout clipping):
  Phase 2.7 で form を 900px のまま据え置いたが、ユーザのディスプレイ可視領域
  (タスクバー控除後 ~800px) より大きく **[Start Backup] / [Start Restore] /
  [Done] ボタンが画面下端に隠れていた**。また Add userdata entry ダイアログが
  640px 幅で **[Include ACL (/COPYALL)] checkbox が右端で見切れ、環境変数の
  ヒスト文字列も "(reso..." で切れていた**。
  - main_form.ps1: form 900 → **780** に縮小
  - backup_view / restore_view / progress_view / mode_select_view: 全 view を
    新 form 高さに合わせて再レイアウト。printer grid 高さを縮小 (216→140 backup
    / 418→350 restore)、Y 位置をタイトな配置にシフト、Start/Done ボタンを Y=624 に移動
  - userdata_edit_dialog: dialog 640 → **720** に拡張。SourcePath textbox /
    Exclude / Description テキストボックスを 544px に拡張、Browse ボタンを
    右シフト、IncludeAcl チェックボックスをダイアログ内に収まる位置に調整、
    環境変数ヒント文字列を簡潔化 (1 行に収まる長さに)
  - 内部: backup_view.ps1 の PSScriptAnalyzer warning も同時に解消
    (`$null` を左辺へ、`$sender` を `$src` にリネーム)

### Added
- apps/fabriq_backuper v0.6.1 → v0.7.0 (Phase 2.7 / Userdata editor + per-user path resolution):
  Userdata エントリの本格 GUI 編集機能と、移行元/先で paths に使用するユーザを
  選択する仕組みを追加 (MINOR / 新規 UI 機能 + sections への後方互換オプション追加)。
  - **新規ヘルパー**:
    - `apps/fabriq_backuper/lib/ui/csv_io.ps1` — `Read-UserdataCsv` /
      `Save-UserdataCsv`。書き込み時に 1-gen `.bak` rotation で破壊編集から保護
    - `apps/fabriq_backuper/lib/ui/user_selector.ps1` — `Get-UserProfileList`
      (Win32_UserProfile, Special 除外), `Get-LoggedOnInteractiveProfilePath`
      (Resolve-HkcuRoot 経由で admin 昇格時の logged-on user を解決),
      `Get-DefaultProfileIndex` (logged-on > current process > 0),
      `Expand-PathWithUser` (case-insensitive 置換で
      `%USERPROFILE%`/`%APPDATA%`/`%LOCALAPPDATA%`/`%USERNAME%` を選択ユーザに合わせて展開)
    - `apps/fabriq_backuper/lib/ui/userdata_edit_dialog.ps1` — Add/Edit
      モーダル (`Show-UserdataEditDialog`)。SourcePath / Folder Browse /
      File Browse / Recurse / ExcludePattern / OnConflict (skip/overwrite/rename) /
      IncludeAcl / Description / Enabled の全項目を扱う
  - **Backup View 拡張**: User Data セクションを編集 UI に格上げ
    - 行ごとの [On] checkbox は **ephemeral** (この実行回限りの enable/disable)
    - [Add] / [Edit] / [Delete] ボタンは押下時に CSV 即時保存 (.bak 取得)
    - [Save changes] ボタンで [On] toggle 状態を `userdata_list.csv` に書き戻し
    - **[Source user] dropdown** (logged-on user 既定) で
      `%USERPROFILE%` 等の展開先を選択可能に。admin 昇格による env var ずれを解消
    - 行 double-click でも Edit dialog 起動
  - **Restore View 拡張**: **[Target user] dropdown** を追加。
    restore 側でも path resolution を別ユーザに向けられるよう拡張
    (cross-user migration)
  - **section logic 拡張** (後方互換 - SectionParams optional key):
    - `userdata/backup.ps1`: `SourceUserProfilePath` を受け取り、
      `Resolve-EntryPath` 内で `Expand-PathWithUser` を優先使用。
      省略時は従来の `[Environment]::ExpandEnvironmentVariables` 動作。
      manifest.json に `sourceUser` ブロックを追加 (trace/debug 用)
    - `userdata/restore.ps1`: `TargetUserProfilePath` を受け取り、
      manifest の `sourcePath` 再展開時に `Expand-PathWithUser` を優先使用。
      省略時は従来通り `[Environment]::ExpandEnvironmentVariables`
  - **3-tier 永続化セマンティクス** (toggle / immediate / explicit save) は
    操作者の意図を反映: バックアップ実行直前の「今回だけこれを除外」と
    CSV ベースラインの編集を明確に分離

### Fixed
- apps/fabriq_backuper v0.6.0 → v0.6.1 (Phase 2.6.1 hotfix / Start button cut-off):
  Phase 2.6 で form height を 660 → 820 に拡張したが、Backup/Restore/Progress
  view の Start ボタン (Y=718, height 44) が form content area (実効 730px)
  を超えて**画面外にはみ出していた**。Form height を 820 → **900** に拡張して
  Start ボタン + 下部マージンを確保。

### Changed
- apps/fabriq_backuper v0.5.2 → v0.6.0 (Phase 2.6 / GUI polish: layout + UNC dialog):
  実機 UI 検証で指摘された 2 件を修正 (MINOR / 新規 dialog 追加 + layout 一新)。
  - **Form size**: 760×660 → **960×820** (printer / userdata grid を贅沢に
    取れるサイズ)。背景の理由: 旧サイズでは "Select All" ボタンが描画上
    重なって消えかけて見えていた + userdata プレビューがほぼ無視できる
    高さしかなかった
  - **layout 一新**: Backup View で printer grid を高さ 270 に拡張、
    userdata grid を 180 に拡張、Start ボタンを右下にきれいに配置。
    ボタン幅も統一 (Select All 96 / None 80 / Refresh 96)。Restore View
    も同形で再配置 (printer grid 440 高さ)
  - **新規: `lib/ui/unc_connect_dialog.ps1`** + `Show-UncConnectDialog`:
    Windows 標準 FolderBrowserDialog で UNC を直接入力できない問題に対応。
    Backup View / Restore View 双方に [Connect UNC...] ボタンを追加し、
    押すと専用 WinForms モーダル (UNC Path / Username / Password / Connect)
    で資格情報入力を一括処理。`New-PSDrive -Credential` で一時マウント、
    疎通確認後に destination root box に貼り付け (Backup 側) または
    後続の [Browse for backup...] が UNC 配下を見られる状態に (Restore 側)
  - lint: `Append-ProgressLog` → `Add-ProgressLog` (PowerShell approved verb)

### Fixed
- apps/fabriq_backuper v0.5.1 → v0.5.2 (Phase 2.5 / duplicate console 解消):
  Fabriq_BackUper.exe を double-click すると **conhost ウィンドウが 2 つ表示**
  されていた (launcher 起源の conhost #1 + self-spawn 起源の conhost #2) のを
  修正。self-spawn の `Start-Process powershell.exe` に **`-NoNewWindow`**
  を追加し、子プロセスが親の console window を共有するように変更
  (process isolation は維持、console window だけ共有)。
  fabriq main と同じ UX (conhost 1 つ + WinForms 1 つ) になる。

### Changed
- apps/fabriq_backuper v0.5.0 → v0.5.1 (default OnConflict 変更):
  `apps/fabriq_backuper/data/userdata_list.csv` の OnConflict 既定値を
  `skip` → `overwrite` に変更。**Migration workflow で機能不全だった既定挙動を修正**。
  - 旧 (skip): target に Desktop / Documents が存在する場合（ほぼ常に存在）
    エントリ全体を skip → 何も復元されない、という非実用的挙動だった
  - 新 (overwrite): robocopy `/E` で **merge** 動作 — target unique
    ファイルは保護、同名ファイルは backup で上書き、backup unique は追加。
    キッティング新 PC で Desktop に既存 shortcut (Edge / OneDrive 等) が
    あっても残しつつ旧 PC の Desktop ファイルを取り込む期待挙動
  - 注意: "overwrite" は merge セマンティクス（/MIR ではない）。
    target unique ファイルを削除したい完全 mirror が必要なら将来 `mirror`
    モード追加（要件次第、現状未実装）
  - operator は引き続き CSV の OnConflict 列で per-entry に
    `skip` / `overwrite` / `rename` を選択可能

### Changed
- apps/fabriq_backuper v0.4.0 → v0.5.0 (Phase 2.4 / UNC destination + Browse restore):
  バックアップ先と復元ソースを **任意の場所 (UNC 含む)** に指定できるよう拡張。
  manifest.json は既に portable descriptor なので、別途 "展開バッチ素材" 生成は
  不要 — 「Browse でフォルダを指定 → FabriqBackUper が manifest を読んで restore」
  だけで Migration workflow が成立する設計。
  - 新規: `lib/ui/unc_helper.ps1` (Test-UncPath / Connect-UncWithCredentials /
    Resolve-UncAccess)。UNC 未認証時に Get-Credential プロンプト
    （Windows 標準）→ New-PSDrive で一時マウント
  - 修正: `lib/ui/backup_view.ps1` 冒頭に "Destination root" テキストボックス +
    [Browse...] ボタン。既定値は `apps/fabriq_backuper/Backup`、UNC で
    `\\server\share\backups` 等に変更可。Start 時に Resolve-UncAccess で
    疎通＋認証を確保してから engine 呼び出し
  - 修正: `lib/ui/restore_view.ps1` に [Browse for backup...] ボタン追加。
    任意フォルダの aggregate manifest を読んで restore plan を構築
    (Hostlist 駆動経路と切り替え可能)。printer リストも内部 manifest 駆動
  - 修正: `lib/engine.ps1`:
    - `Invoke-BackuperBackupCore` に `-DestinationRoot` 追加
    - `Invoke-BackuperRestoreCore` に `-ExplicitAggregateDir` 追加 +
      `-PickedTimestamp` を optional 化
  - 修正: `lib/sections/userdata/restore.ps1` で **sourcePath 再展開**
    （cross-user 対応）: env var を含むパスは target process context で
    再 expand。`%USERPROFILE%` 等が backup-time の suzuki でなく
    restore-time の tanaka に解決される。env var なしの絶対パスは
    resolvedPath にフォールバック
  - form size: 720×540 → **760×660**（destination root + Browse 等のフィールド
    追加分のスペース確保）
  - kernel 公開 API 影響なし、Min Kernel API = 2.0.0
  - 既存 modules への波及: ゼロ
  - 既知: future 候補として「展開先ユーザの選択 GUI」(admin 昇格時の
    env var resolution 問題対応) はメモリに記録、Phase 2.7 以降で対応予定

### Changed
- apps/fabriq_backuper v0.3.1 → v0.4.0 (Phase 2.3 / userdata section 内製化):
  printer に続いて **userdata セクションも内製化**。出力先が
  `Backup/<OldPCname>/<ts>/sections/userdata/` に統合され、printer と
  userdata の保存先不一致を解消。modules/extended/userdata_backup は
  引き続き touched せず並走運用継続。
  - `lib/sections/userdata/backup.ps1`: 既存 module 全機能を内製版に移植
    (robocopy /B + ACL + ExcludePattern + per-entry OnConflict)。
    CSV ソースを **FabriqBackUper-owned** `apps/fabriq_backuper/data/userdata_list.csv`
    に切替（modules 側 CSV はもう参照しない）。`$SectionParams.IncludeEntries`
    で per-entry filter 可能（Phase 2.2.2 GUI editor が活用予定）
  - `lib/sections/userdata/restore.ps1`: manifest 駆動の内製 restore。
    per-entry OnConflict (skip/overwrite/rename) を尊重。
    `$SectionParams.IncludeEntries` でフィルタ
  - 新規: `apps/fabriq_backuper/data/userdata_list.csv`
    (Documents / Desktop / Outlook signatures 等のサンプル付き default)
  - 修正: `lib/ui/backup_view.ps1` userdata プレビューの CSV 参照先を
    `apps/fabriq_backuper/data/userdata_list.csv` に切替
  - **既存 modules への波及**: ゼロ (modules/extended/userdata_backup 並走)
  - 移行注意: operator が modules 側 CSV で行っていたカスタマイズは
    自動移行されない。新規 deployment では FabriqBackUper-owned CSV を
    手動で編集する必要あり（Phase 2.2.2 で GUI 編集機能を提供予定）

### Fixed
- apps/fabriq_backuper v0.3.0 → v0.3.1 (Phase 2.2.1a hotfix):
  Phase 2.2.1 で engine が全 section に `-SectionParams` を渡すように変更したが、
  userdata セクション wrapper の param 宣言を忘れていたため、printer 完了後
  に userdata 実行で "パラメーター名 'SectionParams' に一致するパラメーターが
  見つかりません" でクラッシュしていた。userdata wrapper（backup/restore 両方）
  に `[hashtable]$SectionParams = @{}` を追加して accept できるよう修正
  （wrapper 自身は SectionParams を無視、Phase 2.3 内製化時に活用予定）。
  同時に lint 警告 `$null` 比較順序（`$x -eq $null` → `$null -eq $x`）も修正。

### Changed
- apps/fabriq_backuper v0.2.1 → v0.3.0 (Phase 2.2.1 / printer section 内製化):
  printer の backup/restore セクションを内製化し、GUI 上で **per-printer 選別**
  (checkbox) を可能化。modules/extended/printer_backup は touched せず並走継続。
  - `lib/sections/printer/backup.ps1`: 既存 module 全機能を内製版に移植
    (Get-Printer 列挙 / DEVMODE / hwconfig / payload export 全部入り)。
    `$SectionParams.IncludePrinters` でフィルタ。出力は
    `Backup/<OldPCname>/<ts>/sections/printer/` に統合
  - `lib/sections/printer/restore.ps1`: manifest 駆動の内製 restore。
    Spooler restart を挟んだ 2 pass 構造 / Color/duplex の explicit Set /
    HKLM hwconfig + HKCU DEVMODE (Resolve-HkcuRoot 経由) すべて踏襲。
    `$SectionParams.IncludePrinters` でフィルタ
  - `lib/engine.ps1`: `Invoke-SectionScript` / `Invoke-BackuperBackupCore` /
    `Invoke-BackuperRestoreCore` に **`SectionParams` / `SectionParamsBySection`**
    パラメータ追加。section ごとに任意の追加 param を hashtable で渡せる
  - `lib/ui/backup_view.ps1`: printer DataGridView (checkbox + Select All /
    None / Refresh)。**仮想プリンタ (Microsoft Print to PDF / OneNote /
    XPS / Fax / RDP redirect) は既定 uncheck**、それ以外は check。
    Start 時に checked 名のリストを `SectionParams.printer.IncludePrinters`
    として渡す
  - `lib/ui/restore_view.ps1`: timestamp 選択時に
    `sections/printer/manifest.json` を読んで printer DataGridView に
    populate (manifest 駆動)。RDP redirect は plan 段階で除外、仮想は
    既定 uncheck。同じ idiom で per-printer 選別
  - `lib/manifest_aggregator.ps1`: aggregate manifest に `manifestPath`
    フィールド追加 (internal section 用、`sections/<name>/manifest.json`
    relative path)。`externalOutputDir` も残置で後方互換
  - **userdata section は引き続き wrapper のまま** (Phase 2.3 で内製化予定)
  - kernel 公開 API 影響なし、Min Kernel API = 2.0.0
  - 既存 modules への波及: **ゼロ** (modules/extended/printer_backup 並走)

### Fixed
- apps/fabriq_backuper v0.2.0 → v0.2.1 (Phase 2.1c hotfix):
  Backup 実行時に "引数が空のコレクションであるため、パラメーター 'Warnings'
  にバインドできません" エラーで manifest 書き込みが失敗する不具合を修正。
  原因: `New-AggregateManifest` の `Warnings` パラメータが
  `[Parameter(Mandatory = $true)][array]` 宣言で空配列 `@()` の binding を
  拒否していた (PowerShell の Mandatory+array+empty 仕様)。
  Fix: `Mandatory = $true` を外し、デフォルト値 `@()` を設定する形に変更。
  console 経路でも同じ潜在 bug があったが今まで未実行で顕在化していなかった。

### Changed
- apps/fabriq_backuper v0.1.0 → v0.2.0 (Phase 2.1 / GUI 化):
  console UI を WinForms GUI に置き換え (tonebender 風の Mode Select →
  Backup/Restore View → Progress View の 4 view 構造)。fabriq_operator の
  CentreCOM-inspired light theme を satellite-app 独立原則に従い自前移植
  (`lib/ui/theme.ps1`)、`lib/ui/main_form.ps1` が view 切替と shared
  state を管理。
  - `lib/ui/theme.ps1`: 色定数 / フォント / スタイルヘルパー
    (New-StyledButton, Set-GridStyle, New-StyledLabel 等) を自前定義
    (fabriq_operator/lib/theme.ps1 からの直接コピー + 名前空間維持)
  - `lib/ui/main_form.ps1`: 720x540 Form、ヘッダ (ホスト指示) + 切替可能
    Content Panel、`Switch-View` / `Update-HostHeader` / `Start-FabriqBackuperGui`
  - `lib/ui/mode_select_view.ps1`: hostlist combo + Backup/Restore/Quit ボタン。
    `$env:SELECTED_OLD_PCNAME` 既セット時は自動継承
  - `lib/ui/backup_view.ps1`: section チェックボックス + userdata エントリ
    プレビュー DataGridView (Phase 2.1 は read-only) + Start。Phase 2.2 で
    ephemeral checkbox + Add Folder/File 編集機能を追加予定
  - `lib/ui/restore_view.ps1`: timestamp dropdown + manifest 要約表示 +
    section チェックボックス + Start
  - `lib/ui/progress_view.ps1`: monospace log + Done ボタン (Phase 2.1 は
    synchronous 実行のため完了後表示、Phase 2.2 で Runspace 化予定)
  - `lib/engine.ps1`: UI-agnostic core 関数 `Invoke-BackuperBackupCore` /
    `Invoke-BackuperRestoreCore` を追加 (param 受取式、prompt なし)。
    Legacy console orchestrator `Invoke-BackuperEngine` は fallback として
    保持
  - `fabriq_backuper.ps1`: console menu loop を `Start-FabriqBackuperGui`
    呼び出しに切替
  - **既知の Phase 2.1 限界** (Phase 2.2/2.3 で解決予定):
    1. 過去時点 restore は wrapped module 側で「最新」に依存 (Phase 2.3
       sections 内製化で根本解決)
    2. UI が backup/restore 実行中に freeze (synchronous 実行、Phase 2.2
       Runspace 化で解決)
    3. CSV エントリの GUI 編集は未対応 (Phase 2.2 で追加)
  - kernel 公開 API 影響なし、Min Kernel API = 2.0.0

### Added
- apps/fabriq_backuper (新規 satellite app v0.1.0 / Phase 1 PoC):
  Backup/Restore 専業 satellite として `apps/fabriq_backuper/` を新設。
  FabriqIOS と同型の self-spawning subprocess + kernel dot-source パターンで
  独立 entry を持つ。Operator-facing は `Fabriq_BackUper.exe`（fabriq root、
  56 KB の C# launcher）一系統に絞り、FabriqApps ダイアログからは除外。
  - **Entry**: `apps/fabriq_backuper/fabriq_backuper.ps1`（self-spawning,
    kernel/common.ps1 dot-source, defensive log suppression）
  - **Engine**: `lib/engine.ps1` が section 群を orchestrate。
    hostlist 読み取り（`Import-ModuleCsv` 経由で ENC: 自動復号）→
    host selector → section selector → 各 section 実行 → 集約 manifest
  - **Hostlist 結合**: `kernel/csv/hostlist.csv` を read-only 消費、
    `Test-MasterPassphrase` でパスフレーズ検証、`$global:FabriqMasterPassphrase`
    セット後 `Import-ModuleCsv` が自動復号。`$env:SELECTED_OLD_PCNAME` が
    既セットなら継承（fabriq 親 session から）、なければ console picker
  - **Section 契約** (Phase 1 plug-in interface):
    `lib/sections/<name>/{backup,restore}.ps1` が PSCustomObject (Status /
    ElapsedMs / Summary / Warnings / ExternalOutputDir / ExternalManifestPath)
    を return。Phase 1 では既存 `modules/extended/printer_backup` /
    `modules/extended/userdata_backup` を **薄く wrap**（modules touched なし）
  - **Aggregate manifest** (`fabriq-backuper-snapshot` schemaVersion=1):
    `Backup/<OldPCname>/<yyyy_MM_dd_HHmmss>/manifest.json` に各 section 個別
    manifest への pointer + summary を集約。printer_backup / userdata_backup の
    parallel manifest 構造を上位に統合する形
  - **Console UI** (Phase 1): main menu / section selector / timestamp picker /
    confirm prompt。Phase 3 で WinForms 化予定
  - **既存 modules への波及**: ゼロ（並走運用。Phase 4 で削除予定）
  - kernel 公開 API §1〜§5 のみ依存、Min Kernel API = 2.0.0
- dev/launcher/Launcher_BackUper.cs + app_backuper.manifest + build_backuper.ps1:
  Fabriq_BackUper.exe を csc.exe で生成する launcher 一式
  （Launcher_IOS と同型、UAC requireAdministrator、anycpu .NET Framework 4.x）
- apps/fabriq_operator/lib/apps_dialog.ps1: exclusion list に
  `fabriq_backuper` を追加（FabriqIOS と同様、独自 EXE 起動経路に揃える）

### Added
- modules/extended/userdata_backup (新規モジュール v0.1.0):
  任意のファイル / ディレクトリの portable backup + restore を担う
  試作モジュール。printer_backup と同じ backup フォルダ規約
  （`backup/<OldPCname>/<yyyy_MM_dd_HHmmss>/`）・hostlist 駆動・
  manifest 中心設計を踏襲し、将来の **backup モジュール統合**を見据えて
  manifest を parallel な構造で設計（manifestType=`fabriq-userdata-backup`
  schemaVersion=1）。
  - **Backup (`userdata_backup.ps1`)**: `userdata_backup_list.csv` の各
    enabled 行を 1 entry として robocopy で `entries/<NN>/data/` に複製。
    Per-entry 設定: `SourcePath` (環境変数展開対応) / `Recurse` (再帰 ON/OFF) /
    `ExcludePattern` (`;` 区切り、末尾 `/` で `/XD`、それ以外は `/XF`) /
    `OnConflict` (restore 時挙動) / `IncludeAcl` (`/COPYALL` opt-in)。
    robocopy `/B` backup mode でロックファイルにも対応。Post-Apply
    Verification は manifest 妥当性 + 必須ファイル存在 + 各 entry の
    data/ フォルダ存在を検証。
  - **Restore (`userdata_restore.ps1`)**: manifest を契約として読み込み、
    robocopy で逆方向に復元。ソース解決は **hostlist 駆動**
    (`$env:SELECTED_OLD_PCNAME` → backup/<その名前>/ → 最新 timestamp、
    `userdata_backup_config.csv` の `SourcePcName` / `BackupTimestamp` で
    override 可)。Per-entry `OnConflict`: skip (既存温存) / overwrite /
    rename (既存を `.bak_yyyyMMdd_HHmmss` に退避)。printer_restore と
    異なり osArch / osVersion は hard fail せず情報記録のみ (ファイル系は
    arch 独立)。
  - **将来統合計画への配慮**: manifest 構造を printer_backup と意図的に
    parallel に保ち、`manifestType` と `items` 内のキーだけがモジュール
    固有。3rd backup モジュール出現時 or 明示的統合タスク時に
    `kernel/common.ps1` への共通化を検討する想定 (現時点は premature
    abstraction を避けて scaffold 同形コピーで進める)。
  - kernel 公開 API §1〜§5 への影響なし、Min Kernel API = 2.0.0。

### Changed
- modules/extended/printer_backup v0.4.2 → v0.5.0:
  restore の source backup 解決を **hostlist 駆動**化 (MINOR / 後方互換)。
  - Why: backup 側が hostlist OldPCname でフォルダを切る運用に揃ったため、
    restore も「選択中ホストの OldPCname → backup/<その名前>/」で
    自動マッチさせるのが kitting workflow と整合的
  - 新解決順位:
    1. CSV `SourcePcName` (明示 override、テスト/特殊ケース向け)
    2. `$env:SELECTED_OLD_PCNAME` (hostlist 選択中のホスト)
    3. どちらも空 → **エラーで abort**（"select a host or set CSV"）
  - 旧 v0.4.x の「scan all PCs and pick newest by collectedAt」fallback は
    廃止。どの PC の backup を復元したか不明な事故を防ぐ
  - Plan display と Show-Info で resolution source を明示
    （`PcName=hostlist OldPCname, Timestamp=latest manifest.collectedAt` 等）
  - 公開 API への追加依存: `$env:SELECTED_OLD_PCNAME` (KERNEL_API.md §3.1
    既存、Min Kernel API = 2.0.0 のまま)
  - 既存 CSV-explicit ワークフローは override 経路として完全互換

- modules/extended/printer_backup v0.4.1 → v0.4.2:
  backup 出力フォルダの `<PCName>` 解決順序を `$env:COMPUTERNAME` から
  **`$env:SELECTED_OLD_PCNAME`（hostlist の OldPCname）優先、未設定時は
  `$env:COMPUTERNAME` フォールバック** に変更。
  - Why: kitting workflow では OldPCname が顧客視点のホスト identity
    で、バックアップフォルダもこちらで識別する方が運用上自然
    （Windows 上の現行 COMPUTERNAME はリネーム前後で揺れる可能性あり）
  - Backup Plan 表示と `_restore_notes.txt` にも両方の名前
    （PCName as folder / actual COMPUTERNAME）を表示し追跡可能化
  - manifest schema 変更なし（`computerName` フィールドが OldPCname
    ベースに切り替わるのみ。schemaVersion=1 内の解釈変更）
  - 既存 backup（旧 COMPUTERNAME ベース）の restore は不変

### Fixed
- modules/extended/printer_backup v0.4.0 → v0.4.1:
  Color / Trays が反映されない問題に対する restore 書き込み順序の最適化
  (PATCH / 内部実装変更、公開 API 影響なし)。
  - **症状**: Hardware config restore + Spooler restart 後に Canon
    iR-ADV のような driver で `DeviceOption01` REG_BINARY が部分リセット、
    DEVMODE 復元値も driver 再初期化で上書きされる現象を実機テストで観測
  - **原因の仮説**: pass 1 で DevModePerUser + PrinterDriverData を書いて
    から Spooler restart すると、restart 時の driver 再初期化が我々の
    binary 値を validation → 部分的に再生成してしまう
  - **対応**: Phase 4 を 2 pass に分割。
    1. pass 1: Set-PrintConfiguration (PrintTicket + explicit fields) +
       Set-PrinterProperty のみ
    2. Spooler 再起動 (HW config or DEVMODE が存在する場合のみ、+ 2 秒
       stabilize)
    3. pass 2: PrinterDriverData (HKLM) + DevModePerUser (HKCU) を**最後**
       に書く
  - これにより我々の binary 値が driver 再初期化後の "最後の書き手"
    となり、上書きされにくくなる。なお Canon driver 私的 DEVMODE 拡張は
    driver version 依存のため、本変更でも一部 vendor 固有設定（色モード
    詳細・フィニッシャ等）の cross-version 復元は本質的に困難（Microsoft
    PrintBRM も同じ制約を抱える領域）

### Added
- modules/extended/printer_backup v0.3.1 → v0.4.0:
  実機テストで報告された 2 つの残課題に対応 (MINOR / 後方互換)。
  - **Color 設定が driver default に戻る問題**:
    PrintTicketXML round-trip だけでは一部ドライバ（特に企業向け
    カラー複合機）が `DM_COLOR` を無視するため、PrintTicket 適用後に
    `Set-PrintConfiguration -Color/-Collate/-DuplexingMode/-PaperSize/
    -PaperSource/-PrintQuality` を**個別パラメータ**で明示再適用する
    fallback パスを追加。各 field は個別 try/catch で driver 非対応時
    も他 field の適用を阻害しない
  - **Installable options（トレイ・フィニッシャ）の欠落問題**:
    Add-Printer 直後はドライバ最小ハードウェア構成（カセット 1 のみ等）に
    リセットされる。これは per-user DEVMODE ではなく **HKLM 側の
    `Printers\<name>\PrinterDriverData`** に保存される vendor 固有
    REG_BINARY blob 群が原因。バックアップ時に当該サブキー全 value を
    type 付き JSON でダンプ、復元時に書き戻し → `Restart-Service Spooler`
    で適用する経路を追加。risk としてドライババージョン依存のため
    `RestoreHardwareConfig` 列で opt-in 制御（既定 1）
  - manifest.items.printers[] に `hwConfigFile` フィールド追加
    (additive / schemaVersion=1 内の後方互換変更)
  - CSV 列 `RestoreHardwareConfig` 追加（既定 1）

### Fixed
- modules/extended/printer_backup v0.3.0 → v0.3.1:
  per-user DEVMODE の HKCU アクセスを `Resolve-HkcuRoot` 経由に修正
  (PATCH / 内部実装変更、公開 API 影響なし)。
  - 元の実装は `HKCU:\Printers\DevModePerUser` を直接参照していたが、
    UAC 昇格時（別管理者アカウント）や SYSTEM コンテキスト実行時に
    `HKCU:` が **ログオン中ユーザではなく昇格元アカウント / SYSTEM** を
    指してしまい、backup は誤ったユーザの blob を採取、restore も
    誤ったユーザに書き戻す問題があった
  - reg_hkcu_config / desktop_icon_backup と同じ idiom（共通関数
    `Resolve-HkcuRoot` で HKU\<LoggedOnUserSid> に切り替え）を採用。
    Redirect 発生時は実行時に対象 SID を Show-Info で通知

### Added
- modules/extended/printer_backup v0.3.0:
  per-user DEVMODE のバックアップ/復元 + virtual printer 自動スキップ
  対応 (MINOR / 後方互換)。
  - **per-user DEVMODE 対応**: `HKCU:\Printers\DevModePerUser\<name>` の
    REG_BINARY blob を base64 で `printsettings/<safe>.devmode.b64` に
    保存し、復元時にそのまま書き戻す。これまで `Set-PrintConfiguration
    -PrintTicketXml` のみ実装していたが、これは spooler 側の既定値
    更新であり、アプリの**印刷ダイアログが参照する per-user blob は
    更新されない**。実機テストで「印刷設定が引き継がれない」現象が
    確認されたためこちらを正規ルートに追加。
  - **virtual printer 自動スキップ**: `SkipVirtualPrinters` 列 (既定 1)
    新設。Microsoft Print To PDF / XPS Document Writer / OneNote /
    Fax / OpenXPS など inbox 級の仮想プリンタを driver/port 名パターンで
    検出し plan 段階で除外。Win11 でも標準で同等品が存在するため
    cross-PC restore で再作成する価値が低いケースに対応。
  - manifest.items.printers[] に `devModeFile` フィールド追加
    (additive、schemaVersion=1 内の後方互換変更)。古い manifest を
    新 restore で読んでも `devModeFile=null` で従来どおり動作。
- modules/extended/printer_backup (新規統合モジュール v0.2.0):
  このPCのプリンタ環境（プリンタ、ポート、ドライバ、印刷設定）の
  **バックアップ + 復元**を担う統合モジュール。1 モジュール内に
  backup / restore の 2 スクリプトが同居し、`printer_backup_config.csv`
  を共有制御契約として読む構成（`desktop_icon_config` の backup/restore
  パターンを踏襲）。
  - **Backup (`printer_backup.ps1`)**: `backup/<ComputerName>/<yyyy_MM_dd_HHmmss>/`
    に manifest.json (schemaVersion=1, `fabriq-printer-backup`) 中心の
    portable バックアップを生成。ドライバ payload は
    `pnputil /export-driver` で oemNN.inf 単位に export、DriverName →
    oemNN.inf 逆引きマップを manifest に内包。Post-Apply Verification
    はファイル存在・サイズ・JSON 妥当性・件数一致で実施。
  - **Restore (`printer_restore.ps1`)**: 同一モジュール内の `backup/`
    配下から manifest を読み込み、drivers (pnputil + Add-PrinterDriver) /
    ports (type-aware Add-PrinterPort) / printers (Add-Printer +
    Set-Printer) / print settings (Set-PrintConfiguration -PrintTicketXml)
    / default printer の順で復元。osArch 不一致は hard fail、osVersion
    不一致は `StrictOsVersion=0` 既定で warning 降格（Win10→Win11 amd64
    移行を許容）。RDP redirect printer（`Remote Desktop Easy Print` /
    `TS\d+` ポート）と Local/WSD/Bonjour ポートは自動 skip。
    `ReuseInboxDrivers=1` 既定で Microsoft 供給ドライバの payload
    再注入を回避し Windows 内蔵を再利用。`OnConflict=skip|replace` で
    既存プリンタ衝突挙動を制御。Post-Apply Verification は
    Get-Printer の存在/DriverName/PortName 一致を全件検証。
  - 共有 CSV 列: `Enabled` / `IncludeDriverBinaries` /
    `IncludePrintSettings` / `SourcePcName` / `BackupTimestamp` /
    `StrictOsVersion` / `ReuseInboxDrivers` / `OnConflict` /
    `RestoreDefaultPrinter` / `Description`（10 列、各スクリプトは
    自分が利用する列のみ参照、他列は無視）。
  - kernel 公開 API §1〜§5 への影響なし、Min Kernel API = 2.0.0。

### Changed
- apps/fabriq_operator/lib/execution_toolbar.ps1 + kernel/main.ps1:
  Execution Toolbar に **ホスト情報ライブ照合パネル** を追加し、見た目を
  旧 Status Monitor (3.4.0 で retire) に忠実再現 (予想 PATCH /
  3.4.0 → 3.4.1)。kernel 公開 API §1〜§5 への影響なし。
  - ビジュアル: ダークテーマ (#1E1E1E) + cyan accent (#00C8C8) +
    Consolas monospace、`Fabriq - Status Monitor` native title bar
    (FormBorderStyle = FixedSingle、600×700、右上配置)、TableLayoutPanel
    2 段 (`PC Info Comparison` 60% / `Surkitinisme` 40%)、StatusStrip
    最下部に `Gyotaq` (cyan) / `Skip` (orange) + 状態 + email。
  - PC Info Comparison: kernel が push する target ホスト情報 +
    toolbar runspace 内の `Get-NetIPAddress` / `Get-NetRoute` /
    `Get-DnsClientServerAddress` / `Get-Printer` を tiered polling
    (Hostname 1s / Network・DNS 3s / Printer 10s) し、ID / PC Name /
    [Ethernet] / [Wi-Fi] / [DNS] / [Printers] セクションで current vs
    target を比較表示 (`[OK]` 緑 / `[!!]` 赤 / `[--]` 赤)。ホストリストに
    設定値がないフィールドは行ごと省略 (driven by SELECTED_* env vars)。
  - Surkitinisme art panel: 旧 Status Monitor の typing + glitch +
    flash effect を完全移植。`kernel/json/art_pulse.txt` /
    `kernel/json/status.json` / `kernel/txt/art_sentences.txt` /
    `kernel/txt/silence.flag` の polling は同じ disk file 契約で継続
    (kernel 側の Write-ArtPulse 等は不変)。render 周期は 100ms (10fps、
    旧 40ms から保守的に緩和。in-process なので UI スレッド共有、
    typing/cursor blink は十分滑らか)。
  - `Update-ExecutionToolbar` (§6 内部実装) に `-TargetHostInfo
    <hashtable>` パラメータ追加。既存呼び出しは引数省略で従来動作維持。
  - `Set-SelectedHostEnvironment` (kernel/main.ps1) で
    `Update-ExecutionToolbar -TargetHostInfo (Get-FabriqHostInfoFromEnv)`
    を呼び、`SELECTED_*` env-var を構造化 hashtable として push。
    toolbar 未起動時は no-op (Get-Command guard)。
  - `Show-ExecutionToolbar` 末尾でも同じ self-push を実施。kernel が
    `Set-SelectedHostEnvironment` (fresh-start) や Restore-HostEnvironment
    (resume) で SELECTED_* env vars を populate **した後に** toolbar を
    起動する経路 (kernel/main.ps1 line 1448 / 1305 → 1535) を救済する
    ため必須 (これ無しでは toolbar が "Waiting for host selection..."
    から動かない)。
  - `Get-FabriqHostInfoFromEnv` (新、§6 内部): `SELECTED_*` env vars
    から TargetHostInfo hashtable を構築する DRY ヘルパー。Set-Selected
    HostEnvironment と Show-ExecutionToolbar の両方から呼ばれる。
  - PC Info 描画は change-detect cache (`$script:lastPcInfoText`) で
    内容変化なし tick の RichTextBox 書き換えをスキップ。`Set-Colorized
    Text` は SuspendLayout/ResumeLayout/Refresh の三点セットで bulk
    edit を 1 paint にまとめる。


## [3.4.0] - 2026-05-13

### Changed
- kernel/main.ps1 + kernel/common.ps1 + apps/fabriq_operator/lib/execution_toolbar.ps1:
  Status Monitor 廃止 + Execution Toolbar 新設 (MINOR / 3.3.1 → 3.4.0)。
  詳細は kernel/KERNEL_API.md §8 3.4.0 セクション参照。実装は 4 サブステージ
  (4a〜4d) で段階的に投入し、各ステージで Pester 157/157 pass + 実機 smoke
  で検証。
  - Stage 4a: kernel/common.ps1 `Invoke-SafeCommandAsync` の poll loop に
    `[Application]::DoEvents()` を 16ms 間隔で注入 (sub-sleep ループで
    pollMs を内側分割、Skip 検出応答性も改善)。try/catch でヘッドレス
    環境耐性確保。
  - Stage 4b: apps/fabriq_operator/fabriq_operator.ps1 の lib dot-source
    リストに execution_toolbar.ps1 を追加。
  - Stage 4c: apps/fabriq_operator/lib/execution_toolbar.ps1 新設
    (dedicated STA Runspace ベース、WinForms TopMost ツールバー、
    Skip / Gyotaq ボタン、ヘッダドラッグ可)。
    kernel/main.ps1 への lifecycle wiring (Show / Update Running /
    Update Idle / Hide before __RESTART__ x 4 / Hide before Refabriq)。
    kernel/common.ps1 :: `Exit-Fabriq` に Hide-ExecutionToolbar 追加
    (Get-Command guard 付き、headless 耐性)。
  - Stage 4d: kernel/main.ps1 の `Start-StatusMonitor` 呼び出しを
    comment-out。`status_monitor.ps1` ファイル本体と
    `Start-StatusMonitor` / `Stop-StatusMonitor` 関数は 3.4.x 系で
    残置 (緊急 rollback 用、3.5.0 で削除予定)。
- kernel/KERNEL_API.md §6: `Start-StatusMonitor` / `Stop-StatusMonitor` に
  Deprecated 注記、`Show/Hide/Update-ExecutionToolbar` / `Save-Screenshot`
  を内部実装として追記。
- kernel/KERNEL_API.md §8: 3.4.0 release notes セクション追加。
  動機: recent な Windows Update / Defender / ASR ヒューリスティックが
  hidden child PowerShell process spawn をブロックするケースが増え、
  Status Monitor の Skip / Gyotaq surface が一部端末で失われていた。
  in-process Toolbar への移管で Defender / ASR の影響を受けない構成へ。

### Added
- profiles/easy_template: `easyprofile.csv` の列定義に `Segment` を追加し、
  `easyprofile.ps1` で行ごとの Segment 値を実行直前に
  `$env:FABRIQ_SEGMENT` へ export、`finally` ブロックで clear するように
  変更（`kernel/main.ps1:441-465` の Linear path と同じ env-var contract）。
  これにより EasyProfile からも Segment-aware モジュール
  （`reg_hklm_config` / `app_config` / `acl_config` / `test_harness_config`
  等）を `_list.csv` の Segment フィルタ付きで呼び分け可能。
  実装メモ: profile 級 CSV を `Import-ModuleCsv` で読むと、CSV 自身が
  `Segment` 列の strict-match で絞られてしまう（`kernel/common.ps1:1063-
  1080` の挙動。`_list.csv` 用途には正しいが Profile では行ごとに違う
  Segment を意図するため不整合）。これを避けるため CSV ロードを
  `Import-CsvSafe` + `Test-CsvColumns` + 手動 Enabled フィルタの組み合わせ
  に切り替え（Linear path `Resolve-ProfileModules` が `Import-Csv` を
  直接呼ぶのと同じ路線）。Segment 列を持たない既存 CSV は後方互換で
  そのまま動作（`$entry.PSObject.Properties.Name -contains 'Segment'`
  ガード）。事前表示および実行バナーには Linear/Flex 慣習に合わせて
  `[seg:<value>]` ラベルを併記。公開 API 不変、kernel `KERNEL_VERSION`
  据え置き 3.3.1。検証: 既定 CSV（全 Enabled=0）の早期 exit / Segment=値
  と Segment=空の 2 行同時実行による env-var set/clear 動作を
  `test_harness_config` 経由で smoke 確認済み（Segment filter が指定行
  にのみ適用され、次の Segment 空行には漏れないことを確認）。

## [3.3.1] - 2026-05-12

### Fixed
- kernel/common.ps1 :: `Invoke-SafeCommandAsync`: child runspace の
  inject hashtable に `$global:AutoConfirmMode` を追加（PATCH /
  3.3.0 → 3.3.1）。
  **症状**: 3.3.0 で DefaultAsync 既定 ON 化以降、FlexProfile の `[Run This]`
  ボタン（RunSingle 経路）でモジュールを実行すると、モジュール内
  `Confirm-ModuleExecution` の Y/N プロンプト（例: hkcu_config の
  「Apply the above registry changes?」、network_profile_config の
  「Apply the above network category changes?」など）が auto-confirm
  されず Read-Host へ落ちる regression。
  **原因**: kernel 3.1.0 で導入された公開グローバル `$global:AutoConfirmMode`
  （KERNEL_API.md §2 で宣言済）が `Invoke-SafeCommandAsync` の inject
  リストに含まれておらず、child runspace では common.ps1 init 既定値
  `$false` のまま。3.2.5 以前は RunSingle が sync 経路だったため
  同一スコープで読めて無症状だったが、3.3.0 で全モジュールが async
  経路に切り替わったことで dormant bug が発現。
  **影響範囲**: `Confirm-ModuleExecution` を呼ぶ全モジュール（108 件
  ヒット）の FlexProfile `[Run This]` 経由実行。`[Run Batch]` /
  `[Run: Group]` / Linear `[Execute Profile]` は `AutoPilotMode` が
  inject 済みのため無影響。
  **修正**: inject hashtable に `AutoConfirmMode = $global:AutoConfirmMode`
  を 1 行追加。公開 API 仕様（KERNEL_API.md §2）の遵守。

## [3.3.0] - 2026-05-12

### Changed
- kernel/common.ps1 + kernel/json/async_config.json: `__ASYNC__` マーカー
  挙動の後方互換な拡張（MINOR / 3.2.5 → 3.3.0）。
  `async_config.json` に新フィールド `DefaultAsync` を追加し、shipped
  default を `true` に設定。profile 内のマーカー有無に関わらず全モジュール
  が監視付き Runspace（`Invoke-SafeCommandAsync`）経路で実行されるように
  なる。`__ASYNC__` マーカー自体は idempotent ON-only no-op として
  後方互換保持。kill switch `Enabled=false` は引き続き優先で全モジュール
  を同期経路に降格させる。`DefaultAsync` フィールド欠損時のフォールバック
  は `false`（旧 config 互換）。
  動機: マーカーを profile に書き忘れた場合でも Status Monitor の Skip
  ボタン / timeout 強制中断という安全網が常時利用可能になる。疲弊した
  キッティング作業者の認知負荷軽減を目的とした defensive default。
  `kernel/KERNEL_API.md` §4.2 / §8 を同コミット内で更新済み。
- kernel/common.ps1 :: `Get-FabriqAsyncConfig`: 返却 PSCustomObject に
  `DefaultAsync` プロパティ追加。fallback default は `$false`（config
  欠損／パース失敗時に silent な async 化を防ぐ安全側挙動）。
- kernel/common.ps1 :: `Resolve-ProfileModules`: `$asyncMode` 初期値を
  `$false` ハードコードから `Get-FabriqAsyncConfig` の `DefaultAsync`
  AND `Enabled` の組み合わせから算出するように変更。
- tests/kernel/Resolve-ProfileModules.tests.ps1: `DefaultAsync` 既定 ON
  挙動と kill switch 優先性を検証する 3 ケース追加。

### Security
- modules/standard/robocopy_config v1.0.0 → v1.0.1 (PATCH):
  UNC 認証パスワード（AuthPass）を `net use` のコマンドライン引数経由
  から **stdin パイプ経由** に変更（A3 / T1-B robocopy）。閉じられる
  leak vector:
    - **`Win32_Process.CommandLine`**: net.exe lifetime 中の live
      snapshot。同 logon session の他プロセス（AV / EDR / sysmon /
      `Get-CimInstance Win32_Process` 経由）から password 文字列を
      取得可能だった。本変更で net.exe のコマンドライン引数から
      AuthPass が消えるため遮断
    - **Windows Security Log Event ID 4688**: site で
      `Audit Process Creation` + `IncludeCommandLine` が有効な場合
      （B2G の標準 audit 設定）、永続的にイベントログ記録 → SIEM 転送
      先まで到達する経路。本変更で 4688 にも password が記録されない
    - **PowerShell Module Logging**: site で有効化された場合、native
      command の引数も間接的に記録される実装。本変更で対象外
  実装詳細: L162 の `& net use $share "$pass" /user:U` を
  `"$pass" | & net use $share /user:U` に変更。net.exe は `/user:` 指定 +
  password 引数なしで prompt mode に入り、PowerShell の native command
  pipe からの 1 行を password として読む（標準 net.exe 挙動、Windows
  10/11 で安定）。SMB session 確立後 robocopy は OS の SMB credential
  cache 経由で UNC パスへ透過アクセス。cleanup (`& net use $share /delete
  /y`) は password 不要のため変更なし。
  検討済却下案:
    - PSDrive `-Persist` 化: drive letter mapping を新規導入する副作用
      が現状仕様（drive letter 不使用）と乖離するため見送り
    - `WNetAddConnection2W` P/Invoke: 最も architectural だが Add-Type
      C# block の複雑性大、stdin pipe で同等のセキュリティ効果が得られる
      ため不採用
  Guide.txt の UNC Authentication 節と L107「パスワードはコンソール/
  ログに一切出力されません」記載を、実装と整合する正確な表現に書き換え
  （PS Transcript / Win32_Process.CommandLine / 4688 / Module Logging
  の各経路について明示）。stdin pipe 化以前は L107 記載は
  Win32_Process.CommandLine 経由で実際には漏洩していた虚偽記載で、
  本変更で記載が事実と整合。
  公開 API 不変、`REQUIRES_KERNEL` 据え置き 2.0.0、CSV スキーマ /
  ModuleResult 契約 / 動作フロー operator 視点いずれも不変
- modules/standard/windows_license_config v1.0.0 → v1.0.1 (PATCH):
  Windows プロダクトキーの transcript 漏洩を抑止（A3 / T1-A、3 箇所）。
  共通 redact-map 経路（`Invoke-TelemetryRedact`）はキー値そのものを
  redact map に持たないため適用不能で、表示時マスクが正解。
    - L59 CSV 読み込み後の `Write-Host "  Key: $productKey"` を
      `Get-MaskedKey` 経由に変更。末尾 5 文字のみ可視化（dash 維持）、
      他 24 文字を `*` に置換（`*****-*****-*****-*****-XXXXX`）
    - L90 形式不正時の `Show-Error "Invalid key format: $productKey"`
      から生キー値を除去、形式テンプレ（`XXXXX-XXXXX-XXXXX-XXXXX-XXXXX`）
      のみ表示
    - L118 確認ダイアログ前の `Write-Host "New Key: $productKey ..."`
      を `Get-MaskedKey` 経由に変更
    - ローカルヘルパ `Get-MaskedKey` を script 先頭に追加
      （autologon_config L87-91 のマスク手法と同精神。`common.ps1` の
      `Invoke-TelemetryRedact` は redact-map 型で要件不一致のため独自）
    - 公開 API 不変、`REQUIRES_KERNEL` 据え置き 2.0.0
- modules/standard/office_license_config v1.0.0 → v1.0.1 (PATCH):
  Office プロダクトキーの transcript 漏洩を抑止（A3 / T1-A、2 箇所）。
    - L118 事前一覧表示の `Write-Host "    Key:  $($item.ProductKey)"`
      を `Get-MaskedKey` 経由に変更（windows_license と同マスク仕様）
    - L177 形式不正時の `Show-Skip "Invalid key format: $($item.ProductKey)"`
      から生キー値を除去、形式テンプレのみ表示
    - ローカルヘルパ `Get-MaskedKey` を script 先頭に追加（windows_license
      と同実装）
    - 公開 API 不変、`REQUIRES_KERNEL` 据え置き 2.0.0
- modules/standard/office_license_config/Guide.txt: Security Note 節を新設。
  `office_license_install.ps1` の `cscript OSPP.vbs /inpkey:KEY` が
  Win32_Process.CommandLine / 4688 audit / Module Logging に ProductKey を
  露出させる構造的制約を明記。Microsoft プラットフォーム制約 (OSPP.vbs
  positional 引数のみ、C2R では `OfficeSoftwareProtectionService` WMI
  provider 削除済) で fabriq 側に解がないため accept & document。
  B2G 厳格監査環境向けに Office Customization Tool / KMS-ADBA / 監査
  一時無効化等の代替運用を提示。`project_crypto_security_review.md`
  メモリの A3 T1-B office セクションへの参照を追加
- A3 (transcript 漏洩点検) close-out 整理:
    - **A3 T1-A** (license keys masking, windows / office): ✅ 完了
      (本コミット `[Unreleased]` 参照、commit 5500cc1)
    - **A3 T1-B robocopy** (`net use` AuthPass stdin pipe 化): ✅ 完了
      (commit 2a52eee)
    - **A3 T1-B office cscript** (`cscript /inpkey:KEY` 引数経由): ⚠️
      Microsoft プラットフォーム構造的制約で fabriq 側 unfixable、
      accept & document (Guide.txt Security Note 参照)。Microsoft の
      platform 改善通知があれば再評価
    - **A3 T2** (`Read-Host` manual 入力 echo, `windows_license_install.ps1:78`):
      ⏸ defer。SecureString 化は可能だが 25 文字 product key を
      asterisks 入力する UX 劣化。発火条件 (license_key.csv 不在 or
      Enabled キーゼロ) は fabriq 設計上の例外パスで real pain 低
    - **A3 T3** (catch エラーメッセージ経由漏洩, `Show-Error "...: $_"`):
      ⏸ defer。Add-Computer / Import-Certificate / netsh 等の
      Microsoft 由来 cmdlet error は password 非含有が empirical 慣例、
      理論上のみ。各エラーパスを実機 trigger して検証する運用工数の
      問題で、real pain surface 待ち
  A3 の core security gain は T1-A + T1-B robocopy で達成済。残り
  T1-B office / T2 / T3 は ROI 低 or 構造的制約のため defer。次の
  natural step は A4 (resume_state stale 検出設計議論)、A5 (Studio
  リリース同調)、B 群 (B2G trigger 待ち)。詳細は `project_crypto_security_review.md`
- commands/diag_crypto.ps1: 診断出力からの機密情報漏洩を 2 箇所削減
  （crypto/security review メモ A1 + A2、`project_crypto_security_review.md`）。
  本変更は出力整形のみで、診断機能（passphrase SET 判定、CSV ENC: 値スキャン、
  Unprotect-FabriqValue 経由の復号成否確認）は不変。
    - **A1** L16: `FabriqMasterPassphrase: SET (length=$($pp.Length))` から
      `(length=...)` サフィックスを削除。マスターパスフレーズ長は brute-force
      探索空間を絞る情報のため、PowerShell Transcript に残るのを避ける。
      NOT SET 判定行（L14）は不変
    - **A2** L52: ENC: 値の先頭 30 文字プレビュー出力を削除。fixed salt /
      決定論的 IV と組み合わさると暗号文ヘッダ部の同一性が攻撃材料に
      なり得るため。診断価値（どの列が失敗したか）保持のため列名
      （`Column '$($prop.Name)':`）は出力継続。直後の Decrypt OK / FAILED
      ライン（L57 / L60）は不変
- (framing correction) 前 commit 5500cc1 の A1+A2+A3-T1-A entries で
  「PS Transcript に password が記録される」と記述したが、これは部分的
  に不正確だった。正確な leak vector の整理:
    - A1+A2 (`commands/diag_crypto.ps1`) と A3-T1-A (license modules) の
      `Write-Host` / `Show-*` 経由は **Console output 経由で展開後の値
      が PS Transcript に記録される** 経路（PowerShell Transcript の
      output capture 機能による）。実際の漏洩経路で、修正は妥当
    - 一方 A3-T1-B (今回 robocopy) は PS Transcript ではなく
      Win32_Process.CommandLine と 4688 audit log が leak vector で、
      PS Transcript には source code 行が literal で記録されるのみ
      （変数参照 `$($item.AuthPass)` のまま、展開後の値は出ない）
    - 修正は両方とも妥当かつ価値あり。framing 上「PS Transcript への
      漏洩」と一括りで扱った前 commit のメッセージを精度訂正する記録
- 公開 API 不変（`KERNEL_API.md` §1〜§5 touched せず）。kernel/common.ps1
  も touched せず、Unprotect-FabriqValue / Import-ModuleCsv の挙動・
  シグネチャ完全一致

## [3.2.5] - 2026-05-10

### Added
- tests/: kernel ユニットテスト基盤（Pester v5+）の Phase 0 + Phase 1 着手。
  公開 API は不変、production code 1 行も touched せず、純粋に追加のみ。
    - `tests/_helpers/test_state.ps1`: `Get-FabriqRepoRoot` パス解決ヘルパ
      （Phase 2 で `Set-FabriqTestState` 等の script-scope 設定 helper を
      追加予定）
    - `tests/_helpers/test_csv.ps1`: `New-TestProfileCsv` /
      `Remove-TestProfileCsv` / `New-MockModule` の 3 ヘルパ。一時 Profile
      CSV は `Resolve-ProfileModules` の `Import-Csv -Encoding Default`
      に合わせ ASCII で書き出し、JP locale CP932 環境の BOM 誤読を回避
    - `tests/kernel/Resolve-ProfileModules.tests.ps1`: 6 Context / 14
      ケース。CHANGELOG 上の過去 regression（kernel 3.0.0 / 3.1.3 / 3.1.7
      / 3.2.0 / 2.1.0）を将来検出可能に。AutoPilot Enabled=0 ガード /
      WaitSec=N パース / `__ASYNC__` kill switch（`Get-FabriqAsyncConfig`
      mock 経由）/ Group 列の trim・欠損 fallback / Order int 紐付け
      + 昇順ソート / Sibling-row Segment 分離（PSObject.Copy 経由の
      source 不変性検証）/ `__RESTART__` `_IsRestart` flag /
      `__AUTO_to_<User>__` 展開 / 退役マーカー graceful degradation を
      pin。kernel touch ゼロ
    - `tests/kernel/New-BatchResult.tests.ps1`: 4 Context / 18 ケース。
      Status 自動判定 5x3 マトリクス全分岐 (Success only / Skip only /
      Fail only / Success+Skip / Success+Fail / Skip+Fail / 全部混在 /
      All-zero default branch) + Verified pass-through ($true / $false /
      $null + Partial と組み合わせ) + Message format ("Success: N, Skip:
      M, Fail: K" + MessageSuffix のスペース連結) + ModuleResult shape
      contract (`_IsModuleResult=$true`, `Timestamp` is DateTime,
      `Details` defaults to empty array)
    - `tests/kernel/Import-ModuleCsv.tests.ps1`: 4 Context / 14 ケース。
      Segment 厳密マッチ 5 ケース (列欠損 = フィルタ無効 / 空 vs 空 = 一致
      / 値 vs 同値 = 一致 / 値 vs 別値 = 不一致 / 両側 trim) + Segment
      whitespace-only 正規化 + `-FilterEnabled` Enabled=1 抽出 + 未指定時
      全件返却 + `-RequiredColumns` 必要列存在/欠落分岐 + caller-observable
      boundaries 4 ケース (filter eliminates all = `$null`、Segment filter
      eliminates all = `$null`、file 不存在 = `$null`、header-only 空ファ
      イル = `$null`)。最後の 4 ケースは「`Import-ModuleCsv` が内部的に
      `@()` を返すパスは PowerShell の scalar-unwrap で呼び出し側では
      `$null` に潰れる」という現行挙動を pin。これにより標準モジュール
      テンプレートの `if ($items.Count -eq 0)` 分岐が実質 dead code であ
      ることが明文化された (将来の kernel 改修で `,@()` 化する場合は本
      テストの reframe が必要、回帰検出のシグナル)
    - `tests/_helpers/test_csv.ps1`: `[AllowEmptyCollection()]` 化 +
      `-Rows` 既定値 `@()` で header-only fixture を許可
      (Import-ModuleCsv 空ファイル分岐のテストに必要)
    - `tests/_helpers/test_state.ps1` 拡張: `Set-FabriqTestState`
      ヘルパ追加 (`$script:ResumeStatePath` / `$script:SessionID` /
      `$script:HistoryPath` / `$script:StatusFilePath` を tests から
      override 可能に)。production code は不変、main.ps1 経由の
      Initialize-Session を経由しないテストでも production 既定の
      相対パス (CWD 根拠) で実状態を上書きしないよう隔離
    - `tests/kernel/ResumeState.tests.ps1` (Phase 2): 4 Context /
      15 ケース。`Save-ResumeState` / `Load-ResumeState` の on-disk
      JSON 形式契約と __RESTART__ 越え round-trip を pin。
        - **Linear (v1) 出力 4 ケース**: `schemaVersion` /
          `ExecutionMode` / `SelectedOrders` / `ModuleStates` 4 fields
          が **全て不在** であることを検証 (pre-FlexProfile v1 file との
          byte-compat) / 核心 fields (ProfilePath, ProfileName,
          ResumeAfterOrder, SessionID, AutoPilot 系, EvidenceBasePath)
          シリアライズ / CompletedModules を Order/MenuName/Status
          triple として記録 / ProfileStartTime は ISO 8601 "o" 形式
        - **Flex (v2) 出力 4 ケース** (kernel 3.1.0 導入): `-ExecutionMode
          Flex` で schemaVersion=2 + ExecutionMode='Flex' を emit /
          SelectedOrders 配列シリアライズ / ModuleStates hashtable を
          JSON object 化 (`'10' = @{Status='Success'; Verified=$true}`
          形式の round-trip) / v1 核心 fields も同時保持
        - **Round-trip 4 ケース**: v1 形式の Save → Load 完全回復 /
          v2 形式の Save → Load (schemaVersion / SelectedOrders /
          ModuleStates 全保持) / HostEnvironment スナップショットの
          値保存 (SELECTED_NEW_PCNAME を実 env で設定 → save → load
          → 復元値一致、テスト後に env 復元) / SELECTED_PRINTER_<N>_*
          (10 slots × 3 suffixes = 30 keys) の HostEnvironment 構造
          保証
        - **Load 境界 3 ケース**: ファイル不在で `$null` 返却 /
          壊れた JSON で `$null` 返却 (try/catch 経路) / pre-FlexProfile
          v1 ファイル (schemaVersion 不在) を error 無く読める
          (kernel 3.1.0 後方互換契約)
        - `Mock Write-KernelTelemetryEvent {}` で telemetry 副次副作用
          を抑止。`$global:AutoPilotMode` / `WaitSec` /
          `FabriqMasterPassphrase` / `FabriqEvidenceBasePath` を BeforeEach
          で初期化、tmp resume_state ファイルは BeforeEach で生成
          AfterEach で削除
    - `tests/kernel/New-ModuleResult.tests.ps1` (Phase 3): 5 Context /
      19 ケース。Phase 1b で `New-BatchResult` の wrapper をテスト済の
      foundation 契約 (`KERNEL_API.md` §1.3 / §5) を直接 pin。
        - **Status ValidateSet 6 ケース**: 5 つの正規 status (Success /
          Error / Cancelled / Skipped / Partial) を `-ForEach` で受理 +
          不正値での parameter-binding throw
        - **Defaults 3 ケース**: Message 既定 `''` / Details 既定 `@()`
          (`[System.Object[]]` 型) / Verified 既定 `$null` (`-eq $null`
          の identity 含む)
        - **Shape contract 3 ケース**: `_IsModuleResult=$true` /
          Timestamp が `[datetime]` / property 6 件ちょうど
          (`_IsModuleResult, Status, Message, Details, Verified,
          Timestamp`、それ以上もそれ以下もない契約)
        - **Verified Nullable[bool] pass-through 3 ケース**: $true /
          $false (Success Status と組み合わせ) / 明示的 $null
          (sysprep 等の verify 不能ケース)
        - **Details pass-through 2 ケース**: PSCustomObject 配列の
          verbatim 保存 / string 配列の保存
        - **`$global:_LastModuleResult` side-effect 2 ケース**:
          返却値と global の reference identity 一致 / 後続呼び出しでの
          上書き (FlexProfile AutoConfirmMode が pipeline 取りこぼし
          時に頼る fallback の存在を pin)
    - `tests/kernel/Confirm-Execution.tests.ps1` (Phase 3): 3 Describe
      / 14 ケース。AutoPilot (kernel 2.x 系) と AutoConfirmMode (kernel
      3.1.5 で導入、FlexProfile 単発 [Run] 用) の二段短絡契約
      (`KERNEL_API.md` §1.4 / §2) を pin。回帰のシグネチャは
      "AutoPilot 中に Y/N ダイアログが顔を出す" "Flex 単発 [Run] で
      Press-Enter が止まる" の 2 系統で、いずれも静的解析では検出
      不可能なため Read-Host モックで明示固定。
        - **Confirm-Execution AutoPilot 短絡 1 ケース** + **AutoConfirm
          短絡 2 ケース** (片側のみ true / 両 true): `Mock Read-Host
          { throw }` で「呼ばれていないこと」を意図明示し、`Should
          -Invoke -Times 0 -Exactly` で固定
        - **Confirm-Execution 対話 5 ケース**: Y / y / N / n の case
          insensitivity + invalid → '' → Y の loop 復帰 (Read-Host が
          3 回呼ばれること、最終 return が `$true` であることを同時
          検証)
        - **Wait-KeyPress 3 ケース**: AutoPilot=true で Read-Host 呼ば
          れず / AutoConfirm=true でも同じ / 両方 false で Read-Host
          が 1 回だけ
        - **Confirm-ModuleExecution 3 ケース**: AutoPilot 中は `$null`
          (proceed signal) / Y で `$null` / N で `Cancelled` の
          ModuleResult (Status / Message='User canceled' / `_IsModuleResult`
          marker) を返却
    - `tests/kernel/Set-SelectedHostEnvironment.tests.ps1` (Phase 4):
      4 Context / 16 ケース。Profile 実行の背骨である hostlist 行 →
      `SELECTED_*` env vars 47 個 (3 identity + 3 ethernet + 3 wifi +
      4 DNS + 1 PIN + 30 printer slots) のマッピング契約
      (`KERNEL_API.md` §3.1) と ENC: 復号 routing 契約 (§1.5) を pin。
      kernel/main.ps1 は top-level 起動 side effect (sleep suppression
      / GUI load / `exit 1` if no GUI) を持つため通常 dot-source 不可
      なので、新規 helper `Get-FabriqMainFunctionScriptBlock`
      (`tests/_helpers/test_state.ps1`) で
      `System.Management.Automation.Language.Parser::ParseFile` 経由
      に AST から `FunctionDefinitionAst` を抜き出し scriptblock 化、
      テスト側で `. $sb` して startup 副次副作用を持ち込まずに関数
      本体だけを inject する pattern を新規導入。
        - **Plain identity / network env mapping 5 ケース**:
          identity trio (AdminID → SELECTED_KANRI_NO + OldPCName +
          NewPCName) / Ethernet trio / Wifi trio / DNS 4 slots /
          Pin → SELECTED_PIN。CSV 列名 → env var 名の rename / swap
          regression をピンポイントで検出
        - **Printer slot enumeration 4 ケース**: slot 1 trio / slot 10
          trio (`for ($i = 1; $i -le 10; ...)` の off-by-one guard、
          `-lt 10` への退化が surface する) / 全 30 keys (1..10 ×
          NAME/DRIVER/PORT、cross-slot や cross-suffix の wiring エラー
          を sentinel 値経由で検出) / 欠損 slot 入力での no-throw
          (`Set-Item -Value $null -ErrorAction SilentlyContinue` の
          $null bind 黙殺で slot carryover 挙動が未仕様であることを
          コメントで明文化)
        - **Empty / null value handling 2 ケース**: 空文字列 input は
          decryption に進まず empty env (Mock Unprotect-FabriqValue
          throw で短絡を assert) / 列欠損 PSCustomObject input は
          BeNullOrEmpty で吸収 (PS 5.1 は `$env:X = $null` で unset
          / read-back $null、PS Core は '' に coerce する処理系差を
          test 側で許容)
        - **ENC: decryption routing 5 ケース** (Mock Unprotect-FabriqValue
          で algorithm 切り離し、本体は `Unprotect-FabriqValue.tests.ps1`
          で別途 pin 予定): plain pass-through (passphrase 設定済でも
          ENC: prefix 無しなら decrypt しない) / passphrase 未設定時の
          ENC: pass-through (anti-leak: パスフレーズ無いのに silent
          decrypt は禁) / passphrase 空白のみ時も pass-through
          (`IsNullOrWhiteSpace` guard) / 正規 path での invocation
          (passed-through passphrase の identity 確認 + 復号結果の env 反映)
          / 復号 throw 時の graceful fallback (raw ENC: 値が env に
          残り、catch 経路で Show-Warning が発火)
- dev/run_tests.ps1: Pester v5+ runner。`tests/` と
  `apps/fabriq_ios/tests/` を一括実行。Pester v5 未インストール環境で
  は明示エラー + `Install-Module` ヒント表示で exit 1（v3.4.0 の
  Windows 同梱版では v5 構文の `Should -Be` / `BeforeAll` が動かない
  ため、`-SkipPublisherCheck` 付きインストールが必須）
- README.md: `## テスト` 節を新設（前提 Pester v5+、実行コマンド、
  ディレクトリ構成）
- dev/framework_overlay_rules.json: `bundles.kernel.includePaths` に
  `tests/` を追加。Studio 経由の overlay で kernel bundle と一緒に
  tests/ も配信される（schema 不変、外部 tool 側コード改修不要 —
  JSON 動的読込みのため）
- kernel/KERNEL_API.md §9.2: kernel bundle 構成行に `tests/` を追記
  （overlay rules JSON の真実源と同期）
- CLAUDE.md ルール E（事前宣言）/ ルール F（実装サマリ）テンプレに
  「テスト影響」「テスト実行結果」項目を追加。kernel または
  modules/** を touched する作業では、関連テストの宣言・実行・
  結果報告を必須化

### Fixed
- `apps/fabriq_ios/tests/completer.tests.ps1`: Phase 8 fork (commit
  4ad817d) で hostlist / workerlist 結合を切った際に同期漏れしていた
  stale テスト 2 件を Phase 8 契約の negative assertion に書き換え。
    - `completes hostname candidates from hostlist` → `does not suggest
      hostlist entries after hostname verb (Phase 8 contract)`: `host `
      補完が NEW-PC-01 等の hostlist 項目を **含まない** ことと候補が
      空であることを pin
    - `completes ip address candidates with from-hostlist literal` →
      `does not surface from-hostlist literal after ip address (Phase 8
      contract)`: 廃止された `from-hostlist` literal が **再浮上しない**
      ことを pin
  Phase 8 で `Get-DynamicCompletion` の `hostname` / `ip.address`
  source case を削除した時点でこれらは事実上 dead test だったが、
  期待値が古いまま残って 2 件 fail を継続していた。Get-FabriqIosCompletion
  本体は改修なし、製品挙動は不変。Phase 1 で `__SHUTDOWN__` 退役
  マーカーを `[+] collects retired __SHUTDOWN__ marker as InvalidPath
  (graceful degradation, kernel 3.0.0)` として pin しているのと
  同じパターン。
- `kernel/KERNEL_API.md` L3 `**Current Kernel Version**` ヘッダの drift
  解消（3.2.2 → 3.2.4）。3.2.3 / 3.2.4 の 2 連続 PATCH 昇格時に sync 漏れ
  していたものを retroactive 修正。
- `README.md` モジュール数表記と一覧の整合化。実態と乖離していた箇所を
  一括補正:
    - L15 概要: 「70 種類超」→「76 種類」
    - L30 機能表: 「Standard 59 種、Extended 14 種、計 74 種以上」→
      「Standard 60 種、Extended 16 種、計 76 種」
    - L65 ディレクトリツリー: 「拡張モジュール群（14）」→「（16）」
    - L200 Standard モジュール一覧見出し: 「（59）」→「（60）」
    - L213 Standard System カテゴリに `windows_feature_config` を追記
      （v3.2.3 で新規追加されていたが README 反映漏れ）
    - L221 Extended モジュール一覧見出し: 「（14）」→「（16）」
    - L228 Extended System カテゴリに `server_feature_config` を追記
      （直近新規追加分）
    - L230 Extended ManualWorks カテゴリに `pianist` を追記（v1.6.0 で
      extended モジュール統合済だったが README 反映漏れ）
  bloatware_export 削除に伴う README sync は本エントリに統合。

### Changed
- CLAUDE.md ルール J step 3 の sync list に `kernel/KERNEL_API.md` L3 を
  追加（従来 3 箇所 → 4 箇所）。`KERNEL_API.md` L3 drift 観察を受けた
  恒久対策。`dev/check_version.ps1` も同 L3 ヘッダの drift 検出を追加し、
  リリース時の見落としを構造的に防止。
- modules/standard/{firewall_rule_config, firewall_rule_make_config,
  temp_ipaddress_config}: VERSION 0.1.0 → 1.0.0。2026-04-23 の baseline
  seed (`dev/seed_module_versions.ps1`) は idempotent で既存値保持の
  挙動だったため、上記 3 件は当時すでに 0.1.0 が打刻されていたことに
  より 1.0.0 baseline へ昇格漏れしていた。実装は production 安定運用
  実績ありで、CLAUDE.md ルール H 定義（`0.1.0` = 「開発中・未リリース」
  の目印 / `1.0.0` = baseline）に整合させる訂正。スクリプト本体は
  touched せず VERSION ファイル 1 行のみの変更。
  windows_feature_config / extended/server_feature_config は直近新規
  追加分で開発中扱いを継続するため `0.1.0` 据え置き。
- apps/bloatware_exporter: CSV スキーマを bloatware_remove v1.0.0 の
  5 列形式 (`Enabled, DisplayName, MatchPattern, Description, Segment`)
  に揃える。旧 12 列形式 (Publisher / DisplayVersion / Architecture /
  WindowsInstaller / QuietUninstallString / UninstallString / NoRemove /
  SystemComponent / InstallDate / RegistryKey) は bloatware_remove の
  動的レジストリルックアップ方式 (実行時に HKLM Uninstall ハイブから
  UninstallString 等を取得) 採用以降不要だったが、exporter Save が
  新形式 `bloatware_list.csv` を旧 12 列で全件上書きして破壊する
  unsafe な状態が放置されていたため整合化。`Add to CSV >>` 押下時に
  MatchPattern を `<DisplayName>*` で seed し wildcard 一括削除を
  デフォルト挙動に。Scan grid (上段 7 列スナップショット) は
  Publisher / NoRemove / SystemComponent の警告視認のため維持。
  load → append-with-dedup → 全件 save の操作モデルは不変

### Removed
- modules/standard/bloatware_export: 廃止。役割を
  apps/bloatware_exporter に統合 (機能重複)。profile 参照は皆無
  (3.2.4 までの全 profile で未使用、Order=99 / Enabled=0 の
  デフォルト無効モジュールだった) のため互換性影響なし。
  `apps/fabriq_ios/data/module_categories.json` の settings カテゴリ
  からも該当エントリを除去。README.md 側の Applications カテゴリ表記
  ・モジュール数の同期は下記 Fixed エントリで一括修正

## [3.2.4] - 2026-05-10

### Added
- kernel/common.ps1 + kernel/main.ps1 (PATCH): Verbose stream capture
  channel (`cmdlet.verbose` イベント) — **標準配備、デフォルト ON**。
  `kernel/json/verbose_capture.flag` を git tracked で同梱、起動時に
  presence をチェックして verbose capture を活性化。`Invoke-SafeCommand`
  がモジュール実行を `$VerbosePreference='Continue'` +
  `$PSDefaultParameterValues['*:Verbose']=$true` + `4>&1` redirect で
  wrap し、cmdlet の verbose ストリームを `cmdlet.verbose` 構造化
  イベントに変換してモジュール envelope に append。出力は local-only
  （log_uploader が `logs/telemetry/` を除外、顧客 PC 外へ流出しない）。
  顧客がオプトアウトしたい場合は flag 削除で無効化（escape hatch）。
    - 新規関数: `Enable-FabriqVerboseCapture` /
      `Disable-FabriqVerboseCapture`（いずれも内部実装、KERNEL_API.md §6
      範囲、PATCH 可）
    - 内部実装変更: `Invoke-SafeCommand` の `& $ScriptBlock` 呼び出しを
      flag 有効時のみ `4>&1 | ForEach-Object` でストリーム redirect 化。
      `VerboseRecord` 型の record だけフィルタして telemetry 行き、その他
      の output（ModuleResult 含む）は pass through。`$VerbosePreference`
      は finally で復元
    - 公開 API 不変: シグネチャ・戻り値・コンソール出力（Show-* の
      Write-Host 経路は 4>&1 の影響を受けない Information ストリーム）
      は不変
    - 取得対象: `Set-ItemProperty`, `Rename-Computer`, `Set-NetFirewallProfile`,
      `New-LocalUser` 等の組み込み cmdlet が標準で発行する "Performing
      the operation..." verbose 行。ユーザー定義関数は `Write-Verbose`
      呼び出しがある場合のみ捕捉、外部プロセス（winget / robocopy / dism）
      は対象外
    - redact: 既存 `New-TelemetryRedactMap` を再利用。verbose の `msg`
      フィールドが redact map 通過
    - **process restart 不要**（先行検討した Module Logging E1 trial は
      registry を mid-session に書き換えても既存ロード済みモジュールに
      反映されない致命的制約があり未採用 / 撤退。本機能は
      `$VerbosePreference` のランタイム評価により即時有効化される）
    - Phase 1 は Invoke-SafeCommand 同期パスのみ。`Invoke-SafeCommandAsync`
      の child runspace 内 redirect は Phase 2 で扱う
    - dev/TELEMETRY_INTERNAL.md §6.7 に schema・coverage・trade-off を
      明文化
- kernel/json/verbose_capture.flag (新規): 標準配備で verbose capture を
  活性化するためのフラグファイル（git tracked）。fabriq 配布物に常時
  含まれる。削除が opt-out の唯一の方法
- .gitignore: 旧 `/kernel/json/verbose_capture.flag` 除外エントリを撤去
  （上記方針変更により git tracked へ移行）

- kernel/common.ps1 + kernel/main.ps1 (PATCH 予定): Telemetry Layer の
  カバレッジ拡張（v3.2.3 で導入された AI 開発コーパス機構の機能追加。
  公開 API §1〜§5 不変、KERNEL_API.md 不変、`dev/TELEMETRY_INTERNAL.md`
  schemaVersion=1 内での後方互換追加）。
    - **(A) `csv.load` イベント追加**: `Import-ModuleCsv` の正常 return
      時に `{fileName, path, totalRows, returnedRows, filterEnabled,
      segment, columns}` を発行。CSV 値そのものは含まない構造メタのみ。
      AI が「reg_hklm_config が KeyPath=... の場合に失敗しがち」のような
      decision-trace パターンを追跡可能に
    - **(B) Profile 実行コンテキストの envelope.start 追加フィールド**:
      `Invoke-BatchExecution` (main.ps1) が `$global:_FabriqCurrentProfileContext`
      を per-module で立て、`Start-ModuleTelemetry` (common.ps1) が読み取って
      envelope.start に `profileName`, `profileOrder`, `executionMode`
      (Linear/Flex), `prevModuleName`, `prevModuleStatus` を追加。
      cross-module 依存検出（「IP 設定が Skipped の後で domain_join が
      失敗する」パターン）が AI から見えるように
    - **(C) `_meta.json` に host info 追加**: `Get-TelemetryHostInfo`
      (Win32_OperatingSystem + Win32_ComputerSystem from CIM、1 セッション
      1 回キャッシュ) で `os: {caption, version, build}`, `hardware:
      {manufacturer, model, ram_gb}`, `powershell` を記録。manufacturer /
      model はフリート単位の識別子（ThinkCentre / OptiPlex 等）で
      顧客個体ではないため hash 化せず raw で保存。PC モデル横断で
      "this hardware class has X-class issues" を AI が分析可能に
    - **(D) Kernel events channel `_kernel.jsonl` 新設**: 各セッション
      ルートに module envelope と並列の session-lifecycle イベント
      ファイル。新規 helper `Write-KernelTelemetryEvent` が以下イベントを
      発行:
        - `profile.start` / `profile.end`: `Invoke-BatchExecution`
          (main.ps1) entry/exit。outcome enum (`Success` / `WithErrors` /
          `Partial`) + per-status counts
        - `restart.invoked`: `Save-ResumeState` (common.ps1)。`__RESTART__`
          マーカーと FlexProfile [Restart Now] の両者を捕捉
          (resumeAfterOrder=-1 が後者の sentinel)
        - `resume.consumed`: `Load-ResumeState`。post-reboot session の
          先頭で発火
        - `finalize.start` / `finalize.end`: `Complete-ProfileExecution`
          (common.ps1) wrap。Auto / Manual モード別の HTML+log_uploader
          パイプライン計時
    - 内部実装関数 1 件追加: `_WriteTelemetryMeta` (`_meta.json` 書き込み
      ロジックを `Start-ModuleTelemetry` から抽出、Write-KernelTelemetryEvent
      も使う共通化)
    - **モジュール無変更**: 74 モジュール touched ゼロ。Import-ModuleCsv 経由
      で 73 モジュールが csv.load イベントを自動的に出力（副次的副作用）
- dev/TELEMETRY_INTERNAL.md (更新): §6 イベント仕様の追記
  (csv.load §6.2、kernel events channel §6.6 に profile.start/end /
  restart.invoked / resume.consumed / finalize.start/end 各イベントの
  schema を明文化)。`_meta.json` schema (§5) に host フィールドの仕様
  と redaction rationale を追加

## [3.2.3] - 2026-05-09

### Added
- kernel/common.ps1 (PATCH): Telemetry Layer 追加。AI 開発コーパス
  生成のためのモジュール内部挙動の構造化ログ機構。
    - `Show-Info` / `Show-Success` / `Show-Warning` / `Show-Error` /
      `Show-Skip` 5 関数に副次的副作用を追加（既存の戻り値 / 出力ストリーム
      / コンソール表示は不変）。各呼び出しが `logs/telemetry/{SessionID}/
      modules/{seq:0000}_{module}.jsonl` に show.* イベントとして
      append される
    - `Invoke-SafeCommand` / `Invoke-SafeCommandAsync` の entry / exit
      に envelope hook 追加（`Start-ModuleTelemetry` /
      `Complete-ModuleTelemetry`）。envelope.start / envelope.end
      イベント + 累積した `$Error` を error イベントとして flush
    - 新規関数: `Get-TelemetrySalt`, `New-TelemetryRedactMap`,
      `Invoke-TelemetryRedact`, `Write-TelemetryEvent`,
      `Start-ModuleTelemetry`, `Complete-ModuleTelemetry`,
      `_GetShowTag`, `_TrackShowEvent`, `_HashTelemetryValue`
    - プライバシ契約: `$env:SELECTED_*` / `FABRIQ_WORKER_NAME` /
      `COMPUTERNAME` / `$global:FabriqUniqueId` を salt 付き
      SHA-256 (12 hex) に hash 化、`SELECTED_PIN` は `[REDACTED]`
      固定。Salt は `kernel/json/telemetry_salt.txt` (256 bit、初回
      自動生成、site-specific、`.gitignore`)
    - 公開 API には未昇格。`KERNEL_API.md` 不変。設計は
      `dev/TELEMETRY_INTERNAL.md` (新規、非公開) に集約。
      schemaVersion=1 はファイル内フィールドで管理し、将来公開契約へ
      promote する場合に §11 新設で MINOR 昇格予定
    - 子 runspace（`__ASYNC__`）でも envelope 共有: 親が立てた
      `$global:_CurrentModuleTelemetry` を inject hashtable 経由で
      子に渡し、Show-* / ShowCounts は同じ hashtable を共有
    - **Scope 設計**: telemetry の内部状態変数 (`_TelemetryWriting`,
      `_TelemetrySalt`, `_TelemetrySaltDigest`, `_TelemetryModuleSeq`,
      `_TelemetryUtf8NoBom`) は `$global:` で保持。`$script:` だと
      Show-* が `& module.ps1` 内から呼ばれた瞬間にモジュール側の
      script scope（変数未定義）を参照してしまう PowerShell の
      動的 scope 仕様に対応するための設計判断（test_harness 検証で
      surfacing）
- dev/TELEMETRY_INTERNAL.md (新規): Telemetry 設計メモ。schemaVersion 1 の
  ファイル配置・redact 方針・JSONL イベント書式・実装サーフェス・
  reentrancy / 失敗隔離契約を明文化
- modules/extended/log_uploader v1.0.0 → v1.1.0 (MINOR):
  robocopy `logs/` 転送に `/XD logs\telemetry` 除外を追加。Telemetry
  は AI 開発専用の内部コーパスでありキッティング PC 外への持ち出し禁止
  契約（`dev/TELEMETRY_INTERNAL.md` §4）を実装で担保
- .gitignore: `kernel/json/telemetry_salt.txt` 追加（site-specific salt、
  環境ごとに自動生成、commit 禁止）

- modules/standard/sysprep_config v1.0.0 → v1.1.0 (MINOR):
  unattend.xml の generalize パス `Microsoft-Windows-PnpSysprep` 配下
  にハードコードされていた `DoNotCleanUpNonPresentDevices` /
  `PersistAllDeviceInstalls` を `unattend_list.csv` から制御可能に。
    - 新規 SettingName: `DoNotCleanUpNonPresentDevices`,
      `PersistAllDeviceInstalls` (Value: true / false)
    - 既存挙動の完全互換のため、行が無い場合・Enabled=0 の場合は
      他の SettingName と異なり「要素省略」ではなく「デフォルト true
      を出力」する非対称セマンティクスを採用 (v1.0.x の固定挙動を
      バイト一致で保持)。明示的に false にする場合は Enabled=1,
      Value=false を指定
    - テンプレートに `{{GENERALIZE_DRIVER_BLOCK}}` プレースホルダ追加
    - `unattend_list.csv` に 2 行 seed (どちらも Enabled=1, Value=true)
    - Guide.txt の `[対応 SettingName]` に新カテゴリ
      「generalize パス（ドライバ保持）」を追加し、非対称セマンティクス
      を明記
- kernel/common.ps1 + kernel/ps1/status_monitor.ps1 (PATCH): Status
  Monitor の起動診断ログと子側 defensive try/catch を全段に追加。
  端末によって Status Monitor が立ち上がってこない事象 (子プロセスが
  silently 即死) の根本原因特定が可能になった。
    - **親側 (`Start-StatusMonitor`)**: 子プロセス spawn 直後に 500ms
      sleep + `HasExited` 即死検知。即死時は ExitCode と診断ログ
      パスを Show-Warning で surface。`Test-Path $monitorScript`
      false で明示警告 (従来は無言で `$null` 返却)。`logs\status_monitor_<ts>.log`
      を 1 セッション 1 ファイル生成し、`-DiagLogPath` 引数で子に渡す
    - **子側 (`status_monitor.ps1`)**: `[string]$DiagLogPath = ""`
      param 追加。`Write-DiagLog` ヘルパ (`Add-Content` best-effort、
      throw しない設計) を導入し、起動チェーン全段にログを残す:
      DPIUtil Add-Type / SetProcessDPIAware / `Graphics.FromHwnd` の
      DpiX 取得 / `System.Windows.Forms` ロード / `Native.Win32`
      type 定義 / `NoActivateForm` Add-Type / Console hide /
      common.ps1 dot-source / Form 生成と配置 / `Application.Run`。
      失敗箇所は exit code 11-14 で識別 (System.Drawing / WinForms /
      common.ps1 / Application.Run)
    - **DpiX ゼロ・負値ガード**: `Graphics.FromHwnd([IntPtr]::Zero).DpiX`
      が病的に 0 や負を返す環境で Form Size が `(0,0)` になり「実行
      しているのに不可視」となる経路を遮断 (96 fallback、`dpiScale<=0`
      は 1.0 にクランプ)
    - **`NoActivateForm` Add-Type 失敗時の fallback**: 動的 C# コンパイル
      が AMSI ブロック / `csc.exe` パス問題 / `%TEMP%` 書込権限欠如等で
      失敗した場合に通常の `System.Windows.Forms.Form` および
      `System.Windows.Forms.StatusStrip` で代用。フォーカスを奪わない
      挙動 (`WS_EX_NOACTIVATE`) は失うが、**Form が画面に出る** ことを
      優先。`$script:useFallbackForm` flag で分岐し、失敗理由は
      診断ログに記録
    - **`Application.Run` を try/catch でラップ**: 例外時は
      `ScriptStackTrace` を診断ログに残して exit 14
    - **ウィンドウ存在ポーリングによる成否判定**: 当初は spawn 後
      `HasExited` の 500ms / 2 秒チェックで判定していたが、App Control
      配備済端末では子が 2 秒以内には死なず後段で死亡 / hung するケース
      があり取りこぼしが発生 (2026-05-09 NG 端末再現で確認)。判定の
      主シグナルを **「Fabriq タイトルのメインウィンドウが
      `$monitorProcess.MainWindowTitle` に現れたか」** に変更。最大
      4 秒間 200ms 間隔でポーリングし、(1) ウィンドウ検出 = 成功、
      (2) `HasExited` = 早期死亡、(3) timeout でウィンドウ無し =
      hung と分類 (hung の場合は orphan を防ぐため `Kill()`)。
      ウィンドウ検出時は poll を即座に抜けるため正常時の起動オーバー
      ヘッドは ~1 秒程度
    - **失敗時の WinForms `MessageBox.Show` 通知**: console の
      Show-Warning だけでは dashboard 表示直後に conhost が hide
      されて operator が見落とすため、`Show-MonitorFailureDialog`
      ヘルパで modal ダイアログを表示。`MessageBox` 自体も policy で
      ブロックされる可能性があるため try/catch で best-effort、失敗
      時は console warning のみが残る。OK ボタン押下まで kitting フロー
      は停止 = 「明示的にあきらめる」UX
    - 失敗分類:
        - **早期死亡 + ログ無し**: "host security policy is blocking
          the child PowerShell process (WDAC / AppLocker / Defender
          ASR / equivalent)"
        - **早期死亡 + ログあり**: "exited mid-startup", 最終ログ行
          を modal/console に表示
        - **window 未出現 + alive**: "no window within 4 seconds
          (likely hung)", 子プロセスを `Kill()` で終了
    - 挙動仕様の変更ゼロ (正常時): 既存 OK 端末では window 検出で
      poll を即座に抜けるため挙動・タイミング共に従来と同等。TopMost
      化や Form 位置クランプは見送り (Form が出ないケースに対しては
      Z-order 系の対処は効かないと判定)
    - 副作用: `logs\status_monitor_*.log` が 1 セッション 1 ファイル
      増える (各段ログのみ、数 KB〜十数 KB)。`Reset-FabriqState` の
      logs/*.log clear 対象に自動的に含まれる。失敗時の起動時間が
      最大 +4 秒 (poll timeout)、正常時は概ね +1 秒程度
    - 公開 API への影響なし (KERNEL_API.md §6 で Status Monitor の
      内部実装は PATCH 可と明示済み)、KERNEL_API.md 更新不要

### Fixed
- kernel/common.ps1 (PATCH): `$script:ArtPulseFilePath` /
  `$script:ArtPulseCounter` を `$global:` に変更。`$script:` の動的 scope
  仕様により `Write-ArtPulse` が `& module.ps1` 内から呼ばれた際に変数 null
  → `WriteAllText` 例外を silent 発生していた長期バグを修正。telemetry
  追加で初めて surfacing。修正後はカウンタが全 Show-* 呼び出しで正しく
  累積し、Status Monitor のアート pulse animation が一貫して反応する
  ようになる（pre-existing バグの副次的修正、公開 API 不変）

### Changed
- modules/standard/domain_join **2.0.0** (MAJOR): 失敗時の振る舞いを
  fabriq 標準 ErrorMode 機構へ完全集約。FlexProfile 導入後 (kernel 3.1.x
  以降)、本モジュール内部の無限リトライループは外側の AutoPilot ErrorMode
  分岐 (retry / skip / Show-AutoPilotErrorDialog) を dead path 化させる
  だけで FlexProfile dashboard への戻り契約も踏み越えていたため撤去。
    - `while ($true) { try {...} catch { Show-ErrorDialog; continue } }`
      ループを削除。`Add-Computer` 例外時は `New-ModuleResult Error` を
      即返却するよう変更。失敗時のリトライ・スキップ・ダイアログ提示は
      profile CSV `ErrorMode` 列および FlexProfile dashboard
      (`[Run]` 単発再実行 / Status バッジ視認) に委譲
    - 自前の WinForms `Show-ErrorDialog` 関数および「adminstop」文字列
      入力による中断 UX を撤去 (他モジュールに無いローカル方言、
      無人運用との相性も悪い)
    - `Add-Type System.Windows.Forms / System.Drawing` 不要化により削除
    - DNS 事前チェックを `Wait-NetworkReady` (kernel API、無限ループ)
      から `Test-Connection -Count 2 -Quiet` の bounded local probe に
      置き換え。不到達時は `New-ModuleResult Error -Message "DNS unreachable: <ip>"`
      を即返却し、FlexProfile dashboard の Status バッジで視認可能に
    - `domain.csv` スキーマ・`module.csv`・公開呼び出し契約は不変
      (profile 側参照ゼロ、外部呼び出し影響なし)
    - 推奨 profile 設定: 一過性 DNS 障害向けに `ErrorMode=retry` を併記
      (kernel 既定 AutoPilotMaxRetry=5)
    - Guide.txt も新挙動に合わせて改訂 (リトライダイアログ・adminstop
      記述削除、AutoPilot/FlexProfile 節を新設)

### Added
- modules/standard/evidence_config **1.7.0**: 新規 `evidence_list.csv` で
  各セクション (§01〜§31 + §8b、計 32 種) の取捨選択が可能に。default-on
  policy: CSV 不在 / 該当 Id 行不在 / Enabled 不正値はすべて「有効」扱い
  となり、CSV を編集しない限り従来通り全 32 セクションを収集する非破壊
  設計。無効化されたセクションは manifest.json で
  `status="Skipped"` / `reason="Disabled by configuration (evidence_list.csv)"`
  として記録され、外部 evidence consumer は intrinsic skip (Server-only /
  バッテリ非搭載 等) と区別可能。EVIDENCE_MANIFEST.md schemaVersion=1
  維持 (新 enum 値・新必須フィールドなし)。master log には
  `[Section XX] Title : Skipped (disabled by configuration)` が 1 行
  グレーで残る。preset.csv 新規追加で Studio 側 `Enabled` 列がドロップ
  ダウン編集 UI に。`Test-SectionEnabled` / `Write-DisabledSection` の 2
  ヘルパーをモジュール内に新設 (kernel 改修不要、`REQUIRES_KERNEL=2.0.0`
  維持)。
- modules/extended/server_feature_config 新規 **0.1.0**: Windows Server の
  役割・役割サービス・機能 (ServerManager `Install-WindowsFeature`) の
  インストールを CSV マニフェストから一括制御するモジュール。online
  (稼働中 Windows Server) のみを扱い、offline VHD / ConfigurationFilePath
  経路は scope 外。クライアント OS 用の windows_feature_config と pair。
    - 配置: `modules/extended/`。クライアント SKU では ServerManager 非搭載
      のため、ProductType=1 検出で Status=Skipped 返却 (混在フリート profile
      でも安全に通過)
    - CSV スキーマ: `Enabled` / `Name` / `IncludeAllSubFeature` /
      `IncludeManagementTools` / `Source` / `Description` / `Segment`
    - **`-IncludeManagementTools` 明示制御**: `Install-WindowsFeature` cmdlet
      は default では管理ツール非導入 (Server Manager GUI と挙動異)、CSV で
      明示的に 1 を指定する運用を推奨
    - `Source` 列は windows_feature_config と同形式の 5 形式受容: ドライブ
      レター絶対 / UNC / leading `/` モジュール相対 / leading `\` /
      separator なし。空欄は `-Source` 不渡し = ローカルストア + WU 経路
    - Step 3 dry-run で `Get-WindowsFeature` による InstallState 読み + Source
      事前解決を行い、`[INSTALL]` / `[ALREADY-INSTALLED]` /
      `[PENDING-RESTART]` / `[NEEDS SOURCE OR WU]` / `[NOT FOUND]` の
      5 マーカーで状態を可視化
    - Step 5.5 Post-Apply Verification: `Get-WindowsFeature` で InstallState
      読み戻し、`{Installed, InstallPending}` を PASS、それ以外を FAIL と
      して `-Verified` 付きで `New-BatchResult` を返却 (hostname_config /
      windows_feature_config と同じ pending=受理済み方針)
    - 再起動は profile 側 `__RESTART__` の責務。本モジュールは
      `Install-WindowsFeature` に `-Restart` を渡さず、`RestartNeeded=Yes`
      戻り値は `Show-Info` ログ + 集計サマリの `MessageSuffix` で通知
    - cmdlet 戻り値の `.Success=$false` も failure 扱いで集計 (cmdlet 自体は
      throw せず失敗を返すケースに対応)
    - source-not-found 系エラー (`0x800F081F` / `0x800F0954` /
      "source files could not be found") を検出して、Server ISO build 整合
      ヒントを表示 (英語 + 日本語両方の文言にマッチ)
    - preset.csv で `Name` 列を curated 32 機能 (役割 16 + 機能 8 + RSAT 8) +
      日本語ラベルでドロップダウン化、`Enabled` /
      `IncludeAllSubFeature` / `IncludeManagementTools` も列挙化。Studio 側で
      operator が機能名を暗記しなくても選択可能
    - 公開 API への影響なし、kernel 改修不要 (`REQUIRES_KERNEL=2.0.0`)
- modules/standard/windows_feature_config 新規 **0.1.0**: Windows Optional
  Features (DISM `*-WindowsOptionalFeature` 系) の有効化／無効化を CSV
  マニフェストから一括制御するモジュール。online (稼働中 OS) のみを
  扱い、offline WIM 注入は scope 外 (tonebender 側責務として分離)。
    - CSV スキーマ: `Enabled` / `Action` (Enable/Disable) / `FeatureName`
      / `IncludeAllSubFeatures` / `Source` / `LimitAccess` /
      `Description` / `Segment`
    - **NetFx3 (DisabledWithPayloadRemoved default) 対応**: `Source` 列で
      sxs パスを指定。`LimitAccess=1` (default) で WU フォールバック禁止
      (閉域安全側)、`LimitAccess=0` で WSUS / WU フォールバック許可
    - `Source` 列は **5 形式受容**: ドライブレター絶対 (`D:\sources\sxs`)
      / UNC (`\\nas\share\sxs`) / leading `/` モジュール相対
      (`/payload/dotnetfx35`) / leading `\` (`\payload\dotnetfx35`) /
      separator なし (`payload\dotnetfx35`)。空欄は `-Source` 不渡し =
      Windows Update 経路
    - **module 内 `payload/dotnetfx35/` を fabriq 標準 staging 先として
      推奨**。robocopy /E オーバーレイは site staging を保護
      (`framework_overlay_rules.json` 不変 / `KERNEL_API.md §9` 不変)
    - Step 3 dry-run で `Test-Path` + `*.cab` 存在チェックを行い、
      `[NEEDS SOURCE OR WU]` / `[BAD SOURCE]` / `[NOT FOUND]` /
      `[Current]` / `[APPLY]` / `[INVALID]` の 6 マーカーで状態を可視化
    - Step 5.5 Post-Apply Verification: `Get-WindowsOptionalFeature` で
      State 読み戻し。Enable は `{Enabled, EnablePending}`、Disable は
      `{Disabled, DisablePending, DisabledWithPayloadRemoved}` を PASS
      とし、`-Verified` 付きで `New-BatchResult` を返却 (hostname_config と
      同じ pending=受理済み方針)
    - 再起動は profile 側 `__RESTART__` の責務。本モジュールは常に
      `-NoRestart` を付与し、`RestartNeeded=True` 戻り値はログ表示のみ
    - Source 不存在 / payload removed + Source 空のとき DISM の
      `0x800F081F` / `0x800F0954` 系エラーを検出して、build 整合と
      WSUS バイパス (`RepairContentServerSource=2`) のヒントを表示
    - 公開 API への影響なし、kernel 改修不要 (`REQUIRES_KERNEL=2.0.0`)
- modules/extended/pianist 1.5.0 → **1.6.0**: Run 中の制御ボタン 3 種を追加
  (Stop / Pause / Speed)。既存の実行モデル (Run Phase = 全 Step を順次
  実行) はそのまま維持し、走行中に operator が介入できる手段を増やす。
    - **Stop**: 緊急停止。次の安全な境界 (Step 終了 / Wait 中 / WaitWin
      polling 中) で `Invoke-PianistPhase` の foreach を break。走行中の
      SendKeys / Open は完了させてから停止する (PowerShell からは中断
      不可のため意図的にこの仕様)。停止後の Phase は Auto=Error、log に
      "Stopped by operator before Step N/M" が残り、Manual 判定は通常
      通り operator が `[Phase Status...]` で記録
    - **Pause**: 一時停止/再開の toggle。走行中の Step / Wait が完了
      した時点で hold、もう一度押すと続きから再開。Pause 中は Auto
      バッジが "Paused" 表示 (orange) に切替 — 内部 enum (`Running`)
      は変えず、表示専用の上書き
    - **Speed**: 1.0x ⇔ 1.5x の toggle。`procedure.csv` の Wait 値・
      WaitWin の timeout・Step 後 Wait の 3 箇所に倍率を掛ける。
      `Invoke-PianistAppFocus` / `Invoke-PianistPaste` 内の固定
      settle delay (200-300ms) は機能上の最小値なので scale 対象外
    - 新設 `Wait-PianistResponsive`: blocking `Start-Sleep` を 50ms
      chunk + DoEvents + flag check に置換した chunked sleep。Stop/Pause
      flag が立っていない時の挙動は元の `Start-Sleep` と等価 (50ms
      granularity だけが差分、人間が知覚しない)
    - 新設 `Get-ScaledWaitMs`: `[int]([Math]::Round($Ms * $slowFactor))`
      の薄いラッパー。slowFactor=1.0 の時は no-op
    - 新設 `Update-RunControlButtons`: Stop/Pause は `$script:isRunning`
      の時だけ enable、Speed は常時 enable。Pause ボタンのテキスト/色は
      `pauseRequested` flag に応じて Pause(orange) ⇔ Resume(blue) を切替
    - `Set-AllControlsEnabled` のスコープを既存 5 ボタン (RunPhase /
      Screenshot / PhaseStatus / Prev / Next) に限定し、新 3 ボタンは
      独立した enable rule に分離。これにより「実行中は触らせない」現行
      モデルを維持しつつ、Stop/Pause だけは実行中に押せる
    - `Invoke-PianistPhase` 冒頭で stopRequested/pauseRequested を
      `$false` にリセット (前回 abort 時の flag が漏れないよう defensive)
    - `Invoke-PianistWaitWin`: polling ループ冒頭で stop/pause check、
      pause 中は timeout 消費しない (intuitive: pause 中の経過時間は
      operator の都合なのでカウントしない)
    - 新規モジュール状態: `$script:stopRequested` / `$script:pauseRequested`
      / `$script:slowFactor` (デフォルトはすべて no-op 値)
    - UI レイアウト: action button row (Y=492) の右側に 3 ボタンを追加
      (X=620 / X=740 / X=860、各 Anchor=Bottom,Left)。フォーム本体
      サイズ・既存ボタン位置は変更なし
    - mouse-only ポリシー厳守 (memory: feedback_pianist_mouse_only.md)、
      キーボード accel は持たせない
    - 既存 profile (procedure.csv / values.csv / instructions/) は
      バイト一致で動作。flag=false / slowFactor=1.0 のデフォルト経路は
      v1.5.0 と挙動同一
    - 公開 API への影響なし、kernel 改修不要
- modules/extended/pianist 1.4.0 → **1.5.0**: 「簡易 RPA + 手順書ハイブリッド」
  進化計画の **Phase C**。見本画像表示機能を追加し、3 タブ構成
  (Procedure / Samples / Values) が完成。
    - Phase view に **[Samples] タブ**を新設 (Procedure と Values の間)。
      タブ名は "Screenshots" を避けて **Samples** とすることで、author
      提供の見本画像と operator が撮るエビデンス (procedure.csv の
      Screenshot 列 / Capture-ScreenEvidence 出力) を意味的に分離
    - instructions/<PhaseID>.txt の `[Samples]` section に列挙された
      画像ファイルを `<profile>/screenshots/` から読み込み、サムネイル
      (300×220 のカード型 Panel、PictureBox + caption Label) を
      `FlowLayoutPanel` LeftToRight + WrapContents で並べる
    - サムネイルクリックで `Show-PianistImageViewer` を起動。
      `PictureBox.SizeMode=Zoom` で原寸ズーム表示、リサイズ可。
      **モードレス**で開くため、ビューワを開いたまま Pianist 本体の
      Run Phase / Phase 移動 / Copy Values 等が継続操作可能。
      `Owner = main form` 関係で main form の上に常時 float +
      Pianist 終了時に自動 close
    - 複数枚の見本画像を同時に開いて並べて参照する運用も可
    - 画像読込は `[System.IO.File]::ReadAllBytes` → `MemoryStream` →
      `Image.FromStream` 経由でファイルロックを残さない方式
    - 画像欠損時は `(missing) <filename>` placeholder + dim 色で表示
    - `[Samples]` セクションが空または無い場合は "no entries"
      ヒントメッセージで誘導
    - タブ見出し "Samples (N)" の N は当該 Phase の見本画像数
    - パネル更新時に古い PictureBox.Image を Dispose() してから Clear、
      ファイルハンドルリーク防止
    - `Show-PianistImageViewer` / `New-PianistScreenshotThumbnail` /
      `Update-PianistScreenshotsPanel` 新設
    - サンプル profile の screenshots/ 配下バイナリは非含 (各案件で
      画像を配置する想定)
    - **procedure.csv の `Screenshot` 列を撤去** (v1.0 設計で実装が
      間に合わず vestigial 化していた optional 列、コードからの参照
      ゼロ)。Phase 単位の見本画像参照は新設の `[Samples]` section に
      統一。サンプル profile 2 件の procedure.csv を 9 列 → 8 列へ
      更新。後方互換シムは設けないが、既存 profile に Screenshot 列
      が残っていても Pianist は無視するため runtime での実害なし
    - 公開 API への影響なし、kernel 改修不要
- modules/extended/pianist 1.3.0 → **1.4.0**: 「簡易 RPA + 手順書ハイブリッド」
  進化計画の **Phase B**。Pianist を業務手順書ビューア + 部分的 RPA
  ランナーへ進化させる主要刷新。
    - **instructions/<PhaseID>.txt の section marker** 導入:
      - `[RPA]` Run Phase で自動実行される操作の説明
      - `[Manual]` operator が目視 / クリック / 確認で実施する手順
      - `[Variables]` この Phase で Copy Values に出したい変数を明示宣言
        (procedure.csv の `$VarName` 自動抽出と union される)
      - `[Samples]` 見本画像参照 (parser のみ実装、表示は次版で)
      - 後方互換: marker のないプレーンテキストファイルは全文を Manual として表示
      - `[RPA]` が無い場合は procedure.csv の Step 一覧を fallback 表示
    - **Phase view を TabControl 化** ([Procedure] / [Values] の 2 タブ):
      - [Procedure] タブ: 上段 "RPA" (auto-executed by Run Phase) +
        下段 "Manual" (performed by operator) の 2 段組テキスト表示
      - [Values] タブ: Show-all トグル + Copy Values 行の inline 表示。
        v1.3.0 のモーダルダイアログを廃し、Phase view 内で完結
      - タブ見出し "Values (N)" の N は参照変数の数を反映
    - **アクションボタン行から `[Copy Values...]` を撤去**: Values タブが
      代替するため不要に
    - **Steps preview ListBox を撤去**: [RPA] section の operator 視点
      テキストが代替し、技術的な Step 列挙は不要に。Step 一覧の確認は
      [RPA] section 未記載時の fallback 表示として保持
    - **Show-PianistVariablesDialog 削除**: Values タブで完全代替
    - **`Parse-PianistInstructionFile` 新設**: section marker 対応の
      パーサ。lenient parsing (section marker 前の text は Manual に
      append、未知 section は無視)
    - **Get-PhaseReferencedVariables 拡張**: `[Variables]` section 由来
      の宣言と auto-discovered の union を返す
    - サンプル profile 2 件 (notepad / kintone) を新形式に移行
    - 公開 API への影響なし、kernel 改修不要
- modules/extended/pianist 1.2.0 → **1.3.0**: Copy Values ダイアログに
  **Show all values トグル**を追加（Phase A 拡張）。
    - ダイアログ上部の `[Show all values for this PC]` チェックボックス
      を ON にすると、procedure.csv で参照されていない values.csv 全列も
      含めて全変数を列挙
    - Step では使わないが「電話で読み上げる」「別アプリで paste したい」
      といった用途のコピペ専用変数も values.csv に列を追加するだけで
      ダイアログから取り出せる
    - Default OFF（既存挙動：Phase 参照変数のみ）
    - 副次修正: 2 行目以降の名前ラベルが描画されない問題を修正。原因は
      Label の AutoSize（default $true）と絶対座標 Panel の組合わせ。
      defensive な AutoSize=$false 指定でも一部環境で再現したため、
      ダイアログ内側を **FlowLayoutPanel + 行 Panel container 方式** に
      refactor（`New-PianistVariableRow` ヘルパ新設、各行を self-contained
      Panel として構築）。絶対 Y 座標管理を撤廃し、FlowLayoutPanel が
      自動 stack するので AutoSize / 描画タイミングの影響を受けない
    - `[Copy Values...]` ボタンは N=0 でも openable に変更（Show all
      使用のため）。プロファイルが values.csv に何も持たない場合のみ
      disabled
    - `Get-AllProfileVariables` / `Update-PianistVariablesPanel` 新設
- modules/extended/pianist 1.1.1 → **1.2.0**: 各 Phase に **Copy Values
  ダイアログ**を追加（Phase A: 「簡易 RPA + 手順書ハイブリッド」進化計画
  の第 1 段階）。
    - Phase view のアクションボタン行に `[Copy Values (N)...]` ボタンを
      新設。N は当該 Phase の procedure 行で参照されている `$VarName` の
      数（自動抽出、0 件なら disabled）
    - クリックで Copy Values ダイアログがモーダル表示。各変数行に変数名 +
      解決値（`ENC:` セルは透過復号済み平文）+ `[Copy]` ボタン
    - `[Copy]` 押下で `[System.Windows.Forms.Clipboard]::SetText()` に
      値を転送。operator は target アプリへそのまま paste できる
    - 変数の収集は procedure.csv の現 Phase 行の `Value` / `Note` 列から
      regex `\$([A-Za-z_][A-Za-z0-9_]*)` で抽出 → values.csv 由来の
      ValuesDict から解決
    - values.csv に存在しない変数は `(undefined - not in values.csv)`
      で表示し Copy ボタン disabled
    - `Get-PhaseReferencedVariables` / `Show-PianistVariablesDialog` 新設
    - 既存 profile / Studio / kernel API への影響なし
- modules/extended/pianist/ (new): **Pianist** v1.0.0 — autokey_template
  の進化版として GUI 設定作業を Profile 単位で実行する extended モジュール。
  業務アプリ等の GUI 操作が必須な設定作業を、Phase x Step マトリクスの
  procedure と Values / Shortcuts プールに分けて記述し、operator が
  Phase ごとにマウスクリックで実行・判定する。
    - **Profile 構造**: `profiles/<profile_name>/` 配下に
      `pianist.json` (メタデータ) / `procedure.csv` (Phase x Step) /
      `values.csv` ($VarName 展開・ENC: 対応) / `shortcuts.csv` /
      `instructions/<PhaseID>.txt` (Phase 手順テキスト) /
      `screenshots/` (Screenshot Step 用)
    - **アクション 10 種**: Open / WaitWin / AppFocus / Type / Key /
      Wait / Copy / Paste / Screenshot / Prompt
    - **Phase ヘッダー色分け 9 色**: Blue / Green / Yellow / Orange /
      Red / Purple / Cyan / Pink / Gray
    - **左右 `<`/`>` ナビゲーション**で Phase 間を行き来。最後の
      Phase で `>` が緑の `Done` ボタンに変化、押下で確認 → ModuleResult
      を return
    - **二軸ステータス**: Auto (Run Phase の自動結果) と Manual
      (operator が `[Phase Status...]` で記録した OK/Warning/Error/Skip)
      を独立して保持。Module 終了時に Manual を集計して
      Status (Success/Partial/Error) と Verified (true/false/null) を
      決定し `New-ModuleResult` で kernel に返却
    - **kernel 連携**: ModuleResult を return することで kernel runner が
      `Write-ExecutionHistory` を呼んで `logs/history/execution_history.csv`
      へ記録。SessionID / KanriNo / PCName / WorkerName が自動補完
      される。`Capture-ScreenEvidence` も kernel が前後で auto-capture。
      → 通常のモジュールと同じく HTML チェックリスト / evidence summary
        に Pianist 実行結果が並ぶ
    - **Win32 EnumWindows finder**: ダイアログウィンドウ (Run dialog や
      Notepad の Save As 等) を `Get-Process | Where MainWindowTitle`
      では検出できないため、P/Invoke で `EnumWindows` + `GetWindowTextW` +
      `SetForegroundWindow` を呼んで全トップレベルウィンドウを総当たり
      する `PianistWin32` クラスを実装
    - **UI ポリシー: マウス操作のみ**。キーボードショートカットは
      意図的に持たない。理由は SendKeys との混線防止 / Run 中 race 防止 /
      迷子キー対策 (memory: `feedback_pianist_mouse_only.md`)
    - **レイアウト**: `manual_kitting_assistant` 流の **絶対座標 +
      Anchor** で組む (Dock 入れ子レイアウトの Z-order 問題を回避)
    - **pianist_list.csv**: 利用可能 profile を `Enabled / ProfileName /
      Group / Description / Segment` で記述。Profile CSV 経由実行時は
      Segment フィルタで自動選択、単発実行時は operator が dropdown で
      選択
    - master passphrase は fabriq 本体と共有し、`Unprotect-FabriqValue`
      で values.csv の ENC: 値を透過復号
    - `instructions/<PhaseID>.txt` の改行は LF / CRLF / 混在いずれも
      WinForms TextBox 用に CRLF へ正規化
    - サンプル profile 同梱:
      - `notepad_memo_to_desktop/`: Win+R 相当の Run ダイアログ
        (`shell:::{2559a1f3-21d7-11d4-bdaf-00c04f60b9f0}`) -> notepad
        起動 -> メモ Paste -> Ctrl+S -> `%USERPROFILE%\Desktop\fabriq_memo.txt`
        へ保存 -> Alt+F4 終了 の 5 Phase 構成 (動作確認用)
      - `example_kintone_admin/`: Kintone 管理画面ログイン -> 初期パス
        ワード変更 -> テナント確認 -> ログアウト の 5 Phase 構成 (ひな形)

### Removed
- apps/pianist/ : v0.1.0 〜 v0.2.x の app 版を削除。release されない
  まま module 化に統合された (本 Unreleased 範囲内の作業)。理由:
    - app から `Write-ExecutionHistory` を呼ぶと PowerShell の動的
      `$script:` スコープ規則により kernel 側の `$script:HistoryPath`
      が解決されず履歴記録に失敗
    - SessionID / WorkerName / KanriNo 等のセッション情報が app scope
      からは見えず、自前で履歴行を組み立てる必要があり整合性が崩れる
    - WinForms event handler scope 内で `$PSScriptRoot` が null に
      なる挙動でパス解決が壊れる
    - エラーハンドリング (`Invoke-SafeCommand`) や evidence
      自動取得 (`Capture-ScreenEvidence`) を kernel から受けられない
  → 「実行に専念する operator-driven GUI」という性格は
    `manual_kitting_assistant` の延長線上であり、extended モジュール
    として配置するのが fabriq の設計哲学に整合。

### Changed
- kernel + apps/fabriq_operator: **`FrexProfile` → `FlexProfile` 改名**
  （コードネーム誤記訂正、由来は "Flexible Profile"）。後方互換は意図的に
  非対応（"後方互換は不要" との明示指示）。公開 API サーフェス（KERNEL_API.md
  §1〜§5）への影響はゼロ。touched は内部実装（§6）に区分される識別子・
  値・UI 文字列のみ:
    - apps/fabriq_operator/lib/`frex_dashboard.ps1` → `flex_dashboard.ps1`
      （ファイルリネーム、`fabriq_operator.ps1` の dot-source も追従）
    - 関数: `Show-FrexDashboard` → `Show-FlexDashboard`,
      `Invoke-FrexProfileLoop` → `Invoke-FlexProfileLoop`
    - 文字列値: dashboard 戻り `Action="FrexProfile"` → `"FlexProfile"`,
      `ValidateSet('Linear','Frex')` → `ValidateSet('Linear','Flex')`,
      `resume_state.json` の `ExecutionMode='Frex'` → `'Flex'`
    - 変数: `$isFrexResuming` / `$frexAutoContinue` / `$btnExecFrex` /
      `$frex` / `$frexState` / `$frexResolved` / `$frexRemaining` 等の
      内部ローカル変数すべて
    - UI: フォームタイトル "FrexProfile: ..." → "FlexProfile: ..."、
      Profiles タブのボタン "Execute (Frex)" → "Execute (Flex)"、
      MessageBox タイトル群、console 出力 "FrexProfile Resume Detected"
      等すべて
    - ドキュメント: README.md / KERNEL_API.md (§2 / §4 / §8 内の解説文)、
      modules/extended/pianist/Guide.txt の言及
  CHANGELOG の過去エントリ（kernel 3.1.0〜3.2.2 の "FrexProfile" 表記）は
  歴史記録としてそのまま保持（Keep a Changelog の精神に沿い、過去事実は
  改変しない）。
- modules/extended/pianist/pianist.ps1: 最後の Phase で `Done` ボタン
  のレイアウトが崩れる問題を修正。Anchor=Top,Right,Bottom のボタン
  に `Width=110` を直接代入しても anchor 計算が右側固定を優先して
  Width が反映されず 60px のままで、font 14pt bold の "Done" が
  "Don" / "e" に縦折れしていた。Width=60 を維持して font を 12pt
  bold に縮める方式に変更。テキストが 1 行に収まる。
- modules/extended/pianist/pianist.ps1: **Manual ステータスを設定
  しないと次の Phase へ進めない仕様**を追加。`Update-NavButtons`
  で current phase の `Manual` が `Unset` の間は `>` / `Done`
  ボタンを `Enabled = $false` に。Phase Status ダイアログを開いて
  OK / Warning / Error / Skip のいずれかを選ぶと unlock される。
  status row に hint label "Set Phase Status to advance" を追加し、
  operator にブロック理由を可視化。これで Phase スキップが事故で
  起きるのを防ぎ、毎 phase に operator 判定が確実に付く運用に。
  `<` (戻る) は manual に依存せず常に有効。
- modules/extended/pianist 1.0.0 → **1.1.0**: `values.csv` を **wide
  format** へ刷新（後方互換あり）。
    - 旧スキーマ (`Key,Value,Encrypted,Note`) から
      `NewPCName,<Var1>,<Var2>,...,<VarN>` の hostlist.csv 流横持ち
      に変更。1 行 = 1 ホスト、列 = 変数名で、案件ごとに必要な変数
      プールを Studio から自由に増減できる
    - `NewPCName='*'` (or 空欄) 行が全ホスト共通デフォルト。
      `$env:SELECTED_NEW_PCNAME` 一致行が列ごとに上書き、空セルは
      default 行へフォールバック
    - 暗号化は `ENC:<Base64>` インライン prefix に統一（`Encrypted`
      列は廃止）。fabriq 全体の暗号化規約 (hostlist.csv /
      `Unprotect-FabriqValue`) と整合
    - 列名は `[A-Za-z_][A-Za-z0-9_]*` のみ許可、`NewPCName` は予約
    - 旧 long format の values.csv も自動判別して読み込み可
      （`Build-PianistValuesDict` がヘッダーで分岐、Encrypted=1 の
      旧来挙動を bit-for-bit で保持）。新規 profile は wide format
      で書く
    - サンプル profile 2 件 (`notepad_memo_to_desktop` /
      `example_kintone_admin`) を新スキーマに移行
    - 公開 API への影響なし（kernel 側は `SELECTED_NEW_PCNAME` /
      `Unprotect-FabriqValue` / `$global:FabriqMasterPassphrase` の
      既存契約を消費するのみ）

### Fixed
- modules/extended/pianist 1.1.0 → **1.1.1**: `Invoke-PianistOpen` が
  引用符なしでスペースを含む実行ファイルパス（例: `C:\Program Files\
  Foo\bar.exe`）を渡されると、先頭スペースで split して `C:\Program`
  を起動しようとし失敗していた問題を修正。
    - 修正後は `Test-Path -LiteralPath $Value -PathType Leaf` で全体が
      実在ファイルを指す場合に full path として扱い、引用符なしのスペース
      入りパスを引数なしで起動できるようになった
    - 引用符付き（`"path" args` 形式）の解析も明示化（先頭 `"` から対応
      する閉じ `"` までを FilePath、残りを ArgumentList として剥がす）
    - **引用符なしパス + 引数あり** のケース（例:
      `C:\Program Files\app.exe /flag1`）は引き続き Windows 標準慣習
      どおり引用符必須（破壊的変更を避けるため意図的にスコープ外）
    - URL / `ms-settings:` / `shell:` 系の分岐は従来通り維持
    - 既存の `cmd /c start` フォールバックも維持（PATH 解決・環境変数
      展開の救済として）
    - Guide.txt の Open アクション節に引用符ルールを追記
    - 公開 API への影響なし、後方互換完全維持

### Fixed
- modules/standard/driver_config **1.0.0 → 1.0.1**: Windows Server 2022 上で
  `Export-WindowsDriver -Online` cmdlet が「SafeHandle を Null にすることは
  できません」エラーで失敗していた問題を修正。
    - 根本原因: DISM の PowerShell ラッパー (`Microsoft.Dism.Powershell`)
      が Server 2022 / 2025 SKU で online image の SafeHandle 取得に失敗する
      既知の不具合（Microsoft 側未修正）。Windows 10 / 11 / Server 2019 では
      再現しない
    - 修正方針: cmdlet を `dism.exe /online /export-driver
      /destination:"<path>"` の直接呼び出しに置換。両者は同一の `DismApi.dll`
      を経由するため出力フォルダ構造 (`oem<N>.inf_<arch>_<hash>/`) は
      byte-for-byte 等価。Windows 10 / 11 / Server 2019 環境でのリグレッション
      なし
    - エラー検知: `$LASTEXITCODE -ne 0` で throw、既存 try/catch に流入する
      既存パターンを維持。dism.exe 出力の最終非空行を例外メッセージに含めて
      ログ追跡性を確保。import 側 (`pnputil` 直叩き) と同じ native exe 直接
      呼び出しスタイルで揃った
    - import 側 (`driver_import_config.ps1`) は `pnputil /add-driver
      "<path>\*.inf" /subdirs /install` で `oem*.inf_*/` を再帰スキャン
      するため、出力構造が cmdlet 版と等価である本修正の影響をゼロで吸収
    - Guide.txt の「使用コマンド」「前提条件」「Post-Apply Verification」節を
      更新。Server 2022 互換性に関する注記を追記
    - 公開 API への影響なし、後方互換完全維持
- modules/standard/local_user_config **1.0.0 → 1.1.0**: Profile で
  Segment 別運用（例: `actio.csv` で Order=90 が `Segment=create`、
  Order=130 が `Segment=delete`）を行ったとき、`local_user_list.csv`
  に該当 Segment 行が無いと `Failed to load local_user_list.csv` で
  Error 終了していた問題を修正。
    - 根本原因: `Import-ModuleCsv` は Segment フィルタで 0 件になった
      場合 `@()` を返すが、PowerShell の collection auto-unwrap により
      呼び出し側スカラ変数では `$null` に縮退する。これを既存ロジック
      では「ファイル/列ロード失敗」と区別できず一律 Error 終了していた
    - 修正方針: `local_user_list.csv` のロードが `$null` でも、ファイル
      が存在すれば空集合として継続し、`local_user_host_list.csv` 側の
      合成結果で実行内容を決める。ファイル不在のみ Error 終了
    - 影響対象: `local_user_config.ps1` (Create) と
      `local_user_delete.ps1` (Delete) の両方
    - 副次効果: 「PC 固有単独モード」（base CSV を空にして host_list
      側だけで作成/削除する運用）が事実上有効化される。Guide.txt の
      動作モード節に第 3 モードとして追記
    - エラーステータス挙動の変化: 必須列タイポ等の本質的な CSV 破損は
      従来 Error 返却 → 修正後は kernel 側 `Show-Error` ログ出力のうえ
      Skipped 返却に降格。Profile の `ErrorMode=Stop` で停止していた
      ケースで停止しなくなる点に注意（コンソール/ログでは引き続き赤字
      で Error メッセージが表示される）
    - 公開 API への影響なし。kernel 側は無変更

### Notes
- `__PIANIST_to_<profile>__` マーカー追加は不要になった。Pianist は
  通常のモジュールとして Profile CSV の `ScriptPath` 列に
  `extended/pianist/pianist.ps1` を直接書けば呼べる。Segment 列で
  どの profile を実行するか指定する。

## [3.2.2] - 2026-05-02

### Changed
- apps/fabriq_operator/lib/frex_dashboard.ps1: Group 列の薄水色塗り
  （3.2.1 で導入）を撤去。Operator から「他の無着色列との対比で
  悪目立ちして気になる」フィードバックを受けての調整。Group 列は
  既定スタイル（無着色、太字でない）に戻し、ヘッダー "Group" +
  セル内のテキスト表示だけで identification を成立させる方針に。
  Status / Verified の badge 色だけが視覚アクセントとして残るので、
  全体の色数が減って情報の優先順位が明確化される副次効果も。
  KERNEL_API.md §6 内部実装、KERNEL_VERSION 影響なし。

## [3.2.1] - 2026-05-02

### Fixed
- apps/fabriq_operator/lib/frex_dashboard.ps1: Groups バー追加時に
  `[Complete]` ボタンが下端で見切れる問題を修正。3.2.0 のフォーム
  高さ計算式 `$formH = $footerY + 80 + 20` の `+20` がフォーム下部
  chrome（タイトルバー + 枠）に対して不足していた（元 layout は
  `+40` 相当）。`+40` に修正してフォーム下端に十分な余白を確保。
  Groups バー無しの profile も同じ式で 660 高（旧版と同一）に保たれる。

### Added
- apps/fabriq_operator/lib/frex_dashboard.ps1: Grid に **`Group` 列**
  を追加。各行が CSV Profile の `Group` 列値を表示し、operator が
  Groups バーボタンと grid 行の対応関係を一目で把握できる。空欄は
  「グループ無所属」を意味する。
    - 列位置: `[Checked][Order][Group][Module][Status][Verified][Run]`
      （Order 直後）。幅 90px、中央揃え、ReadOnly
    - `CellFormatting` ハンドラに新分岐を追加: Group 値が非空のセル
      には薄い水色 (`RGB(210,230,240)`) の背景塗り。Status / Verified
      の badge 色とは異なる中間色なので視覚的に競合しない
    - 空 Group のセルは default styling（塗りなし）で「ungrouped 行」
      が暗黙的に区別可能
  KERNEL_API.md §6 内部実装、KERNEL_VERSION 影響なし。

## [3.2.0] - 2026-05-02

### Added
- kernel/KERNEL_API.md §4.1: Profile CSV スキーマに **任意列 `Group`**
  を追加（後方互換、列順末尾追加）。同一 `Group` 値の行群を
  FrexProfile dashboard の Groups バー上の `[Run: <Group>]` ボタンで
  1 クリック実行できる。空文字列 / 列欠落は「グループ無所属」を
  意味し、ボタン化されない。
- apps/fabriq_operator/lib/frex_dashboard.ps1: ヘッダー直下に
  **Groups バー**を新設。Profile に少なくとも 1 行 `Group` 値が
  ある場合のみ render（無ければ layout shift 無し）。各 Group ごとに
  `[Run: <Group>]` ボタンを配置（青アクセント色、太字）。クリックで
  当該 Group のモジュール群を即時 batch 実行（`RunGroup` action 経由）。
- kernel/main.ps1: `Invoke-FrexProfileLoop` に新 case `"RunGroup"`
  追加。`RunBatch` と同じ `Invoke-BatchExecution` パイプラインを
  共有（`-AutoPilot:$true` `-FinalizeOnComplete:$false`
  `-ExecutionMode 'Frex'`）、`SelectedOrders` の source が group
  filter になるだけ。完走後 `$pendingFinalize=$true` 設定。
- kernel/common.ps1: `Resolve-ProfileModules` が `Group` 列を読んで
  module オブジェクトに `_Group` プロパティを付与。3 か所
  （`__AUTO_to_<User>__` / 特殊マーカー / 通常モジュール）すべてに
  対応。Linear 経路は `_Group` を参照しないため挙動完全互換。

### Changed
- kernel/KERNEL_VERSION : 3.1.9 → **3.2.0**（**MINOR** bump、Profile CSV
  スキーマへの後方互換な任意列追加）。版表記同期: README.md L1
  `ver3.1` → `ver3.2`、kernel/common.ps1 L2 `v3.1.9` → `v3.2.0`、
  kernel/main.ps1 L3 `ver3.1` → `ver3.2`。
- kernel/KERNEL_API.md §8 に `### 3.2.0` エントリ追加（Group 列の
  契約、literal interpretation 仕様、Linear 不参照を明記）。

### Notes (literal-Group contract)
- Group 内に `__RESTART__` を含めることは可（reboot → resume →
  group の残りモジュールのみ実行）。
- Group 跨ぎの `__RESTART__`（Group 値が空 or 異なる）は当該 Group
  実行時に **skip** される。Operator が RESTART を group 実行に
  含めたい場合は明示的に Group 値を打つ必要がある。

### Existing flow compatibility
- Linear `[Execute Profile]` : 完全不変
- Frex `[Run Selected]` : 完全不変
- Frex 行ごと `[Run]` : 完全不変
- AutoPilot resume after `__RESTART__` : `SelectedOrders` 機構を共有、
  group 実行中に RESTART が発火しても resume は group の残り orders
  のみ自動継続
- 旧 fabriq が新 Profile CSV を読む : `Group` 列は header-driven
  Import-Csv で無視される（旧挙動維持）

## [3.1.9] - 2026-05-02

### Changed
- apps/fabriq_operator/lib/frex_dashboard.ps1: Status / Verified
  セルの視覚表現を **文字色** から **背景塗りつぶしバッジ** に
  変更。視認性を大幅に強化。
    - Success → 緑背景（`$script:bgAdd`）+ 白文字
    - Partial → 黄背景（`$script:stripeYellow`）+ 暗文字
    - Error   → 赤背景（`$script:bgDelete`）+ 白文字
    - Skipped / Cancelled → 中灰背景 + 白文字
    - Pending → 薄灰背景 + 暗灰文字
    - PASS / FAIL（Verified 列） → 緑 / 赤 + 白文字
  実装: `CellFormatting` ハンドラで `$e.CellStyle.BackColor` /
  `SelectionBackColor` / `ForeColor` / `SelectionForeColor` /
  `Font = $script:fontBold` を一括設定。`SelectionBackColor` を
  `BackColor` と同色にすることで行選択時もバッジが消えない。
  既存テーマ色（`$script:bgAdd` / `$script:bgDelete` /
  `$script:stripeYellow`）を再利用してテーマ整合性を維持。
  内部 dispatch / state map / per-Order tracking には影響なし、
  純粋に視覚改善。
  KERNEL_API.md §6 内部実装、KERNEL_VERSION 影響なし。

## [3.1.8] - 2026-05-02

### Changed
- apps/fabriq_operator/lib/frex_dashboard.ps1: 単発実行 UI を
  per-row `[Run]` ボタンに移行。各 module 行に `DataGridViewButtonColumn`
  で `[Run]` ボタンを配置し、クリックで該当 row の `RunSingle`
  action を即時 dispatch。footer の `[Run This: M]` ボタンを撤去
  （per-row 化により冗長）。
    - Grid 末尾に `RunBtn` 列追加（W=56、`UseColumnTextForButtonValue`）
    - `CellContentClick` ハンドラで列名 `RunBtn` の場合のみ
      `result.Action = "RunSingle"` をセットして form.Close()
    - 同 row の checkbox 列クリックは既存 `CellValueChanged` /
      `CurrentCellDirtyStateChanged` で処理されるため干渉なし
    - `[Run This: M]` ボタン関連コード（btnRunThis 宣言・Click
      ハンドラ・updateRunThisLabel scriptblock・SelectionChanged
      ハンドラ・初期同期呼び出し）を完全撤去
    - footer Row 1 のレイアウト調整: `[Run Selected (N)]` を
      W=160 → W=340 に拡張して空きを埋め、Row 2 の `[Complete]`
      (W=340) と視覚的に整合
- 内部挙動の互換性: per-row Run も legacy `[Run This]` も同じ
  `RunSingle` action を `Invoke-FrexProfileLoop` に渡す → 同じ
  `Invoke-BatchExecution` 呼び出しに帰着。AutoConfirmMode、
  per-Order tracking、PendingFinalize 状態管理など下層ロジックは
  全て unchanged。UI 層の入れ替えのみ。
- マーカー行（[RESTART] / [REEXPLORER]）も同様に per-row `[Run]`
  ボタンが押せる（Profile-internal RESTART を ad-hoc 実行する操作が
  直感的に）。
  KERNEL_API.md §6 内部実装、KERNEL_VERSION 影響なし。

## [3.1.7] - 2026-05-02

### Fixed
- apps/fabriq_operator/lib/frex_dashboard.ps1 + kernel/common.ps1
  (Export-HtmlChecklist): 同じ MenuName を持つ Profile 行が複数
  ある場合（例: `_test_harness.csv` の Order 80 と Order 110 が
  両方 `[seg:error_basic]`）、片方だけ実行すると未実行側にも
  実行側のステータスがミラーされる残バグを修正。
  3.1.3 で `Order` 列を一級識別子として導入したが、Order 一致
  ヒットが無いときの **MenuName fallback が候補の Order を見ず
  に flat に照合**していた。結果、未実行 row が兄弟 row の
  実行結果を流用してしまう状態だった。
  修正: state map / HTML checklist の MenuName fallback を
  STRICT 化。候補エントリの Order が
    (a) 0 = legacy / 非 Profile → 採用（旧 history.csv との互換）
    (b) row の Order と一致 → 採用（防御的）
    (c) 別の非ゼロ Order → **不採用**（兄弟 row のエントリと
        判断、自分は Pending のまま維持）
  これにより同名 MenuName の複数 row が真に独立した状態を
  持てる。Order 列を持たない旧 history.csv との互換性は維持。
  KERNEL_API.md §6 内部実装、KERNEL_VERSION 影響なし。

## [3.1.6] - 2026-05-02

### Fixed
- apps/fabriq_operator/lib/frex_dashboard.ps1: 初期状態（全 Pending /
  全 unchecked）で `[Complete]` ボタンが緑表示になり、operator が
  実行ゼロのまま finalize できる misleading な挙動を修正。
  根本原因: 3.1.4 のルール "Pending は checked のみカウント" が、
  「初期状態 = 全 unchecked = issue ゼロ = 緑 Complete」を許容して
  しまっていた。`Select All` 直後は警告に切り替わるが、押さずに
  Complete を押すと空の HTML チェックリストを finalize 可能だった。
  修正: `$updateCounters` と `btnComplete.Add_Click` の両方に
  `$hasExecuted` フラグ判定を追加（「Status が Pending 以外の行が
  1 つでもあるか」）。`$issueCount=0` でも `$hasExecuted=$false` の
  場合は黄色 "Complete (nothing executed)" を表示。確認ダイアログも
  "No modules have been executed in this session yet." を冒頭に
  含める。`$issueCount>0` と同時発生時は両方の警告を併記。
  KERNEL_API.md §6 内部実装、KERNEL_VERSION 影響なし。

## [3.1.5] - 2026-05-02

### Changed
- apps/fabriq_operator/lib/frex_dashboard.ps1 + kernel/main.ps1:
  FrexProfile の実行モデルを「実行 = 常に AutoPilot 挙動 / 完了 =
  常に手動」に simplify。AutoPilot toggle を撤去し、operator が
  AutoPilot か否か判断する余地を取り除いて、checkbox による
  module 選択だけで意思決定が完結する形に統一。
    - Dashboard footer の `[AutoPilot ☐]` checkbox を撤去
    - 代わりに `[Select All]` / `[Clear All]` の bulk-select 2 ボタン
      を追加。`[Select All]` は CSV `Enabled=1` 行のみ check（CSV
      作成者の opt-in 意図を尊重して `Enabled=0` 行は unchecked）、
      `[Clear All]` は全 uncheck
    - `[Run Selected]` は `result.AutoPilot = $true` を hardcode、
      `Invoke-FrexProfileLoop` の "RunBatch" ハンドラも
      `-AutoPilot:$true -FinalizeOnComplete:$false` を hardcode して
      条件分岐を撤去
    - Frex resume 実行ブロックの `-FinalizeOnComplete:$true` を
      `:$false` に変更。auto-continue 完走後は dashboard へ復帰
      （operator が `[Complete]` を押下するまで HTML / log_uploader
      は発火しない）
- WaitSec NumericUpDown は据え置き（常時 unattended 実行のため inter-
  module 待機制御として常時必要）。位置は X=255 へ移動

### Added
- apps/fabriq_operator/lib/frex_dashboard.ps1 + kernel/main.ps1:
  「PENDING FINALIZE」状態追跡 UI を追加。`[Run Selected]` /
  `[Run This]` / `[Mark as Pending]` 後 dashboard を再表示すると、
  ヘッダの "Last finalized" 行が **赤バッジ "PENDING FINALIZE"** に
  切り替わる。`[Complete]` 押下で解除されて元の "Last finalized:
  HH:MM:SS" 表示に戻る。
    - `Show-FrexDashboard` に `[bool]$PendingFinalize = $false`
      パラメータ追加
    - `Invoke-FrexProfileLoop` に `[bool]$InitialPendingFinalize =
      $false` パラメータ追加（Frex resume 入口から auto-continue
      経由で開く際に `$true` を渡して、operator が flagged 状態の
      dashboard で復帰するよう制御）
    - Frex dashboard の `[Back]` ボタン / X ボタンクローズ時に
      `$PendingFinalize=$true` の場合は確認ダイアログ表示。No で
      閉鎖キャンセル、Yes で操作続行（成果物未生成のまま離脱）
- 運用注意: 完了 phase が常に手動になるため、operator が
  `[Complete]` を押し忘れると HTML 未生成 / log_uploader 未発火の
  状態で離脱可能。pending badge と close 確認の 2 段階の UI 警告で
  認知支援を提供するが、最終的な press 動作は operator 責務。

KERNEL_API.md §6 内部実装、KERNEL_VERSION 影響なし。

## [3.1.4] - 2026-05-02

### Fixed
- apps/fabriq_operator/lib/frex_dashboard.ps1: Frex dashboard の
  `[Complete]` ボタン警告判定が unchecked 行の Error / Partial を
  見落としていた問題を修正。原因: P11 で「default は全 uncheck」に
  変えた結果、AutoPilot 後の単発再実行で dashboard が再表示される
  と全行 unchecked → checked のみを issue 計数する旧ロジック
  （`if (-not $isChecked) { continue }`）が Error 行を全て無視 →
  ボタンが緑の "Complete" 表示になっていた（GUI 行と HTML には
  正しい Status が出ていたため、ボタンラベルだけ実態と齟齬）。
  修正: Error / Partial は **常にカウント**（実行結果の事実を
  uncheck で隠せない）、Pending は **checked 時のみカウント**
  （operator の "やる" 意思表示）に分離。Skipped / Cancelled は
  従来通り無視。`btnComplete.Add_Click` の確認ダイアログも同じ
  ルールで生成、ダイアログ文言を「checked rows」→「rows」に
  変更。AutoPilot を ON にすると bulk-check で Pending 行が
  追加カウントされて警告が出る挙動も同時に整合。
  KERNEL_API.md §6 内部実装、KERNEL_VERSION 影響なし。

## [3.1.3] - 2026-05-02

### Fixed
- kernel/main.ps1 + kernel/common.ps1: FrexProfile で単発実行
  （[Run This]）を行うと、その後の HTML チェックリスト / dashboard で
  他の Profile エントリが NotRun として表示される問題を修正。
  原因: `Invoke-BatchExecution` 先頭の `Clear-ExecutionResults` が
  `IsRestored=$true` 行のみ保護し、直前バッチで追加された
  非 IsRestored 行を wipe していた。`__RESTART__` を跨いだ Profile
  では「跨いだ前は session-start Restore で IsRestored 化された
  ため生存／跨いだ後は wipe される」非対称が発生。
  対処: `Invoke-BatchExecution` の `Clear-ExecutionResults` 直後に
  `Restore-ExecutionHistory -SessionIDFilter $script:SessionID` を
  呼んで現セッションの履歴を IsRestored 化、wipe を相殺。

### Added
- kernel/common.ps1: `execution_history.csv` に `Order` 列を追加
  （末尾追加、後方互換）。Profile CSV 行の `Order` を執行履歴の
  一級識別子として記録し、同名 MenuName の複数行を per-row で
  区別可能に。
    - `Write-ExecutionHistory` に `-Order [int]` パラメータ追加。
      0 = "Profile 行に紐付かない" (markers / ad-hoc / log uploader)、
      正整数 = Profile CSV 行の Order
    - `Add-ExecutionResult` に `-Order [int]` パラメータ追加、結果
      オブジェクトに保存
    - `Restore-ExecutionHistory` が `Order` 列を解釈して結果オブジェクトに
      復元。旧 CSV（`Order` 列なし）は Order=0 として扱う
    - `Save-ResumeState` の `CompletedModules` reshape に `Order`
      フィールド追加（resume 後の per-row 状態追跡継続性確保）
    - `Show-FrexDashboard` の state map を `(SessionID, Order)`
      タプル照合に切替。Order=0 のエントリは MenuName fallback
    - `Export-HtmlChecklist` の照合を Order ベースに切替、Order=0 は
      MenuName fallback。これにより同名 MenuName の複数 Profile
      行が異なる Status を表示可能（旧来は最終エントリで上書きされ
      まとめ表示されていた pre-existing issue を解消）
    - `Invoke-BatchExecution` の全 Add/Write 呼び出しサイト
      （`__RESTART__` / `__REEXPLORER__` / 通常モジュール）に
      `-Order $module.Order` を配線
    - `Invoke-FrexProfileLoop` の `[Restart Now]` / `[Mark as Pending]`
      ハンドラに `-Order` を配線
    - Linear / Frex resume bootstrap の `CompletedModules` 再生に
      `-Order` を配線

- kernel/common.ps1: `Restore-ExecutionHistory` に
  `-SessionIDFilter [string]` パラメータ追加。指定時は該当
  SessionID のエントリのみ pull、Limit 50 解除、Separator 行
  非生成、結果が 0 件の場合は明示的に空配列で `$script:ExecutionResults`
  を置換（cross-session の IsRestored 残骸を eviction）。
  既存呼び出しは引数省略で legacy 挙動維持（top 50 cross-session、
  Separator 付き）。

### Changed
- kernel/common.ps1: `execution_history.csv` のスキーマが拡張
  （`Order` 列追加）。外部 evidence consumer ツールは Import-Csv
  で読み込めば自動対応。列順依存があるツールのみ要追従（fabriq
  本体・fabriq_evidence_manager・fabriq_studio はいずれも header-driven
  読み込みのため影響なし）。
  KERNEL_API.md §6 内部実装、公開 API 影響なし、KERNEL_VERSION 影響なし。

## [3.1.2] - 2026-05-02

### Changed
- apps/fabriq_operator/lib/frex_dashboard.ps1: Frex dashboard の
  チェック状態既定値と AutoPilot トグル挙動を仕様変更:
    - **既定値**: 全モジュール unchecked（旧: CSV `Enabled` を反映）。
      Frex の哲学（"default mode = pick from blank"）と整合
    - **AutoPilot ON**: Enabled=1 の行のみ bulk-check（CSV 作成者の
      opt-in 意図を尊重して Enabled=0 行は unchecked のまま残す）
    - **AutoPilot OFF**: 副作用なし、現在のチェック状態を保持。
      "AutoPilot で全選択 → AutoPilot を外して通常実行" という
      bulk-select-only ワークフローを支援
  実装: 行追加時の初期 check を `$false` に固定、`row.Tag` に
  `IsCheckedDefault` フィールドを追加して AutoPilot bulk-check が
  CSV `Enabled` 値を参照可能に。`$chkAutoPilot.Add_CheckedChanged`
  ハンドラ新設で bulk-set を実装。`$frexState.BulkUpdating` flag で
  bulk 中の `CellValueChanged` 連発による `updateCounters` 過剰
  呼び出しを抑止（70 モジュール × N 連発の性能影響を防止）。
  KERNEL_API.md §6 内部実装、KERNEL_VERSION 影響なし。

## [3.1.1] - 2026-05-01

### Fixed
- apps/fabriq_operator/lib/frex_dashboard.ps1: Frex dashboard の
  チェックボックス列が編集不能だった問題を修正。`Set-GridStyle` が
  `$grid.ReadOnly = $true` を設定し、WinForms 仕様により per-column の
  `ReadOnly = $false` が無視されていた。`Set-GridStyle` 呼び出し直後に
  `$grid.ReadOnly = $false` でグリッドレベル設定を解除し、Order /
  Module / Status / Verified の各列を個別に `ReadOnly = $true` で
  ロックすることで checkbox 列のみ編集可能化。
- kernel/main.ps1: AutoPilot ON で `__RESTART__` マーカーを含む Frex
  バッチ実行が再起動後にダッシュボードへ戻ってしまい unattended
  契約が破られていた問題を修正。3.1.0 設計時に D5（`[Restart Now]`
  ボタン要件）を `__RESTART__` マーカー全般に過剰適用していた認識
  誤り。Frex resume を 3 経路に分岐:
    - AutoPilot=true + ResumeAfterOrder≥0（mid-batch __RESTART__）
      → countdown + auto-continue 実行（Linear 対称）
    - AutoPilot=false + ResumeAfterOrder≥0（manual mid-batch）
      → dashboard 復帰
    - ResumeAfterOrder=-1（`[Restart Now]` sentinel、任意の AutoPilot）
      → dashboard 復帰
  Auto-continue 経路は SelectedOrders ∩ (Order > ResumeAfterOrder) で
  残モジュールを構築し、`Invoke-BatchExecution -ExecutionMode 'Frex'
  -SelectedOrders ... -FinalizeOnComplete:$true` で完走させる。
  Wait-KeyPress + 完了メッセージは Linear 対称。

## [3.1.0] - 2026-05-01

### Added
- apps/fabriq_operator/lib/dashboard_form.ps1: Profiles タブに
  `[Execute (Frex)]` ボタンを追加（success-green アクセント色、
  `[View Details]` と `[Execute Profile]` の中間に配置）。
  クリックで `$result.Action = "FrexProfile"` を返却し main.ps1 が
  `Invoke-FrexProfileLoop` を起動。AutoPilot は FrexProfile dashboard
  側の checkbox（既定 OFF）が独立に制御するため、main dashboard の
  AutoPilot 設定は Frex 経路に伝播させない。`[View Details]` は
  W=110 → W=92 に縮小、X=378 → X=200 に左シフトしてレイアウト調整。
  `[Execute Profile]` は X=498/W=150 のまま不変（Linear 経路の
  operator muscle memory を保護）。これで P1〜P7 を貫く配線が
  完成し、operator が Profile を選んで `[Execute (Frex)]` を押す
  と FrexProfile dashboard が起動する。

- kernel/main.ps1: `Invoke-FrexProfileLoop` ヘルパー関数を新設。
  FrexProfile dashboard sub-loop を駆動する内部関数で、main loop の
  `"FrexProfile"` action と Frex resume bootstrap の両方から呼び出される
  単一の真実の源。dashboard が返す Action を以下に dispatch する:
  - `RunSingle`  → AutoConfirmMode ON で Invoke-BatchExecution 単発呼び出し
                   （`-FinalizeOnComplete:$false`、`-ExecutionMode 'Frex'`）
  - `RunBatch`   → Invoke-BatchExecution 複数呼び出し。AutoPilot ON 時は
                   finalize 自動発火（Linear 対称）、OFF 時は operator が
                   `[Complete]` で finalize
  - `Complete`   → Complete-ProfileExecution -Mode 'Manual'
  - `RestartNow` → 現在の checked subset と execution_history.csv 由来の
                   ModuleStates を schemaVersion=2 / ExecutionMode='Frex'
                   resume_state.json に保存 → Register-FabriqRunOnce →
                   Invoke-CountdownRestart。`ResumeAfterOrder=-1` sentinel
                   で「dashboard 復帰のみ、auto-continuation なし」を表現
  - `ResetState` → execution_history.csv に Status='Pending' 行追記で
                   特定 Order の状態を Pending に戻す
  - `Close`      → Remove-ResumeState で Frex resume state を消費して終了
- kernel/main.ps1: main loop switch に `"FrexProfile"` action ハンドラを
  追加（Invoke-FrexProfileLoop を呼ぶだけ）。P7 で main dashboard 側に
  ボタンが追加されるまでは dead code として安全待機
- kernel/main.ps1: 起動時 resume 検知ロジックに schemaVersion 分岐追加。
  schemaVersion>=2 + ExecutionMode='Frex' を `$isFrexResuming` として
  検出し、Linear の auto-resume countdown / Y-N プロンプトをスキップして
  常に dashboard 復帰経路へ進む。Linear resume execution block には
  `-and -not $isFrexResuming` ガードを追加して同時起動を排他化
- kernel/main.ps1: Frex resume execution block を追加（Initialize-
  ModuleSystem 後に Invoke-FrexProfileLoop を呼ぶ）。Linear resume
  execution block と並列構造で配置

### Fixed
- kernel/common.ps1: Export-HtmlChecklist の Status switch に明示的な
  `Pending` 分岐を追加。従来は default 分岐に落ちて `notRunTotal` が
  加算されず、HTML フッタの集計合計が `DefinedModules.Count` と
  食い違う pre-existing バグだった（Linear では Pending Status が
  通常生成されないため顕在化しなかったが、FrexProfile の `[Mark as
  Pending]` 操作で表面化）。修正後は Pending 行が `notRunTotal++` に
  寄与し overall 判定 "Incomplete" / Not Run chip / Pending ラベル
  表示が一貫。Linear 既存 HTML 出力には影響なし。
  KERNEL_API.md §6 内部実装、KERNEL_VERSION 影響なし。

### Changed
- kernel/main.ps1: Invoke-BatchExecution に FrexProfile 用の
  pass-through パラメータ 3 つ追加: `-ExecutionMode 'Linear'|'Frex'`,
  `-SelectedOrders int[]`, `-ModuleStates hashtable`。`__RESTART__`
  ハンドラ内の `Save-ResumeState` 呼び出しに透過的に渡される。
  Linear 既存呼び出しは省略時 v1 出力（byte-for-byte 互換）、Frex
  経由で渡されると schemaVersion=2 出力。KERNEL_API.md §6 内部実装、
  公開 API 影響なし、KERNEL_VERSION 影響なし。

- apps/fabriq_operator/lib/frex_dashboard.ps1: FrexProfile dashboard 形式の
  WinForms 新規追加。`Show-FrexDashboard` 関数 1 個を提供。
  - profile CSV を `Resolve-ProfileModules -IncludeDisabled` で全行
    取得（Enabled=0 含む）し、CSV 既定のチェック状態を初期表示
  - 状態列は `execution_history.csv` の現セッション ID フィルタ +
    オプション `$LastBatchResults` 上書きで構築（D2「履歴=真実の源」
    決定の実装）
  - 行ダブルクリック / `[Run This: N]` ボタンで個別実行、
    `[Run Selected (N)]` でチェック分一括、Order 昇順
  - `[Complete]` ボタンは常に押下可能。Error / Partial / Pending が
    checked 行に残ると "Complete with N issues"（黄）に切替、
    押下時に確認ダイアログ。確認後 main.ps1 が
    Complete-ProfileExecution を呼ぶ（P6/P7 で配線）
  - `[Restart Now]` ボタンは Profile 外から再起動を発火（D5 「いつでも
    実行できるリスタートマーカー」要件）。confirm dialog → main.ps1 が
    Save-ResumeState v2 + Register-FabriqRunOnce + Invoke-CountdownRestart
  - 行右クリック → "Mark as Pending (reset state)" で個別行のみ Pending
    リセット要求を返す
  - AutoPilot 既定 OFF（FrexProfile での約束）、WaitSec NumericUpDown
  - 戻り値 hashtable: `{ Action, SelectedOrders, TargetOrder, AutoPilot,
    AutoPilotWaitSec, ResetTargetOrder }`。Action は
    `RunSingle` / `RunBatch` / `Complete` / `RestartNow` / `ResetState` / `Close`
  - 本フェーズでは何も実行 dispatch しない（intent を返却するだけ）。
    実行配線と Frex resume 検知は P6 で main.ps1 に追加予定
- apps/fabriq_operator/fabriq_operator.ps1: lib dot-source に
  frex_dashboard.ps1 を追加（fabriq_operator GUI ロード時に自動展開）

- kernel/common.ps1: Save-ResumeState に schemaVersion=2 書き出し対応
  を追加。新 optional パラメータ `-ExecutionMode 'Linear'|'Frex'`、
  `-SelectedOrders int[]`、`-ModuleStates hashtable` を追加し、
  `-ExecutionMode 'Frex'` のときのみ `schemaVersion` / `ExecutionMode` /
  `SelectedOrders` / `ModuleStates` の v2 拡張フィールドを JSON に
  書き出す。Linear 既存呼び出しは新パラメータを渡さないため出力は
  pre-P4 の v1 と byte-for-byte 互換。Load-ResumeState は P4 では
  不変（v1 / v2 両方が ConvertFrom-Json でそのまま読める。Frex resume
  分岐は P6 で main.ps1 に追加予定）。
  P4 単独では v2 を書き出す呼び出し元が無いため実ファイル生成は
  発生しない（純粋に capability の追加）。
  KERNEL_API.md §6 内部実装、公開 API 影響なし、KERNEL_VERSION 影響なし。
  互換性 stance: アップグレード方向のみ互換（旧 fabriq が v2 ファイルを
  読むときは未知 schemaVersion を検知せず v1 として処理 → 副作用は
  実害なし。逆ダウングレードは未保証）。

- kernel/common.ps1: `$script:LastBatchResults` 配列を新設。
  Invoke-BatchExecution が完走 / cancel / mid-throw のどの経路でも
  finally 経由で publish される per-module 結果スナップショット。
  各エントリは hashtable `{ Order, MenuName, Status, Verified, Message }`。
  FrexProfile dashboard が `execution_history.csv` の SessionID
  全 import を回避して直前 1 バッチの差分だけを即座に拾うための
  揮発チャンネル。`Reset-FabriqState` でクリア。
  KERNEL_API.md §6 内部実装に該当、公開 API 影響なし。

### Changed
- kernel/main.ps1: Invoke-BatchExecution の `$completedResults` を
  hashtable 拡張（Order / Verified / Message を追加）。Save-ResumeState
  は射影で {MenuName, Status} のみシリアライズするため `resume_state.json`
  形式は不変。`__REEXPLORER__` ハンドラの結果記録を try/catch の実結果
  に修正（従来は失敗時も "Success" がハードコードされていた pre-existing
  バグ。resume 表示の Status 文字列のみが影響、ロジックには無影響）。
  `$completedResults` の宣言を try ブロック前に移動して finally から
  参照可能に。挙動互換、KERNEL_VERSION 影響なし。

- kernel/main.ps1: Invoke-BatchExecution に `[bool]$FinalizeOnComplete`
  パラメータを追加（既定 `$true`）。`$true` 時のみループ完走後に
  `Complete-ProfileExecution` を発火させる。Linear 既存呼び出しは
  既定値で従来挙動を維持（呼び出し側無修正）。FrexProfile バッチ
  / 個別実行は `:$false` を渡して finalize を operator 主導の
  `[Complete]` ボタンに委譲する。`__RESTART__` 早期抜けには影響なし
  （resume 時の二脚目で改めて評価される）。
  switch ではなく [bool] を採用してデフォルト $true を許容
  （PSAvoidDefaultValueSwitchParameter 回避、既存 `:$false` 構文は
  両方の型で互換）。KERNEL_API.md §6 内部実装。
  公開 API 影響なし、KERNEL_VERSION 影響なし。

- kernel/common.ps1 + kernel/main.ps1: post-profile pipeline（execution
  history エクスポート → HTML チェックリスト生成 → log_uploader →
  view_report 起動）を `Complete-ProfileExecution` 関数に集約。
  従来 `Invoke-BatchExecution` 末尾と main.ps1 の `RegenerateChecklist`
  action にコピーペーストで重複していた実装を 1 関数に統合し、
  - `-Mode 'Auto'`   = Linear AutoPilot 完走（silent upload / viewer 後）
  - `-Mode 'Manual'` = `[cl]` 再生成（"Log Upload (cl)" 履歴記録 / viewer 前）
  の二経路を ValidateSet で明示。挙動は両経路とも byte-for-byte 互換。
  単一の真実の源を確保することで、FrexProfile `[Complete]` ボタン
  (P5/P7) が同じ関数を呼ぶだけで finalize できる。
  KERNEL_API.md §6 内部実装。公開 API 影響なし、KERNEL_VERSION 影響なし。

### Added
- kernel/common.ps1: `$global:AutoConfirmMode` フラグ導入。
  Confirm-Execution / Wait-KeyPress に短絡条件を追加し、Y/N プロンプトと
  Press-Enter 待機をスキップする AutoPilot のサブセット動作を提供。
  AutoPilot の他の副作用（inter-module wait、ErrorMode 分岐、
  Show-AutoPilotErrorDialog、ErrorMode retry loop）は発火しない。
  FrexProfile の個別モジュール実行で「ボタンクリック1回で実行」を
  実現するための土台。AutoPilot と同時 ON の場合は AutoPilot 表示が
  優先（防御的、実用上両立しない設計）。
  KERNEL_API.md §2 への正式記載は FrexProfile 機能群完成時に MINOR
  bump と同時に行う方針のため、現時点では文書化保留。
  KERNEL_VERSION 影響なし。

- kernel/common.ps1: Resolve-ProfileModules に `-IncludeDisabled` switch
  を追加。指定時は `Enabled=0` 行も返却対象に含め、各エントリへ
  `_IsCheckedDefault` プロパティ（CSV `Enabled` 値の bool 反映）を
  付与する。Linear 経路は switch 省略で完全互換（フィールド非付与）。
  `__AUTOPILOT__` / `__ASYNC__` マーカーの作用は `Enabled=1` 必須に
  内部で縛り、無効化マーカーが global state を flip することを防ぐ。
  FrexProfile dashboard が CSV 既定のチェック状態を再現するための
  土台。Resolve-ProfileModules は KERNEL_API.md §6 の内部実装の
  ため公開 API への影響なし、KERNEL_VERSION 影響なし。
  （FrexProfile 機能の Phase 1 / 全体は完成時に MINOR で集約 bump
  予定）

- modules/standard/evidence_config: 5 new sections covering inventory
  items that PCView (a legacy 2011 inventory tool) captured but fabriq
  evidence previously did not, raising audit-pack completeness:
    - §27 Environment Variables (CSV) — Machine + User scopes,
      27_EnvironmentVariables.csv. Process scope intentionally excluded
      as volatile (depends on running shell, not system state).
    - §28 Startup Items (CSV) — 28_StartupItems.csv with a Source
      column. Win32_StartupCommand is captured in full (Run / RunOnce /
      Startup folder, PCView-compatible). ScheduledTask is filtered to
      logon-triggered, non-Disabled, non-`\Microsoft\Windows\*` entries
      to keep the CSV evidence-relevant.
    - §29 Memory Slots (CSV) — 29_MemorySlots.csv per-slot detail from
      Win32_PhysicalMemory (BankLabel / DeviceLocator / Capacity_GB /
      Speed_MHz / Manufacturer / PartNumber / SerialNumber / FormFactor
      / SMBIOSMemoryType / DataWidth / TotalWidth) with FormFactor and
      SMBIOSMemoryType numeric codes translated to strings.
    - §29b Memory Array Summary (CSV) — 29b_MemoryArraySummary.csv from
      Win32_PhysicalMemoryArray (MaxCapacity_GB / MemoryDevices /
      MemoryErrorCorrection). Split mirrors §8b Disks/Partitions
      convention. Sub-collection failure marks §29 Partial without
      failing the whole section.
    - §30 PnP Devices (CSV) — 30_PnpDevices.csv from `Get-PnpDevice`
      without `-PresentOnly` (past-connected devices retained for audit
      traceability). DriverVersion / DriverDate are queried per-instance
      via `Get-PnpDeviceProperty`; per-device query failures fall back
      to blank cells without failing the section.
    - §31 Hardware Identifiers (TXT) — 31_HardwareIdentifiers.txt
      aggregating Win32_ComputerSystem / Win32_ComputerSystemProduct /
      Win32_BaseBoard / Win32_SystemEnclosure. ChassisTypes numeric
      codes translated to strings. Complements §10 (PC serial) and §23
      (BIOS / TPM) without duplication.
  evidence_config 1.5.0 -> 1.6.0 (MINOR, backward-compatible additions).
  manifest schemaVersion stays at 1 per kernel/EVIDENCE_MANIFEST.md §4.2
  ("new section addition is OK within schemaVersion=1"). Section IDs
  27 / 28 / 29 / 29b / 30 / 31 are new, evidence_manager will fall back
  to UnknownSection raw display until its KnownSections dictionary
  catches up (separate task; no fabriq-side blocker).

### Fixed
- apps/fabriq_ios/ inline `?` no longer hides the alphabetically-first
  candidate. The chord handler in lib/completer.ps1 wrote candidate
  rows starting from the input row's current cursor column, then
  called InvokePrompt() which redrew prompt+buffer at PSReadLine's
  _initialY = the input row, overwriting whatever the first Write-Host
  had appended. With six Group enums for `(config-mod)# set Enabled 1
  Group ?`, only Backup Operators..Users were visible; Administrators
  silently vanished even though Tab returned all six (Tab handler
  already had a leading sacrificial Write-Host "" that the ? handler
  was missing). The empty-list "(no candidates)" message was hidden
  by the same mechanism.

  Fix: emit the same sacrificial blank line at the start of the ?
  candidate emission so candidates land on rows below the input row
  and survive the InvokePrompt redraw. Orphan-tail blanking math is
  unchanged: $tailY (captured after all writes) shifts down by one
  to match, and $extra = $prevRows - $rowsWritten only depends on the
  relative diff, which both presses share. VERSION 0.3.4 -> 0.3.5
  (PATCH, visual fix in the same family as 0.3.2-0.3.4 inline ?
  refinements; covers visual subtlety #1 from the file-level comment
  block which was documented but never implemented).

- apps/fabriq_ios/ inline `?` no longer leaves merged garbage
  candidates from previous presses ("Passwordperators" =
  new "Password" overlaid on stale "Operators...something"). Two
  underlying causes:
  
  (a) Per-row overlap: each `?` press writes its candidates to the
      same row range as the previous press because PSReadLine's
      _initialY does not move. A new short candidate landing on a
      row that previously held a longer one left the trailing
      characters of the old visible after the new.
      Fix: pad every help row to full window width so leftover
      trailing chars are blanked.
  
  (b) Per-list tail: rows beyond the new list's length still held
      old candidates that were never overwritten, producing a
      stale tail underneath.
      Fix: track the previous press's row count and explicitly
      blank any orphan tail rows. Reset the count when the input
      row changes (user submitted a command and a fresh prompt
      sits on a new row), since rows below now hold unrelated
      console output we must not overwrite.
  
  Also disabled PSReadLine prediction (PredictionSource None) in
  Initialize-FabriqIos. Default since PSReadLine 2.2 is InlineView,
  which renders history-based suggestions in the same screen area
  used by our manual candidate rows. Subprocess-scoped so the
  parent shell's prediction setting is not affected. Wrapped in
  try/catch for PSReadLine < 2.1 compatibility.
  
  Replaces the 0.3.3 single-row clear which addressed the wrong
  failure mode. VERSION 0.3.3 -> 0.3.4 (PATCH, visual refinement).

### Fixed (superseded)
- apps/fabriq_ios/ inline `?` no longer leaves the original prompt
  row visible above the candidate list. The previous fix preserved
  the buffer correctly, but each `?` press painted the new
  prompt+buffer below the help via InvokePrompt while leaving the
  REPL's manually-rendered prompt row (from [Console]::Write before
  ReadLine) intact, producing a stale duplicate row above the help
  on every redraw ("？で描画されなおすたびに前の表示と重なる").
  
  PSReadLine's InvokePrompt re-anchors _initialY at the post-help
  cursor row but does not erase rows above it, since they were
  written outside PSReadLine's own input-area tracking. The `?`
  chord handler now SetCursorPosition + space-fills the current
  input row before writing candidates, so the help and the
  re-rendered prompt below stack cleanly with nothing left over.
  
  Single-row clear only - multi-row wrapped buffers still leave
  the upper wrap rows as visual artifact (PSReadLine state stays
  correct, just cosmetic). Acceptable for fabriq_ios since most
  commands fit on one line; revisit if long quoted values become
  common. Tab handler unchanged (user has not reported the same
  artifact for Tab; same pattern would apply if needed).
  
  VERSION 0.3.2 -> 0.3.3 (PATCH, visual-only refinement of the
  Cisco-style `?` introduced in 0.3.2).

### Changed
- apps/fabriq_ios/ inline `?` now preserves the buffer in Cisco IOS
  fashion. Previously, pressing `?` after typing something
  (`(config)# host?`) showed candidates but then redrew an empty
  prompt - the user had to retype the partial command. Now the
  buffer is restored after the candidate list, so editing
  continues from where it left off:
  
    (config)# host?
      hostname
    (config)# host_      <- cursor here, ready to keep typing
  
  Implementation: the `?` chord handler in lib/completer.ps1 now
  branches on buffer contents. Non-empty buffer takes the same
  inline path as Tab (Write-Host candidates + InvokePrompt, which
  preserves the input line by design). Empty buffer still defers
  to the REPL via AcceptLine because the full mode-level help can
  run 30+ lines and would push the prompt out of InvokePrompt's
  redraw window. The original "重なる" overlap problem from the
  earlier inline attempt was caused by Show-FabriqIosHelp's
  long-form output, not by InvokePrompt itself - candidate lists
  are short enough that the issue does not recur.
  
  REPL pending-help branch in fabriq_ios.ps1 simplified
  accordingly: the non-empty-buffer arm is dead code under the new
  handler and was removed. VERSION 0.3.1 -> 0.3.2 (PATCH: UX
  refinement, no API change). Smoke tests unchanged - PSReadLine
  chord behaviour is not exercised in the smoke suite (requires a
  real terminal); manual verification only.

### Fixed
- apps/fabriq_ios/ ModuleConfig `set` / `add` Tab completion no
  longer goes silent past the first column. Previously
  Get-FabriqIosCompletion routed any 2+ token prefix into the
  fixed-arity sub-vocabulary path (designed for `show running-config`
  style verbs), so `set Enabled <Tab>`, `set Enabled 1 <Tab>`,
  `set Enabled 1 Type REG_<Tab>` etc. all returned no candidates.
  
  Added Get-SetAddPositionalCompletion: a dedicated branch for
  set / add that decides candidate kind by position parity. Even
  args after the verb -> next token is a column (filtered to
  hide already-named columns, case-insensitive). Odd -> next is
  an enum value for the column at TokensBefore[-1], pulled from
  ConfigModuleSchema.Enums (preset.csv). Columns without preset
  enums silently return zero candidates rather than erroring.
  
  Behaviour matches Cisco IOS chained-arg completion within a
  single `set` command and removes the friction of having to
  re-check `show` for each subsequent column. The named
  `set <col> <val> [<col> <val>...]` syntax is preserved as-is;
  no parser change.
  
  VERSION 0.3.0 -> 0.3.1 (PATCH: completion engine completing
  cases that were always intended to work). 18 new assertions
  added to tests/_phase7_smoke.ps1 covering enum-value position,
  prefix narrowing, used-column filtering (case-insensitive),
  multi-pair chains (3 pairs deep), add-verb parity with set,
  and enum-less columns. All phase smokes pass: 339/339.

### Added
- apps/fabriq_ios/ Phase 9b: JSON object-form entries for
  multi-script / multi-CSV modules. data/module_categories.json
  schemaVersion 1 now accepts each `modules[]` item as either a
  string (implicit defaults: dir=name, script=name.ps1, csv
  auto-detected via *_list.csv glob) OR an object with explicit
  overrides: `{ name, dir?, script?, csv?, label? }`. This lets a
  single underlying module directory expose multiple logical
  entries with their own prompts and CSVs - the joke shell now
  surfaces operations that string form cannot represent.
  
  Nine object-form entries added across 4 modules:
    settings: local_user_create (-> local_user_config\local_user_config.ps1
              + local_user_list.csv), sysprep_main / sysprep_unattend /
              sysprep_setupcomplete (all -> sysprep_config\sysprep_config.ps1
              with sysprep_list / unattend_list / setupcomplete_list CSVs)
    cleanup : local_user_delete (-> local_user_config\local_user_delete.ps1,
              shares local_user_list.csv with create), destroy_history
              (-> history_destroyer\history_destroyer.ps1 + destroy_list.csv),
              destroy_ssid (-> history_destroyer\history_destroyer.ps1
              + ssid_list.csv)
    install : printer_driver_install (-> printer_driver_config\
              printer_driver_install.ps1), printer_register
              (-> printer_driver_config\printer_config.ps1 + printer_list.csv)
  
  This silently fixes two bugs from Phase 9a's string-only model:
  (1) `printer_driver_config` had no `printer_driver_config.ps1`
  so the verb never resolved to anything; (2) `local_user_delete.ps1`
  was completely unreachable from the shell.
  
  NEW lib/commands/categories.ps1: Resolve-ModuleEntry normalises
  both string and object forms into a uniform hashtable
  { Name; Dir; Script; Csv; Label; Category }. The exclusion list
  in lib/commands/module.ps1 still takes precedence.
  Find-ModuleEntryAcrossCategories searches every category when
  the dispatch path knows the name but not the category.
  Get-FabriqIosEntryName is the private accessor for the display
  name regardless of entry shape.
  
  lib/commands/module.ps1 updated: Find-ModulePath gained an
  optional -CategoryId so it honours per-entry Dir / Script
  overrides when scoped (and falls back to cross-category search
  otherwise). Get-ModuleCsvSchema gained matching -CategoryId
  and now returns ScriptPath in the result so downstream callers
  do not need a second Find-ModulePath round-trip. When an explicit
  csv override is declared but the file is missing on disk, schema
  is reported as unavailable (no silent substitution). String-form
  entries continue to auto-detect via *_list.csv glob.
  
  Enter-CategoryConfigMode now passes $cat.id to both Find-ModulePath
  and Get-ModuleCsvSchema so the right entry is resolved when the
  same name might exist across categories. Invoke-ModuleEphemeralRun
  reads $schema.ScriptPath instead of recomputing the path.
  
  VERSION bumped 0.2.0 -> 0.3.0 (MINOR: schema extension is a
  backward-compatible addition - all existing string entries
  continue to work unchanged). tests/_phase9b_smoke.ps1 NEW
  (82 assertions, all PASS) covers Resolve-ModuleEntry for both
  shapes, override resolution, cross-category search, dispatcher
  routing of object-form entries, prefix completion narrowing,
  and wrong-category rejection. tests/_phase9_smoke.ps1 updated
  to reflect new entry counts (cleanup 7->9, install 11->12).
  All phase smokes pass: 3=26, 4=17, 5=32, 6=41, 7=49, 9=74,
  9b=82 -> 321/321.

- apps/fabriq_ios/ Phase 9a: category-driven verb dispatch.
  Single (config-mod)# universe replaced with five
  category-specific verbs in (config)#:
    `module <name>`   -> (config-mod)#       settings (43 modules)
    `cleanup <name>`  -> (config-clean)#     cleanup (7 modules)
    `copy <name>`     -> (config-copy)#      copy (2 modules)
    `install <name>`  -> (config-install)#   install (11 modules)
    `script <name>`   -> (config-script)#    scripting (5 modules)
  Inside any of these the same `set <col> <val> [<col> <val>...]`
  / `add` / `show` / `exit` / `end` commands work - only the
  prompt suffix and the candidate module list change.
  
  NEW data/module_categories.json (schemaVersion 1): the single
  source of truth for the category-to-modules mapping. Each
  category declares `id`, `verb`, `promptSuffix`, `intro`, and a
  `modules[]` list. To reclassify or to remove "modules that don't
  work well via CLI", edit this file - no PowerShell change needed.
  
  NEW lib/commands/categories.ps1: JSON loader (cached) +
  Get-FabriqIosCategoryByVerb / Get-FabriqIosCategoryById /
  Get-CategoryModuleCompletion / Test-ModuleInCategory. The
  exclusion list in lib/commands/module.ps1 still takes precedence
  (modules listed in JSON but globally excluded are silently
  filtered out).
  
  ShellState gained CurrentCategoryId. lib/prompt.ps1 reads it
  from the bound state and looks up the matching promptSuffix on
  every render, so each (config-mod)/(config-clean)/(config-copy)/
  etc. mode shows the right prompt out of the box.
  
  GlobalConfig vocabulary expanded from
  ('hostname','interface','module','exit','end','help','?')
  to add 'cleanup','copy','install','script'. Each verb routes
  through Invoke-VerbModeEntry -> Enter-CategoryConfigMode, which
  validates the verb-to-category mapping AND that the supplied
  module name is listed under that category. Wrong-category
  invocations (e.g. `cleanup reg_hklm_config`) are rejected before
  any module-side work happens.
  
  G option (immediate strict): the `module` verb no longer accepts
  modules outside the settings category - this is a deliberate
  Phase 9 break. To run a cleanup module use `cleanup <name>`,
  not `module <name>`. Phase 6 / 7 tests updated accordingly
  (Enter-ModuleConfigMode renamed to Enter-CategoryConfigMode
  with -Verb parameter; "not found" wording for category mismatch
  changed to "is not in the '<category>' category").
  
  data/help_text.csv updated with 4 new GlobalConfig entries
  (cleanup / copy / install / script). VERSION bumped 0.1.0 -> 0.2.0
  (Phase 9 is a feature MINOR addition over the standalone Phase 8
  fork). tests/_phase9_smoke.ps1 NEW (69 assertions, all PASS)
  covers JSON loading, verb-to-category lookup, per-category
  filtering, mode entry per verb, prompt suffix variation,
  wrong-category rejection, parser abbreviations
  (cl/co/ins/s -> cleanup/copy/install/script).
  
  All five other phase smokes still pass:
  Phase 3=26, 4=17, 5=32, 6=41, 7=49, 9=69 -> 234/234.

- Fabriq_IOS.exe (NEW, top-level): standalone launcher for the
  fabriq_ios sub-project. Mirrors the Fabriq.exe pattern - tiny C#
  wrapper (dev/launcher/Launcher_IOS.cs + app_ios.manifest +
  build_ios.ps1), auto-elevates via UAC, runs `conhost.exe
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File
  apps\fabriq_ios\fabriq_ios.ps1` from the fabriq root. Reuses
  fabriq.ico. AssemblyVersion 0.1.0.0 mirroring fabriq_ios's own
  SemVer line. ~56 KB binary.
- apps/fabriq_ios/VERSION (NEW): independent SemVer for the joke
  shell sub-project (initial 0.1.0). Decoupled from kernel
  version. show version / show running-config now read this.

### Changed
- apps/fabriq_ios/: Phase 8 fork - established as a standalone
  sub-project. The fabriq_operator launch path is removed (the
  `[Fabriq IOS]` Quick Actions button is gone, FabriqApps dialog
  excludes fabriq_ios from auto-discovery), and the joke shell
  now boots only via Fabriq_IOS.exe at the fabriq root. The
  development tree stays at apps/fabriq_ios/ (same repo) but
  hostlist / workerlist / profile / fabriq-logging coupling is
  severed:
    - lib/parser.ps1 show subvocabularies dropped `host` and
      `hosts`. UserExec.show is now @('version','manifesto');
      PrivilegedExec.show is @('version','running-config',
      'profiles','modules','evidence','manifesto'). `show profiles`
      and `show evidence` are read-only informational displays
      (no execution / no writing).
    - lib/commands/show.ps1: removed Show-FabriqIosHost and
      Show-FabriqIosHosts. Show-FabriqIosVersion / -RunningConfig
      now use a new Get-FabriqIosVersion helper that reads
      apps/fabriq_ios/VERSION. show running-config dropped the
      `session worker <name>` line (workerlist coupling) and the
      `version <kernel>` line is now `version <fabriq_ios>`.
    - lib/commands/hostname.ps1: simplified Invoke-HostnameSelection
      (no hostlist lookup). Takes the name as a positional ad-hoc
      value, sets SELECTED_NEW_PCNAME directly with adhoc
      OldPCName / KanriNo when no host is bound, runs
      hostname_config. Get-HostnameCompletionFromHostlist removed.
    - lib/commands/ip_address.ps1: Invoke-IpAddressFromHostlist
      and Get-IpAddressCompletionFromHostlist removed. The only
      surviving form is the Phase-8-extended positional
      `ip address <ip> <mask> [<gw> [<dns1> [<dns2>]]]`.
    - lib/modes/interface_config.ps1: dropped the `ip address
      from-hostlist` branch and updated usage hint.
    - lib/completer.ps1: dropped `hostname` and `ip.address`
      Get-DynamicCompletion sources.
    - lib/dispatch.ps1: dropped Set-FabriqIosHostEnvironment,
      Find-HostlistRowByNewName, and the duplicate
      ConvertFrom-SubnetMaskToPrefix (canonical copy stays in
      lib/commands/ip_address.ps1). dispatch.ps1 now hosts only
      Invoke-FabriqIosModule.
  fabriq_ios.ps1 startup additionally shadows the fabriq logging
  functions (Initialize-ExecutionHistory / Write-ExecutionHistory /
  Add-ExecutionResult / Capture-ScreenEvidence /
  Initialize-EvidenceBasePath / etc.) at global scope with no-ops,
  so any module that accidentally calls them is silently dropped
  (defence in depth - the current dispatch path never calls them).
  Reads apps/fabriq_ios/VERSION into $script:FabriqIosVersion.
- apps/fabriq_operator/lib/dashboard_form.ps1: Quick Actions
  rearranged after [Fabriq IOS] removal. Front row now hosts
  [Open CSV Editor] / [Windows Update] / [Refabriq]; second row
  [System Launcher] / [And More...].
- apps/fabriq_operator/lib/apps_dialog.ps1: auto-discovery now
  also excludes fabriq_ios alongside fabriq_operator.
- apps/fabriq_operator/lib/dashboard_form.ps1: renamed
  Populate-ModuleGrid -> Update-ModuleGrid (PSScriptAnalyzer
  warning PSUseApprovedVerbs - "Populate" is unapproved, "Update"
  is approved).
- tests/_phase3_smoke.ps1, _phase4_smoke.ps1, _phase5_smoke.ps1:
  removed assertions for deleted hostlist-coupled functions and
  show subcommands. Phase 4 smoke rewritten to cover the
  ad-hoc hostname / 6-arg ip address forms.

- apps/fabriq_ios/: `(config-if)# ip address` extended to accept up
  to 5 positional args, allowing fully ad-hoc IP configuration
  without requiring a host to be bound. Previous form
  `<ip> <mask>` (Gateway / DNS still inherited from the bound host)
  remains. New trailing arguments are all optional:
    ip address from-hostlist                              (existing)
    ip address <ip> <mask>                                (existing)
    ip address <ip> <mask> <gw>                           (NEW)
    ip address <ip> <mask> <gw> <dns1>                    (NEW)
    ip address <ip> <mask> <gw> <dns1> <dns2>             (NEW)
  Invoke-IpAddressManual extended to accept Gateway / Dns1 / Dns2
  parameters and to seed an `(adhoc)` SELECTED_NEW_PCNAME identity
  (with $env:COMPUTERNAME for OldPCName, '0' for KanriNo) when no
  host is bound, so ipaddress_config still has display data even
  without prior `(config)# hostname <NewName>`. All env-var
  mutations remain saved/restored in finally. Use case: setting a
  freshly-thought-up IP that does not correspond to any hostlist
  row.

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
