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
