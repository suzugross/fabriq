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
