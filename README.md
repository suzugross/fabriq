# Fabriq ver3.6

**Manifeste du Surkitinisme**

Windows PC キッティング自動化フレームワーク（PowerShell + WinForms GUI）

## 概要

Fabriq は、Windows 11 PC の初期セットアップ（キッティング）を自動化する PowerShell ベースのフレームワークです。

- **CSV 駆動 + モジュール型アーキテクチャ** — コード改修なしに対象 PC・設定値・実行順序を切り替え
- **GUI ダッシュボード**（WinForms）から個別モジュール実行・プロファイル一括実行・セッション切替を操作
- **AutoPilot** による省人化キッティング。モジュール単位のエラー処理ポリシー（`skip` / `retry`）、再起動跨ぎ自動再開、HTML チェックリスト出力、スクリーンショット自動取得（運用上は仕上げ・確認のため operator の立ち会いを想定）
- **AES-256-CBC + PBKDF2** による CSV 内機密値の暗号化保持
- **79 種類の標準・拡張モジュール**を同梱

Fabriq には、本体（PowerShell フレームワーク）と連携する 2 つの GUI アプリケーション（いずれも別リポジトリ・WPF / .NET 8.0）があります。設定・運用は本体だけでも完結しますが、これらを併用することで「設定の整備 → キッティング実行 → エビデンスの納品」までを一貫してカバーできます。

- **Fabriq Studio** — Fabriq ワークスペースの設定管理 GUI（ホスト・モジュール設定・プロファイル・暗号化・パスフレーズ設定）。<https://github.com/suzugross/fabriq_studio>（後述「[Fabriq Studio との関係](#fabriq-studio-との関係)」）
- **Fabriq Evidence Manager** — キッティング実行で出力されたエビデンスを読み込み、顧客納品用フォーマットへ一括エクスポートする GUI。<https://github.com/suzugross/fabriq_evidence_manager>（後述「[Fabriq Evidence Manager との関係](#fabriq-evidence-manager-との関係)」）

また、本体・上記アプリを含む Fabriq シリーズ全体の技術ドキュメントは、ドキュメント専用リポジトリに一元管理されています。

- **Fabriq Doc** — Fabriq シリーズの統合ドキュメント集（仕様・利用方法・カーネル/モジュール解説）。<https://github.com/suzugross/fabriq_doc>（後述「[Fabriq Doc との関係](#fabriq-doc-との関係)」）

## デモ動画

Fabriq で 2 台の PC を一括キッティングし、Fabriq Evidence Manager でエビデンスを確認するまでの一連の流れを収録しています。

▶ **<https://youtu.be/eiZHYYbUUKY>**

動画内で適用している主な設定:

- ホスト名変更 / IP アドレス変更 / ドメイン参加（`fabriq.group`）
- デフォルト BitLocker 解除 → 再起動 → BitLocker 設定（PIN）
- 電源設定（高パフォーマンス）/ オートログオン設定
- アプリインストール（Chrome / 7-Zip / Office）/ プリンタ（Canon）
- 壁紙変更（FabriqWallPaper）/ レジストリ（Spotlight 無効・高速スタートアップ無効・タスクバー左寄せ）
- スタートピン留め（エクスプローラ・設定・Word・Excel・PowerPoint）/ タスクバーピン留め（エクスプローラ・Chrome）
- 履歴削除 / エビデンス収集

## 前提条件

- **Windows 11**
- **PowerShell 5.1** 以降
- **管理者権限**（`Fabriq.exe` / `Fabriq.bat` が自動昇格）
- **Fabriq Studio** によるマスターパスフレーズ設定（後述）

> **ランチャの `.exe` / `.bat` 並列同梱**: ルートには `Fabriq.exe` と `Fabriq.bat`（`Fabriq_IOS` も同様）を併置しています。`.bat` は `.exe` の代替手段で、ウイルス対策ソフト（Defender 等）のヒューリスティックが未署名の EXE ランチャを誤検知・隔離した場合でも起動できるようにするためのものです。両者は同じ挙動（管理者へ自動昇格 → 作業ディレクトリを fabriq ルートへ固定 → `kernel/main.ps1` を起動）で、どちらか一方が使えればよく、互いを置き換えるものではありません。

> **重要**: Fabriq の起動時にはマスターパスフレーズの入力が必須です。パスフレーズ検証トークン（`kernel/txt/passphrase_verify.txt`）が存在しない場合、Fabriq は起動できません。このトークンは **Fabriq Studio** でパスフレーズを設定することで生成されます。必ず初回起動前に Fabriq Studio でパスフレーズを設定してください。

## 主な機能

| 機能 | 説明 |
|------|------|
| **モジュールシステム** | Standard 61 種・Extended 18 種・計 79 種（ホスト名、IP、レジストリ、アプリ、BitLocker、Sysprep 等。うち `windows_update` は GUI ボタン専用） |
| **GUI ダッシュボード** | `Fabriq.exe` 起動後、WinForms ダッシュボードから全操作を実施。CLI モードは廃止 |
| **プロファイル実行** | 複数モジュールを順序付きで一括実行。`__AUTOPILOT__` マーカー以降は確認ダイアログをスキップして自動進行 |
| **AutoPilot ErrorMode** | プロファイル CSV の `ErrorMode` 列でモジュール単位に `skip` / `retry`（最大 5 回）を宣言し、AutoPilot 中のエラー対応を自動化 |
| **再起動跨ぎ** | `__RESTART__` マーカーにより、再起動後に RunOnce 経由で自動再開 |
| **`__AUTO_to_<User>__`** | `autologon_config` の `autologon_list.csv` から `User` 列一致のエントリを指定して呼び出す専用マーカー（例: `__AUTO_to_admin01__`） |
| **CSV 駆動** | `hostlist.csv` で PC 毎の設定を定義。各モジュールの設定も CSV で管理 |
| **暗号化** | AES-256-CBC + PBKDF2-SHA256（10 万回、固定ソルト）で CSV 中の機密値を `ENC:<Base64>` 形式で保持 |
| **エビデンス自動取得** | モジュール実行ごとにスクリーンショットを自動保存 |
| **HTML チェックリスト** | プロファイル実行後に実行結果・ネットワーク照合・プリンタ照合を HTML レポート出力 |
| **Execution Toolbar** | ダッシュボード内蔵（in-process）の実行状況バー。旧・別プロセスの Status Monitor は kernel 3.4.0 で Defender/ASR ブロック対策のため撤去し in-process 化 |
| **セグメント** | 同じモジュールをプロファイル内で設定値別に呼び分け可能（厳密マッチ） |
| **Post-Apply Verification** | 設定適用後に読み返し検証を行い、実行履歴に `Verified` 列として記録 |
| **プリセット UI** | 各モジュールの `preset.csv` により、設定 CSV の列挙型列を Fabriq Studio 側でドロップダウン編集可能 |
| **ログ管理** | PowerShell Transcript + 実行履歴 CSV + `log_destinations.csv` による外部共有フォルダへの自動アップロード |

## ディレクトリ構成

```
fabriq/
├── Fabriq.exe              # エントリーポイント（管理者自動昇格、GUI 起動）
├── Fabriq.bat              # 同上の .bat 版（AV が未署名 EXE を隔離した場合の代替・同挙動）
├── Fabriq_IOS.exe          # fabriq_ios サブプロジェクト用ランチャ（独立 SemVer）
├── Fabriq_IOS.bat          # Fabriq_IOS.exe の .bat 版（代替・同挙動）
├── kernel/
│   ├── main.ps1            # メインスクリプト（Fabriq.exe から呼び出し）
│   ├── common.ps1          # 共通関数ライブラリ（96 関数）
│   ├── KERNEL_VERSION      # カーネル API SemVer（真のソース）
│   ├── KERNEL_API.md       # 公開 API サーフェスの明文化
│   ├── EVIDENCE_MANIFEST.md # evidence manifest 公開契約（外部 consumer 向け）
│   ├── csv/                # マスタ CSV（categories, hostlist, workers, log_destinations, manifesto）
│   ├── json/               # ランタイム状態（session, status, resume_state, art_pulse, async_config）
│   ├── ps1/                # カーネルサブスクリプト（manifesto, view_report）
│   └── txt/                # パスフレーズ検証トークン、アート文言、silence フラグ
├── modules/
│   ├── standard/           # 標準モジュール群（60）
│   └── extended/           # 拡張モジュール群（16）
├── profiles/               # 実行プロファイル CSV
├── apps/                   # FabriqApps：GUI アプリツール群
│   ├── fabriq_operator/    # メインダッシュボード GUI
│   ├── fabriq_ios/         # Cisco IOS 風シェル（芸術部門サブプロジェクト）
│   ├── csv_editor/         # CSV 編集 GUI
│   ├── system_launcher/    # OS 機能ランチャ
│   ├── bloatware_exporter/ # インストール済みアプリ一覧エクスポート
│   ├── desktop_icon_backup_app/
│   ├── local_user_setup/
│   ├── storeapp_editor/
│   └── winget_gui/
├── commands/               # ユーティリティコマンド（diag_crypto, get_evidence, gpupdate 等）
├── evidence/               # エビデンス（スクリーンショット + HTML チェックリスト）出力先
├── logs/                   # ログ出力先（Transcript + 実行履歴 CSV）
└── dev/                    # 開発用ツールチェーン
    ├── template/           # 新規モジュール用テンプレート（VERSION/REQUIRES_KERNEL 含む）
    ├── launcher/           # Fabriq.exe / Fabriq_IOS.exe の C# ソース
    │                       # (fabriq_backuper の launcher は分離先 E:\fabriq_backuper\ 側で管理)
    ├── ico/                # ランチャーアイコン素材
    ├── framework_overlay_rules.json  # 更新オーバーレイ契約（KERNEL_API.md §9）
    ├── build_framework_patch.ps1     # フレームワーク更新パッチ生成
    ├── seed_module_versions.ps1      # 全モジュール VERSION/REQUIRES_KERNEL の baseline seed（idempotent）
    ├── check_version.ps1             # KERNEL_VERSION と版表記の整合チェック
    └── verify_comments_only.ps1      # スクリプトコメント英語化検証
```

## クイックスタート

### 1. Fabriq Studio でパスフレーズを設定

Fabriq Studio を起動し、ワークスペースとして Fabriq フォルダを開き、マスターパスフレーズを設定します。検証トークン `kernel/txt/passphrase_verify.txt` が生成され、Fabriq が起動可能になります。

### 2. デプロイ

Fabriq フォルダを対象 PC へ直接配置します（USB メモリ等からコピー）。

### 3. 起動

`Fabriq.exe` をダブルクリックします（UAC により管理者権限へ自動昇格 → PowerShell コンソールを経由してダッシュボード GUI が立ち上がります）。ウイルス対策ソフトが `Fabriq.exe` を隔離するなどして起動できない場合は、同じ挙動の `Fabriq.bat` を代わりに使用します。

### 4. セッション開始

1. マスターパスフレーズを入力
2. 作業者を選択（`workers.csv` から選択、または手入力）
3. 対象 PC を選択（`hostlist.csv` から選択。PC 名が一致すれば自動選択）

### 5. モジュール実行

ダッシュボード GUI の各ボタンから実行します。

## 使い方

### GUI ダッシュボード

ダッシュボードは複数タブで構成され、代表的なボタンは以下の通りです。

| ボタン | 動作 |
|---|---|
| **Execute Profile** | プロファイル CSV を選択して Linear モードで一括実行 |
| **Execute (Flex)** | 選択中プロファイルを **FlexProfile** ダッシュボードで開く（状態追跡型・部分実行・後述） |
| **View Details** | 選択中プロファイルの構成モジュールを事前確認 |
| **Execute** | 個別モジュールをチェック選択して連続実行 |
| **Open Folder**（Evidence） | エビデンス保存先をエクスプローラで開く |
| **Open CSV Editor** | 付属 CSV エディタで hostlist / モジュール設定を編集 |
| **Export History** | 実行履歴をエビデンスとして明示エクスポート |
| **Regenerate Checklist** | 実行履歴から HTML チェックリストを再生成 |
| **Windows Update** | Windows Update を再起動跨ぎで実行 |
| **Restart PC** | 再起動（fabriq を継続するかは確認） |
| **Refabriq** | セッション・履歴・トランスクリプトをリセットして再開 |
| **System Launcher** | Windows の設定系ショートカット集 |
| **FabriqApps** | `apps/` 配下の GUI ツール群を起動 |
| **New Session** | 作業者・対象 PC を切り替えて新セッション開始 |
| **Manifeste du Surkitinisme** | マニフェスト表示 |

### プロファイル実行

プロファイル CSV で実行するモジュールと順序を定義し、AutoPilot モードで自動進行させられます（再起動跨ぎを含む長尺シーケンスでも operator の都度確認は不要）。

**プロファイル CSV の書式:**

```csv
Order,ScriptPath,Enabled,Description,Segment,ErrorMode,Group
10,__AUTOPILOT__,1,WaitSec=3,,,
20,standard/hostname_config/hostname_config.ps1,1,ホスト名設定,,,Network
30,standard/ipaddress_config/ipaddress_config.ps1,1,IP アドレス設定,,retry,Network
40,__RESTART__,1,再起動,,,
50,standard/reg_hklm_config/reg_hklm_config.ps1,1,レジストリ設定,,skip,Tweaks
```

| 列 | 説明 |
|---|---|
| `Order` | 実行順（整数、昇順）。実行履歴の一級識別子 |
| `ScriptPath` | `standard/<module>/<script>.ps1` 形式。区切りは `/` / `\` どちらも可 |
| `Enabled` | `1`=実行 / `0`=スキップ |
| `Description` | プロファイル UI 表示用コメント |
| `Segment` | 同モジュールを設定値別に呼び分ける際のセグメント名（省略可。`_list.csv` 側の `Segment` 列と厳密マッチ） |
| `ErrorMode` | AutoPilot 実行時のエラー処理ポリシー（省略=ダイアログ確認 / `skip`=自動スキップ / `retry`=最大 5 回自動リトライ） |
| `Group` | 任意。FlexProfile ダッシュボードの **Groups バー**で `[Run: <Group>]` ボタンとして集約される名前。Linear `[Execute Profile]` は本列を参照しない（kernel 3.2.0 以降） |

**特殊マーカー:**

| マーカー | 動作 |
|---|---|
| `__AUTOPILOT__` | 以降のプロファイル実行を AutoPilot モードで自動化。`Description` に `WaitSec=N` でモジュール間ウェイト秒を指定可能 |
| `__ASYNC__` | （任意）以降のモジュールを監視付き Runspace で実行。kernel 3.3.0 以降は DefaultAsync が既定 ON のため通常は不要。ハング時は Execution Toolbar の **Skip** ボタン、または `async_config.json` の `DefaultTimeoutSec` で強制スキップ（`Enabled: false` で全体無効化） |
| `__RESTART__` | Windows を再起動し、RunOnce 経由で次モジュールから自動再開 |
| `__REEXPLORER__` | Explorer を再起動（レジストリ変更の即時反映等） |
| `__GATE__` | 前進バリア（kernel 3.6.0〜）。直前ゲート〜本マーカの窓に `Error`/`Partial` または Post-Apply Verification 失敗（`Verified=False`）のモジュールが残る間、本マーカ以降の `Order` 実行を拒否（動的評価）。FlexProfile では該当行をグレーアウト。窓が解消すると解除 |
| `__AUTO_to_<User>__` | `autologon_config` の `autologon_list.csv` から `User` 列一致のエントリを呼び出す（例: `__AUTO_to_admin01__`） |

> kernel 3.0.0 で旧マーカー `__SHUTDOWN__` / `__PAUSE__` / `__STOPLOG__` / `__STARTLOG__` を破壊的に削除しました。これらを含む旧プロファイルは graceful degradation（"module not found" 警告として降格、他モジュールの実行は継続）で動作します。

### FlexProfile（状態追跡型実行）

**FlexProfile** は kernel 3.1.0 で導入された、プロファイルを **state-aware に部分実行**できる WinForms ダッシュボードです。Linear `Execute Profile` の「先頭から末尾まで一気通貫」モデルに対し、Flex は「現セッションで何が成功／失敗／未実行か」を実行履歴から復元してグリッド表示し、operator が任意の組み合わせで段階的に進める運用を可能にします。

ダッシュボードの **Profiles** タブで対象プロファイルを選択し、`[Execute (Flex)]` ボタンで起動します。

| 機能 | 動作 |
|---|---|
| **Status / Verified バッジ** | 各行が `Success` / `Partial` / `Error` / `Skipped` / `Pending` を背景色付きバッジで表示。`Verified` 列は Post-Apply Verification の `PASS` / `FAIL` を緑/赤で表示 |
| **`[Run]`（行ごと）** | 各行末尾の `[Run]` ボタンで該当モジュール 1 件を即時単発実行（AutoConfirmMode で Y/N プロンプト・Press-Enter 待機をスキップ） |
| **`[Log]`（行ごと）** | 各行の `[Log]` ボタンで該当モジュールの実行ログ（Show-* 出力）をその場で色分け表示。conhost 窓や生 transcript を追わずに確認（kernel 3.6.0〜） |
| **`[Run Selected (N)]`** | チェックボックスで選択した行を Order 昇順で一括実行（AutoPilot 挙動 + finalize は手動委譲） |
| **`[Run: <Group>]`（Groups バー）** | プロファイル CSV の `Group` 列で集約された行群を 1 クリックで一括実行。Group 跨ぎの `__RESTART__` は当該 Group 実行時にスキップ（literal interpretation） |
| **`[Select All]` / `[Clear All]`** | bulk-select。`[Select All]` は CSV `Enabled=1` 行のみチェック |
| **`[Mark as Pending]`（行右クリック）** | 該当行の Status を Pending にリセット（再実行候補に戻す） |
| **`[Restart Now]`** | プロファイル外から `__RESTART__` を発火。Flex resume 経由で再起動後に自動でダッシュボードへ復帰 |
| **`[Complete]`** | finalize phase（HTML チェックリスト生成 + log_uploader）を手動発火。Error / Partial / Pending 行があれば黄色バッジで警告 |
| **PENDING FINALIZE バッジ** | バッチ実行後 `[Complete]` 未押下のままダッシュボードを離脱しようとすると赤バッジ + 確認ダイアログで警告 |

実行モデルは「**実行 = 常に AutoPilot 挙動 / 完了 = 常に手動**」（kernel 3.1.5 以降）に統一されています。AutoPilot トグルは無く、operator は「どのモジュールを動かすか」だけを意思決定します。完了処理（HTML 生成・log_uploader 発火）は `[Complete]` 押下まで保留されるため、operator が成果物を確認してから明示的に finalize する運用に最適化されています。

Linear 経路（`Execute Profile`）も並走運用しており、従来の「先頭から最後まで自動」フローはそのまま使えます。FlexProfile が安定したのち Linear は撤去予定です。

## モジュール一覧

### Standard モジュール（61）

| カテゴリ | モジュール |
|---|---|
| **Network** | `hostname_config`, `ipaddress_config`, `temp_ipaddress_config`, `domain_join`, `ssid_config` |
| **Display** | `brightness_config`, `dpi_api_config`, `resolution_api_config` |
| **Desktop** | `wallpaper_config`, `taskbar_config`, `startlayout_config` |
| **Security** | `bitlocker_config`, `firewall_config`, `firewall_rule_config`, `firewall_rule_make_config`, `cert_config`, `credential_config`, `office_license_config`, `windows_license_config` |
| **User Management** | `local_user_config`, `profile_delete` |
| **Printer** | `printer_driver_config`, `printer_delete` |
| **Applications** | `app_config`, `winget_install`, `bloatware_remove`, `storeapp_config`, `odt_config`, `browser_addon_config`, `fabriq_app_launcher` |
| **Power** | `power_config` |
| **Maintenance** | `acl_config`, `copyfile_config`, `file_delete`, `office_update`, `partition_config`, `robocopy_config`, `system_finalize` |
| **System** | `autologon_config`, `default_app_config`, `driver_config`, `generic_process_runner`, `ppkg_config`, `process_killer`, `restart_config`, `restore_point`, `scheduled_task_config`, `signout_config`, `spi_config`, `sysprep_config`, `time_sync_config`, `volume_config`, `windows_feature_config` |
| **Registry** | `reg_hklm_config`, `reg_hkcu_config` |
| **Scripts** | `generic_batch_runner`, `startup_command_config` |
| **Evidence** | `evidence_config` |
| **Test** | `test_error_module`, `test_harness_config` |

`windows_update` は GUI ダッシュボードの **Windows Update** ボタン専用で、`module.csv` を持たず Script Menu には表示されません。

### Extended モジュール（18）

| カテゴリ | モジュール |
|---|---|
| **Network** | `ipv6_config`, `network_profile_config` |
| **Display** | `display_config`, `dpi_config` |
| **Desktop** | `desktop_icon_config` |
| **User Management** | `builtin_admin_config`, `group_config` |
| **Maintenance** | `directory_cleaner`, `history_destroyer` |
| **System** | `azure_ad_join_check`, `reg_template`, `server_feature_config` |
| **Scripts** | `script_looper` |
| **ManualWorks** | `manual_kitting_assistant`, `pianist` |
| **Evidence** | `log_uploader` |
| **Backup** | `printer_backup`, `userdata_backup` |

## モジュール構成

各モジュールは以下のファイルで構成されます。

| ファイル | 役割 |
|---|---|
| `module.csv` | メニュー名・カテゴリ・表示順・有効無効（1 モジュール内に複数エントリ可） |
| `<name>.ps1` | 実行スクリプト本体。`dev/template/_template_script.ps1` をベースに実装 |
| `<name>_list.csv` | 設定データ（対象リスト等） |
| `Guide.txt` | 使い方ガイド（日本語） |
| `preset.csv` | 設定 CSV のドロップダウン UI 定義（Fabriq Studio が検出。列挙型列に候補を提示） |

### 結果ステータス

全モジュールは `New-ModuleResult` または `New-BatchResult` を通じて統一ステータスを返却します。

| ステータス | 意味 |
|---|---|
| `Success` | 正常完了 |
| `Partial` | 一部成功・一部失敗 |
| `Skipped` | 全件スキップ（対象が無い等） |
| `Cancelled` | ユーザーがキャンセル |
| `Error` | エラー発生 |

Post-Apply Verification を実装したモジュールは、`-Verified $true/$false` で実行履歴に検証結果を記録します。

### preset.csv の書式

Fabriq Studio のモジュール編集画面は、`preset.csv` を検出すると該当列を編集可能なドロップダウン（ComboBox）UI に切り替えます。

```csv
Column,Value,Label
Enabled,1,有効
Enabled,0,無効
Type,REG_DWORD,32 ビット整数 (DWORD)
Type,REG_SZ,文字列 (SZ)
```

| 列 | 意味 |
|---|---|
| `Column` | 対象 CSV の列名（`module.csv` 以外の CSV ヘッダと一致） |
| `Value` | セルに実際に書き込まれる値（fabriq PowerShell が受け付ける文字列そのまま） |
| `Label` | 表示用ラベル |

エンコーディングは **UTF-8 BOM**、改行は **CRLF**。

## カスタマイズ

### hostlist.csv

対象 PC ごとの設定を定義します。

```csv
AdminID,OldPCName,NewPCName,EthernetIP,EthernetSubnet,EthernetGateway,...,Printer1Name,Printer1Driver,Printer1Port,...
1,OLD-PC-01,NEW-PC-01,192.168.1.100,255.255.255.0,192.168.1.1,...
```

機密性のあるフィールドは Fabriq Studio で `ENC:<Base64>` 形式に暗号化できます。実行時は `Import-ModuleCsv` がマスターパスフレーズで透過的に復号します。

### 新規モジュール作成

1. `dev/template/` フォルダを `modules/standard/` または `modules/extended/` にコピー
2. フォルダ名をモジュール名にリネーム（snake_case、例: `my_new_config`）
3. `module.csv`、実行スクリプト、設定 CSV を編集
4. 必要に応じて `preset.csv` を追加
5. `kernel/common.ps1` の共通関数（`Show-Info` / `New-BatchResult` / `Import-ModuleCsv` 等）を必ず使用
6. 可能な限り Post-Apply Verification（Step 5.5）を実装

開発時のガイドラインは [CLAUDE.md](CLAUDE.md) を参照してください。

### 暗号化仕様

CSV 中の機密値は `ENC:<Base64>` 形式で保持されます。

- **鍵導出**: PBKDF2-HMAC-SHA256、100,000 回、固定ソルト `fabriq-fixed-salt-2024`
- **暗号化**: AES-256-CBC、PKCS7 パディング
- **エンコード**: UTF-8（平文）、Base64（暗号文）

Fabriq Studio でフィールド単位・行単位・列単位で暗号化／復号できます。実行時にはマスターパスフレーズにより自動復号されます。

### ログ配送

`kernel/csv/log_destinations.csv` に宛先（UNC 共有フォルダ / ローカルパス）と認証情報を登録すると、`extended/log_uploader` モジュールが Transcript・実行履歴・エビデンスを外部へ自動アップロードします。

## Fabriq Studio との関係

Fabriq Studio（別プロジェクト。WPF / .NET 8.0 / リポジトリ: <https://github.com/suzugross/fabriq_studio>）は Fabriq の **GUI 管理ツール** です。本体 Fabriq との接点は以下に限定されます。

| 機能 | 本体との関係 |
|---|---|
| パスフレーズ設定 | `kernel/txt/passphrase_verify.txt` を生成（**Fabriq 起動の必須条件**） |
| ホスト管理 | `kernel/csv/hostlist.csv` を編集、`ENC:` 暗号化／復号 |
| モジュール設定管理 | 各モジュールの `_list.csv` を編集。`preset.csv` を検出してドロップダウン UI を提供 |
| プロファイル管理 | `profiles/*.csv` の作成・編集 |
| レジストリカタログ | レジストリ設定のライブラリからワークスペースへエクスポート |

本体 Fabriq は Studio のバージョン・機能に依存しません。Studio 側の具体的な機能セットは Studio リポジトリを参照してください。

## Fabriq Evidence Manager との関係

Fabriq Evidence Manager（別プロジェクト。WPF / .NET 8.0 / リポジトリ: <https://github.com/suzugross/fabriq_evidence_manager>）は、キッティング実行で本体 Fabriq が `evidence/` 配下へ出力したエビデンス（スクリーンショット・各種ログ・マニフェスト・HTML チェックリスト・実行履歴 CSV）を読み込み、GUI 上で整理・確認したうえで、顧客への**納品用フォーマット（Excel 一覧表 + 個別 PC 詳細シート + 収集アーティファクト）へ一括エクスポート**する **GUI 管理ツール** です。

エビデンス本体には一切変更を加えず、**外側から読み取るだけの consumer** として設計されています。本体との接点は、`kernel/EVIDENCE_MANIFEST.md` で明文化された **evidence manifest 公開契約（schemaVersion 1）** に限定されます。

| 機能 | 本体との関係 |
|---|---|
| マニフェスト検出 | 各 PC の `manifest.json`（`manifestType: "fabriq-evidence-manifest"` / `schemaVersion: 1`）を起点にエビデンスを認識（`kernel/EVIDENCE_MANIFEST.md` 準拠） |
| エビデンス読み込み | `pc_information/`（システム情報・ライセンス・セキュリティベースライン等の各種ログ）、`auto_capture/`（スクリーンショット）、`bitlocker/`、`checklist/`（HTML チェックリスト）、`export_history/`（実行履歴 CSV）を読み取り専用で参照 |
| ホスト突合 | `hostlist.csv` を読み込み、出力されたエビデンスと照合 |
| 納品エクスポート | `{timestamp}_fabriq_evi/` に Excel の PC 情報一覧表・個別詳細シートを生成し、収集アーティファクトを併せて出力 |

本体 Fabriq は Evidence Manager に依存しません。両者の連携は `kernel/EVIDENCE_MANIFEST.md` の公開契約を介してのみ行われ、Evidence Manager はその外部 consumer です。Evidence Manager 側の具体的な機能セット・バージョンは Evidence Manager リポジトリを参照してください。

## Fabriq Doc との関係

Fabriq Doc（別プロジェクト。リポジトリ: <https://github.com/suzugross/fabriq_doc>）は、**Fabriq シリーズ全体の技術ドキュメントを一元管理する「ドキュメント専用リポジトリ」**です。実行されるコードは含まず、各プロジェクトのソースを read-only で参照しながら、仕様・利用方法・カーネル/モジュール解説を Markdown に整備します。

- **対象**: 本体 Fabriq に加え、Fabriq Studio / Fabriq Evidence Manager / Tonebender / Tonebender Controller の計 5 プロジェクトを横断。
- **構成**: 全 md をリポジトリ直下にフラット配置し、`<project>__<category>__<name>.md` というファイル名のプレフィックスでプロジェクトとカテゴリを判別。NotebookLM 等への一括投入を一次運用として想定。
- **入口**: [`INDEX.md`](https://github.com/suzugross/fabriq_doc/blob/main/INDEX.md) が全プロジェクト統合インデックス。

本体 Fabriq は Fabriq Doc に依存しません。Fabriq Doc は本体を含む各ソースの外部 consumer（ドキュメント整備側）であり、本体側のコードや動作には一切影響しません。

## テスト

カーネルの公開関数群に対する Pester ベースのユニットテストが `tests/` 配下に存在します。

### 前提

- **Pester v5+**（Windows 同梱の v3.4.0 では実行不可）

```powershell
Install-Module -Name Pester -MinimumVersion 5.0.0 -Scope CurrentUser -Force -SkipPublisherCheck
```

### 実行

```powershell
powershell.exe -File ./dev/run_tests.ps1   # Windows PowerShell 5.1（正・カーネル実運用と同一エンジン）
pwsh ./dev/run_tests.ps1                   # PowerShell 7+ でも可（導入済み環境のみ）
```

`tests/` と `apps/fabriq_ios/tests/` 以下の `*.tests.ps1` を一括実行します。

### CI（GitHub Actions）

push（main）/ PR で `.github/workflows/ci.yml` が Windows PowerShell 5.1 上で `run_tests.ps1` / `check_version.ps1` / `check_ps1_encoding.ps1` を exit code 判定で実行します（3 チェックを独立ステップで全実行）。

### 構成

| パス | 内容 |
|---|---|
| `tests/_helpers/` | テスト共通ヘルパ（CSV 生成・モックモジュール・パス解決） |
| `tests/kernel/` | `kernel/common.ps1` 公開関数のユニットテスト |
| `apps/fabriq_ios/tests/` | fabriq_ios 用 Pester テスト |

## ライセンス

[MIT License](LICENSE) — fabriq 本体。

### サードパーティ同梱物

本リポジトリは以下のサードパーティ製ソフトウェアをバイナリ同梱しています。各コンポーネントの著作権・ライセンス条件は [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md) および [LICENSES/](LICENSES/) を参照してください。

- **7-Zip 25.01** (GNU LGPL v2.1+ ほか / Copyright © 1999-2025 Igor Pavlov / https://www.7-zip.org/) — `modules/standard/printer_driver_config/tools/` に `7z.exe` / `7z.dll` を同梱。プリンタドライバ展開用

## Author

yuki.suzuki@suzugross.com
