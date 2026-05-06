# ディレクトリ構成全体図

fabriq リポジトリのファイルツリーを上から眺めた全体像。配備時の意味づけ・更新オーバーレイの境界・ランタイム生成物の配置を網羅する。

---

## トップレベル構造

```
fabriq/
├── Fabriq.exe                    ── C# エントリ（UAC 自動昇格 → main.ps1 起動）
├── Fabriq_IOS.exe                ── fabriq_ios サブプロジェクト用ランチャ（独立 SemVer）
├── Deploy.bat                    ── USB → 対象 PC へのデプロイヘルパ
├── README.md                     ── L1 に "# Fabriq verX.Y" の版表記
├── CHANGELOG.md                  ── Keep a Changelog 1.1.0 形式（[Unreleased] + リリース版）
├── CLAUDE.md                     ── Claude 開発時の絶対遵守ルール（テンプレ厳守、命名、SemVer）
├── LICENSE                       ── MIT License
├── THIRD_PARTY_NOTICES.md        ── 7-Zip 25.01 (LGPL) 等のサードパーティライセンス
│
├── kernel/                       ── ★カーネル（更新オーバーレイの中核）
├── modules/                      ── ★モジュール群（独立 SemVer）
├── apps/                         ── ★GUI ツール群（kernel bundle と同期）
├── commands/                     ── ★ユーティリティコマンド（kernel bundle と同期）
├── profiles/                     ── ★site-specific プロファイル（overlay 絶対保護）
├── dev/                          ── ★開発ツールチェーン（kernel bundle と同期）
│
├── evidence/                     ── (runtime) エビデンス出力先
└── logs/                         ── (runtime) ログ出力先
```

---

## kernel/

```
kernel/
├── KERNEL_VERSION                ── 3.2.2（カーネル API SemVer の真のソース、1 行）
├── KERNEL_API.md                 ── 公開 API サーフェス（§1〜§11）
├── EVIDENCE_MANIFEST.md          ── manifest.json 公開契約（schemaVersion=1）
├── common.ps1                    ── 90+ 共通関数ライブラリ（4371 行）
├── main.ps1                      ── メインスクリプト（1913 行、FlexProfile sub-loop / WU loop 含む）
│
├── csv/                          ── マスタ CSV（site-specific は overlay 除外）
│   ├── categories.csv            ── カテゴリ + 表示順 (framework)
│   ├── hostlist.csv              ── 対象 PC マスタ (site-specific, ENC 暗号化対応)
│   ├── workers.csv               ── 作業者 (site-specific)
│   ├── log_destinations.csv      ── ログ配送先 (site-specific, ENC)
│   └── manifesto.csv             ── マニフェスト本文 (framework, 演出)
│
├── json/                         ── ランタイム状態（overlay 除外）
│   ├── status.json               ── (runtime) Status Monitor のライブ状態
│   ├── session.json              ── (runtime) 現セッション情報
│   ├── resume_state.json         ── (runtime) 再起動跨ぎ状態（v1/v2 schema）
│   ├── async_config.json         ── (framework) __ASYNC__ Runspace 制御パラメータ
│   ├── art_pulse.txt             ── (runtime) 動作鼓動カウンタ
│   └── skip_request.flag         ── (runtime) async モジュール強制スキップ要求
│
├── ps1/                          ── カーネルサブスクリプト
│   ├── status_monitor.ps1        ── 別プロセス WinForms モニタ
│   ├── view_report.ps1           ── HTML チェックリスト単体ビューア
│   ├── manifesto.ps1             ── マニフェスト表示 GUI
│   └── art_display.ps1           ── ART 演出（status_monitor に統合済）
│
└── txt/                          ── テキストアセット（site-specific は overlay 除外）
    ├── passphrase_verify.txt     ── (site-specific) パスフレーズ検証トークン（Studio 生成、必須）
    ├── art_sentences.txt         ── (framework) ART pulse 表示文
    └── silence.flag              ── (site-specific) 演出抑制 flag（存在チェックのみ）
```

---

## modules/

```
modules/
├── standard/                     ── 標準モジュール群（60 件）
│   └── <name>/
│       ├── module.csv            ── (framework) MenuName, Category, Order, Script, Enabled
│       ├── preset.csv            ── (framework, optional) Studio 用ドロップダウン UI 定義
│       ├── <name>.ps1            ── 実行スクリプト本体（dev/template ベース）
│       ├── <other>.ps1           ── 補助スクリプト（_install / _uninstall / _backup / _restore 等）
│       ├── <name>_list.csv       ── (site-specific) 設定データ
│       ├── Guide.txt             ── (framework) 使い方ガイド（日本語）
│       ├── VERSION               ── (framework) モジュール SemVer（1 行 X.Y.Z）
│       └── REQUIRES_KERNEL       ── (framework) 要求最小カーネル版（1 行 X.Y.Z）
│
└── extended/                     ── 拡張モジュール群（15 件）
    └── <name>/                   ── 同上
```

### モジュール件数

- Standard: 60 件
- Extended: 15 件（README には 14 とあるが、現状は pianist 含む 15 件）
- 合計: 75 モジュール

---

## apps/

`fabriq_operator` を中心に、kernel bundle と同期する GUI サブプロジェクト群。

```
apps/
├── fabriq_operator/              ── ★メインダッシュボード GUI（main.ps1 から . source）
│   ├── fabriq_operator.ps1
│   └── lib/
│       ├── theme.ps1             ── 色・サイズ定数
│       ├── session_form.ps1      ── 起動時の worker / host / passphrase 一括入力
│       ├── dashboard_form.ps1    ── メインダッシュボード（タブ + ボタン）
│       ├── flex_dashboard.ps1    ── FlexProfile ダッシュボード（Groups バー含む）
│       ├── quickactions_dialog.ps1
│       └── apps_dialog.ps1       ── FabriqApps 起動ダイアログ
│
├── fabriq_ios/                   ── ★Cisco IOS 風シェル（独立 SemVer の "art" project）
│   ├── fabriq_ios.ps1
│   ├── VERSION                   ── 独立 SemVer
│   ├── SPEC.md
│   ├── README.md
│   ├── data/
│   │   ├── version_banner.txt
│   │   ├── syslog_messages.csv
│   │   ├── help_text.csv
│   │   └── module_categories.json
│   ├── lib/
│   │   ├── parser.ps1 / dispatch.ps1 / completer.ps1 / prompt.ps1
│   │   ├── shell_state.ps1 / help.ps1 / syslog.ps1
│   │   ├── modes/                ── user_exec / privileged_exec / global_config / interface_config / module_config
│   │   └── commands/             ── enable_disable / show / hostname / interface / ip_address / categories / module
│   └── tests/                    ── _phase3_smoke.ps1 .. _phase9_smoke.ps1, parser/prompt/completer の unit tests
│
├── csv_editor/                   ── 汎用 CSV 編集 GUI
├── system_launcher/              ── Windows 設定ショートカット
├── bloatware_exporter/           ── インストール済アプリ一覧 export
├── desktop_icon_backup_app/      ── デスクトップアイコン backup
├── local_user_setup/             ── ローカルユーザー作成 GUI
├── storeapp_editor/              ── storeapp_list.csv 編集
└── winget_gui/                   ── winget 操作 GUI
```

---

## commands/

```
commands/
├── gpupdate_command.ps1          ── gpupdate /force ラッパー
├── temp_command.ps1              ── 一時的なコマンド枠
├── explore_restart_command.ps1   ── Explorer 再起動（ad-hoc）
├── diag_crypto.ps1               ── 暗号化機能の診断
├── get_evidence.ps1              ── 現セッションのエビデンス収集
└── system_launcher.ps1           ── System Launcher 経由から呼ばれる
```

---

## profiles/

```
profiles/
├── Master_Pre01.csv              ── マスタ pre-config phase 1
├── Master_Pre02.csv              ── pre-config phase 2
├── Master_Config01.csv           ── マスタ config phase 1
├── Master_Config02.csv           ── phase 2
├── Master_Config03.csv           ── phase 3
├── Master_Config04.csv           ── phase 4
├── Custom Plan.csv               ── 顧客カスタマイズプラン
├── sysprep.csv                   ── Sysprep 実行用
├── _test_harness.csv             ── テストハーネス（test_harness_config 用）
├── _test_harness_async.csv       ── async モード検証用
└── easy_template/                ── EasyProfile（簡易プロファイル実行）
    ├── easyprofile.bat
    ├── easyprofile.ps1
    └── easyprofile.csv
```

profiles/ は overlay の **excludeDirsRecursive** で完全保護される（顧客カスタムが上書きされない）。

---

## dev/

```
dev/
├── template/                     ── 新規モジュールスケルトン
│   ├── _template_script.ps1      ── 7-step 正典スケルトン（Step 1..7 構造）
│   ├── _template_list.csv        ── 設定 CSV のテンプレ（Enabled, Path, Type 等）
│   ├── module.csv                ── テンプレ用 module.csv
│   ├── Guide.txt                 ── 使い方ガイドのテンプレ
│   ├── VERSION                   ── 0.1.0（開発中・未リリースの目印）
│   └── REQUIRES_KERNEL           ── 現行カーネル版に同期
│
├── framework_overlay_rules.json  ── 更新オーバーレイ契約（schemaVersion=1）
├── build_framework_patch.ps1     ── フレームワーク更新パッチ生成
├── seed_module_versions.ps1      ── 全モジュール VERSION/REQUIRES_KERNEL の baseline seed (idempotent)
├── check_version.ps1             ── KERNEL_VERSION と版表記の整合チェック
├── verify_comments_only.ps1      ── スクリプトコメント英語化検証
│
├── launcher/                     ── Fabriq.exe / Fabriq_IOS.exe の C# ソース
│   ├── README.md
│   └── (.csproj, .cs sources)
│
└── ico/                          ── アイコン素材
```

---

## evidence/（runtime）

```
evidence/
└── {Timestamp}_{PCName}_{SerialNumber}_evidence/
    └── evidence/
        ├── auto_capture/             ── Capture-ScreenEvidence の PNG（モジュール毎）
        ├── gyotaku/                  ── Save-Screenshot の手動 PNG（Status Monitor ボタン）
        ├── checklist/                ── Export-HtmlChecklist の HTML 監査レポート
        ├── export_history/           ── Export-ExecutionHistory のフルダンプ CSV
        └── pc_information/           ── evidence_config モジュールの収集結果
            └── {date}_{uid}_{pc}/
                ├── 01_SystemInfo.txt .. 22_OfficeLicense.txt
                ├── manifest.json     ── EVIDENCE_MANIFEST.md 公開契約準拠
                └── manifest.json.bak ── 前回分（1 世代保持）
```

evidence ディレクトリは overlay の **excludeDirsTopLevel** で除外（顧客固有の成果物として保護）。

---

## logs/（runtime）

```
logs/
├── {Timestamp}_{uid}_{hostname}.log  ── PowerShell Transcript（セッション毎）
└── history/
    ├── execution_history.csv         ── 全セッション通算の実行履歴
    ├── execution_history.csv.bak     ── 起動時自動 backup
    └── history_export_{ts}.csv       ── プロファイル完了時の export
```

logs/ も overlay 除外。

---

## .git / .claude（開発時のみ）

```
.git/                              ── Git リポジトリ
.claude/                           ── Claude 関連の cache / settings
```

両者とも overlay の **excludeDirsTopLevel** で除外。

---

## 配備パッケージとしての境界

| 区分 | 含まれるか |
|---|---|
| **Framework**（kernel bundle で overlay 対象） | `kernel/` (csv/json/txt の site-specific は除く), `apps/`, `commands/`, `dev/`, `Fabriq.exe`, `Deploy.bat`, README, CHANGELOG, CLAUDE.md, LICENSE |
| **Module Bundle**（per-module overlay） | `modules/{std,ext}/<name>/` (module.csv + preset.csv + scripts + Guide.txt + VERSION + REQUIRES_KERNEL のみ、`_list.csv` 等は除く) |
| **Site-Specific**（絶対保護） | `profiles/` 全体, `kernel/csv/{hostlist,workers,log_destinations}.csv`, `kernel/txt/passphrase_verify.txt`, `modules/*/_list.csv`, `silence.flag` |
| **Runtime**（除外） | `kernel/json/*`, `evidence/`, `logs/`, `.git/`, `.claude/` |

詳細は `contracts/overlay_contract.md` を参照。
