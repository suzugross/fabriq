# カーネル全体像

**現行版**: `kernel/KERNEL_VERSION` = `3.2.2`（fabriq ver3.2 — *Manifeste du Surkitinisme*）

## カーネルとは何か

`kernel/` は fabriq フレームワークの**運転系**である。モジュールを動かす実行基盤・GUI・状態管理・暗号化・ロギング・エビデンス取得・再起動跨ぎ・更新オーバーレイ契約のすべてを抱え込む層であり、モジュール側からは「公開 API」（`KERNEL_API.md`）で宣言されたサーフェスのみが利用可能。

**役割の分担**:

```
Fabriq.exe (C#ランチャ)
   ↓ UAC 自動昇格
   ↓ 管理者権限の PowerShell コンソールを開く
kernel/main.ps1
   ↓ . source
kernel/common.ps1（90+ 関数の共通ライブラリ）
   ↓ dot-source
apps/fabriq_operator/fabriq_operator.ps1（WinForms ダッシュボード）
   ↓ 操作
modules/{standard,extended}/<name>/<name>.ps1（実モジュール）
```

## カーネルが担う責務（カテゴリ別）

| カテゴリ | 責務 | 主な関数・ファイル |
|---|---|---|
| **エントリポイント** | UAC 昇格 → 管理者権限取得 → main.ps1 起動 → GUI ダッシュボードへ受け渡し | `Fabriq.exe`（C#）, `kernel/main.ps1` |
| **共通ライブラリ** | 表示・確認ダイアログ・CSV 読み込み・結果オブジェクト生成・暗号化・履歴・エビデンス取得・スリープ抑制・コンソール制御の共通関数群（90+） | `kernel/common.ps1` |
| **モジュールシステム** | `modules/{standard,extended}/<name>/module.csv` の自動検出・カテゴリ別グルーピング・カテゴリ順序管理 | `Initialize-ModuleSystem`, `Build-CategoryMenu`, `kernel/csv/categories.csv` |
| **プロファイル解決** | profile CSV の `Order` / `ScriptPath` / `Enabled` / `Segment` / `ErrorMode` / `Group` を読み込み、特殊マーカーを解釈してモジュールリストへ展開 | `Resolve-ProfileModules`, `Load-Profiles` |
| **オーケストレーション** | プロファイル一括実行・モジュール単発実行・AutoPilot ループ・FlexProfile sub-loop・Windows Update リブートループ | `Invoke-BatchExecution`, `Invoke-KittingScript`, `Invoke-FlexProfileLoop`, `Invoke-WindowsUpdateLoop` |
| **エラー制御** | モジュール例外を吸収して `ModuleResult` を取り出し、AutoPilot 中の `ErrorMode` (skip/retry/ask) を分岐 | `Invoke-SafeCommand`, `Invoke-SafeCommandAsync`, `Show-AutoPilotErrorDialog` |
| **CSV 読み込み + 暗号化透過復号** | `Import-CsvSafe` → 必須列検証 → `ENC:` 値の AES-256-CBC 復号 → `Segment` フィルタ → `Enabled` フィルタ | `Import-ModuleCsv`, `Unprotect-FabriqValue`, `Test-MasterPassphrase` |
| **再起動跨ぎ** | `__RESTART__` 検出 → 状態保存（`resume_state.json`）→ RunOnce 登録 → 再起動 → 復帰時に `Wait-SystemReady` → `Invoke-AutoResumeCountdown` → 環境復元 → 残モジュール継続 | `Save-ResumeState`, `Load-ResumeState`, `Register-FabriqRunOnce`, `Invoke-CountdownRestart`, `Invoke-AutoResumeCountdown`, `Restore-HostEnvironment` |
| **エビデンス自動収集** | モジュール実行ごとにスクリーンショット PNG 保存 + 実行履歴 CSV 追記 + プロファイル完了時に HTML チェックリスト生成 | `Capture-ScreenEvidence`, `Save-Screenshot`, `Write-ExecutionHistory`, `Export-HtmlChecklist`, `Initialize-EvidenceBasePath` |
| **ステータスモニタ** | 別プロセス（`status_monitor.ps1`）を起動し、`status.json` をリアルタイム書き込み → 別ウィンドウで進捗・PC 情報・ART pulse を可視化 | `Write-StatusFile`, `Start-StatusMonitor`, `Stop-StatusMonitor`, `Write-ArtPulse` |
| **AutoPilot / AutoConfirm** | プロファイル一括実行で Y/N 確認を自動承認、モジュール間ウェイトを設定 / FlexProfile 単発実行では Y/N と Press-Enter のみ短絡 | `$global:AutoPilotMode`, `$global:AutoConfirmMode`, `Confirm-Execution`, `Wait-KeyPress` |
| **セッション管理** | 作業者選択・媒体シリアル取得・`session.json` 保存・パスフレーズ検証 | `Initialize-Session`, `Test-MasterPassphrase`, `Reset-FabriqState` |
| **公開契約** | カーネル公開 API（§1〜§5）/ 更新オーバーレイ契約（§9）/ Evidence Manifest 契約（§10） | `KERNEL_API.md`, `dev/framework_overlay_rules.json`, `kernel/EVIDENCE_MANIFEST.md` |

## 保管場所マップ（kernel 配下）

```
kernel/
├── KERNEL_VERSION          ── 3.2.2（カーネル API SemVer の真のソース）
├── KERNEL_API.md           ── 公開 API サーフェスの明文化（§1〜§11）
├── EVIDENCE_MANIFEST.md    ── manifest.json 公開契約（外部 evidence consumer 向け）
├── common.ps1              ── 90+ 関数の共通ライブラリ（4371 行）
├── main.ps1                ── エントリスクリプト・FlexProfile sub-loop・Windows Update ループ（1913 行）
├── csv/
│   ├── categories.csv      ── カテゴリと表示順マスタ
│   ├── hostlist.csv        ── 対象 PC マスタ（暗号化フィールド対応）
│   ├── workers.csv         ── 作業者マスタ
│   ├── log_destinations.csv── ログ配送先マスタ（log_uploader 用）
│   └── manifesto.csv       ── マニフェスト本文（演出機能）
├── json/
│   ├── status.json         ── ステータスモニタ用ライブ状態（atomic write）
│   ├── session.json        ── 現セッション情報（worker, media serial, start time）
│   ├── resume_state.json   ── 再起動跨ぎ時の状態スナップショット（v1/v2 schema）
│   ├── async_config.json   ── __ASYNC__ Runspace 制御パラメータ
│   ├── art_pulse.txt       ── 動作鼓動カウンタ（演出用、Show-* で +1）
│   └── skip_request.flag   ── async モジュール強制スキップ要求の flag ファイル
├── ps1/
│   ├── status_monitor.ps1  ── 別プロセス WinForms モニタ（status.json を polling）
│   ├── view_report.ps1     ── HTML チェックリストの単体ビューア
│   ├── manifesto.ps1       ── マニフェスト表示 GUI
│   └── art_display.ps1     ── ART 演出（status_monitor に統合済）
└── txt/
    ├── passphrase_verify.txt ── パスフレーズ検証トークン（Studio で生成、起動必須）
    ├── art_sentences.txt   ── ART pulse で表示する一文集
    └── silence.flag        ── 演出抑制 flag（存在すれば ART を黙らせる）
```

## カーネル API のポリシー

- **真のソースは `kernel/KERNEL_VERSION`**。`README.md` L1 / `common.ps1` L2 / `main.ps1` L3 の版表記は `X.Y` 桁で同期。
- **公開 API は `KERNEL_API.md` の §1〜§5 のみ**。これに記載のない `common.ps1` 関数は内部実装で、PATCH バージョンでも予告なく変更されうる。
- **モジュール側は `REQUIRES_KERNEL` ファイル（1 行 `X.Y.Z`）で要求カーネル版を宣言**。更新オーバーレイ時に `REQUIRES_KERNEL > 現行 KERNEL_VERSION` ならカーネル先行更新を強制。
- **API 変更は `KERNEL_API.md` の同コミット更新が必須**（CLAUDE.md ルール G）。MINOR/MAJOR 昇格に必ず追従。
