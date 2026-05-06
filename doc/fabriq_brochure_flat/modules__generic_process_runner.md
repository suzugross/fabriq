# generic_process_runner (Standard)

**カテゴリ**: System
**メニュー名**: Process Runner
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（起動 EXE の処理内容は観測不能）
**サブスクリプト**: `process_runner.ps1`（メイン処理 1 本のみ）

## 目的
任意の EXE を CSV 定義から順次起動する汎用ランナー。
Office 更新（OfficeC2RClient.exe）、サイレントインストーラ、コマンドラインツール等、
EXE パスと引数で完結する処理全般に使用します。
親プロセスが即終了して子プロセスが本処理を続けるブートストラップ型 EXE 向けに
`WaitProcessName` でポーリング待機する仕組みを持ち、`generic_batch_runner` の EXE 版に相当します。

## 入力 (CSV)
- `Enabled`: 有効フラグ
- `Description`: 表示用説明
- `ExecutablePath`: EXE パス（絶対 / 環境変数展開 / 相対のいずれも可）
- `Arguments`: 起動引数
- `WorkingDirectory`: 作業ディレクトリ（空欄時はモジュールフォルダ、相対 EXE パスの解決基点にも使用）
- `TimeoutSec`: タイムアウト秒数（0 / 空欄 = 無制限）
- `SuccessCodes`: 正常 ExitCode（カンマ区切り、例 `0,3010`）
- `NoNewWindow`: ウィンドウ制御（1=コンソール内実行、空欄/0=新規ウィンドウ）
- `WaitProcessName`: ポーリング待機する子プロセス名（拡張子なし）
- `Segment`: Segment フィルタ

## 主要ステップ
1. `process_list.csv` 読み込み + 環境変数展開
2. 各エントリの存在確認とドライラン表示（パス / Timeout / SuccessCodes / Window / WaitFor を表示）
3. 実行確認（AutoPilot は自動 Y）
4. `Start-Process` で EXE 起動 + PID 表示
5. `Wait-Process -Timeout` でタイムアウト監視。超過時は `Stop-Process -Force`
6. `WaitProcessName` 指定時は 5 秒間隔で当該プロセスの消滅をポーリング待機
7. ExitCode を SuccessCodes と照合 → `New-BatchResult` で集計

## 注意点・運用メモ
- パス解決順は「環境変数展開 → 絶対パスならそのまま → 相対なら WorkingDirectory or `$PSScriptRoot`」
- GUI アプリは `NoNewWindow=0`、コンソールツールは `NoNewWindow=1` 推奨
- ブートストラップ型 EXE（OfficeC2RClient 等）は必ず `WaitProcessName` を設定し、
  早すぎる「完了判定」で次モジュールに進む事故を防ぐ
- 冪等性なし。Enabled=1 のエントリは毎回 EXE が再起動されるため、
  多重実行リスクは EXE 側で処理する設計

## 検証
Post-Apply Verification は未実装。EXE 起動後の状態変化はモジュール側からは観測不能で、
ExitCode のみが成否判定根拠。`New-BatchResult` に `-Verified` は渡さない。
事後検証が必要な場合は別モジュール（例: Office 更新後に `office_license_config` の `/dstatus` 解析）で
状態を読み返す設計が推奨。
