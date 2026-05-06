# generic_batch_runner (Standard)

**カテゴリ**: Scripts
**メニュー名**: Batch Runner
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（バッチ内部処理は観測不能）
**サブスクリプト**: `batch_runner.ps1`（メイン処理 1 本のみ。`assets/` 配下に対象 .bat を同梱）

## 目的
任意の `.bat` ファイル群を CSV 定義に従って順次実行する汎用ランナー。
事前ダウンロード済みドライバー導入バッチや顧客固有のレガシー .bat 資産を
そのまま fabriq の Profile フローに差し込みたいケース向けで、CSV にタイムアウトと
成功判定 ExitCode を持たせ、複数 .bat を一括で結果集計します。
バッチ内部のロジックには介入しないため「ブラックボックス資産の橋渡し」用途に最適です。

## 入力 (CSV)
- `Enabled`: 有効フラグ（1=実行, 0=スキップ）
- `Description`: 表示用の説明
- `BatchPath`: .bat ファイルパス（相対は `$PSScriptRoot` 基準、絶対パスも可）
- `Arguments`: バッチへ渡す引数（省略可）
- `TimeoutSec`: タイムアウト秒数（0 / 空欄 = 無制限）
- `SuccessCodes`: 成功とみなす ExitCode（カンマ区切り、省略時は 0 のみ）
- `Encoding`: エンコーディング（保留列、現状は表示のみ）
- `Segment`: Segment フィルタ（共通機構が暗黙適用）

## 主要ステップ
1. `batch_list.csv` を `Import-ModuleCsv -FilterEnabled` で読み込み
2. 実行対象一覧の存在確認とドライラン表示（[NOT FOUND] は警告）
3. 実行確認（`Confirm-ModuleExecution`、AutoPilot は自動 Y）
4. `cmd /c "path"` を `Start-Process -NoNewWindow -PassThru` で起動、PID 表示
5. `TimeoutSec` 経過時は `Stop-Process -Force` で強制終了し失敗扱い
6. ExitCode を `SuccessCodes` リストと照合し成否判定
7. `New-BatchResult` で Success/Skip/Fail を集計返却

## 注意点・運用メモ
- バッチ自身が管理者権限を必要とする場合、Profile 全体を管理者で起動する必要あり
- `Encoding` 列は予約済みだが、現実装では cscript/cmd 起動時の codepage 制御は未実装
- 多重起動防止はバッチ側責務。再実行による副作用がある .bat は冪等化を推奨
- 出力はリダイレクトせずコンソール直行のため、ログ取得はバッチ側で `>>` に書く

## 検証
Post-Apply Verification は未実装。バッチ内部の事後状態（レジストリ／ファイル／サービス）は
モジュール側からは不可視のため、`-Verified` パラメータは渡さず、実行履歴の Verified 列は空欄。
詳細検証が必要な場合は、各 .bat に `reg query` / `sc query` 等の自己検証ステップを組み込み、
ExitCode 経由で fabriq に結果を伝える設計が推奨。
