# scheduled_task_config (Standard)

**カテゴリ**: System
**メニュー名**: Scheduled Task Enable / Scheduled Task Disable
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 部分実装（事前 State 比較のみ。`-Verified` 未渡し）
**サブスクリプト**:
- `scheduled_task_enable_config.ps1` … タスク有効化
- `scheduled_task_disable_config.ps1` … タスク無効化
- 共有 CSV `task_list.csv`

## 目的
Windows タスクスケジューラのタスクを CSV ベースで一括有効化 / 無効化するモジュール。
2 つのスクリプトが同じ `task_list.csv` を参照し、CSV の `Enabled` 列は「行を処理対象にするか」
であって「タスクを有効化したいか」ではない（アクションは実行スクリプトで決まる）共有 CSV 設計。
冪等で、再実行は Skip 扱いになりエラーにならない。

## 入力 (CSV)
`task_list.csv`
- `Enabled`: 1=処理対象 / 0=スキップ
- `TaskPath`: Task Scheduler フォルダパス（先頭・末尾に `\` 必須、ルートは `\` のみ）
- `TaskName`: タスク名（完全一致）
- `Description`: 表示用説明
- `Segment`: Segment フィルタ（任意）

デフォルト同梱: `Pre-staged app cleanup`（Enabled=1）/ Schedule Scan / Scheduled Start /
ScheduledDefrag / SilentCleanup（いずれも Enabled=0、参考行）

## 主要ステップ
**[Enable]**
1. `task_list.csv` を `Import-ModuleCsv -FilterEnabled` で読み込み
2. 前提チェック（`Get-ScheduledTask` cmdlet は Windows 標準のため依存なし）
3. 各タスクの現状表示 (`[Disabled]` / `[Ready]` / `[NOT FOUND]`)
4. 実行確認（AutoPilot 自動 Y）
5. 各タスクをループ: 状態が `Disabled` なら `Enable-ScheduledTask`、それ以外は Skip、
   存在しなければ Fail
6. `New-BatchResult` で集計返却

**[Disable]**
同じ流れで `Disable-ScheduledTask`、`Disabled` なら Skip。

## 注意点・運用メモ
- 管理者権限必須
- `TaskPath` の末尾 `\` 漏れは `Get-ScheduledTask` で見つからず Fail になる頻出ミス。
  Guide でも明示的に注意喚起
- 対象タスクが OS にない場合（OS バージョン差や日本語版 / 英語版差）は Warning 扱いで Fail カウント
- `task_list.csv` は Enable / Disable 両スクリプトが共有するため、誤って同じタスクを
  両方の Order に含めると最後に実行されたほうの状態に収束する（運用注意）
- AutoPilot は `Confirm-ModuleExecution` 経由で確認スキップ

## 検証
独立した Step 5.5 は実装されていないが、ステップ 3 の State 表示自体が事前検証として機能する：
2 回目実行時は全タスクが `Already enabled` / `Already disabled` で Skip と表示されるため、
これが事実上の Post-Apply Verification（同じスクリプトを再実行して全件 Skip になることで
意図状態が確定）。`New-BatchResult` に `-Verified` は渡していないため履歴の Verified 列は空欄。
タスクが見つからない (`[NOT FOUND]`) ケースは「対象 OS にそのタスクが存在しない」可能性が
あるため、Fail でも CSV 側の Enabled=0 化で運用回避するのが推奨。
