# process_killer (Standard)

**カテゴリ**: System
**メニュー名**: Process Killer
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（再起動型常駐プロセス想定で読み返し検証は意図的に非実装）
**サブスクリプト**: `process_killer.ps1`（メイン処理 1 本のみ）

## 目的
指定したプロセスを強制終了するシンプルな構造のモジュール。
キッティング作業中に開いたままになりがちな Edge / Notepad / OEM 常駐ツール等を、
Profile の終盤で一括クローズしクリーンな状態にする用途。
`Get-Process -Name <ProcessName>`（拡張子なし）を入力契約とし、
既に停止している場合は Skip で冪等性を保ちます。

## 入力 (CSV)
`process_list.csv`:
- `Enabled`: 有効フラグ
- `ProcessName`: プロセス名（.exe なし。例: msedge, notepad）
- `Description`: 表示用説明
- `Segment`: Segment フィルタ

## 主要ステップ
1. `process_list.csv` 読み込み（Enabled=1 のみ、`Description` 必須）
2. ドライラン: 各エントリで `Get-Process` を実行し、[Running] / [Not Running] と
   インスタンス数を表示
3. 実行確認（AutoPilot は自動 Y）
4. 適用ループ:
   - 冪等性チェック: 該当プロセス 0 件なら Skip（再実行で副作用なし）
   - 実行中なら `Stop-Process -Name -Force` で全インスタンス強制終了
   - 例外時は失敗計上
5. `New-BatchResult` で Success / Skip / Fail 集計

## 注意点・運用メモ
- 管理者権限推奨（他ユーザー所有プロセスを終了する場合に必要）
- ProcessName は `.exe` を含めない（`Get-Process -Name` の契約）
- 常駐監視型プロセス（OneDrive, Defender 関連等）は終了直後に再起動する設計のため、
  本モジュールでは検証しない方針
- 重要なシステムプロセス（lsass, csrss 等）を指定すると OS が不安定化するリスクあり、
  CSV のレビューは慎重に

## 検証
Post-Apply Verification は意図的に未実装。
終了後すぐに同名プロセスが再起動するケース（常駐監視型サービス、autorun）を
想定した設計判断。`-Verified` を渡さないため実行履歴の Verified 列は空欄。
「kill した瞬間の事実」のみが本モジュールの責任範囲、という割り切り。
