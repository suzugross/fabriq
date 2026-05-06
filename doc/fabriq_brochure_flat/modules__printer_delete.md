# printer_delete (Standard)

**カテゴリ**: Printer
**メニュー名**: Delete Printers
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 実装あり（`Get-Printer` で削除済み確認）
**サブスクリプト**: `printer_delete.ps1`（メイン処理 1 本。GUI 込み Windows Forms 実装）

## 目的
不要プリンタを削除するモジュールで、3 つの動作モードが共存します。
- **KeepList mode**: hostlist の Printer1..10Name + `printer_driver_config/printer_list.csv`
  TargetHost マッチ行を「残すリスト」として、それ以外の TCP/IP プリンタを削除候補化
- **Explicit mode**: `printer_delete.csv` 列挙のプリンタを名前完全一致で削除（仮想プリンタも削除可）
- **Manual mode**: フォールバック。GUI で手動選択して削除

マスタイメージに全部署プリンタを一括登録 → 各 PC は hostlist で「使うプリンタ」を宣言、
KeepList mode が自動的に他部署プリンタを掃除する部署展開シナリオが代表的ユースケース。

## 入力 (CSV)
`printer_delete.csv`（任意・Explicit mode 用）:
- `Enabled`: 有効フラグ
- `TargetHost`: 対象ホスト名（空=全ホスト、値=NewPCName 完全一致、大小文字非区別）
- `PrinterName`: プリンタ名完全一致
- `Description`, `Segment`

クロス参照: `..\printer_driver_config\printer_list.csv`（Keep List 拡張ソース）

## 主要ステップ
1. Keep List 収集
   - (a) hostlist 環境変数 `SELECTED_PRINTER_1..10_NAME`
   - (b) `printer_driver_config/printer_list.csv` の TargetHost マッチ行
2. Explicit 削除リスト収集（`printer_delete.csv` の TargetHost マッチ行）
3. インストール済みプリンタ列挙 + TCP/IP ポート（`PrinterHostAddress` 持ち）判定
4. 削除候補計算: 各プリンタに Reason ラベル（Keep / KeepList / Explicit / KeepList+Explicit / Manual）と
   PreCheck 状態を付与
5. AutoPilot 分岐:
   - AutoPilot: GUI スキップ、PreCheck=$true を一括 `Remove-Printer`、
     候補 0 件なら Skipped（全削除事故を構造的に防止）
   - Manual: ダーク基調の Windows Forms GUI（DataGridView + Select All ボタン）を表示し、
     ユーザーがチェックボックス選択 → MessageBox 確認 → 削除実行
6. Step 5.5: 削除済みプリンタを `Get-Printer -Name` で再検索し null 確認
7. `New-ModuleResult -Verified` で集計返却

## 注意点・運用メモ
- 管理者権限必須
- KeepList mode は TCP/IP ポート紐付きプリンタのみ対象。仮想プリンタ
  （Microsoft Print to PDF / OneNote / XPS / Fax）は安全装置で対象外、
  削除したい場合は Explicit mode で名前明記
- `printer_driver_config` 配下の `printer_list.csv` をクロス参照することで、
  Register Printers で登録したプリンタが自動的に KeepList にも入る二重管理防止設計
- AutoPilot では `$global:AutoPilotMode` が kernel 側で設定される
- GUI 配色は fabriq 標準のダークテーマ（bgDark=ARGB(30,30,30) 等）

## 検証
Post-Apply Verification を実装。削除実行した各プリンタ名について
`Get-Printer -Name <name>` を実行し、`$null` 返却（つまり削除済み）を確認。
1 件でも残存していれば `-Verified $false`。削除実行 0 件（GUI でキャンセル等）の場合は
Cancelled を返し、Verification はスキップ。
