# bloatware_export (Standard)

**カテゴリ**: Applications
**メニュー名**: Bloatware Export
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（出力系モジュール。システム状態を変更しない）
**サブスクリプト**: なし

## 目的
PC にインストールされているデスクトップアプリケーションを `HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall`（64bit / 32bit 両ハイブ）から列挙し、CSV としてエクスポートするインベントリ採取モジュールです。出力 CSV は `bloatware_remove` モジュールの入力フォーマットに揃えてあるため、「OEM 出荷時の不要アプリを 1 度棚卸して、削除対象だけ Enabled=1 に編集する」という現場ワークフローを支援します。納品エビデンスとしての「出荷時アプリ一覧」記録としてもそのまま使えます。

## 入力 (CSV)
入力 CSV は **無し**。出力 CSV のスキーマ:
- `Enabled` … 全行 0 で出力（編集して 1 にする）
- `DisplayName` … レジストリから取得したアプリ名
- `MatchPattern` … 空欄出力（必要に応じて `McAfee*` 等のワイルドカードに編集）
- `Description` … 空欄出力（任意ラベル）
- `Segment` … 空欄出力（Segment フィルタ用）
- `Publisher` / `Version` … 参考情報（DisplayName だけでは特定しづらい時の手掛かり）

## 主要ステップ
1. （CSV 読み込みなし）
2. Pre-flight: `Test-AdminPrivilege`
3. スキャン対象（HKLM Uninstall 64bit/32bit）と出力先表示
4. 実行確認（AutoPilot 時は自動 Y）
5. レジストリ列挙 → CSV ファイルへ出力
6. 結果サマリ

## 注意点・運用メモ
- **管理者権限必須**（HKLM 読み取り権限のため）
- 出力先は `evidence\inventory\app_inventory_<日時>_<PC名>.csv`
- ファイル名の PC 名サフィックス優先順: `$env:SELECTED_NEW_PCNAME` > `$env:COMPUTERNAME`
- `evidence\inventory\` ディレクトリは自動作成
- `module.csv` の Enabled は **0**（メニューに常時出さない運用）。必要時だけ手動実行する想定
- ストアアプリ（AppX/MSIX）は対象外。デスクトップアプリ（クラシック Win32 アンインストールエントリ）のみ

## 検証
Post-Apply Verification は **不要**（出力系のためシステム状態を変更しない）。`-Verified` は未渡しで Verified 列は空欄になります。出力 CSV のファイル生成成否は `New-ModuleResult` の Status / Message で判定する設計です。CSV 内容そのものは「現状の HKLM Uninstall を反映した snapshot」であり、検証ではなくキャプチャの位置づけです。
