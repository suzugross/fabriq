# firewall_config (Standard)

**カテゴリ**: Security
**メニュー名**: Firewall Settings
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 実装あり（`Get-NetFirewallProfile` で読み返し）
**サブスクリプト**: なし

## 目的
Windows ファイアウォールの **3 プロファイル（Domain / Private / Public）を個別に有効／無効** に設定するモジュールです。Profile レベルの粗い粒度に専念しており、ルール単位の管理は `firewall_rule_config`（全体 backup/restore）と `firewall_rule_make_config`（個別作成）に分離されています。CSV を使う標準モード以外に、**Manual mode**（CSV が空ならインタラクティブメニューでトグル選択）と **Legacy CSV mode**（旧形式の `status` 列のみ）の 3 形態を統一スクリプトで吸収するため、レガシー資産も保護されます。

## 入力 (CSV)
`firewall_list.csv` の主な列:
- `Enabled` … 1=適用 / 0=スキップ
- `Profile` … `Domain` / `Private` / `Public`
- `Status` … `on`（有効）/ `off`（無効）
- `Description` / `Segment`

## 主要ステップ
1. 現在の 3 プロファイル状態取得＋表示
2. CSV または Manual メニューで目標状態決定
3. **冪等性チェック**: 既に期待値と一致するプロファイルは Skip（`Skipped` + `Verified=$true` で返却）
4. 変更必要なプロファイルのみ `Set-NetFirewallProfile` で適用
5. **Post-Apply Verification**: `Get-NetFirewallProfile` で再取得し期待値一致を検証
6. `New-BatchResult ... -Verified $verified` で返却

## 注意点・運用メモ
- 管理者権限必須（`Set-NetFirewallProfile` 実行のため）
- Manual mode では複数選択（カンマ区切り）や `[4] All OFF` / `[5] All ON` の一括操作にも対応
- ルール単位の制御が必要な場合は `firewall_rule_config` / `firewall_rule_make_config` を併用
- 併用時の推奨順序: `firewall_rule_config (Import)` → `firewall_rule_make_config` → `firewall_config`（profile on/off は最終調整）

## 検証
Post-Apply Verification は **実装あり**。`Set-NetFirewallProfile` 後に `Get-NetFirewallProfile` で 3 プロファイルの現状態を読み返し、CSV／Manual で指定した期待値と完全一致するか検証します。`-Verified` 付きで `New-BatchResult` に返却。Step 3 の冪等性スキップ時は早期 return で `New-ModuleResult -Status Skipped -Verified $true` を返す（既に期待状態であることが確認されているため Verified=true 確定）設計になっており、Profile レベルの状態が確実にエビデンス化されます。
