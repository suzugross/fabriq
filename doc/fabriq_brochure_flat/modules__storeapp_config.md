# storeapp_config (Standard)

**カテゴリ**: Applications
**メニュー名**: Remove Store Apps
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 実装あり（AppxPackage / ProvisionedPackage の残存を再チェック）
**サブスクリプト**: なし

## 目的
不要な Microsoft Store / UWP アプリ（Cortana, Bing News, Solitaire, Xbox 系など）を一括削除する
モジュール。「現在のユーザーから削除」(`Remove-AppxPackage`) と「プロビジョニング済み
パッケージ削除」(`Remove-AppxProvisionedPackage`) の両方を行うことで、現ユーザーだけでなく
将来作成される新規ユーザーにも反映されるよう二段階で抹消する。

## 入力 (CSV)
`storeapp_list.csv`
- `No`: 表示順番号
- `AppName`: パッケージ名（例 `Microsoft.BingNews`、`Microsoft.549981C3F5F10`=Cortana 等）
- `Enabled`: 1=削除対象 / 0=残す
- `Description`: 説明
- `Segment`: Segment フィルタ（任意）

デフォルト同梱（既定で削除対象 Enabled=1）: Cortana, Bing News/Weather/Search, Gaming App,
Get Help, Get Started, Office Hub, Solitaire Collection, People, Power Automate Desktop,
Store Purchase, Todos, Alarms, Mail & Calendar, Feedback Hub, Maps, Xbox 系 4 種,
Phone Link, Groove Music, Movies & TV, Outlook for Windows, MSTeams, Quick Assist など多数。

`Get-AppxPackage | Select Name` で実環境のパッケージ名を確認可能。

## 主要ステップ
1. `storeapp_list.csv` 読み込み（Enabled=1 のみ）
2. Appx 関連 cmdlet（`Get-/Remove-AppxPackage` / `Get-/Remove-AppxProvisionedPackage`）の存在確認
3. ドライラン表示（インストール済 / プロビジョニング済の現状を表示）
4. 実行確認（AutoPilot 自動 Y）
5. アプリループ:
   - 5-1. `Get-AppxPackage -Name $AppName -AllUsers` → `Remove-AppxPackage`
   - 5-2. `Get-AppxProvisionedPackage` → `Remove-AppxProvisionedPackage -Online`
6. Step 5.5: 削除後に `Get-AppxPackage` / `Get-AppxProvisionedPackage` で再確認、
   両方残存なし → `[VERIFIED]`、片方でも残れば `[VERIFY FAILED]`
7. `New-BatchResult -Verified $verified` 返却

## 注意点・運用メモ
- 管理者権限必須（ProvisionedPackage 操作のため）
- 一部 OS 内蔵アプリは Remove-AppxPackage を拒否することがあり、その場合は
  `[VERIFY FAILED]` で計上される（OS バージョン依存）
- AppName のワイルドカード非対応（厳密一致）。新製品が出たら CSV 追記が必要
- MSTeams のような新形式パッケージは旧版 (`Microsoft.Teams`) と命名が違うため要注意
- Microsoft.WindowsCalculator のような業務必須アプリは Enabled=0 のままにする運用判断が必要
- 既定の CSV は「業務 PC のキッティング前提でのデフォルト削除セット」として整備されている

## 検証
Step 5.5 でアプリごとに `Get-AppxPackage -Name $AppName -AllUsers` と
`Get-AppxProvisionedPackage -Online` を再実行し、両方の結果が空（残存なし）の場合のみ
`[VERIFIED]` とカウント。1 件でも `[VERIFY FAILED]` があれば `-Verified=$false` で
`New-BatchResult` に渡す。

ただし「現在のユーザーセッション」と「Default プロファイル」の両方の状態は
`-AllUsers` / `-Online` フラグで一括判定しているため、特定ユーザーに残るパッケージは
検出されないケースがある。本モジュールが扱うのはあくまでマシンレベルの
プロビジョニング解除 + 全ユーザー削除であり、既存ユーザープロファイルに残るアプリの
完全クリーンアップは Profile delete 系モジュールに委ねる設計。
