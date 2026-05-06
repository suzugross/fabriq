# bloatware_remove (Standard)

**カテゴリ**: Applications
**メニュー名**: Bloatware Remove
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 実装あり（HKLM Uninstall を再検索して残存エントリを判定）
**サブスクリプト**: なし

## 目的
OEM 出荷時にプリインストールされている不要アプリ（McAfee 系 / Lenovo Now / Microsoft 365 体験版 等）を CSV ベースで一括アンインストールします。アンインストール情報（QuietUninstallString / MSI ProductCode / UninstallString）はレジストリから **実行時に動的取得** するため、端末ごとのバージョン差異やインストールパス差異に左右されないのが特長です。`MatchPattern` のワイルドカードで「`McAfee*` ＝ McAfee 系全部」を 1 行で書けるため、CSV メンテが現実的に保てます。

## 入力 (CSV)
`bloatware_list.csv` の主な列:
- `Enabled` … 1=削除対象 / 0=スキップ
- `DisplayName` … アプリ名（MatchPattern 省略時は完全一致で検索）
- `MatchPattern` … ワイルドカードパターン（例: `McAfee*`）。1 パターンが複数アプリにマッチした場合は全削除
- `Description` … 表示ラベル
- `Segment` … Segment フィルタ

## 主要ステップ
1. `bloatware_list.csv` 読み込み + レジストリ照合（HKLM Uninstall 64bit/32bit ハイブ）
2. Pre-flight: `Test-AdminPrivilege`
3. 検出結果の一覧表示（dry-run）
4. 実行確認（AutoPilot 時は自動 Y）
5. アンインストール実行（`QuietUninstallString` → `MSI ProductCode` → `UninstallString` の優先順）
5.5 **Post-Apply Verification**: `Find-RegistryUninstallEntry` で再スキャンし、`NoRemove=1` / `SystemComponent=1` を除く残存エントリが無いか検証
6. `New-BatchResult ... -Verified $verified` で結果返却

## 注意点・運用メモ
- **管理者権限必須**
- `NoRemove=1` または `SystemComponent=1` のアプリはレジストリ値から自動判定し、対象から除外（自動回復不能なシステムコンポーネントを誤削除しないため）
- アプリのアンインストールパスを CSV に書く必要なし（実行時にレジストリから取得）
- ストアアプリ（AppX）は対象外。同種需要は `storeapp_config` モジュールで対応
- 1 つの MatchPattern が複数アプリにマッチした場合、すべて削除されるため広めのワイルドカードは事前 `bloatware_export` で確認推奨

## 検証
Post-Apply Verification を **実装あり** とする数少ないモジュールの 1 つ。Step 5.5 で `Find-RegistryUninstallEntry` をもう一度走らせ、削除対象だった行が `NoRemove`/`SystemComponent` を除いて完全に消えているかを判定します。残存があれば `[VERIFY FAILED]` をカウントし、`-Verified $verified` フラグ付きで `New-BatchResult` を返却。Evidence Manager 側で「Verified 列で削除完了が突合できる」状態になります。
