# browser_addon_config (Standard)

**カテゴリ**: Applications
**メニュー名**: Browser Addon Config
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 実装あり（HKLM ForceList を再検索）
**サブスクリプト**: なし

## 目的
Chrome / Edge の拡張機能を **グループポリシー（HKLM レジストリ）経由** で強制インストール対象に登録するモジュールです。書き込み先は Chrome が `HKLM\SOFTWARE\Policies\Google\Chrome\ExtensionInstallForcelist`、Edge が `HKLM\SOFTWARE\Policies\Microsoft\Edge\ExtensionInstallForcelist`。Chrome Web Store URL を貼り付けるだけで 32 文字 ID を自動抽出し、既存 Forcelist の最大インデックス +1 を計算して衝突なく追記する設計のため、CSV メンテが極めて軽量です。ブラウザ再起動または `gpupdate /force` で拡張がダウンロード／インストールされます。

## 入力 (CSV)
`browser_addon_list.csv` の主な列:
- `Enabled` … 1=登録 / 0=スキップ
- `Browser` … `Chrome` または `Edge`
- `ExtensionId` … 32 文字 ID 直接、または Chrome Web Store URL（自動抽出）
- `Description` … 表示ラベル
- `Segment` … Segment フィルタ

ExtensionId 解決は 3 段階フォールバック: (1) 32 文字 a-p の ID 直接 → (2) `/detail/` 以降から抽出 → (3) 文字列中の最初の 32 文字 a-p 系列を抽出。

## 主要ステップ
1. `browser_addon_list.csv` 読み込み
2. 前処理（Browser 値検証、`Resolve-ExtensionId`、不正は ERROR フラグ）
3. Dry-run 表示（`[Current]` / `[Change]` / `[ERROR]` で色分け、書き込み先パスも併記）
4. 実行確認（AutoPilot 時は自動 Y）
5. 設定適用ループ（既登録は Skip、新規はキー作成 + `Get-NextForcelistIndex` で次インデックスを採番し `<id>;https://clients2.google.com/service/update2/crx` 形式で書き込み）
5.5 **Post-Apply Verification**: 全対象を `Test-ExtensionInForcelist` で再検証し `[VERIFIED]` / `[VERIFY FAILED]`
6. `New-BatchResult ... -Verified $verified` で返却

## 注意点・運用メモ
- **管理者権限必須**（HKLM 書き込み）
- ブラウザ自体は本モジュールではインストール／確認しない（レジストリのみ）
- 反映にはブラウザ再起動 or `gpupdate /force` が必要
- **同一 ID の二重登録は発生しない**（冪等性あり: `Test-ExtensionInForcelist` で判定）
- update_url は `clients2.google.com/service/update2/crx` 固定。プライベート CRX レジストリには非対応
- 既存 Forcelist に他拡張が登録されていてもそれらは保持し、追記のみ
- 削除機能は本モジュールにはない（`reg_hklm_delete` 等で対応）
- Edge 拡張も Chrome Web Store の update_url を使用（Microsoft 推奨動作）

## 検証
Post-Apply Verification は **実装あり**。Step 5.5 で `Test-ExtensionInForcelist` をもう一度実行し、CSV で対象とした ID 全件がレジストリに登録されていることを検証します。`-Verified $verified` で `New-BatchResult` に返却。ただし「ブラウザが実際に拡張をダウンロード／インストールしたか」までは本モジュールでは検証範囲外で、最終確認は `chrome://extensions` / `edge://extensions` で行う運用です。
