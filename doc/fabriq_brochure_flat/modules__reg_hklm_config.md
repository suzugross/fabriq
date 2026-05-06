# reg_hklm_config (Standard)

**カテゴリ**: Registry
**メニュー名**: Registry Config (HKLM) / Registry Delete (HKLM)
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 実装あり（Config 側のみ。Delete 側は未実装）
**サブスクリプト**:
- `reg_hklm_config.ps1` … HKLM への値書き込み（メイン）
- `reg_hklm_delete.ps1` … HKLM の値削除

## 目的
HKEY_LOCAL_MACHINE 配下のレジストリ値を CSV ベースで一括適用・削除するモジュール。
ファイアウォール無効化、Ctrl+Alt+Del 必須化解除、高速スタートアップ無効化など、
マシン全体に効くポリシー系・OS 内部設定系のレジストリを一元管理する。
冪等性ヘルパー `Test-RegistryValueMatch` を内蔵し、6 つの値種（REG_DWORD/SZ/QWORD/BINARY/
MULTI_SZ/EXPAND_SZ）ごとに型に応じた比較を行う。

## 入力 (CSV)
`reg_hklm_list*.csv` のパターンに一致する全ファイルが自動で読み込まれ集約される
（ジャンル分割管理可能、例: `reg_hklm_list_security.csv`, `reg_hklm_list_ui.csv`）。
- `Enabled`: 有効フラグ（1=実行 / 0=スキップ）
- `AdminID`: 管理番号（表示用）
- `SettingTitle`: 設定タイトル（表示用）
- `KeyPath`: `HKEY_LOCAL_MACHINE\...` のフルパス
- `KeyName`: 値名
- `Type`: REG_DWORD / REG_SZ / REG_EXPAND_SZ / REG_QWORD / REG_BINARY / REG_MULTI_SZ
- `Value`: 設定値
- `Segment`: Segment フィルタ（任意）

## 主要ステップ
1. `reg_hklm_list*.csv` を全件読み込み（複数ファイル集約）
2. Enabled=1 のみ抽出 → 0 件なら Skipped
3. ドライラン表示（`Test-RegistryValueMatch` で `[Current]`/`[Change]` 色分け、
   `$FORCE_OVERWRITE=$true` の場合は FORCE MODE 表示）
4. 実行確認（AutoPilot 時は自動 Y）
5. 書き込みループ（型別キャスト後 `Set-ItemProperty` / `New-ItemProperty`、
   `$FORCE_OVERWRITE=$false` 時は一致エントリを Skip）
6. Step 5.5: Post-Apply Verification（全エントリを再読み出しして期待値一致を確認）
7. `New-BatchResult -Verified $verified` で結果返却

## 注意点・運用メモ
- 管理者権限必須
- CSV は Shift-JIS（ANSI）保存推奨。BOM 付き UTF-8 はヘッダー読み込みで不具合の可能性
- 先頭フラグ `$FORCE_OVERWRITE`（既定 `$true`）で冪等性挙動を切り替え。`$false` 時は
  目標値と一致するエントリは `[Skip]` 扱い、Verification は常に実行
- 新規キーの作成（`New-Item -Path -Force`）と既存値の上書き両方をサポート
- Binary 値は CSV 上のアスキー HEX を `[byte[]]` に変換
- DWord/QWord は `[int]` キャスト前提（10 進数表記）
- Delete 側 (`reg_hklm_delete.ps1`) は Verification 未実装で履歴の Verified 列は空欄

## 検証
Step 5.5 で全エントリに対し `Test-RegistryValueMatch` を再実行。型別比較ロジックは
HKCU 版と同一で、DWord/QWord は数値比較、Binary は HEX 比較、MultiString は改行 join 比較、
その他は文字列比較。1 件でも失敗があれば `-Verified=$false` で `New-BatchResult` に渡す。
ファイアウォール無効化のように OS 側で同等の項目が複数経路で制御される設定でも、
このモジュールはレジストリ値の書き込み一致のみを保証し、ランタイム挙動の検証は行わない
（その層は対応する `firewall_config` 等に分担させる設計）。
