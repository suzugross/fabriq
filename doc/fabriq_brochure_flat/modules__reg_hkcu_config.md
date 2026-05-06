# reg_hkcu_config (Standard)

**カテゴリ**: Registry
**メニュー名**: Registry Config (HKCU) / Registry Delete (HKCU)
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 実装あり（Config 側のみ。Delete 側は未実装）
**サブスクリプト**:
- `reg_hkcu_config.ps1` … HKCU + Default プロファイルへの値書き込み（メイン）
- `reg_hkcu_delete.ps1` … HKCU + Default プロファイルからの値削除
- 配置物: `C:\ProgramData\fabriq\apply_hkcu.ps1`（Active Setup / Startup Batch 有効時）

## 目的
HKEY_CURRENT_USER 配下のレジストリ値を、ログオン中ユーザーの HKCU と
`C:\Users\Default\ntuser.dat`（新規ユーザーひな型）の両方に同時適用するモジュール。
Default ハイブを load して書き込むことで、本モジュール実行後に作成される新規ユーザーは
最初から同じ既定値を持つ。Explorer 初期化による上書きが問題となるキー向けに、Active Setup
登録と Startup Batch 配置（いずれもデフォルト無効）の二重補完機構も用意されている。

## 入力 (CSV)
`reg_hkcu_list*.csv` のパターンに一致する全ファイルが自動で読み込まれ集約される
（ジャンル分割管理可能、例: `reg_hkcu_list_ui.csv`）。
- `Enabled`: 有効フラグ（1=実行 / 0=スキップ）
- `AdminID`: 管理番号（表示用）
- `SettingTitle`: 設定タイトル（表示用）
- `KeyPath`: `HKEY_CURRENT_USER\...` のフルパス
- `KeyName`: 値名（`@` で既定値）
- `Type`: REG_DWORD / REG_SZ / REG_EXPAND_SZ / REG_QWORD / REG_BINARY / REG_MULTI_SZ
- `Value`: 設定値
- `Segment`: Segment フィルタ（任意）

## 主要ステップ
1. `Resolve-HkcuRoot` でログオンユーザーの HKCU 書き込み先（`HKCU:` または `HKU:\<SID>`）を解決
2. `reg_hkcu_list*.csv` を全件読み込み → Enabled=1 を抽出
3. Default ハイブ load（`reg load HKEY_USERS\Hive C:\Users\Default\ntuser.dat`）
4. ドライラン一覧表示（各エントリを `Test-RegistryValueMatch` で `[Current]`/`[Change]` 色分け）
5. 実行確認（AutoPilot 時は自動 Y）
6. 書き込みループ（HKCU 側 → HIVE 側、`$FORCE_OVERWRITE=$false` 時は一致エントリを Skip）
7. Step 5.5: Post-Apply Verification（HKCU と HIVE 双方を再読み出し、全件一致時のみ `-Verified=$true`）→ HKU:\Hive unload（失敗時 1 回リトライ）→ 必要なら Active Setup / Startup Batch 配置 → `New-BatchResult` 返却

## 注意点・運用メモ
- 管理者権限必須（Default ハイブ load/unload、HKLM への Active Setup 登録、ProgramData 配下への配置）
- 昇格セッション対策の HKCU リダイレクトを内蔵。実際の書き込み先は `Current User` または
  `<username> (via HKU)` と表示されるため意図したユーザーへの書き込みかを確認可能
- スクリプト先頭フラグ:
  - `$FORCE_OVERWRITE`（既定 `$true`）= 常時上書き、`$false` で冪等 Skip
  - `$ENABLE_ACTIVE_SETUP`（既定 `$false`）= 新規ユーザー初回ログオン時の HKCU 補完登録
  - `$ENABLE_STARTUP_BATCH`（既定 `$false`）= Explorer 起動後の Startup ショートカット経由補完
- Active Setup / Startup Batch を有効化した場合、CSV 内容から `apply_hkcu.ps1` が自動生成されるため、
  CSV を更新したらモジュール再実行で再生成する必要がある
- ハイブ unload は `[gc]::Collect()` + 1 回リトライで failsafe。失敗時は Warning のみで処理継続
- Delete 側はスクリプト構造は Config と同じだが Verification 未実装（Verified 列は空欄）

## 検証
Step 5.5 で HKCU 側と HIVE 側の各エントリを `Test-RegistryValueMatch` で読み返し、
DWord/QWord は数値比較、Binary は HEX 比較、MultiString は改行 join 比較、その他は文字列比較で
期待値一致を判定。HIVE 側は load に成功している場合のみ検証対象。1 件でも失敗すれば
`-Verified=$false` で `New-BatchResult` に渡す。Active Setup / Startup Batch の効果（新規ユーザー
初回ログオン時の挙動）はその場では検証不可。
