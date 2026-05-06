# spi_config (Standard)

**カテゴリ**: System
**メニュー名**: SPI Config (SystemParametersInfo)
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 意図的に非対応（事前 GET 比較で代替、`project_verification_exclusions` 整合）
**サブスクリプト**:
- `spi_config.ps1`（メイン）
- 配置物: `C:\ProgramData\fabriq\apply_spi.ps1`（Active Setup / Startup Batch 有効時に CSV から動的生成）

## 目的
Win32 `SystemParametersInfo` API 経由で、レジストリ直接書き換えでは安全に制御できない
設定（特に `UserPreferencesMask` のビットフラグ群）を CSV 駆動で変更するモジュール。
視覚効果のオン・オフ、マウス速度、キーリピート速度など SPI でしか正しく扱えない項目を
担当する。Default プロファイルへの反映には Active Setup（HKLM Active Setup 登録 →
新規ユーザー初回ログオン時 cmd 経由で apply_spi.ps1 実行）と Startup Batch（Explorer
起動後トリガー）を併用する belt-and-suspenders 設計。

## 入力 (CSV)
`spi_list.csv`
- `Enabled`: 1=実行 / 0=スキップ
- `SpiAction`: SET 用 SPI 定数（16 進、例 `0x1043`）
- `ValueMode`: `bool` / `uiParam` / `pvParam`（値の渡し方）
- `Value`: 設定値（bool は 0/1）
- `Description`: 表示用
- `Segment`: Segment フィルタ（任意）

デフォルト同梱: 視覚効果 8 項目（アニメーション、影、コンボボックススライド、ヒント、
マウスポインタ影、メニュースライド、フェードアウト、リストボックススクロール）を全て OFF。

## 主要ステップ
1. `spi_list.csv` を `Import-ModuleCsv -FilterEnabled` で読み込み
2. ドライラン表示（GET action = SET action - 1 で現在値を取得し `[SKIP]` / `[APPLY]` 色分け）
3. 実行確認（AutoPilot 自動 Y）
4. SET ループ:
   - `bool` → `pvParam` に 0/1
   - `uiParam` → `uiParam` に整数
   - `pvParam` → `pvParam` に整数
5. CSV bool 項目から `apply_spi.ps1` を動的生成 → `C:\ProgramData\fabriq\` 配置
6. `Register-FabriqActiveSetup -GUID {fabriq-spi-config}` で HKLM Active Setup 登録
7. `$ENABLE_STARTUP_BATCH=$true` ならランチャー / Startup トリガー配置（`Deploy-FabriqUserSetupLauncher` /
   `Deploy-FabriqStartupTrigger`）→ `New-BatchResult` 返却（`-Verified` 未渡し）

## 注意点・運用メモ
- 管理者権限必須（HKLM Active Setup 書き込み、ProgramData 配下への配置）。
  権限不足時はカレントユーザー適用のみ
- Active Setup 対象は `bool` のみ（`pvParam`/`uiParam` はセッション固有のため除外）。
  Startup Batch は全 ValueMode 対象
- `0x0049` (MinAnimate) は SPI ではなくレジストリ書き込みで制御可能なため `reg_hkcu_config` 側で管理
- `reg_hkcu_config` と Startup Batch / Active Setup の補完機構を共有する設計（同じ
  `apply_*.ps1` ランチャーを `Deploy-FabriqUserSetupLauncher` 経由で配置）
- 初回ログオン時に Explorer が一瞬再起動する副作用あり（タスクバーが瞬間的に消える）
- WPA3SAE のような OS バージョン依存ではないが、SPI 定数は OS バージョンによって挙動差あり

## 検証
意図的に Post-Apply Verification 非対応（`project_verification_exclusions.md` の方針と整合）:
1. SPI は SET 前の GET 比較が事実上の事前検証として機能する
2. SET 直後の GET は API 呼び出し順や Active Setup / Startup Batch の後続適用との競合で
   false PASS / FAIL を返すリスクが高い
3. 新規ユーザーへの反映は Active Setup / Startup Batch に委ねる設計のため、
   カレントセッションだけ読み返してもモジュール全体の効果は保証できない

このため履歴の Verified 列は空欄。SPI GET 比較で不一致が発生した場合は SET 自体を
エラー報告する形で品質を担保する。
