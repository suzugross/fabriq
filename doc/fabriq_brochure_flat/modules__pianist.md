# pianist (Extended)

**カテゴリ**: ManualWorks
**メニュー名**: Pianist
**VERSION**: 1.6.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 実装あり（Phase ごとの Manual ステータス集計から `-Verified $true/$false/$null` を返却）
**サブスクリプト**: なし。Profile 単位の構造を `profiles/<profile_name>/` 配下に持つ:
- `pianist.json` — メタデータ (label, description, target_app, default_phase)
- `procedure.csv` — Phase × Step のアクション手続マトリクス
- `values.csv` — 設定値プール（wide format、行=PC 名、列=変数）
- `shortcuts.csv` — 起動先ショートカット集（v1.0 では参照のみ）
- `instructions/<PhaseID>.txt` — Phase 手順テキスト（section marker 対応）
- `screenshots/` — author 提供の見本画像置き場（Samples タブ用）

## 目的
業務アプリの GUI 操作を伴う設定作業（ブラウザのホームページ設定、Kintone 管理画面の初期化、Office アプリの個別設定など、レジストリ/CLI で自動化できない領域）を、operator が **Profile 単位**で半自動的に進めるための GUI モジュール。`autokey_template` の正統進化版。

**位置付けの歴史**: 当初 v0.2.x は `apps/pianist/` の独立アプリだったが、履歴・スコープ問題により破棄され、2026-05-02 の決定で `modules/extended/pianist/` 配下のモジュールとして再生し fabriq 統合を完成させた経緯がある。Profile の作成・編集は Fabriq Studio に一任、本モジュールは「既存 profile を operator が Phase 単位で進めて ModuleResult を返却」する実行側に専念する。

**実行経路は 2 系統**:
1. Profile CSV 経由（推奨）: `extended/pianist/pianist.ps1, Segment=<segment_name>` を Profile に並べると、kernel が `pianist_list.csv` を Segment フィルタで絞り込み 1 profile を直接 GUI に流し込む
2. Script Menu から単発実行: ダッシュボード > Execute Module > ManualWorks > Pianist でプロファイル選択ダイアログから選ぶ

**UI ポリシー**: 全操作マウスのみ（F2 完了のみ例外）。キーボードは SendKeys でターゲットアプリへ送る専用チャネルとして温存し、Run 中の race / 迷子キー事故を防ぐ。

## 入力 (CSV)
`pianist_list.csv`（モジュール直下）:
- **Enabled**: 候補に含めるか
- **ProfileName**: `profiles/` 配下のフォルダ名（1 ProfileName = 1 phase 構造）
- **Group**: 表示用グループ名（Sample / Production など自由記述）
- **Description**: 候補一覧での説明文
- **Segment**: Profile CSV 経由の segment 名

`profiles/<name>/procedure.csv` 列: PhaseID, PhaseLabel, Color (Blue/Green/Yellow/Orange/Red/Purple/Cyan/Pink/Gray の 9 色), StepNo, Action, Value, Wait, Note。
`profiles/<name>/values.csv` は wide format（NewPCName 列 + 変数列群、`*` 行が default、PC 固有行が override、`ENC:` セル単位暗号化対応）。

## アクション 10 種
`Open` / `WaitWin` / `AppFocus` / `Type` (SendKeys) / `Key` (特殊キー) / `Wait` / `Copy` / `Paste` / `Screenshot` / `Prompt` (MessageBox 介入待ち)。

## 主要ステップ
1. `pianist_list.csv` 読み込み（Segment フィルタ適用済み）
2. 参照 profile フォルダの実在検証
3. ドライラン表示
4. `Confirm-ModuleExecution`
5. プロファイル選択（1 件なら自動、複数なら `Show-PianistProfileSelector` ダイアログ）
6. `Load-PianistProfileData` で procedure/values/shortcuts/instructions を全読み込み
7. **Step 6**: メイン GUI を絶対座標 + Anchor で構築（modeless）
8. **Step 7**: form 表示 + メインループ（DoEvents ポーリング、`UserAction` で done/cancel 判定）
9. **Step 8**: form クリーンアップ
10. **Step 9**: Phase ステータス集計 → `New-ModuleResult` を `-Verified` 付きで返却

## タブ構成（v1.5.0 以降、3 タブ）
- **[Procedure]** タブ: `instructions/<PhaseID>.txt` の `[RPA]` / `[Manual]` section を解釈して 2 段組表示。上段「RPA でやること」（Run Phase が自動実行）、下段「手動でやること」（operator が目視/クリック）
- **[Samples]** タブ (since v1.5.0): `[Samples]` section で宣言された author 提供画像のサムネイル一覧。クリックで原寸ズームのモードレスビューワを開き、Pianist 本体は引き続き操作可能。タブ見出し `Samples (N)` の N は当該 Phase の見本画像数。画像欠損時は `(missing) <filename>` placeholder
- **[Values]** タブ: 当 Phase の参照変数 + `[Show all]` トグルで全変数表示。各行 `[Copy]` ボタンで `ENC:` 復号後の値をクリップボードへ転送

## section marker（since v1.4.0、`instructions/<PhaseID>.txt` 内）
- `[RPA]` — Run Phase で自動実行される操作の説明
- `[Manual]` — operator が実施する手順
- `[Variables]` — Copy Values で表示したい変数を明示宣言（procedure.csv の `$VarName` 自動抽出と union）
- `[Samples]` — `screenshots/` 配下の見本画像ファイル名 + キャプション

後方互換: section marker が 1 つもないプレーンテキストは全文を `[Manual]` として扱う（v1.3.x 以前のサンプル無修正動作）。`[RPA]` 不在時は procedure.csv の Step 一覧を fallback 表示。

## Run-time 制御（since v1.6.0）
- **Pause**: 走っている Step / Wait が完了した時点で一時停止、再押下で再開
- **Stop**: 次の安全境界（Step 終了 / Wait 中 / WaitWin polling）で実行中断、中断後の Phase は Auto=Error 記録
- **Speed: 1.0x ⇔ 1.5x** toggle: procedure.csv の Wait・WaitWin timeout・Step 後 Wait に倍率適用（内部 settle delay 200-300ms は scale 対象外）
- `Wait-PianistResponsive`: 50ms チャンクで DoEvents をポンプしながら Sleep する責任ある待機ヘルパー（Stop/Pause を mid-wait で honor）

## 二軸ステータス
各 Phase が独立して保持する 2 軸:
- **Auto**: Run Phase の自動実行結果（NotRun / Running / Success / Partial / Error）
- **Manual**: operator 判定（Unset / OK / Warning / Error / Skip）

両方を Phase 画面下部にバッジ表示。

## 集計ロジック → ModuleResult
Manual 内訳に基づき集計:
- Error >= 1 → Status=Error, Verified=$false
- Warning >= 1 or Unset >= 1 → Status=Partial, Verified=$null
- 全て OK / Skip → Status=Success, Verified=$true
- 操作キャンセル → Status=Cancelled

return 後、kernel が Write-ExecutionHistory / Capture-ScreenEvidence / Invoke-SafeCommand を通常モジュールと同じ流儀で処理し、HTML チェックリスト・evidence summary に並ぶ。

## ベストプラクティス（Guide.txt より）
- Phase 最初の Step は `AppFocus` を必須に（Phase ボタン押下時にフォーカスが Pianist に移るため）。例外: `Open` で新規ウィンドウを起こす Phase / `WaitWin` 含む Phase
- 日本語や記号の `Type` は IME 影響を受けやすいので `Paste` を優先
- ウィンドウ待機は `WaitWin` で行い `Wait` の固定スリープに依存しない
- 機密値は Studio で `ENC:` 暗号化して values.csv に
- 重要な分岐点で `Prompt` を置き operator に目視確認させる

## 制限・既知事項
- UIA / AutomationId 操作は非対応（将来検討）
- 録画機能なし
- IF/LOOP 等の制御構文なし、線形シーケンスのみ
- Mouse 系アクションなし（座標クリックは冪等性なしのため非採用）
- 並列実行不可（1 Phase 実行中は他全ボタン disable）
- Phase 状態は session-only（Pianist 終了でリセット、永続化なし）
- Tools (ad-hoc Copy/Type, Shortcuts ランチャ) は v1.0 で非搭載
- Stop/Pause は **Step の途中**（走っている SendKeys / Open / Screenshot）では効かず、次の境界で初めて応答

## 検証
**Verification 実装モジュール**（Profile 単位の operator 判定が事実上の検証）。`pianist.ps1` 末尾で各 Phase の Manual ステータスを集計し、`Error=0 ∧ Warning=0 ∧ Unset=0` で `$verified=$true`、Error あれば `$false`、それ以外（Partial）は `$null` を `New-ModuleResult -Verified` で返却。
