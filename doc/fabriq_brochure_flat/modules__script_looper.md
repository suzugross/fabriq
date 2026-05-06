# script_looper (Extended)

**カテゴリ**: Scripts
**メニュー名**: Script Looper
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（対象スクリプト側が自身の Verified を返すべきという責務分離。ループ側では読み返しなし）
**サブスクリプト**: なし

## 目的
任意の `.ps1` スクリプトを「条件付きでリトライ・ループ実行」する汎用モジュール。
NIC 設定の DHCP 解放待ち、プリンタドライバ DL の一時失敗、Azure AD Join の完了待ち、外部サービスへの接続不安定などを「時間経過＋リトライで解決」させるためのフレームワーク。`azure_ad_join_check` モジュールはこの looper との組み合わせを主用途として設計されている。

## 入力 (CSV)
`looper_list.csv`:
- **Enabled**: 有効フラグ（必須）
- **ScriptPath**: 対象 `.ps1` パス（必須、fabriq ルート相対パス推奨 / 絶対パス可）
- **MaxRetry**: 最大実行回数（必須、初回含む。1=リトライなし、3=初回+2 回リトライ）
- **IntervalSec**: リトライ前の待機秒数（必須、0 以上）
- **Condition**: `OnError`（Error 時のみリトライ） / `Always`（結果問わず MaxRetry まで実行）
- **Description**: 説明
- **Segment**: Segment フィルタ

## 主要ステップ
1. CSV 読み込み
2. 各エントリで ScriptPath 解決（絶対パスならそのまま、相対なら `Get-Location` 起点で `Join-Path`）+ パラメータ検証（MaxRetry≥1, IntervalSec≥0, Condition∈{OnError,Always}）。結果を `_PathValid` / `_ValidParams` プロパティとして付与
3. ドライラン表示（`[READY]`/`[NOT FOUND]`/`[INVALID]` 色分け）
4. `Confirm-ModuleExecution`
5. 各エントリのリトライループ:
   - 対象スクリプトを `& $scriptPath` で実行
   - **Dual ModuleResult detection**: パイプライン出力から `_IsModuleResult=$true` を持つオブジェクト探索 → なければ `$global:_LastModuleResult` を fallback 参照（kernel の `Invoke-KittingScript` と同じパターン）
   - レガシースクリプト（ModuleResult 返却なし）は例外なし=Success / 例外発生=Error として扱う
   - Condition=OnError なら Error 時のみ次試行、Always は常に次試行（最終試行は必ず終端）
   - `IntervalSec > 0` なら `Start-Sleep`、0 なら即時リトライ
6. 各エントリの最終 `lastStatus` で集計（Error なら failCount、それ以外なら successCount）
7. `New-BatchResult` で全体集計

## 注意点・運用メモ
- 無限ループ防止のため MaxRetry は必須。0 以下や非数値は `[INVALID]` でスキップ
- ScriptPath 不在は `[NOT FOUND]` でスキップ（Error にしない）
- 対象スクリプトは `New-ModuleResult` を返す fabriq 準拠形式が推奨。レガシー対応は fallback 機能だが正確なステータス検出はできない
- 集計ロジック: 全成功→Success / 混在→Partial / 全失敗→Error / 全スキップ→Skipped

## 検証
未実装。Guide.txt 内に明示された設計判断として「対象スクリプトの責務として検証を行うべき（各モジュールが自身の Verified を返すべき）」「ループ側では読み返しを行わない」。`-Verified` 未渡しで Verified 列は空欄。
