# test_harness_config (Standard)

**カテゴリ**: Test
**メニュー名**: Test Harness
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 実装あり（CSV から `Verified` 列の値を集計し true/false/null を再現）
**サブスクリプト**: なし

## 目的
fabriq の動作検証用シミュレーションモジュール。ファイルやレジストリに副作用を出さずに、
Status / Verified / AutoPilot ErrorMode / Resume といった全機能パターンを再現できる
動作シミュレーター。Profile 設計時のリハーサルや、main.ps1 / common.ps1 の振る舞い変更時の
回帰テストに使う。Verified の集計、FailFirstN によるリトライ成功シナリオ、cancel Behavior による
Confirm-bypass の Cancelled 再現など、test_error_module よりも表現力が高い。

## 入力 (CSV)
`test_harness_list.csv`
- `Enabled`: 1=実行 / 0=スキップ
- `Segment`: Profile 側 Segment 値で絞り込むキー
- `TestName`: 表示名 + 状態保存キー
- `Behavior`: `success` / `fail` / `skip` / `cancel`
- `Verified`: `true` / `false` / 空欄（item ごとの Verified 結果）
- `FailFirstN`: 整数。同 (Segment+TestName) の N 回目までは強制 fail
- `DelaySec`: 各 item 処理前のスリープ秒（進捗 UI 確認用）
- `Description`: 説明

デフォルト同梱: success_verified / success_verifail / success_no_verify / skipped / partial
(2 行) / cancelled / error_basic / retry_success (FailFirstN=2) / retry_exhaust /
delay_demo (3s) / hang_sim (120s, `__ASYNC__` + Skip ボタン検証用) / async_ok など。

## 主要ステップ
1. `test_harness_list.csv` 読み込み（Segment フィルタは `Import-ModuleCsv` 共通機構）
2. ドライラン表示（各 item の Behavior / Verified / FailFirstN / DelaySec）
3. 実行確認（AutoPilot 自動 Y）
4. item ループ:
   - DelaySec 指定時はスリープ
   - cancel Behavior は即 Cancelled で全体抜ける（Confirm を介さず確実に Cancelled 再現）
   - FailFirstN 指定 + `$global:FabriqTestHarnessState["${Segment}::${TestName}"]` 累積 ≤ FailFirstN
     なら強制 fail
   - Behavior に応じて Show-Success / Show-Error / Show-Skip と count 更新
5. Verified 集計（全 true→PASS / 1 つでも false→FAIL / 空欄混在→null）
6. `New-BatchResult -Verified $verified` 返却

## 注意点・運用メモ
- 副作用ゼロ（管理者権限不要）。FailFirstN カウンタは `$global:FabriqTestHarnessState`
  ハッシュテーブル上のメモリのみ。`__RESTART__` で PowerShell プロセスが終わると自動リセット
  → ファイル / レジストリ汚染なし
- main.ps1 の `AutoPilotMaxRetry=5` と組み合わせて「N 回失敗 → N+1 で成功」シナリオ再現可能
- cancel Behavior は Confirm を介さないため Stop on Error モード時に Profile 全体停止を
  起こしうる。検証用途以外で使用しない
- 完全なサンプル Profile は `profiles/_test_harness.csv`
- 本番 Profile 厳禁

## 検証
Verified 集計が本モジュールの検証ロジックの中核。CSV の `Verified` 列値を全 item で集約し:
- 1 つでも false → モジュール全体 Verified=FAIL
- すべて true のみ → Verified=PASS
- 空欄混在（false なし） → Verified=null（検証なし扱い）

これは他モジュールの Verified 集計挙動を test_harness 側で意図的に再現できる仕組みで、
common.ps1 の `New-BatchResult -Verified` ロジックをそのまま使う。fabriq 全体の Verified 機構の
動作確認に使う「自作ルーラー」のような立ち位置。
