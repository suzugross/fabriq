# windows_update (Standard, Standalone)

**カテゴリ**: なし（standalone module。`module.csv` を持たず、Profile 直接登録不可）
**メニュー名**: メインメニューの `[wu]` ショートカットからのみ起動
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（次回スキャンで 0 件確認が事実上の検証、`-Verified` 未渡し）
**サブスクリプト**:
- `windows_update.ps1` … 1 回分の scan/download/install パス（メイン）
- `wu_launcher.bat` … RunOnce / 再起動後の自動起動用ランチャー
- 状態ファイル: `wu_state.json`（再起動ループ中）/ `wu_completed.json`（完了サマリ、main.ps1 が消費）

## 目的
Windows Update を Microsoft.Update.Session COM API ベースで完全自動化する standalone モジュール。
scan → download → install → reboot → 再 scan のループを自前で持ち、追加更新が無くなるまで
連続適用する。本モジュールは「fabriq の Profile 内モジュール」ではなく独立サイドカー設計で、
ループ制御（RunOnce 登録、状態 JSON ハンドオフ、再起動後再開）は `kernel/main.ps1` の
`Invoke-WindowsUpdateLoop` が担う。fabriq セッション統合のため、完了時には
`wu_completed.json` を書き、次回 main.ps1 起動時に `execution_history.csv` に取り込まれる。

## 入力 (CSV)
`windows_update_list.csv`（キー・バリュー形式）
- `MaxRebootLoops`: 最大再起動回数（既定 5）
- `ScanTimeoutMinutes` / `DownloadTimeoutMinutes` / `InstallTimeoutMinutes`: COM API タイムアウト
- `SuspendBitLocker`: 再起動前 BitLocker 一時停止 (0/1)
- `RebootCountdownSeconds`: 再起動カウントダウン
- `AutoLaunchFabriq`: 完了後 Fabriq.bat 自動起動 (0/1)
- `AutoLogonEnabled`: ワンタイム AutoLogon 設定 (0/1、autologon_list.csv 連動)
- `IncludeOptionalUpdates`: Optional Quality Updates を含める (0/1)
- `IncludeSeekerUpdates`: OptionalInstallation（段階的ロールアウト）を含める (0/1)
- `AutoInstall`: 確認スキップで全更新自動適用 (0/1)

## 主要ステップ
1. ネットワーク疎通 + COM Session 作成
2. `IUpdateSearcher::Search` で更新スキャン（Optional / Seeker フラグに応じて DeploymentAction
   `'Installation' | 'OptionalInstallation'` を加える）
3. 結果一覧表示（Optional は `[Optional]` タグ付き）
4. 確認（`-AutoConfirm` ON / `AutoInstall=1` 時はスキップ）
5. `IUpdateDownloader::Download` でダウンロード
6. `IUpdateInstaller::Install` でインストール
7. 再起動要否判定 → 必要なら `wu_state.json` に保存 + RunOnce(`FabriqWindowsUpdate`) 登録 +
   AutoLogon 設定（CSV ON 時）+ BitLocker 一時停止 → `Invoke-CountdownRestart`
8. 再起動後: `wu_launcher.bat` から本モジュール再起動 → 再 scan → 0 件で完了
9. 完了時: `wu_completed.json` 書き出し → `Invoke-WindowsUpdateLoop` 終了 →
   `AutoLaunchFabriq=1` なら Fabriq.bat 起動

## 注意点・運用メモ
- 管理者権限必須、ネット必須
- `restart_config` と並行使用禁止（両者の RunOnce 登録が競合）
- `MaxRebootLoops`（既定 5）が無限ループ防止セーフティバルブ
- AutoLogon 連動: `windows_update_list.csv` で `AutoLogonEnabled=1` のとき、
  `modules/standard/autologon_config/autologon_list.csv` から `$env:USERNAME` 一致行を引いて
  ワンタイム AutoLogon を設定（不一致なら警告のみで AutoLogon スキップ）
- COM API 操作は同期。ScanTimeout/DownloadTimeout/InstallTimeout が個別保護
- main.ps1 の `Invoke-WindowsUpdateLoop` から呼ぶときに AutoConfirm が立つ。
  AutoInstall=1 を CSV で立てれば main.ps1 を経由しない直接実行でも同じ挙動

## 検証
本モジュールに Post-Apply Verification は実装されていない。
各パスの成否は COM API の install-result HRESULT で判定し、`-Verified` は渡さない。
fabriq の strict な意味での verification は次回 scan で「0 件」が返ることが事実上の検証。
履歴の Verified 列は空欄。

設計判断としては、Windows Update 適用結果の「個々の KB が正しく入ったか」を読み返すのは
WU 自身の差分検出に委ねる方が信頼性が高く（同じ COM API を使うため）、独自 verification を
挟むと重複・矛盾リスクがある、という整理。`wu_completed.json` には
"installed N updates over M reboots" のような要約が入り、`execution_history.csv` 1 行に
畳まれる。
