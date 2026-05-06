# winget_install (Standard)

**カテゴリ**: Applications
**メニュー名**: Winget App Installer Update / Winget App Install / Winget App Upgrade
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（ExitCode 判定のみ。`winget list` 再確認なし、`-Verified` 未渡し）
**サブスクリプト**:
- `winget_update.ps1` … winget 自体（Microsoft.AppInstaller）の最新化
- `winget_install.ps1` … `app_list.csv` の未インストールアプリ一括導入
- `winget_upgrade.ps1` … `app_list.csv` の既インストールアプリ一括 upgrade

## 目的
winget (Windows Package Manager) を使ったアプリケーション一括導入・更新モジュール。
3 スクリプトが 1 つの `app_list.csv` を共有し、install と upgrade で「未インストール対象」/
「インストール済対象」を自動振り分けする。winget 自体の更新を別スクリプトに切り出している
のは、古い winget では新しいパッケージリポジトリ形式が読めない問題への対策で、Profile に
組み込むときは 1) winget_update → 2) `__RESTART__` → 3) install/upgrade の順序を推奨。

## 入力 (CSV)
`app_list.csv`（install / upgrade 共有）
- `Enabled`: 1=実行 / 0=スキップ
- `AppID`: winget パッケージ ID（例 `Google.Chrome`、`Adobe.Acrobat.Reader.64-bit`）
- `Options`: 追加オプション（例 `--override "/VERYSILENT"`、空可）
- `Description`: 説明
- `Segment`: Segment フィルタ（任意）

デフォルト同梱: Google.Chrome / Adobe.Acrobat.Reader.64-bit (Enabled=1) /
Microsoft.VisualStudioCode / Git.Git (Enabled=0)。

## 主要ステップ（共通）
1. `Wait-NetworkReady` でネット接続確認
2. winget (Microsoft.AppInstaller) の利用可否確認
3. `winget source reset --force` で source リセット
4. `app_list.csv` 読み込み → 振り分け（install: 未インストール対象、upgrade: 既インストール対象）
5. ドライラン表示（[INSTALL]/[SKIP インストール済]/[UPGRADE]/[NOT INSTALLED] 色分け）
6. 実行確認（AutoPilot 自動 Y）
7. 順次実行（`winget install/upgrade --id $AppID --exact --silent --accept-source-agreements
   --accept-package-agreements [Options]`）
8. ExitCode 判定 → 結果集計 → `New-BatchResult` 返却

**振り分けマトリクス**:
| 状態 | install | upgrade |
|---|---|---|
| 未インストール | 対象 | スキップ |
| インストール済み | スキップ | 対象 |
| 最新済み | スキップ | Skipped (`-1978335212`) |

## 注意点・運用メモ
- ネット必須、管理者権限必須
- winget の `winget list --id <AppID> --exact` の出力に AppID が現れるかで冪等性判定
- ExitCode 特殊値:
  - `0` 成功 / `3010` 成功（再起動保留）
  - `-1978335212` NO_APPLICATIONS_FOUND → Skipped 扱い（upgrade で「既に最新」）
  - `-1978335189` UPDATE_NOT_APPLICABLE → Skipped 扱い
- `Options` 列でアプリごとのサイレントインストール引数を渡す（一部アプリは winget の
  既定 silent では完全無人にならないため）
- Profile 例: 初回キッティングは update→reboot→install、定期メンテは update→reboot→upgrade、
  フル更新は update→reboot→install→upgrade

## 検証
本モジュールに Post-Apply Verification は実装されていない。
install/upgrade 後の `winget list` 再確認は行わず、ExitCode のみで Success/Failed/Skipped を
決定する。`New-BatchResult` / `New-ModuleResult` に `-Verified` を渡しておらず履歴 Verified 列は
空欄。

理由としては:
1. winget の ExitCode 体系が成功/失敗/「既に最新」を明確に区別する設計のため、
   ExitCode 判定で実用上十分
2. 検証目的の `winget list` 再呼び出しはネットワーク往復を含むため遅く、
   バッチサイズが大きい Profile では実時間に影響する

手動再確認は `winget list --id <AppID>` で可能。
