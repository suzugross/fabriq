# app_config (Standard)

**カテゴリ**: Applications
**メニュー名**: App Installation
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（インストーラー ExitCode のみで成否判定）
**サブスクリプト**: なし（メイン `app_config.ps1` のみ。インストーラー本体は `file/` 配下に配置）

## 目的
キッティング工程で導入する任意のサードパーティアプリケーション（Chrome / Acrobat Reader / 7-Zip など）を、CSV に記述された順序で一括サイレントインストールするためのモジュールです。インストーラー実体は `file/` ディレクトリに配置し、CSV の `FileName` 列がそれを参照します。`exe` / `msi` / `bat` の 3 種別を統一インターフェースで扱えるため、ベンダーごとに違うサイレントオプションをすべて CSV に集約できます。

## 入力 (CSV)
`app_list.csv` の主な列:
- `Enabled` … 1=実行 / 0=スキップ
- `AppName` … アプリ識別名
- `FileName` … `file/` 内のインストーラーファイル名
- `Type` … `exe` / `msi` / `bat`
- `SilentArgs` … サイレントインストール用引数文字列
- `Description` … 表示用ラベル（指定時は AppName の代わりに表示）
- `Segment` … Segment フィルタ

`Type` ごとの実行形態:
- `exe` … `& <Installer> <SilentArgs>`
- `msi` … `msiexec /i <Installer> <SilentArgs>`
- `bat` … `cmd /c <Installer>`

## 主要ステップ
1. `app_list.csv` を読み込み、Enabled=1 の行を抽出
2. `file/` ディレクトリと CSV の `FileName` 実在を突合
3. インストール対象の一覧表示（dry-run）
4. 実行確認（AutoPilot 時は自動 Y）
5. 各インストーラーを順次起動し ExitCode を判定（0=成功 / 3010=成功・再起動保留 / その他=失敗）
6. 結果集計を `New-BatchResult` で返却

## 注意点・運用メモ
- 多くのインストーラーは管理者権限を要求するため、Fabriq セッションを管理者で起動しておくこと
- **既インストール判定は実装していません**（毎回起動）。冪等性は各インストーラーのサイレントオプションに委ねる方針
- ExitCode 3010 は「成功・再起動保留」として Success 扱い。後続の `restart_config` で集約再起動する運用が想定
- MSI 実行には `msiexec.exe`、BAT 実行には `cmd.exe` が必要（標準 Windows 環境では常時 OK）
- `file/` 配下の実体配布は別途行う（GitHub にバイナリを置かない設計）

## 検証
Post-Apply Verification は **未実装**。アプリのインストール先パスや製品レジストリ存在は本モジュールでは検証しません。アプリ単位で検証要件が大きく異なる（HKLM レジストリにエントリを残すか／App-V／MSIX か等）ため、汎用検証ロジックを書くと誤判定が増える性質を持つためです。インストール痕跡の確認は `bloatware_export` で生成されるアプリインベントリ CSV を Evidence として残すことで間接的にカバーする運用です。
