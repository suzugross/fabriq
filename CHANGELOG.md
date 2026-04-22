# Changelog

本プロジェクトのすべての注目すべき変更はこのファイルに記録されます。

フォーマットは [Keep a Changelog 1.1.0](https://keepachangelog.com/ja/1.1.0/) に基づき、
バージョニングは [Semantic Versioning 2.0.0](https://semver.org/lang/ja/) に従います。

カテゴリの意味:
- **Added**: 新機能
- **Changed**: 既存機能の変更（後方互換）
- **Deprecated**: 将来削除予定の機能
- **Removed**: 削除された機能
- **Fixed**: バグ修正
- **Security**: セキュリティ関連

## [Unreleased]

### Changed
- バージョン管理をカーネル API とモジュール独立の 2 軸に再設計
  - 新規: `kernel/KERNEL_VERSION`（カーネル API 真のソース、初期値 `2.0.0`）
  - 新規: `kernel/KERNEL_API.md`（公開 API サーフェスの明文化）
  - 新規: `dev/template/VERSION`（新規モジュール用、初期値 `0.1.0`）
  - 廃止: ルート直下の `VERSION` ファイル（「全体版」概念の廃止）
  - `CLAUDE.md`: バージョン管理ルールを全面改訂（実装前宣言ルール E、
    実装サマリ報告ルール F、`KERNEL_API.md` 同期ルール G、モジュール
    `VERSION` 運用ルール H を追加）
  - `dev/check_version.ps1`: `KERNEL_VERSION` 基準に切り替え
  - `README.md` L1 / `kernel/common.ps1` L2 / `kernel/main.ps1` L3 /
    `main.ps1` 起動表示 / HTML チェックリストフッター を `2.0` に同期

### Notes
- `KERNEL_VERSION=2.0.0` は現行カーネルの状態を遡及的に formal SemVer
  の出発点として定義した値。過去の `VERSION=2.2.0` は「全体ディストリ版」
  という別概念だったため単純な規格合わせではなく、意味論的に新しい系列
- ランタイムでのモジュール互換性チェックは導入しない（誤判定で現場が
  止まるリスクを避けるため）。代わりに Claude が実装前後で
  `KERNEL_API.md` を参照して手動で整合性を担保する

## [2.2.0] - 2026-04-22

### Added
- `VERSION` ファイルをルート直下に追加（単一の真実のソース）
- `CHANGELOG.md` を新設（Keep a Changelog 形式）
- `dev/check_version.ps1` を追加（版番号整合性チェックスクリプト）
- `CLAUDE.md` に「バージョン管理ルール」セクションを追加

### Changed
- `README.md` / `kernel/common.ps1` / `kernel/main.ps1` の版番号表記を `2.2.0` に統一
  （従来 README と main.ps1 が `ver2.1`、common.ps1 が `v2.2` で不整合）

### Notes
- 本バージョンは過去コミット（Git 237 件）をまとめた初回タグ基準点。
  以降、機能追加・変更は必ず `[Unreleased]` へ追記し、リリース時に本形式で昇格する。
