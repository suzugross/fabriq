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
