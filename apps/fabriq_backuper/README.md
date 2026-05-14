# Fabriq BackUper

Windows ユーザデータ・プリンタ環境のバックアップ／リストア専業 satellite app。

- **Operator-facing entry**: `Fabriq_BackUper.exe`（fabriq root）
- **Internal entry**: `apps/fabriq_backuper/fabriq_backuper.ps1`
- **VERSION**: 独立 SemVer（kernel と独立）
- **設計仕様**: [SPEC.md](SPEC.md)

## 役割

fabriq セッションと**独立して**駆動し、PC キッティングの前後で実施する
backup / restore を担う。fabriq の kernel（公開 API §1〜§5）と hostlist を
read-only で再利用する疎結合 satellite。

## 起動

```
Fabriq_BackUper.exe をダブルクリック
  ↓
パスフレーズ入力（必要なら）
  ↓
ホスト選択（既セット時は skip）
  ↓
Backup or Restore モード選択
  ↓
section 選択（printer / userdata / ...）
  ↓
実行 → 結果表示
```

## バックアップ出力

```
apps/fabriq_backuper/Backup/<OldPCname>/<yyyy_MM_dd_HHmmss>/
  manifest.json                       (fabriq-backuper-snapshot schemaVersion=1)
  sections/
    printer/                          (printer section 出力)
    userdata/                         (userdata section 出力)
  _execution_log.txt
  _restore_notes.txt
```

## 実装ロードマップ

| Phase | 内容 | 状態 |
|---|---|---|
| 0 | 設計確定（SPEC.md / README.md / メモリ） | **進行中** |
| 1 | PoC scaffold + 既存 modules wrap + EXE 化 | 未着手 |
| 2 | sections 内製化（modules からの正式移植） | 未着手 |
| 3 | WinForms UI + 公開契約 docs + overlay 除外 | 未着手 |
| 4 | stable + 既存 modules 削除 | 未着手 |

## 関連リソース

- 詳細仕様: [SPEC.md](SPEC.md)
- 親フレームワーク: [Fabriq](../../README.md)
- 公開 API 契約: [kernel/KERNEL_API.md](../../kernel/KERNEL_API.md)
- 既存 backup modules（Phase 4 で削除予定）:
  - [modules/extended/printer_backup/](../../modules/extended/printer_backup/)
  - [modules/extended/userdata_backup/](../../modules/extended/userdata_backup/)
