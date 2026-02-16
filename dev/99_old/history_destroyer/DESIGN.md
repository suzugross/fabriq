# History Destroyer — Fabriq 拡張モジュール設計書

## 概要

`dev\history_destroyer\histroy_destroyer.ps1` を Fabriq 拡張モジュールとして `modules/extended/history_destroyer/` に取り込む。

## 原本の機能 (8項目)

1. Explorer プロセス停止
2. Explorer 履歴削除 (Recent, JumpList, Registry MRU)
3. イベントビューア全ログ消去
4. Office MRU 削除
5. IME 変換履歴キャッシュ削除
6. 一時ファイル・クリップボード・DNS キャッシュクリア
7. ごみ箱を空にする
8. Explorer 再起動

## 追加提案項目

9. ブラウザ履歴削除 (Edge / Chrome)
   - Cache, History, Cookies, Session 等
   - `$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\` 配下
   - `$env:LOCALAPPDATA\Google\Chrome\User Data\Default\` 配下
10. Windows Search インデックスリセット
    - `WSearch` サービス停止 → `Windows.edb` 削除 → サービス再開
11. Thumbnail Cache 削除
    - `$env:LOCALAPPDATA\Microsoft\Windows\Explorer\thumbcache_*.db`
12. Prefetch 削除
    - `$env:windir\Prefetch\*`

## Fabriq 準拠要件

- `New-ModuleResult` でステータス返却 (Success / Partial / Error / Cancelled)
- `Confirm-Execution` で Y/N 確認 (Enter 空打ち対策済み)
- 管理者権限チェックは不要 (Fabriq 自体が管理者で起動)
- `$ErrorActionPreference = "SilentlyContinue"` → 各操作で try/catch + 個別カウント
- 日本語メッセージ → 全て英語
- `explorer.exe` 停止/再起動はオプション化 (ファイルロック対策として必要だが影響が大きい)

## ファイル構成

```
modules/extended/history_destroyer/
├── module.csv              # モジュール定義
└── history_destroyer.ps1   # メインスクリプト
```

### module.csv
```csv
MenuName,Category,Script,Order,Enabled
History Destroyer,System,history_destroyer.ps1,80,1
```

## 実装フェーズ

### Phase 1: 基本構造 + Explorer 履歴 (原本 Step 1-2) ✅ 完了
### Phase 2: システム系 (原本 Step 3, 5, 6, 7) ✅ 完了
### Phase 3: アプリケーション系 (原本 Step 4 + 追加) ✅ 完了
### Phase 4: 追加項目 + 仕上げ ✅ 完了
