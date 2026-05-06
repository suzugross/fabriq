# history_destroyer (Extended)

**カテゴリ**: Maintenance
**メニュー名**: History Destroyer
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（キャッシュ系は OS が即時に再生成し得るため false FAIL リスクから意図的に非対応）
**サブスクリプト**: なし（CSV ドリブンで 7 種類の Special ハンドラを内包）

## 目的
キッティング完了直前の最終クリーンアップとして、Windows の各種履歴・キャッシュ・一時ファイル・ブラウザデータ・Wi-Fi プロファイル等を一括破壊するモジュール。
CSV ドリブン設計で、`ActionType` 列により `DeletePath` / `ClearRegistry` / `Command`（Invoke-Expression）/ `Special`（専用ハンドラ）の 4 種類に分岐。Special には `clear-all-eventlogs` / `recycle-bin`（Shell32.dll P/Invoke の `SHEmptyRecycleBin`）/ `office-mru` / `edge-cleanup` / `chrome-cleanup` / `search-index` / `wifi-ssid` の 7 種類が用意される。

## 入力 (CSV)
`destroy_list.csv` — 削除ターゲット定義（Enabled, GroupName, TargetName, ActionType, TargetPath, IfNotFound, Description, Segment）。
`ssid_list.csv` — `wifi-ssid` ハンドラが参照する Wi-Fi プロファイル削除リスト（Enabled, SSID, Description, Segment）。

## 主要ステップ
1. P/Invoke (`SHEmptyRecycleBin`) を `Add-Type` で取り込み
2. **Step 0**: Explorer プロセス停止（ファイルロック解除）
3. CSV 読み込み + 環境変数展開（DeletePath 行のみ）
4. ドライラン表示（GroupName でグルーピング表示）
5. `Confirm-ModuleExecution`
6. メイン処理ループ: ActionType で switch 分岐、Special は `Invoke-DestroyHandler` でルーティング
7. **最終ステップ**: Explorer 自動再起動を最大 15 秒待機、復活しなければ `Start-Process explorer.exe` を強制実行
8. `New-BatchResult` 集計

## 注意点・運用メモ
- 13 系統のクリーンアップ: Explorer 履歴 / イベントログ / IME / Temp / クリップボード / DNS / ごみ箱 / Office MRU / Edge / Chrome / Search Index / サムネイル/アイコンキャッシュ / Wi-Fi
- 標準で `WSUS PingID/AccountDomainSid/SusClientId/SusClientIDValidation` の削除も含まれる（マスター展開後の重複検出回避）
- 管理者権限必須（イベントログ消去・Prefetch 削除・Wi-Fi プロファイル削除）
- Edge/Chrome のキャッシュ削除時は対応プロセスを `Stop-Process -Force`、その後 14 種類のターゲット（Cache, History, Cookies, Web Data 等）を削除
- 作業中は GUI が一時的に消える（Explorer 停止のため）

## 検証
キャッシュ系は OS が即時再生成するため読み返し検証で false FAIL を起こしやすく、設計上意図的に Verification 非対応。`-Verified` 未渡しで履歴 Verified 列は空欄。
