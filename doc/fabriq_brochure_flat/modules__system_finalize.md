# system_finalize (Standard)

**カテゴリ**: Maintenance
**メニュー名**: System Finalize
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（Explorer 再起動 + キャッシュ再生成で false FAIL を招くため意図的に未実装）
**サブスクリプト**: なし

## 目的
キッティング作業の総仕上げとして、シェル環境のリフレッシュを行うモジュール。
shell32.dll の再登録（File Type Association などのデフォルトを再構築）、Explorer 停止、
アイコン / サムネイル / レガシーアイコンの各キャッシュファイル削除、Explorer 再起動を
1 連で実行する。多数の個別設定モジュール適用後に「OS 側のシェルキャッシュが古い」状態を
クリーンアップして反映を確実にする位置づけ。

## 入力 (CSV)
設定 CSV なし（操作内容が完全固定のため）。

## 主要ステップ
1. 対象キャッシュファイルの存在状況を表示（`%LOCALAPPDATA%\Microsoft\Windows\Explorer\`
   配下の `iconcache*.db` / `IconCache*.db` / `thumbcache*.db` / レガシー `IconCache.db` 等）
2. 実行確認（AutoPilot 自動 Y）
3. 順次実行:
   - 3-1. `regsvr32 /s /i:U shell32.dll` で shell32.dll 再登録
   - 3-2. Explorer 停止 (`Stop-Process` / taskkill)
   - 3-3. iconcache*.db / IconCache*.db 削除
   - 3-4. thumbcache*.db 削除
   - 3-5. レガシー IconCache.db 削除
   - 3-6. Explorer 再起動
4. 結果集計表示

## 注意点・運用メモ
- 管理者権限必須（shell32.dll 再登録、システム Explorer の停止 / 再起動に必要）
- 対象キャッシュは現在のユーザー (`%LOCALAPPDATA%`) 配下のみ。他ユーザーは対象外
- 実行中はタスクバーとデスクトップが一時的に消える（Explorer 停止のため）副作用あり
- ロックされたファイルは削除できないことがあり、その場合は Warning 表示で他は継続
- Order=99 は `restart_config` (99) や `signout_config` (100) と並ぶ「最終位置」帯。
  Profile では Sysprep の前 / リスタート前に置いて使うのが想定運用
- ICONCACHE 系の不可視属性 (Hidden) ファイルは Force 指定で削除

## 検証
本モジュールに Post-Apply Verification は実装されていない。
主処理が「Explorer 再起動」であり、Explorer は再起動直後に OS 側でアイコンキャッシュ等を
即座に再生成し始めるため、削除直後にファイル存在チェックを行うと「再生成された新ファイル」
を検出して false FAIL 扱いになるリスクがある。`-Verified` を `New-BatchResult` に
渡していないため履歴の Verified 列は空欄。

実効性確認はオペレーターの目視（タスクバーアイコンが正しく更新されたか、
ピン留めアプリが意図通り表示されるか）に委ねる設計。
