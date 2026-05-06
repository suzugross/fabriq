# profiles カタログ

`e:/fabriq/profiles/` には fabriq の Profile (= モジュール実行手順書 CSV) のサンプル / テンプレートが収められています。実運用では各 Profile は **現場固有の編集物**であり、`framework_overlay_rules.json` により `profiles/` ツリー全体が overlay 時に **保持対象 (preserved)** に指定されています。つまりここに置かれているファイルは「あくまで出発点」で、実際のキッティング案件では現場で書き換えられて使われます。

## Profile CSV のスキーマ

すべての Profile CSV は最低限 `Order,ScriptPath,Enabled,Description,Segment` の 5 列を持ちます (`Segment` は省略可)。一部の Profile (sysprep.csv, _test_harness.csv) は `Note` や `ErrorMode` 列を追加で持ち、3.2.0 以降の FlexProfile 対応 Profile は `Group` 列を持つこともあります。

特殊マーカー:

- `__RESTART__` — 再起動を挿入。fabriq は再起動後 RunOnce で resume する
- `__AUTOPILOT__` — AutoPilot モード ON、`Description` に `WaitSec=N` を書くと wait 秒数指定可
- `__ASYNC__` — このマーカー以降のモジュールを非同期 (Runspace) 実行に切替

## 同梱されている Profile

### Master_Pre01.csv (1 行 / 単発)

```
10,standard/driver_config/driver_export_config.ps1,1,Driver Export
```

**シナリオ**: マスター機 (golden master) からドライバを **エクスポート**する単発作業。Pre01 は「マスター機作成側」で 1 度だけ走らせる位置付け。

### Master_Pre02.csv (9 行)

```
10  builtin_admin_config             ローカル Administrator 有効化など
20  autologon_config                 自動ログオン設定
30  __RESTART__
40  local_user_delete                既存ローカルユーザー整理
50  reg_hklm_config / 60 reg_hkcu_config
70  scheduled_task_disable_config
80  bitlocker_disable
90  driver_import_config             Pre01 でエクスポートしたドライバを適用
```

**シナリオ**: 配布先 PC の **キッティング前段階** (OS 入れ替え直後)。再起動を挟んで管理者・自動ログオン・レジストリベースライン・ドライバ適用までを行う。

### Master_Config01.csv (17 行)

ライセンスアクティベーション → 壁紙 → 時刻同期 → ファイアウォール → 電源 → ローカルユーザー作成 → レジストリ HKLM/HKCU → アプリインストール (winget) → ODT → DPI → 解像度 → ファイルコピー → プリンタドライバ / プリンタ登録、と続く **本体構成の主要 Profile**。

**シナリオ**: キッティングの中核作業。1 台あたり最も時間がかかるフェーズ。

### Master_Config02.csv (5 行)

```
10  startlayout_backup_config
20  startlayout_build_config
30  default_app_config / export_app_associations
40  desktop_icon_backup
50  fabriq_app_launcher
```

**シナリオ**: マスター機側で **スタートレイアウト・既定アプリ・デスクトップアイコン**を採取して `app_associations.xml` 等を作る。Pre01 と同じく golden master 系の作業。

### Master_Config03.csv (5 行)

```
10  taskbar_config
20  startlayout_import_config
30  startlayout_delete_config
40  desktop_icon_restore
50  default_app_config
```

**シナリオ**: Config02 で採取した成果物を配布先で **インポート / 復元**する Profile。

### Master_Config04.csv (5 行)

```
10  evidence_config           証跡採取
20  directory_cleaner         不要ディレクトリ消去
30  history_destroyer         履歴/キャッシュ消去
40  storeapp_config           Store アプリ削除
50  sysprep_config            Sysprep 構成
```

**シナリオ**: キッティング **完了直前 / 出荷前**の仕上げ。証跡を採ってから掃除して Sysprep。

### Custom Plan.csv (5 行)

```
10  reg_hklm_config / 20 reg_hklm_delete
30  __RESTART__
40  reg_hkcu_config / 50 reg_hkcu_delete
```

**シナリオ**: 案件固有のレジストリ調整を入れる **カスタムスポット**枠。再起動を挟んで HKLM 系と HKCU 系を分離。

### sysprep.csv (9 行)

`Order,ScriptPath,Enabled,Description,Segment,Note,ErrorMode` の拡張スキーマ。app_config (Enabled=0)、reg_hklm_delete、storeapp_config、taskbar_config、default_app_config、generic_process_runner、`__RESTART__` (Enabled=0)、history_destroyer、sysprep_config の順。

**シナリオ**: 出荷直前の **Sysprep 専用ライン**。`__RESTART__` をオフにしておくのは、sysprep_config 自体が再起動 / シャットダウンを最終ステップで担うため。

### _test_harness.csv (11 行)

`__AUTOPILOT__` で WaitSec=1 にした上で、`test_harness_config` モジュールの各セグメント (success_verified, success_verifail, success_no_verify, skipped, partial, cancelled, error_basic, retry_success, retry_exhaust) を順に呼ぶ統合テスト Profile。`__RESTART__` を中盤に挟んで resume 動作の検証もする。

**シナリオ**: kernel / FlexProfile / Status Monitor のリグレッションテスト。各 Status badge 描画と ErrorMode (skip / retry) の挙動、resume 動作を 1 本で検証する。

### _test_harness_async.csv (7 行)

`__AUTOPILOT__` + `__ASYNC__` を併用し、非同期実行パスのテストを行う。`hang_sim` セグメントで Status Monitor の手動 Skip も検証対象。

**シナリオ**: 非同期 Runspace 経路のリグレッションテスト。

### easy_template/ (3 ファイル)

```
easyprofile.bat       管理者昇格 + powershell -File easyprofile.ps1
easyprofile.ps1       AutoPilot 軽量ランナー (history なし / evidence なし / checklist なし)
easyprofile.csv       Enabled,Script,Description の 3 列 (Order なし、上から順実行)
```

**シナリオ**: fabriq の正規ダッシュボードを通さずに「**選択された数モジュールだけ即実行する 1-shot ランナー**」を作るためのテンプレート。`profiles/easy_profile_<X>/` という規約で複数並べて運用できる。Hostname 設定 / ローカルユーザー作成 / 削除のような単発タスクを、エビデンス無しで素早く流すユースケース。デフォルト 3 行はすべて Enabled=0 で安全側。

## 運用ルール

- `framework_overlay_rules.json` により `profiles/` ツリー全体が overlay 時に保持される (= フレームワークアップデートで現場 Profile が消えない)
- 同梱されているこれらのファイルは「**出発点 (starting points)**」であり、現場では命名 / Order / Enabled / Segment を改変して使われる
- `_test_harness*.csv` は test_harness_config モジュールに依存するため、本番現場の Profile としては使わない (開発用)
- `easy_template/` はディレクトリごとコピーして `profiles/easy_<案件名>/` の形で運用するためのスケルトン
