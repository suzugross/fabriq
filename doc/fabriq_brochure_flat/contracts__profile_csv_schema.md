# Profile CSV スキーマ契約

`KERNEL_API.md §4` で公式宣言。fabriq の最重要契約のひとつ。プロファイル CSV はキッティングシナリオを宣言する DSL であり、kernel と fabriq_studio の双方が依存する。

---

## ファイル配置

```
profiles/<ProfileName>.csv
```

例: `profiles/Master_Pre01.csv`, `profiles/Custom Plan.csv`, `profiles/sysprep.csv`

エンコーディングは Default（PS5.1 の `Import-Csv -Encoding Default` で読める SJIS / UTF-8 BOM のいずれか）。改行は CRLF。

---

## 列定義

| 列 | 必須 | 用途 | 由来版 |
|---|---|---|---|
| `Order` | 必須 | 整数・昇順実行。実行履歴の一級識別子（同 MenuName が複数行ある場合の per-row state matching に使う） | 2.0.0 baseline |
| `ScriptPath` | 必須 | `{standard,extended}/<module>/<script>.ps1` 形式 or 特殊マーカー。区切りは `/` `\` どちらも可 | 2.0.0 baseline |
| `Enabled` | 必須 | `1`=実行 / `0`=スキップ | 2.0.0 baseline |
| `Description` | 任意 | プロファイル UI 表示用コメント。`__AUTOPILOT__` 行では `WaitSec=N` 形式で wait 秒指定 | 2.0.0 baseline |
| `Segment` | 任意 | `Import-ModuleCsv` の Segment フィルタ値として渡される。`<name>_list.csv` の Segment 列と厳密マッチ | 2.0.0 baseline |
| `ErrorMode` | 任意 | AutoPilot 時のエラー処理（空=ダイアログ確認 / `skip` / `retry` 最大 5 回） | 2.0.0 baseline |
| `Group` | 任意 | FlexProfile dashboard の Groups バー集約名。Linear `Execute Profile` は無視 | **3.2.0** |

---

## 行の例

```csv
Order,ScriptPath,Enabled,Description,Segment,ErrorMode,Group
10,__AUTOPILOT__,1,WaitSec=3,,,
20,standard/hostname_config/hostname_config.ps1,1,ホスト名設定,,,Network
30,standard/ipaddress_config/ipaddress_config.ps1,1,IP アドレス設定,,retry,Network
40,__RESTART__,1,再起動,,,
50,standard/reg_hklm_config/reg_hklm_config.ps1,1,レジストリ設定,office,skip,Tweaks
60,standard/reg_hklm_config/reg_hklm_config.ps1,1,レジストリ設定（home），home,skip,Tweaks
70,__AUTO_to_admin01__,1,管理者で AutoLogon 設定,,,
80,__ASYNC__,1,以降を Runspace 化,,,
90,standard/winget_install/winget_install.ps1,1,Winget アプリ,,,Apps
100,standard/evidence_config/evidence_config.ps1,1,Evidence 収集,,,Evidence
```

---

## 特殊マーカー（5 種、kernel 3.0.0 で 4 種を破壊的削除）

### 現行マーカー

| マーカー | 動作 | 由来版 |
|---|---|---|
| `__AUTOPILOT__` | 以降を AutoPilot 化（Y/N 自動承認 + 指定 wait 秒のモジュール間スリープ）。`Description` に `WaitSec=N` で wait 秒指定 | 2.0.0 |
| `__ASYNC__` | 以降を Runspace 実行に切り替え。Status Monitor の Skip ボタン or `async_config.json` の `DefaultTimeoutSec` で強制中断可能 | 2.1.0 |
| `__RESTART__` | Windows 再起動 → RunOnce 経由で resume | 2.0.0 |
| `__REEXPLORER__` | Explorer 再起動（HKCU レジストリ変更の即時反映等） | 2.0.0 |
| `__AUTO_to_<User>__` | `autologon_config` を該当 User で呼び出し | 2.0.0 |

### 削除済みマーカー（kernel 3.0.0 / MAJOR）

`__SHUTDOWN__` / `__PAUSE__` / `__STOPLOG__` / `__STARTLOG__` の 4 種を削除。

#### 削除理由

実運用での参照ゼロ（`__PAUSE__` / `__STOPLOG__` / `__STARTLOG__`）または唯一の使用箇所も廃止済み（`__SHUTDOWN__`）。fabriq_studio のマーカーパレットでも既に除外されており、UX 上は事実上 deprecated だった。

#### 既存プロファイル互換

削除後のマーカーを含む旧プロファイルは `Resolve-ProfileModules` の `$invalidPaths` 経由で「module not found」warning として降格、kernel はクラッシュせず他モジュールの実行を継続する（**graceful degradation**）。

---

## マーカー単位の挙動詳細

### `__AUTOPILOT__`

- `Enabled=1` のときのみ効果発動（`IncludeDisabled` モードでも disabled 行は無視される、安全弁）
- ValidModules には入らず、戻り値の `AutoPilot` フラグを `$true` にする
- `Description` に `WaitSec=3` のような形式で待機秒指定（regex `WaitSec=(\d+)`）
- 効果範囲: プロファイル末尾まで（off にするマーカーは無い）

### `__ASYNC__`

- `Enabled=1` かつ `async_config.json` の `Enabled=true` のときのみ効果発動（kill switch）
- ValidModules には入らない
- 以降のすべての通常モジュールに `_IsAsync=$true` を attach（sticky）
- 効果範囲: プロファイル末尾まで

### `__RESTART__`

- ValidModules に `_IsRestart=$true` の専用 PSCustomObject として入る
- `Invoke-BatchExecution` の loop 内で検出されると：
  1. `Save-ResumeState` で現状態を json 保存
  2. `Register-FabriqRunOnce` で次回起動を予約
  3. `Invoke-CountdownRestart -Seconds 5` で再起動（return）
- 再起動後の resume で「Order > restart 行の Order」のモジュールから続行

### `__REEXPLORER__`

- ValidModules に `_IsReexplorer=$true` で入る
- 検出時: `Stop-Process explorer -Force` → 最大 15 秒待機 → 自動復活しなければ `Start-Process explorer.exe`
- HKCU レジストリ変更（reg_hkcu_config）後の即時反映に使う

### `__AUTO_to_<User>__`

- 正規表現 `^__AUTO_to_(.+)__$` でマッチ
- 抽出された User 名が `_AutoLogonUser` として attach
- `autologon_config` モジュールが内部で `$env:FABRIQ_AUTOLOGON_USER` を読んで該当 user を `autologon_list.csv` から探す
- MenuName は `[AUTO:User] AutoLogon Configuration` に書き換えられる

---

## Group 列セマンティクス（kernel 3.2.0+）

### 動作モデル

- 同一 `Group` 値の行群を FlexProfile dashboard の **Groups バー** で `[Run: <Group>]` ボタンとして集約
- 1 クリックで該当 Group の全モジュールを Order 昇順で `Invoke-BatchExecution` する
- 実行は `AutoPilot=$true` + `FinalizeOnComplete:$false`（完了は operator が `[Complete]` で手動）
- Group 内の Order 順序は保たれる
- 空文字列 / 列自体の欠落 = 「グループ無所属」

### Linear への影響

Linear `[Execute Profile]` は本列を参照しない（旧来挙動維持）。Linear から見れば `_Group` 属性が module オブジェクトに増えるだけで無害。

### __RESTART__ との関係（literal interpretation）

Group 跨ぎ間の `__RESTART__`（Group 値が異なる）は当該 Group 実行時には **skip** される。**literal interpretation**: Group 列が batch を厳密に決定する契約。

operator が RESTART を含めたい場合は明示的に Group 値を打つ：

```csv
Order,ScriptPath,...,Group
10,standard/hostname_config/...,Network
20,standard/ipaddress_config/...,Network
30,__RESTART__,Network          ← Network グループで RESTART を含める
40,standard/domain_join/...,Network
```

`[Run: Network]` クリックで 10 → 20 → 30 (RESTART) → 再起動 → 40 まで一気通貫。

---

## __RESTART__ + AutoPilot の組み合わせ

最も強力な組み合わせ。再起動跨ぎを含む長尺シーケンスを完全 unattended 化（operator 立ち会い前提だが介入不要）：

```csv
10,__AUTOPILOT__,1,WaitSec=3
20,standard/hostname_config/hostname_config.ps1,1,
30,standard/ipaddress_config/ipaddress_config.ps1,1,
40,__RESTART__,1,
50,standard/domain_join/domain_join.ps1,1,
60,__RESTART__,1,
70,standard/bitlocker_config/bitlocker_config.ps1,1,
80,standard/evidence_config/evidence_config.ps1,1,
```

各 `__RESTART__` で resume_state.json + RunOnce + AutoLogon（autologon_config 経由）が組み合わさり、operator の介入なしで再起動 → 自動ログオン → 続行 → 再起動 → 自動ログオン → 続行 → 完了 まで進む。

---

## エラー処理マトリクス（ErrorMode 列）

| ErrorMode | AutoPilot=on | AutoPilot=off |
|---|---|---|
| 空文字 / `ask` | `Show-AutoPilotErrorDialog`（Retry/Skip 選択） | エラー記録のみ、続行 |
| `skip` | warning 出力 → エラー記録のまま続行 | 同左（off でも適用） |
| `retry` | autoRetryCount を最大 5 回まで増やして再実行、超えたらエラー記録 | 同左 |

ErrorMode は AutoPilot 中に最も意味を持つ（無人運用での自動復旧）。skip / retry は AutoPilot off 時もそのまま動く（operator が立ち会っていても自動 retry が便利な場合に対応）。

---

## Segment フィルタ仕様

- profile CSV 行の `Segment` 値が `$env:FABRIQ_SEGMENT` に渡される
- モジュール内の `Import-ModuleCsv` が `Segment` 列を持つ `_list.csv` に対して **厳密一致**（empty vs empty も match）でフィルタ
- 同モジュールを設定値別に呼び分けるパターンに使う（同じスクリプト、別データ）

例: 上の例の `Order=50` と `Order=60` は同じ `reg_hklm_config.ps1` を走らせるが、`reg_hklm_list.csv` の `Segment=office` 行と `Segment=home` 行を別々に適用する。

UI 上の MenuName 表示も `Registry HKLM Configuration [seg:office]` のようにサフィックス付きになる。

---

## デフォルトプロファイルの自動生成

`profiles/` ディレクトリが空のとき、`Create-DefaultProfiles` が以下 2 つを生成：

### Basic Setup.csv

```csv
Order,ScriptPath,Enabled,Description
10,standard\hostname_config\hostname_config.ps1,1,Change Hostname
20,standard\ipaddress_config\ipaddress_config.ps1,1,IP Address Settings
30,standard\domain_join\domain_join.ps1,1,Domain Join
```

### Full Setup.csv

検出された全モジュールを Order 10, 20, ... で並べた巨大プロファイル。

これらはあくまで初期値で、operator が Studio で site-specific プロファイルを作るのが本来の運用。
