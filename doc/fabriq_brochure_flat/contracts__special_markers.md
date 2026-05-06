# 特殊マーカー（Special Markers）契約

profile CSV の `ScriptPath` 列に書ける特殊識別子。`Resolve-ProfileModules` がこれらを解釈してプロファイル全体の挙動を制御する。

---

## 現行マーカー（5 種、kernel 3.2.x）

### 1. `__AUTOPILOT__`（kernel 2.0.0〜）

```csv
Order,ScriptPath,Enabled,Description,Segment,ErrorMode,Group
10,__AUTOPILOT__,1,WaitSec=3,,,
```

**動作**:
- `Enabled=1` のとき `$global:AutoPilotMode = $true` を立てる
- `Description` 列に `WaitSec=N` 形式があれば `$global:AutoPilotWaitSec` を `N` に設定（regex `WaitSec=(\d+)`）
- 効果は **以降のすべてのモジュール**に及ぶ（プロファイル末尾まで）
- ValidModules には**入らない**（メタデータ抽出のみ）

**AutoPilot の意味**:
- `Confirm-Execution` / `Confirm-ModuleExecution` が自動 Y を返す（Y/N プロンプトをスキップ）
- `Wait-KeyPress` が即 return（Press-Enter 待機をスキップ）
- 各モジュール開始前に `AutoPilotWaitSec` 秒の inter-module wait
- Error/Partial 時に `_ErrorMode` 列で分岐（`skip` / `retry` / `ask`）
- `Show-AutoPilotErrorDialog` が `ask` モードで Retry/Skip ダイアログ表示

**注意**: AutoPilot は「**確認スキップ + auto-resume**」であり「完全無人」ではない。operator が脇で見ていて状況に応じて Esc できる前提（feedback memory `feedback_autopilot_wording`）。

### 2. `__ASYNC__`（kernel 2.1.0〜）

```csv
20,__ASYNC__,1,以降を Runspace 化,,,
```

**動作**:
- `Enabled=1` かつ `async_config.json` の `Enabled=true` のときのみ効果発動（kill switch）
- 以降のすべての通常モジュールに `_IsAsync=$true` を attach（sticky）
- `Invoke-BatchExecution` が `_IsAsync` を見て `Invoke-SafeCommandAsync`（Runspace + monitoring）を選ぶ
- ValidModules には入らない
- 効果範囲: プロファイル末尾まで（途中で sync に戻すマーカーは無い）

**意味**: モジュール内ハングや長時間処理に対して、Status Monitor の Skip ボタン or `DefaultTimeoutSec` で強制中断できるようにする。詳細は §08_async_execution.md。

### 3. `__RESTART__`（kernel 2.0.0〜）

```csv
40,__RESTART__,1,再起動,,,
```

**動作**:
- ValidModules に `_IsRestart=$true` の専用 PSCustomObject として入る（MenuName=`[RESTART]`, Category=`System`）
- `Invoke-BatchExecution` の loop が検出すると：
  1. `Save-ResumeState` で `kernel/json/resume_state.json` に状態保存
  2. `Register-FabriqRunOnce` で次回起動を予約
  3. `Add-ExecutionResult [RESTART] Success` で履歴記録
  4. `Invoke-CountdownRestart -Seconds 5` で再起動（loop から return）
- 再起動後の resume で「Order > restart 行の Order」のモジュールから続行

**前提**: `autologon_config` モジュールが事前に AutoLogon を組んでいないと、再起動後に手動ログオンが必要になる（無人運用が成立しない）。

### 4. `__REEXPLORER__`（kernel 2.0.0〜）

```csv
50,__REEXPLORER__,1,Explorer 再起動,,,
```

**動作**:
- ValidModules に `_IsReexplorer=$true` で入る
- `Invoke-BatchExecution` 検出時:
  ```powershell
  Stop-Process -Name explorer -Force
  最大 15 秒、1 秒間隔で Get-Process explorer の復活を待つ
  自動復活しなければ Start-Process explorer.exe を強制発火
  Add-ExecutionResult [REEXPLORER] Success/Error
  ```
- 結果は `$completedResults` に正しい Status で記録される（kernel 3.x で改善、それ以前は hardcoded "Success" だった）

**用途**:
- HKCU レジストリ変更（reg_hkcu_config）を即時反映する
- タスクバー設定 / Start Layout / Explorer 動作系の変更後に必須
- Active Setup / Startup Batch 系のリフレッシュ

### 5. `__AUTO_to_<User>__`（kernel 2.0.0〜）

```csv
70,__AUTO_to_admin01__,1,管理者で AutoLogon 設定,,,
80,__AUTO_to_user01__,1,ユーザーで AutoLogon 設定,,,
```

**動作**:
- 正規表現 `^__AUTO_to_(.+)__$` でマッチ、`$Matches[1]` を `$autoLogonUser` として抽出
- `autologon_config` モジュールを参照し、PSCustomObject を copy
- `_AutoLogonUser` 属性を attach
- MenuName を `[AUTO:User] AutoLogon Configuration` に書き換え
- `Invoke-BatchExecution` 実行直前に `$env:FABRIQ_AUTOLOGON_USER = $module._AutoLogonUser`、終了後に `$null` でクリア
- `autologon_config` 内で `$env:FABRIQ_AUTOLOGON_USER` を読んで `autologon_list.csv` から該当行を選択

**用途**:
- 同じプロファイル内で複数の AutoLogon 切替が必要なシナリオ（admin で WU → user で BitLocker など）
- profile に `autologon_config` を直書きすると user 引数が固定されるが、`__AUTO_to_<User>__` で動的選択可能

---

## 削除済みマーカー（kernel 3.0.0 で MAJOR 破壊削除）

### `__SHUTDOWN__` / `__PAUSE__` / `__STOPLOG__` / `__STARTLOG__`

| マーカー | 旧動作 | 削除理由 |
|---|---|---|
| `__SHUTDOWN__` | プロファイル末尾でシャットダウン | 唯一の使用箇所（profile 末尾）が廃止済み、`Invoke-CountdownShutdown` 内部関数も削除 |
| `__PAUSE__` | profile 中で `Wait-KeyPress` を強制発火 | 実運用での参照ゼロ |
| `__STOPLOG__` / `__STARTLOG__` | Transcript の一時停止/再開 | 実運用での参照ゼロ |

### graceful degradation

旧プロファイルがこれらを含んでいても fabriq はクラッシュせず、`Resolve-ProfileModules` の `$invalidPaths` 経由で「module not found」warning として降格、他モジュールの実行は継続する。fabriq_studio のマーカーパレットでも既に除外されているため、新規 profile 作成時には誤って書けない。

---

## 同 Group 内 / Group 跨ぎでのマーカー挙動（kernel 3.2.0+）

FlexProfile Groups バーの `[Run: <Group>]` ボタン経由で実行する場合：

- **同 Group 内**: マーカー（`__RESTART__` / `__REEXPLORER__` / `__AUTOPILOT__` / `__ASYNC__` / `__AUTO_to_<User>__`）はすべて **そのまま実行**
- **Group 跨ぎ間のマーカー**: 当該 Group 実行時には **skip**（**literal interpretation**）

operator が RESTART を含めたい場合は明示的に Group 値を打つ：

```csv
10,standard/hostname_config/...,1,,,,Network
20,standard/ipaddress_config/...,1,,,,Network
30,__RESTART__,1,,,,Network          ← Network グループで RESTART を含める
40,standard/domain_join/...,1,,,,Network
```

`[Run: Network]` クリックで 10 → 20 → 30 (RESTART) → 再起動 → 40 まで一気通貫。

Linear `[Execute Profile]` は Group 列を無視して全マーカーを順序通り実行する（旧来挙動維持）。

---

## マーカーと `IncludeDisabled` スイッチ

`Resolve-ProfileModules -IncludeDisabled` モード（FlexProfile 専用）でも：

- `__AUTOPILOT__` / `__ASYNC__` の効果発動には `Enabled=1` が必須（disabled 行は無視）
- `__RESTART__` / `__REEXPLORER__` / `__AUTO_to_<User>__` は ValidModules に入るが `_IsCheckedDefault=$false` がつく（dashboard で初期 unchecked）

「**disabled marker 行が global state を勝手に flip しない**」安全弁。FlexProfile でプロファイル全体を可視化しても、Enabled=0 のマーカー行が誤動作することはない。

---

## マーカー選択の運用知見

| 場面 | 推奨マーカー組み合わせ |
|---|---|
| 全自動キッティング（再起動含む） | `__AUTOPILOT__` + 各モジュール + `__RESTART__` 数箇所 + `autologon_config` 事前配置 |
| 短時間モジュールが連続するセクション | 単純な順序付けのみ（マーカー不要） |
| 長時間で hang リスクある winget / WU 系 | `__ASYNC__` を冒頭に置いて以降を runspace 化 |
| Operator 立ち会い・確認しながら実行 | マーカーなし（CLI モード廃止後は Linear `[Execute Profile]` で個別 Y/N） |
| Site-specific 段階的実行 | FlexProfile + Group 列で論理グループ化 |
| HKCU レジストリ変更後の即時反映 | reg_hkcu_config の直後に `__REEXPLORER__` |

---

## マーカー追加・削除のバージョン影響

- **新マーカー追加**: 公開 API への後方互換な追加 → kernel **MINOR** 昇格（例: 2.0.0 → 2.1.0 で `__ASYNC__` 追加）
- **既存マーカー削除**: 破壊的変更 → kernel **MAJOR** 昇格（例: 2.x → 3.0.0 で 4 マーカー削除）
- **マーカーの動作変更**: 後方互換でない場合は MAJOR / 互換なら MINOR

`KERNEL_API.md §4` がマーカー仕様の真のソース。Studio のマーカーパレットは `KERNEL_API.md` を参照して動的構築すべき設計。
