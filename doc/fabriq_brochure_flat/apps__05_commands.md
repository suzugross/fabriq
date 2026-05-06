# commands/ — 手動操作コマンド集

`e:/fabriq/commands/` は、**Status Monitor の手動操作**または **System Launcher 経由**で操作員が任意のタイミングで呼び出すユーティリティスクリプト群です。Profile 経由で自動実行されるモジュールとは異なり、ここに置かれるスクリプトは「操作員が必要だと判断したとき」に走らせる手元の道具箱としての位置付けで、エビデンス採取や履歴記録の対象外であることが多いのが特徴です。

## ファイル一覧

| ファイル | 役割 | 起動経路 |
|---|---|---|
| `gpupdate_command.ps1` | グループポリシー強制更新 | Status Monitor / System Launcher |
| `temp_command.ps1` | カスタム差し込み用テンプレート (空) | Status Monitor |
| `explore_restart_command.ps1` | Explorer 再起動 | Status Monitor |
| `diag_crypto.ps1` | 暗号化 / passphrase 状態の診断 | Status Monitor / 開発者手動 |
| `get_evidence.ps1` | PC 情報の手動採取 (split log 形式) | Status Monitor / FabriqApps |
| `system_launcher.ps1` | Windows 設定ショートカットパレット (apps/system_launcher と同一実装) | Status Monitor |

---

## gpupdate_command.ps1

### Role
`gpupdate /force` を実行してグループポリシーを即時適用する。ドメイン参加直後に GPO ベースのポリシーを反映させたい場合に手動で呼ぶ。

### 中身
- `gpupdate /force` を try/catch で実行
- 結果に応じて [SUCCESS] / [ERROR] を色付き出力
- 末尾 `pause` でユーザーが結果を確認してから閉じる

### 起動コンテキスト
Status Monitor の手動ボタンから、または System Launcher の管理ツール枠経由。Profile に組み込むケースは少ない (組み込むなら専用モジュール化される)。

---

## temp_command.ps1

### Role
**空のテンプレート**。中身は `# ========================================` の枠 + `# template` の 1 行のみ。現場ごとに必要な手作業をその場で書き込んで使う「使い捨てスロット」。

### 中身
```
# ========================================
# template
# ========================================
```

### 起動コンテキスト
Status Monitor の手動枠に「temp」ボタンとして固定で表示される。ここに案件固有の臨時処理を書き込んで実行する想定。

---

## explore_restart_command.ps1

### Role
`explorer.exe` を停止 → Windows の自動再起動を 15 秒まで待機 → 起動したか確認するコマンド。レジストリ変更の即時反映、Start レイアウト適用後のタスクバー / デスクトップ再描画などに使う。

### 中身
- `Confirm-Execution` で Y/N 確認 (common.ps1 経由)
- `Stop-Process -Name explorer -Force`
- 1 秒間隔で 15 秒まで起動を polling
- 復帰しなかった場合は警告

### 起動コンテキスト
Status Monitor。デスクトップやタスクバーが一時的に消える操作なので、ユーザー確認必須。

---

## diag_crypto.ps1

### Role
fabriq の暗号化機能 (`Unprotect-FabriqValue` / DPAPI / passphrase) が正しく動いているかを **診断**する。CSV 内に `ENC:` プレフィックスの暗号化値があるかを scan し、復号成否を試す。

### 中身
1. `$global:FabriqMasterPassphrase` の有無確認
2. `Unprotect-FabriqValue`, `Import-ModuleCsv`, `Import-CsvSafe` 関数の存在確認
3. fabriq 配下の `*.csv` を再帰スキャンして `ENC:` パターンを含むファイルを列挙
4. 各 ENC 値に対して復号試行、成否を表示

### 起動コンテキスト
- 開発者の手動診断
- Status Monitor からトラブルシュート用に呼ぶ (パスフレーズ未投入で `[ENCRYPTED]` のまま動いてしまう問題の切り分け)

CLAUDE.md memory: `project_crypto_security_review.md` で挙げた懸念の確認用。

---

## get_evidence.ps1

### Role
PC 情報を手動採取して `evidence/pc_information/` 配下にテキストログとして保存する。`evidence_config` モジュールが定型化したエビデンス採取とは別経路の **緊急時 / 個別採取**用。

### 中身
- 出力先: `$global:FabriqEvidenceBasePath\pc_information` (未設定時は `evidence/pc_information/<COMPUTERNAME>_<yyyyMMdd_HHmmss>/`)
- マスターログ `_ALL_<COMPUTERNAME>_Log.txt` + セクション別 split log (例: `01_SystemInfo.txt`)
- `Out-Log` 関数で「Console / Master / Split」3 箇所同時出力
- `Start-Section -Title -FileName` でセクション切替
- 採取対象: 基本情報 (hostname / OS / spec)、その他 PC 構成

### 起動コンテキスト
Status Monitor、または FabriqApps から「evidence_config が動かなかったとき / その前後で個別採取したいとき」用に手動起動。

---

## system_launcher.ps1

### Role
`apps/system_launcher/system_launcher.ps1` と **同一実装**。commands/ にも置かれているのは、Status Monitor から手動操作枠で 1-click 起動できるようにするため。

### 中身
- 34 項目の Windows ツール (ms-settings: 系 / *.cpl / *.msc / shell:::{GUID} / cmd / powershell / runas)
- `Invoke-Tool` で Type に応じた Start-Process 起動分岐

### 起動コンテキスト
Status Monitor の手動ボタン枠。apps/ 側の利用と機能は同じ。
