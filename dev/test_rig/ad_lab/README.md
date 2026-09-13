# AD Lab Builder

検証用の Active Directory ドメインを、JSON 定義から一発で構築する dev ツール。
`domain_join` / `group_config` / `credential_config` など **AD が無いと検証できないモジュール**のためのラボを用意する。

**検証環境専用。** パスワードは JSON に平文で書く前提（暗号化しない）。本番 AD には絶対に向けないこと。

出荷物ではない（`dev/` 配下 = Deploy に含めない）。

---

## できること / できないこと

| | |
|---|---|
| ✅ やる | 静的 IP 設定 → AD DS 役割追加 → フォレスト昇格 → **自動再起動** → OU ツリー / サービスアカウント / **OU への委任** / 事前作成コンピュータ / `ms-DS-MachineAccountQuota` / DNS フォワーダ |
| ❌ やらない | VM の作成、OS のインストール、Windows Update、スナップショット取得（すべて手作業） |

---

## 使い方

### 0. 事前準備（手作業）

1. VM を作り、**Windows Server**（評価版で可・**デスクトップエクスペリエンス**推奨。ADUC で目視確認するため）をインストール
2. 検証用クライアント VM と**同じ仮想ネットワーク**に接続
3. **Windows Update を当てる** — 2023/9 以降の更新が無いと、コンピュータアカウント再利用ブロック（KB5020276 / status `0xaac`）の挙動が再現しない
4. このフォルダ（`ad_lab.ps1` / `ad_lab.json` / `ad_lab.bat`）をサーバーにコピー

### 1. `ad_lab.json` を編集

最低限 `domain.fqdn` / `domain.netbios` / `domain.dsrmPassword` / `network.ip` を自分の環境に合わせる。

### 2. `ad_lab.bat` を実行（管理者）

ダブルクリックでよい（管理者権限が無ければ自動で昇格を促す）。

```
ad_lab.bat
```

以降は自動:

```
フェーズ1  NIC設定 → AD DS 役割追加
             ↓ 再起動保留があれば、ここで一度再起動して自動で戻ってくる
           フォレスト昇格 → 再起動
             ↓ （起動時タスクが自動で続きを実行）
フェーズ2  OU作成 → ユーザー作成 → 委任付与 → コンピュータ事前作成 → 設定 → 完了レポート
```

途中の再起動が 1 回増えることがあるのは、`Install-ADDSForest` の**前提条件チェックが「再起動保留中」を理由に昇格を拒否する**ため
（役割追加そのものと Windows Update が保留フラグを立てる）。保留を検出したら昇格せずに再起動し、起動時タスクで戻ってきてから昇格する。

ログは `C:\ad_lab\ad_lab.log`（追記）。フェーズ 2 の完了レポートで、定義した項目が `[x]` / `[ ]` で一覧表示される。

### 3. 完了したらスナップショットを取る

**このスナップショットがリセットボタン**。ラボを壊す検証（ドメイン参加テストなど）のたびにここへ戻す。

---

## 日常の操作

| やりたいこと | コマンド |
|---|---|
| 今の状態を見る（変更しない） | `powershell -File ad_lab.ps1 -Status` |
| JSON に項目を足して**追加適用** | `ad_lab.bat` をもう一度実行（冪等 — 既にある物は SKIP） |
| 別の定義で作る | `ad_lab.bat -Config other_lab.json` |
| 昇格だけして再起動を待つ | `ad_lab.bat -NoReboot` |
| ラボをリセット | **スナップショットに戻す**（スクリプトに削除機能は無い） |

状態ファイルは持たない。「まだ DC でない → フェーズ 1」「もう DC → フェーズ 2」と**実機の状態から判定**するので、何度実行しても壊れない。

---

## ad_lab.json の書き方

```jsonc
{
  "domain": {
    "fqdn": "lab.fabriq.local",       // ドメイン FQDN
    "netbios": "LABFABRIQ",           // 15 文字以内
    "dsrmPassword": "LabP@ssw0rd!",   // DSRM パスワード（複雑性要件あり）
    "forestMode": "WinThreshold",     // 省略可
    "domainMode": "WinThreshold"      // 省略可
  },

  "network": {                        // ip を空にすると NW 設定を一切触らない
    "ip": "10.1.10.20",
    "prefixLength": 27,
    "gateway": "10.1.10.30",
    "dnsForwarder": "8.8.8.8"         // 省略可（ラボから外部名前解決したい場合）
  },

  "ous": [                            // ドメインルートからの相対で書ける（DC= は書かなくてよい）
    "OU=PC",
    "OU=Kitting,OU=PC",               // 親から順に自動作成される
    "OU=Sales\\,EMEA"                 // OU 名にカンマを含む場合は \\, でエスケープ
  ],

  "users": [
    { "name": "svc_join_ok", "password": "LabP@ssw0rd!", "description": "...", "groups": [] }
  ],

  "delegations": [
    { "ou": "OU=Kitting,OU=PC", "user": "svc_join_ok", "right": "CreateComputer" }
  ],

  "computers": [
    { "name": "PRESTAGED-01", "ou": "OU=Kitting,OU=PC", "createdBy": "svc_creator" }
  ],

  "settings": {
    "machineAccountQuota": 10,
    "domainAdmins": [ "FabriqAdmin" ]  // 昇格後も DC にログオンさせたいアカウント
  }
}
```

### 昇格後のログオン（`settings.domainAdmins`）

第一 DC への昇格では、スタンドアロン時代のローカル SAM がディレクトリへ移行され、**ビルトイン Administrator がドメインの Administrator**（Domain Admins）になる。
一方、**それ以外のローカルアカウントは「ただのドメインユーザー」**として移行され、ただのドメインユーザーは**ドメインコントローラーにログオンできない**。

つまり「普段 `FabriqAdmin` のようなローカル管理者で作業していた」場合、昇格後にその名前でログオンできなくなる。`settings.domainAdmins` に入れておくと、フェーズ2 が Domain Admins に追加してログオンできる状態にする。

フェーズ1 は実行前にビルトイン Administrator が**有効かつパスワード必須**かを確認する:

- 「パスワード不要」設定 → **停止**（この状態だと昇格自体が前提条件チェックで失敗する）
- 無効 かつ `domainAdmins` が空 → **停止**（誰もログオンできなくなるため）
- 無効 だが `domainAdmins` に指定あり → 警告のみで続行

### `delegations[].right`

| 値 | 付与される権限 | 用途 |
|---|---|---|
| `CreateComputer` | 対象 OU に computer オブジェクトを**作成**する権限のみ | 通常のドメイン参加（新規アカウント作成） |
| `FullJoin` | 上記 + 削除 / パスワードリセット / `dNSHostName`・SPN の検証済み書き込み / アカウント制限の書き込み | **既存アカウントの再利用**を伴う参加 |

権限は `dsacls` ではなく **スキーマ GUID を使った ACE** で付与している。`dsacls` は権利名が**ローカライズされる**（日本語 Server では「パスワードのリセット」）ため、スクリプトが OS 言語に依存してしまうのを避けている。

### `computers[].createdBy`

指定したユーザーの資格情報でオブジェクトを作る = **そのユーザーが所有者になる**。
KB5020276 の再利用ブロックは所有者を見て判定するため、「別アカウントが作ったオブジェクトが残っている」状況を作るのに使う。

---

## domain_join 検証との対応

| JSON の要素 | 効いてくる検証 |
|---|---|
| `OU=Kitting,OU=PC` + `svc_join_ok` の `CreateComputer` 委任 | 正常系（OU 指定でそこに作られるか） |
| `svc_join_plain`（委任なし） | 権限不足で Access denied になるか |
| `PRESTAGED-01`（`createdBy: svc_creator`） | 再利用ブロック `0xaac` / 事前 staging 運用 |
| `OU=Sales\,EMEA` | DN 内のエスケープ済みカンマが実 AD で通るか |
| `OU=NoDelegation` | 委任の無い OU を指定したときの失敗形 |
| `settings.machineAccountQuota` | 既定コンテナ参加のクォータ挙動 |

---

## 既知の注意点

- **時刻同期**: DC は PDC エミュレータになるため、ハイパーバイザのゲスト時刻同期と競合すると Kerberos が壊れる。VM 側のホスト時刻同期を切るのが無難
- **ネットワークプロファイル**: クライアント側が Public のままだと通信が通らないことがある（既存リグで踏んだ罠と同じ）
- **評価版の期限**: 180 日。構築直後のスナップショットを残しておけば実害は小さい
- **`network.ip` を設定する場合**、接続済みアダプタが 1 枚であることが前提（複数あると警告を出して NW 設定をスキップする）
- スクリプトは初回実行時に自分自身と JSON を `C:\ad_lab\` へコピーし、再起動後の続きはそこから走る（共有フォルダや ISO から起動しても再開できるようにするため）
- **再起動が 2 回になることがある**（役割追加 → 再起動 → 昇格 → 再起動）。`PendingFileRenameOperations` のような benign な保留フラグでも 1 回余計に回る。保留が消えないまま 3 回再起動したら、ループを避けるためエラーで停止する（`C:\ad_lab\reboot_count.txt` がカウンタ）

---

## テスト

AD に触らない純ロジック（相対 DN 解決 / DN の分解 / 設定検証 / 委任 ACE の構成）は
`tests/dev/AdLabLogic.tests.ps1` でカバーしている。

```
powershell.exe -File ./dev/run_tests.ps1
```
