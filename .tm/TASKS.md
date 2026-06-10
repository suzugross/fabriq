# タスク管理表 — fabriq

<!-- このファイルは TM アプリが .tm/tasks.json から自動生成します。
     直接編集しないでください（次回保存で上書きされます）。
     タスクの追加・更新は tasks.json か TM アプリから行ってください。 -->
最終更新: 2026-06-10 11:30

## 未着手 (2)

### [t-0002] evidence_config取得範囲拡大

**内容:**

・資格情報
・Outlookメールアカウント（Outlook2016~2024）
それぞれ、Fabriq_backuperの実装が参考になります。
情報を採取し、エビデンスとして残せるようにしたい。

<sub>更新: 2026-06-09 00:23 ／ 作成: 2026-06-09 00:19</sub>

### [t-0004] リファクタリング

**内容:**

common.ps1ですが肥大化が激しい状況。おそらく保守性が悪くなっている状況。但し、本プロジェクトの実装は100%AIが行うため、それを踏まえたうえで、common.ps1の分割リファクタリングが必要の評価も必要。

<sub>更新: 2026-06-09 23:17 ／ 作成: 2026-06-09 23:15</sub>

## 完了 (2)

### [t-0001] example設定値の統一

**内容:**

domain_configのCSVに設定されている、ツインピークス風のドメイン名を確認し、ほかのモジュールのCSVで例として挿入されているドメイン名をそのツインピークス風のものに差し替えてほしい。一旦、モジュールのみで大丈夫。

**Claudeメモ:**

調査(2026-06-09 確認): 基準ドメイン=domain_join/domain.csv の town.twinpeaks.blacklodge.local。
他モジュールCSVの例示ドメイン置換候補:
- credential_config/credential_list.csv L2 sql01.corp.local + CORP\administrator / L3 CORP\svc_kit(target=fileserver01) / L4 internal-portal.corp.local + svc_app
- autologon_config/autologon_list.csv L3 example.local
- group_config/group_list.csv L2 city.osaka.example.jp (L3/L4は既にblacklodge)
- manual_kitting_assistant/step_list.csv L3 https://intranet.example.com (URL/判断要)
除外: time_sync_config の NTP(time.windows.com等=実サーバ), '(Example)'ラベル/example.pfx等のファイル名/Example App等(ドメインでない)。
要決定: (1)CORP\ のNetBIOS名(TOWN/TWINPEAKS/BLACKLODGE), (2)FQDNのホストラベル(sql01./internal-portal./intranet.)を残しドメイン部のみ差替か, (3)intranet URLを含めるか。scope=modulesのみ(profiles除外)。版/CHANGELOG不要(例示データのみ・スキーマ不変)。
実施(2026-06-09): 決定=NetBIOS:TWINPEAKS / intranet URL含む。置換完了(各x1): credential L2 sql01.town.twinpeaks.blacklodge.local + TWINPEAKS\administrator, L3 TWINPEAKS\svc_kit, L4 internal-portal.town.twinpeaks.blacklodge.local / autologon L3 / group L2 / manual_kitting_assistant L3 https://intranet.town.twinpeaks.blacklodge.local。FQDNはホストラベル保持しドメイン部のみ差替。NTP(time_sync)除外。BOM/CRLF/JP保全・旧ドメイン残存ゼロ確認。コミット済(157e46b)。
履歴秘匿: ユーザ確認の上で見送り(public repo・偽プレースホルダ・force-pushしてもGitHub側に残存のため費用対効果低)。

<sub>更新: 2026-06-09 10:19 ／ 作成: 2026-06-09 00:16</sub>

### [t-0003] 自己参照モードの実装

**内容:**

ホストリストに特別な変数を用意し、それが入力されているカラムについては、そのPCの実際の値（ホスト名やIPアドレスなど。一旦プリンタやbitlockerは除外）をホストリストの値として代入する。常に正解モード。ホストリストに縛られないキッティングや端末調査を行うときに利用するイメージ。エビデンスなどのファイル名が実際に設定されているPC名を使用できたりするイメージです。変数への値の代入はそのリストを選んでパスフレーズを入力し、キッティング画面に入った時をそのタイミングとしてください。それ以降、PC名変更などを行った際は追従する必要はありません。その運用は非推奨とするイメージです。
HOSTLISTCSVにもFabriqリポジトリの一部として、そのダイナミックエントリをデフォルトで書き込んでおいてください。
変数名の例：
OldPCName＝DynamicOldName
NewPCName=DynamicNewName
IPaddress=DynamicIPaddress
など

**Claudeメモ:**

検討(2026-06-09): 自己参照モード=hostlistの特定列にDynamicトークンを置くと入室時にそのPCのlive値へ解決。
既存機構への適合: (1)解決点=Set-SelectedHostEnvironment(main.ps1:80)のResolve-HostValue。ENC:透過復号と同構造でトークン検出→置換。(2)live値源=Get-CurrentPCInfo(common.ps1:4286)がComputerName/Eth・Wifi{IP,Subnet,Gateway}/DNSをhostlist列と1:1で返す→再利用。(3)『入室時1回解決・以降追従しない』はSave-ResumeState(common.ps1:3361-)が解決後のSELECTED_*値をsnapshot→Restore-HostEnvironmentでbaked値復元、で既存アーキが自動充足。__RESTART__跨ぎ後も再解決されない。
タイミング: main.ps1:1417-1418(list選択+passphrase後/メニュー前)。再選択経路1861-1862はfresh selectionで再解決OK、resume経路1275はRestore(baked)でOK。
対象列: OldPCName/NewPCName/Eth{IP,Subnet,Gateway}/Wifi{IP,Subnet,Gateway}/DNS1-4。除外: Pin(live源なし)/Printer/bitlocker(タスク指定)。
版影響: kernel MINOR(後方互換な新トークン解釈。SELECTED_*契約は不変=モジュール透過)。KERNEL_API.md追記。設計ゲート=フル版。
注意: hostlist.csvはoverlay除外(site-specific)のためrepoのdefault dynamic行は新規配備のtemplateのみ(既存サイトはkernel更新でfeature入手・行は各自追加)。
要決定: (1)トークン形式=単一列文脈sentinel(推奨)vs名前付き(ユーザ例), (2)解決不能時=空+warn(推奨)vsトークン保持, (3)対象列/Pin除外確定, (4)default行のAdminID。
決定(2026-06-09): トークン=__SELF__(単一列文脈sentinel) / 解決不能=空+Show-Warning。設計ゲート(フル版:概要+stateDiagram+敵対検証+変更スコープ宣言)提示済。
実装(2026-06-09): 完了・検証済(未コミット)。main.ps1 Set-SelectedHostEnvironment に __SELF__ 解決(列文脈・Get-CurrentPCInfo 1回キャッシュ・解決不能=空+Show-Warning・ENC:と独立)。hostlist.csv に AdminID=SELF のdefault __SELF__行(BOM保全)。KERNEL_API.md §3.1/§8(3.5.0 Unreleased)・CHANGELOG[Unreleased]追記。新規テスト10件(__SELF__)→全体167 passed/0 failed。check_version OK(KERNEL_VERSION 3.4.1据置=実昇格はリリース指示時)。encoding 新規違反ゼロ・BOM保全。
敵対レビュー(2026-06-10・多観点): critical/high ゼロ。再追従しない契約・回帰・モジュール透過を call-site 追跡で証明済。確定欠陥3件(medium)。ユーザ判断で finding1 のみ修正: apps/fabriq_operator/lib/session_form.ps1 の host-picker で __SELF__ セルを (this PC) 表示+SELF行を auto-select(Row.Tag は生行保持=Start時の解決不変)。finding2(__SELF__ 大文字小文字非依存=他マーカーと同規約)/finding3(toolbar PC Info が SELF行で恒真)は見送り。167 passed/0 failed・encoding違反ゼロ。デスクトップに patch 再生成(session_form込・hostlist除外)。コミット指示待ち。

<sub>更新: 2026-06-10 11:30 ／ 作成: 2026-06-09 00:24</sub>

