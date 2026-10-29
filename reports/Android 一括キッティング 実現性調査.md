# Android 一括キッティング実現性調査：ADB が設定を担い、Device Owner が強制を補う

fabriq をお手本にした Android 一括キッティングツールについて、本調査では次の構成が現実解であることが確認されました。**Windows ホスト PC から USB 経由の ADB で N 台を並列に操作する方式（tier a）を基盤とし、ADB では届かない証明書・EAP Wi-Fi・ユーザー制限だけを自作 Device Owner エージェント（tier b）で補完する構成**です。「MDM なし」の前提では、この構成が最も多くの項目を自動化できます。ADB の shell ユーザー（uid 2000）は `WRITE_SECURE_SETTINGS`・`INSTALL_PACKAGES`・`GRANT_RUNTIME_PERMISSIONS`・`MANAGE_ROLE_HOLDERS`・`NETWORK_SETTINGS`・`SET_TIME_ZONE`・`ENABLE_TEST_HARNESS_MODE` などの権限を保持しています ([Shell AndroidManifest.xml](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/packages/Shell/AndroidManifest.xml))。このため、本書で評価した典型的なキッティング項目 27 件のうち **11 件が ADB 単独で ◎（無人で適用し、読み返しまで可能）** になり、自作 DO を加えると **◎ は 18 件に増えます**。一方、shell は `MANAGE_USERS`・`MASTER_CLEAR`・`MANAGE_CA_CERTIFICATES`・`CONTROL_VPN` を保持しません ([core AndroidManifest.xml](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/core/res/AndroidManifest.xml))。そのため、CA 証明書のサイレント導入、VPN、ユーザー制限による「強制」は ADB 単独ではサポートされません。さらに、**開発者オプションと USB デバッグの有効化、および RSA 鍵の許可という前段の手作業は、ADB を使うすべての方式で端末ごとに残ります** ([Android Debug Bridge](https://developer.android.com/tools/adb))。この前段を不要にできるのは QR プロビジョニング、AMAPI、ゼロタッチです。しかし 2025〜2026 年にかけて次の変更があり、自作ツールがこれらを直接使える道は大きく狭まっています。Play Protect による未承認 DPC のブロック、Play EMM API の新規登録停止、AMAPI の利用者の商用 EMM 等への限定と、新規プロジェクトの既定クォータ 0 台です ([Build a DPC](https://developer.android.com/work/dpc/build-dpc); [AMAPI Permissible Usage](https://developers.google.com/android/management/permissible-usage); [Bayton: The DPC allowlist](https://bayton.org/blog/2025/12/the-dpc-allowlist/))。fabriq の資産のうち、ModuleResult 契約、Step 1〜6＋5.5 のモジュール構造、プロファイル CSV、`__GATE__`、エビデンス、ガバナンスはそのまま移植できます。ただし、`SELECTED_*` 環境変数のようにプロセス単位で 1 つしか持てない状態と、RunOnce による再起動跨ぎは、「1 ホスト対 N 台」向けに作り直す必要があります ([kernel/main.ps1](../kernel/main.ps1))。初版の目標は約 12 モジュール、1 ホストあたり 10〜20 台とすることが推奨されます。あわせて、testharness による初期化、日本メーカー端末での `dpm set-device-owner`、自作エージェントに対する Play Protect の挙動といった高リスク項目は、実機で先に確定させてください。

## 本書の前提：評価記号、確度表示、一次情報の欠落

本書は、先行する Mac 版の実現性調査と比較できるよう、同じ評価記号と章立てを使います。章立ては、方式比較、項目別の表、提案アーキテクチャ、制約、判断事項、実機検証項目、次のステップです。ただし前提となる実行モデルが異なります。Mac 版は、キッティング対象の Mac 上で USB キットの bash を実行する 1 対 1 モデルでした。本書は、ホスト PC が USB と ADB で多数の端末を同時に操作する **1 対多モデル**を前提とします。fabriq 側の事実は、ローカルリポジトリ（調査時点の HEAD は `928a4c9`、`kernel/KERNEL_VERSION` は `3.6.2`）に基づきます。

| 記号 | 意味 |
|---|---|
| ◎ 自動 | 無人で適用し、読み返し検証まで可能 |
| ◯ 半自動 | 自動で適用できるが、前提操作・ヘルパー APK・一部の UI 操作が必要、または読み返しに制約がある |
| △ 補助 | ツールが行うのは、画面を開く・ファイルを置く・手順を案内する・エビデンスを記録するところまで。操作は人が行う |
| ✕ 手動 | 自動化の手段がない、またはサポートされない |
| — | その方式では不要 |
| 要確認 | 本調査では、機能の有無を確認できていない |
| ※ | **要一次情報確認**。公式ページの本文に到達できず、検索スニペット・第三者情報・調査者の学習知識で代替した |

確度は 3 段階で示します。【確度: 高】は、AOSP ソースまたは公式文書を直接確認したものです。【確度: 中】は、公式ページの検索スニペット、または複数の第三者情報の一致によるものです。【確度: 低】は、単一の報告、報道、推論によるものです。

**調査上の制約に注意してください。** 調査環境のネットワーク制限により、android.googlesource.com、cs.android.com、source.android.com、developers.google.com、support.google.com、docs.samsungknox.com、bayton.org、androidenterprise.community、および複数の報道サイトには直接アクセスできませんでした。AOSP は GitHub 上の aosp-mirror と LineageOS の `lineage-23.0`（Android 16 ベース）で代替しており、aosp-mirror の `main` は Android 16 リリース前の開発スナップショットです ([Build.java main](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/core/java/android/os/Build.java); [Build.java lineage-23.0](https://github.com/LineageOS/android_frameworks_base/blob/lineage-23.0/core/java/android/os/Build.java))。したがって、**Android 17 の AOSP ソース（Shell 権限、DPMS の前提条件、SettingsProvider の shell 免除）は未確認**です。AMAPI の Permissible Usage の原文、Play Protect の DPC allowlist を定めた Google ヘルプ 16694822、FRP と開発者検証に関する Google ヘルプ、および Knox の詳細は検索スニペットに依拠しており、Bayton の記事は GitHub 上のソース Markdown で読んでいます。また、AMAPI のフィールド名の一部（`applications[].installType`、`passwordPolicies`、`deviceRadioState`、`additionalData` など）は調査者の学習知識に基づくもので、再検証されていません。これらに依拠する結論には ※ を付けています。

## 方式比較：「MDM なし」で成立するのは ADB 単独と ADB＋自作 Device Owner の二方式

本調査では、方式を 6 つに分けて評価しました。**ADB のみ（tier a）**は、ホスト PC が `adb shell` の `settings`・`cmd`・`pm`・`locksettings` などを直接実行する方式です。shell が保持する権限の範囲であれば、設定値・アプリ・権限・既定アプリ・Wi-Fi（PSK 系）・画面ロック・タイムゾーン・棚卸しについて、「適用＋読み返し」まで CLI で組めます ([Shell AndroidManifest.xml](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/packages/Shell/AndroidManifest.xml))【確度: 高】。ただし、小売端末の「ビルド番号を 7 回タップ → USB デバッグを ON → RSA を『常に許可』」という手順は省略できません。ADB が動く前に ADB で設定することはできない、という鶏と卵の問題があるためです ([Android Debug Bridge](https://developer.android.com/tools/adb))。この手作業は 3 つの方法で減らせます。第一に、ホスト鍵を全キッティング PC で共有すれば、端末側の許可は 1 回で済みます。第二に、初回接続の直後に `settings put global adb_allowed_connection_time 0` を実行すれば、既定 7 日の無通信失効が止まります。値 0 は AOSP 上「無期限」を意味します ([AdbDebuggingManager.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/services/core/java/com/android/server/adb/AdbDebuggingManager.java); [Settings.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/core/java/android/provider/Settings.java))。第三に、再キッティングでは `cmd testharness enable` を使うと ADB 鍵を保持したまま初期化できます ([TestHarnessModeService.java](https://github.com/LineageOS/android_frameworks_base/blob/lineage-23.0/services/core/java/com/android/server/testharness/TestHarnessModeService.java))。

**ADB＋自作 Device Owner（tier b）**は、自作アプリを `adb shell dpm set-device-owner` で DO 化する方式です。AOSP では、shell または root から呼ばれた場合、セットアップウィザード（SUW）の完了後でも DO を設定できます。条件は、既存の DO／PO がないこと、既定以外のユーザー（ワークプロファイル、プライベートスペース、アプリのクローンを含む）がいないこと、アカウントがないこと、対象ユーザーが実行中であることの 4 つです ([DevicePolicyManagerService.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/services/devicepolicy/java/com/android/server/devicepolicy/DevicePolicyManagerService.java))【確度: 高（Android 16 系ソース）】。本番署名の（testOnly でない）DPC は、アカウントが 1 つでもあると失敗します。そのため手順は「初期化 → アカウントを追加せずに SUW を通過 → USB デバッグを有効化 → `adb install` → `dpm set-device-owner`」になります。DO を解除できるのは、testOnly ビルドを `dpm remove-active-admin` で外す場合に限られ、それ以外は原則として初期化が必要です ([DevicePolicyManagerServiceShellCommand.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/services/devicepolicy/java/com/android/server/devicepolicy/DevicePolicyManagerServiceShellCommand.java))。また、DO を設定するとバックアップサービスが恒久的に停止され、managed profile の追加が制限されます。日本メーカー（AQUOS、arrows、京セラ、Xperia）やキャリア版端末に固有の阻害要因については、信頼できる資料が見つかっていません。参考として、Xiaomi と OPPO での失敗事例と、Samsung の「user limit reached」（文脈が曖昧）が OSS の README に報告されています ([OwnDroid](https://github.com/BinTianqi/OwnDroid))。

**自作 DPC の QR プロビジョニング（tier c）**は、Android 7 以降でサポートされます。Android 9 以降では、SUW の Welcome 画面を 6 回タップすると内蔵の QR リーダーが起動し、Android 13 以降は Wi-Fi 情報を payload に含めないとネットワークの選択を求められます ([Bayton: provisioning methods](https://bayton.org/android/android-enterprise-provisioning-methods/))。Android 12 以降の DPC には `ACTION_GET_PROVISIONING_MODE` と `ACTION_ADMIN_POLICY_COMPLIANCE` の 2 つの Activity が必須で、これがないとプロビジョニングは失敗します ([Android 12 enterprise](https://developer.android.com/work/versions/android-12))。実例として、自作 DPC「muster」は Pixel 6a でこの 2 つを欠いたため APK のダウンロード後に初期化を求められ、修正後の 2026-10-19 に DO へ到達しています ([quadseven/muster #89](https://github.com/quadseven/muster/issues/89))。GET_PROVISIONING_MODE には admin extras bundle に加えて端末の IMEI とシリアル番号が渡されるため ([DevicePolicyManager.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/core/java/android/app/admin/DevicePolicyManager.java))、**拠点単位の QR を 1 枚だけ用意し、端末側でシリアルを使って hostlist の行を引く**という fabriq 流の設計が可能です【推論】。最大のリスクは Play Protect です。Google は「承認済み DPC 以外のインストールは enrollment 中にブロックされる」としていますが ([Google Help 16694822](https://support.google.com/work/android/answer/16694822) ※)、実際の挙動についてはブロックされた報告とされなかった報告が**混在**しています（ポリシー変更の章を参照してください）。

**AMAPI 直結または商用 EMM（tier d1）**は、技術的には理想的なバックエンドです。enrollment token から QR を生成でき、ポリシーを JSON で宣言でき、compliance の読み返しをそのまま検証に使えます ([AMAPI enrollmentTokens](https://developers.google.com/android/management/reference/rest/v1/enterprises.enrollmentTokens) ※)。ただし、Permissible Usage により利用者は商用 EMM 開発者、Device Trust 提供者、OEM に限定され、2025-11 以降の新規プロジェクトは既定のクォータが 0 台です ([AMAPI Permissible Usage](https://developers.google.com/android/management/permissible-usage) ※; [Bayton FAQ: AMAPI usage limits](https://bayton.org/android/android-enterprise-faq/amapi-usage-limits/))。Android Device Policy が DO となって Google のクラウドへ継続的にチェックインするため、実質的には MDM です【推論】。現実的な使い方は、顧客の商用 EMM の API を本ツールから操作し、本ツールを「EMM の前段のオーケストレータ＋エビデンス生成器」と位置付ける形で、この場合 AMAPI の承認は EMM ベンダーが保持します【推論】。

**ゼロタッチ（ZTE）と Knox Mobile Enrollment（tier d2）**は、どちらも MDM への登録の入口にすぎません。ZTE は認定リセラーから購入した端末が前提で、端末を登録できるのはリセラーだけです ([Zero-touch for IT admins](https://support.google.com/work/android/answer/7514005?hl=en) ※; [Bayton: zero-touch](https://bayton.org/android/what-is-android-zero-touch-enrolment/))。国内では SoftBank、KDDI、NTT ドコモビジネス、T-Gaia が ZTE を扱っています ([SoftBank BCDM](https://www.softbank.jp/biz/services/security/bcdm/function/android/); [KDDI Android Enterprise サポート](https://biz.kddi.com/service/gyoumushien/android-enterprise-support/); [T-Gaia](https://www.t-gaia.co.jp/biz/service/catalog/Android))。KME は無償で、ライセンスキーも不要です ([Knox licenses](https://docs.samsungknox.com/admin/fundamentals/knox-licenses/) ※)。

**OEM ツール（tier d3）**のうち、**Knox Configure（KC）**は有償の端末単位ライセンスで提供されます。Google・Samsung・キャリアの初期設定画面のスキップ、アプリの自動導入、プリインストールアプリの削除、ブランディングに対応し ([Knox Configure](https://www.samsungknox.com/en/solutions/it-solutions/knox-configure) ※)、1 回限りの適用となる Setup エディションが「MDM ではない」と主張しやすい唯一の Samsung 製品です【推論】。日本での提供については情報が食い違っています。公式の対応地域ページは日本を掲載していますが、SEAP のページは日本を含めていません（古い記述である可能性があります） ([KC supported locations](https://www.samsungknox.com/en/solutions/it-solutions/knox-configure/supported-locations) ※; [SEAP: Knox Configure](https://seap.samsung.com/sdk/knox-configure))。**Zebra StageNow** は無償の Windows ツールで、MX のバージョンに合わせたプロファイルをバーコードや NFC に変換します。順序付きの設定ステップを端末の能力版（MX）に照らして束ねる点で、**fabriq のプロファイル概念に最も近い既存製品**です ([StageNow 5.18 install](https://techdocs.zebra.com/stagenow/5-18/installing); [StageNow Profiles](https://techdocs.zebra.com/stagenow/latest/stagingprofiles/))。Honeywell Enterprise Provisioner も、Windows 上で 2D バーコードを生成する同種のツールです ([Honeywell EP guide](https://support.honeywellaidc.com/servlet/fileField?entityId=ka02K000000c8deQAA&field=File_1__Body__s))。**京セラ Device Control** は利用料 0 円で、PC もネットワークも使わず、管理用端末から QR または NFC でポリシーを複製します。対象は京セラ端末に限られます ([京セラ Device Control アプリ](https://www.kyocera.co.jp/prdct/telecom/office/phone/app-service/service/mdm/dcapp/))。

| 方式 | 主な前提条件 | MDM 扱い | 費用（調査時点） | 前段の手作業 | 主なリスク | fabriq の流用度 |
|---|---|---|---|---|---|---|
| (a) ADB のみ | Windows ホスト、OEM の USB ドライバ、版を固定した platform-tools。端末ごとに開発者オプション・USB デバッグ・RSA 許可が必要 | 該当しない | ツールは無償。platform-tools の再配布条件に注意が必要 ([SDK License](https://developer.android.com/tools/releases/platform-tools)) | 必要（SUW 完走＋USB デバッグ＋RSA） | 証明書・VPN・EAP・強制は不可。OEM 差。USB デバッグを残置した場合の攻撃面 | 高 |
| (b) ADB＋自作 DO | (a) に加え、アカウント 0・追加ユーザー 0・既存 DO／PO なし。署名鍵の永続管理 | 解釈次第（クラウド常駐管理はないが、組織管理端末になる）。**顧客との合意が必要** | ライセンス費なし。開発・保守工数がかかる | (a) と同じ | 日本メーカーでの DO 化の可否が未確認。解除は原則初期化。年次の targetSdk 追従 | 高（ホストは共通。エージェントを追加） |
| (c) 自作 DPC の QR | Android 7 以降（内蔵リーダーは 9 以降）。12 以降は Activity 2 つが必須。SUW 中に到達できる APK 配布先 | (b) より MDM 寄り | (b) に加えてホスティング費用 | 6 タップ＋QR 読み取り（USB デバッグは不要） | **Play Protect の allowlist によるブロック（証拠は混在）**。managed Google Play は使えない | 中（ホスト側は QR 生成と検証に縮小） |
| (d1) AMAPI／商用 EMM | GCP プロジェクト、enterprise binding、Google の承認（新規は 0 台から）。商用 EMM 限定 ※ | MDM そのもの | API はコミュニティ情報では無償。国内 MDM は約 ¥300〜550／台／月 ([next-sfa](https://next-sfa.jp/journal/mdm/mdm-comparison/))。Intune Plan 1 は約 ¥1,199／ユーザー／月 ([locked.jp](https://locked.jp/article/mdm/intune-mdm-pricing)) | QR／afw#／ZTE（USB デバッグは不要） | 社内用途は規約に抵触するおそれ。承認待ちは数週間。未承認プロジェクトでは FRP・初期化ポリシーを使えない | 中（ツールは EMM の前段に回る） |
| (d2) ZTE／KME | 認定リセラーからの購入（ZTE）、Knox 対応 Galaxy（KME）、EMM の DPC | MDM への登録経路 | KME は無償。ZTE の顧客側費用は未確認 | 電源投入のみ | MDM が必須。自作 DPC は対象外とみられる | 低〜中 |
| (d3) OEM ツール（KC／StageNow／Honeywell EP／京セラ DC） | 単一ベンダーの端末と、ベンダーのポータル・Windows ツール・管理用端末 | KC Setup・StageNow・EP・DC はクラウド常駐管理なし。KC Dynamic は軽い MDM に近い | KC Dynamic は 1 年 $10〜12.50／台（米国） ([CDW](https://www.cdw.com/product/knox-configure-dynamic-edition-license-1-year-1-license/4721934))。StageNow は無償。京セラ DC は ¥0 | バーコード／NFC／QR | ベンダーロックイン。KC にはプロファイル API が見当たらない。DC のポリシーは端末上で作成 | 低（StageNow は設計の参考になる） |

この比較から、Mac 版と同じ「MDM なし」の条件で現実に残るのは (a) と (b) の 2 方式です。(c) は Play Protect の判定次第で成否が変わり、(d1)〜(d2) は MDM そのものか、MDM への登録経路です。(d3) は単一ベンダーの端末群に限り、例外的に選択肢になります。

## 項目別ケイパビリティ：ADB 単独で設定値の大半が ◎、証明書と「強制」は Device Owner で ◎

各方式の境界は権限で決まります。ADB の shell は system・secure・global のほぼすべての設定キーを読み書きでき、Android 12 で導入された @hide キーの読み取り制限も uid ≥ 10000 のアプリが対象で、shell には適用されません ([SettingsProvider.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/packages/SettingsProvider/src/com/android/providers/settings/SettingsProvider.java))【確度: 高】。これに対し DO の `setGlobalSetting`・`setSecureSetting`・`setSystemSetting` は allowlist 方式で、ADB の `settings put` より範囲が狭くなります。Global は `ADB_ENABLED`・`AUTO_TIME`・`AUTO_TIME_ZONE`・`STAY_ON_WHILE_PLUGGED_IN`・`PRIVATE_DNS_MODE` など、System は輝度・輝度モード・消灯タイムアウトに限られ、`BLUETOOTH_ON`・`DEVELOPMENT_SETTINGS_ENABLED`・`WIFI_ON` への書き込みは**黙って無視されます** ([DevicePolicyManagerService.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/services/devicepolicy/java/com/android/server/devicepolicy/DevicePolicyManagerService.java))。逆に、`settings put` は値を書くだけで、ユーザーは後から自由に変更できます。したがって tier (b) では、**ADB が設定値を書き、DO がポリシー・証明書・制限による強制を担う**というハイブリッドが最も合理的です。

表の (d) 列は主に AMAPI を、補足として OEM ツールを示します。(d) 列の AMAPI フィールドには一次情報で確認できていないものが多いため、※ と「要確認」が多くなっています。

| # | 項目 | (a) ADB | (b) ADB＋DO | (c) QR | (d) AMAPI／OEM | 適用・検証の手段と制約 |
|---|---|---|---|---|---|---|
| 1 | セットアップウィザード（SUW） | ✕ 初回／◯ 再キッティング | ✕ 初回／◯ 再キッティング | ◯ | ◎ ZTE・KME・KC／◯ QR | 新品は USB デバッグが OFF のため、SUW 中に ADB は使えない。再キッティング時は `cmd testharness enable`（Android 10 以降）で全消去・ADB 鍵の保持・`device_provisioned`/`user_setup_complete` の自動設定が行われる ([TestHarnessModeService.java](https://github.com/LineageOS/android_frameworks_base/blob/lineage-23.0/services/core/java/com/android/server/testharness/TestHarnessModeService.java))。(c) は 6 タップ＋QR。KC は Google・Samsung・キャリアの初期画面をスキップできる |
| 2 | 開発者オプション・USB デバッグ・RSA 許可 | ✕ | ✕ | — | — | 端末ごとの手作業。軽減策は `adbkey` の共有、`settings put global adb_allowed_connection_time 0`、testharness による鍵の保持。Samsung は Auto Blocker の OFF が前提 ([Samsung: Auto Blocker](https://www.samsung.com/uk/support/mobile-devices/protect-your-galaxy-device-with-the-new-auto-blocker-feature/)) |
| 3 | Wi-Fi（open／OWE／WPA2／WPA3／WEP） | ◎ | ◎ | ◎ | ◎ ※ | 適用は `cmd wifi add-network "<SSID>" wpa2 "<PSK>" [-h] [-m] [-d] [-r none]`、検証は `cmd wifi list-networks`・`cmd wifi status` ([WifiShellCommand.java](https://github.com/LineageOS/android_packages_modules_Wifi/blob/lineage-23.0/service/java/com/android/server/wifi/WifiShellCommand.java))。PSK は読み返せないため、更新は forget→add とする。DO は `addNetwork`（DO は deprecation の対象外）と `getCallerConfiguredNetworks()`（API 31）を使う ([WifiManager](https://developer.android.com/reference/android/net/wifi/WifiManager))。(c) は `EXTRA_PROVISIONING_WIFI_*`、AMAPI は `openNetworkConfiguration` |
| 4 | Wi-Fi（EAP／802.1X） | ✕ | ◎ | ◎ | ◎ ※ | `WifiShellCommand` には `eap` と `WifiEnterpriseConfig` の参照がない。DO は `installCaCert`＋`installKeyPair`＋`grantKeyPairToWifiAuth`（API 31）を使う ([API diff 31](https://developer.android.com/sdk/api_diff/31/changes/android.app.admin.DevicePolicyManager))。(c) の EAP 系 extras は API 29 で追加。AMAPI は ONC で EAP-TLS／PEAP などに対応 ([AMAPI configure networks](https://developers.google.com/android/management/configure-networks) ※)。ただし証明書が反映されないという報告がある |
| 5 | 端末名 | ◯ | ◯ | ✕ | 要確認 | 適用は `settings put global device_name "PC-001"`、検証は同じキーの `get`。Bluetooth 名などへの反映は OEM 差がある【推論】。DO の global 設定 allowlist に `device_name` は含まれない |
| 6 | 画面消灯・輝度・自動回転・充電中のスリープ | ◎ | ◎ | ◯ | 要確認 | `settings put system screen_off_timeout 600000`、`screen_brightness`、`screen_brightness_mode 0`、`accelerometer_rotation 0`、`settings put global stay_on_while_plugged_in 7`。DO の `setSystemSetting`（API 28）が扱えるのは輝度・輝度モード・タイムアウトのみ。輝度の値域には OEM 差がある |
| 7 | ロケール（システム言語） | ◯ | ◯ | ◎（プロビジョニング時のみ） | 要確認 | 専用の shell コマンドはない。`system_locales` について、AOSP は「直接使うな」と注記している ([Settings.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/core/java/android/provider/Settings.java))。手段は、`CHANGE_CONFIGURATION` を `pm grant` で付与したヘルパー APK か、shell uid で動く `app_process` ヘルパー（未検証）。検証は `getprop persist.sys.locale`。アプリ別のロケールは `cmd locale set-app-locales`（13 以降）。DO から後でロケールを変える公開 API はなく、使えるのは `EXTRA_PROVISIONING_LOCALE` だけ |
| 8 | タイムゾーン・時刻・自動同期 | ◎ | ◎ | ◎ | 要確認 | `cmd alarm set-timezone Asia/Tokyo`、`cmd time_zone_detector`／`cmd time_detector`（12 以降。サブコマンド名は `help` で確定させる） ([TimeZoneDetectorShellCommand.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/services/core/java/com/android/server/timezonedetector/TimeZoneDetectorShellCommand.java))。検証は `getprop persist.sys.timezone`。DO の `setTimeZone` は自動 TZ が有効だと false を返すため、`setAutoTimeZoneEnabled`（API 30）を先に呼ぶ。`ntp_server` を書き換えたときの反映は未検証 |
| 9 | 無線系（機内モード・BT・NFC・位置情報・モバイルデータ） | ◎ | ◎ | ◯ | 要確認 | `cmd connectivity airplane-mode enable`、`cmd bluetooth_manager enable`、`svc nfc enable`、`cmd location set-location-enabled true`、`cmd phone data enable` ([Svc.java](https://github.com/LineageOS/android_frameworks_base/blob/lineage-23.0/cmds/svc/src/com/android/commands/svc/Svc.java); [LocationShellCommand.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/services/core/java/com/android/server/location/LocationShellCommand.java))。`svc usb` は ADB の切断要因になるため使用禁止 |
| 10 | 画面ロック（PIN） | ◎ | ◎ | ◯ | ◯ | 適用は `locksettings set-pin <PIN>`。検証は `locksettings verify --old <PIN>`・`get-disabled`。既に設定済みなら `--old` が必須で、管理ポリシーに適合しない資格情報は拒否される ([LockSettingsShellCommand.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/services/core/java/com/android/server/locksettings/LockSettingsShellCommand.java))。testharness はロック画面があると拒否するため、PIN は最後に設定する。DO は `setRequiredPasswordComplexity`（API 31）。AMAPI は `RESET_PASSWORD`（Android 14 では数字 6 桁以上） ([issueCommand](https://developers.google.com/android/management/reference/rest/v1/enterprises.devices/issueCommand) ※) |
| 11 | APK 導入（単体・split APK） | ◎ | ◎ | ◯ | ◎ ※ | `adb install -r [-g]`、`adb install-multiple base.apk split_*.apk`、`pm install-create/-write/-commit`。検証は `pm list packages --show-versioncode` ([PackageManagerShellCommand.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/services/core/java/com/android/server/pm/PackageManagerShellCommand.java))。Android 15 は targetSdk 24 未満をブロックし、回避は `--bypass-low-target-sdk-block` ([Android 15 behavior changes](https://developer.android.com/about/versions/15/behavior-changes-all))。DO はユーザー操作なしでインストールを完了できる ([PackageInstallerSession.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/services/core/java/com/android/server/pm/PackageInstallerSession.java))。AMAPI は SDK 1.6.0 から APK の直接配布に対応 ([Bayton](https://bayton.org/blog/2025/08/amapi-apk-deployment/)) |
| 12 | 不要アプリの無効化・削除（debloat） | ◎ | ◎ | ◎ | ◯ | 既定は `pm disable-user --user 0`、削除は `pm uninstall -k --user 0`。戻すときは `pm enable`／`cmd package install-existing`。`pm hide` は `MANAGE_USERS` が必要なため shell では使えない ([PackageManagerService.java](https://github.com/LineageOS/android_frameworks_base/blob/lineage-23.0/services/core/java/com/android/server/pm/PackageManagerService.java))。DO は `setApplicationHidden` と `setUninstallBlocked`。QR で DO 化すると、`LEAVE_ALL_SYSTEM_APPS_ENABLED` を指定しない限りシステムアプリが無効化される |
| 13 | ランタイム権限・AppOps・既定アプリ・電池最適化 | ◎ | ◎ | ◯ | 要確認 | `pm grant`／`pm revoke`、`appops set <pkg> SYSTEM_ALERT_WINDOW allow`、`cmd role add-role-holder android.app.role.BROWSER <pkg>`（検証は `get-role-holders`） ([RoleShellCommand.java](https://github.com/LineageOS/android_packages_modules_Permission/blob/lineage-23.0/service/java/com/android/role/RoleShellCommand.java))、`dumpsys deviceidle whitelist +<pkg>`。Samsung などの独自の省電力機能は別管理で、未調査。DO は `setPermissionGrantState`（値の固定）、`setDefaultDialerApplication`（API 34） |
| 14 | ユーザー補助・通知リスナー | ◯ | ◯ | ✕ | 要確認 | `settings put secure enabled_accessibility_services <pkg>/<svc>`、`cmd notification allow_listener`。Android 13 以降の「制限付き設定」がこの経路を阻むかは一次資料がない（要実機検証） ([Kaspersky](https://kaspersky.com/blog/android-restricted-settings/49991))。DO の secure 設定 allowlist にこのキーは含まれない |
| 15 | CA 証明書 | △ | ◎ | ◎ | ◯ ※ | Android 11 以降、CA は「設定」からしかインストールできない ([KeyChain.java](https://github.com/LineageOS/android_frameworks_base/blob/lineage-23.0/keystore/java/android/security/KeyChain.java))。(a) では `adb push` で Download に置き、`am start` で設定画面を開き、人がファイルを選ぶ。検証はスクリーンショット。DO は `installCaCert` |
| 16 | クライアント証明書（PKCS#12） | △ | ◎ | ◎ | ◯ ※ | (a) では CertInstaller の UI を使う。DO は `installKeyPair`、`grantKeyPairToApp`（API 30）、`grantKeyPairToWifiAuth`（API 31） ([API diff 30](https://developer.android.com/sdk/api_diff/30/changes/android.app.admin.DevicePolicyManager)) |
| 17 | APN | ◯（14 以降）／✕（13 以前） | ◎ | ◎ | 要確認 | 適用は `content insert --uri content://telephony/carriers --bind name:s:CorpAPN --bind apn:s:corp.example ...`、検証は `content query`。shell に `WRITE_APN_SETTINGS` が付与されたのは A14 r1 から ([Shell manifest A13](https://github.com/aosp-mirror/platform_frameworks_base/blob/android-13.0.0_r1/packages/Shell/AndroidManifest.xml); [A14](https://github.com/aosp-mirror/platform_frameworks_base/blob/android-14.0.0_r1/packages/Shell/AndroidManifest.xml); [TelephonyProvider.java](https://github.com/aosp-mirror/platform_packages_providers_TelephonyProvider/blob/main/src/com/android/providers/telephony/TelephonyProvider.java))。DO は `addOverrideApn`（API 28） |
| 18 | VPN | ✕ | ◯ | ◯ | 要確認 | shell は `CONTROL_VPN` を持たない。DO は `setAlwaysOnVpnPackage`（API 24。lockdown 対象外リストは API 29） ([API diff 29](https://developer.android.com/sdk/api_diff/29/changes/android.app.admin.DevicePolicyManager))。VPN プロファイルの中身は、VPN アプリ側の取り込み機能に依存する |
| 19 | 壁紙・ホーム画面 | ◯ 壁紙／✕ ホーム | ◯／✕ | ◯／✕ | 要確認（KC はブランディングに対応） | `cmd wallpaper` で操作できるのは dim 系だけ ([WallpaperManagerShellCommand.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/services/core/java/com/android/server/wallpaper/WallpaperManagerShellCommand.java))。`SET_WALLPAPER` は normal 権限なので、ヘルパー APK で設定できる。検証は `dumpsys wallpaper`＋スクリーンショット。ホーム画面の配置はランチャー固有 |
| 20 | Google アカウント | ✕ | ✕ | ✕ | ◯（managed Google アカウント） | UI 操作が必要。FRP 対策として、既定は「追加しない」。(b) はアカウントがあると DO 化に失敗する。(c) は Play EMM API が新規登録を受け付けていないため、managed Google Play を使えない |
| 21 | OS 更新 | △ | ◯ | ◯ | ◯ | ADB から OTA を実行させる手段は見当たらない。(a) は Verify-only とし、`getprop ro.build.version.security_patch`・`ro.build.fingerprint` を期待値と照合する。DO は `setSystemUpdatePolicy` と `installSystemUpdate`（API 29）。AMAPI は `SystemUpdate`・`FreezePeriod` ([AMAPI policies](https://developers.google.com/android/management/reference/rest/v1/enterprises.policies) ※) |
| 22 | 棚卸し（シリアル・機種・OS・IMEI・MAC） | ◎（IMEI は ◯） | ◎ | ◎ | ◎ ※ | `adb get-serialno`、`getprop ro.serialno`／`ro.product.model`／`ro.build.version.release`、`pm list packages --show-versioncode`、`dumpsys battery`。IMEI は shell が `READ_PRIVILEGED_PHONE_STATE` を持つが、`service call iphonesubinfo` の番号は版や OEM によって異なり不安定。Wi-Fi の MAC は接続先ごとにランダム化されるため、factory MAC と区別して記録する。DO は `Build.getSerial()` が使える ([Build.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/core/java/android/os/Build.java))。(c) では IMEI とシリアルが DPC に渡される |
| 23 | スクリーンショット・エビデンス | ◎ | ◎ | △ | ◯ | `adb exec-out screencap -p > x.png`、`screenrecord --time-limit 180`（最大 180 秒、音声なし）、`uiautomator dump` ([Android Debug Bridge](https://developer.android.com/tools/adb))。ADB がない場合はエージェントが MediaProjection を使う必要があり、ユーザーの同意が求められる【推論】 |
| 24 | 再起動・再開 | ◎ | ◎ | ◯ | ◯ | `adb reboot` → `adb -s <serial> wait-for-device` → `getprop sys.boot_completed` が 1 になるまでポーリング。その後 `pm path android` の応答でヘルスチェックする。AMAPI の `REBOOT` は fully managed 端末のみとの報告がある ([fleetdm #23232](https://github.com/fleetdm/fleet/issues/23232)) |
| 25 | 初期化・再キッティング | ◯ | ◎ | ◎ | ◯ | shell は `MASTER_CLEAR` を持たないため、`FACTORY_RESET` のブロードキャストは使えない。`cmd testharness enable` は shell から全消去を実行できるが、ロック画面があると拒否され、端末はテストハーネスモードになる。DO は `wipeData`／`wipeDevice`（API 34）。AMAPI の未承認プロジェクトでは、2025-05 以降 FRP と初期化のポリシーを使えない ([Bayton notes](https://bayton.org/notes/99/)) |
| 26 | ユーザー制限・強制（ユーザーが後から変更できない状態） | ✕ | ◎ | ◎ | ◎ ※ | `settings put` は値を書くだけで、ユーザーは後から変更できる。`pm set-user-restriction` は `MANAGE_USERS` が必要 ([UserManagerService.java](https://github.com/LineageOS/android_frameworks_base/blob/lineage-23.0/services/core/java/com/android/server/pm/UserManagerService.java))。DO は `addUserRestriction` で `DISALLOW_FACTORY_RESET`／`DISALLOW_USB_FILE_TRANSFER`／`DISALLOW_DEBUGGING_FEATURES`／`DISALLOW_CONFIG_WIFI` などを設定でき、`setFactoryResetProtectionPolicy`（API 30）も使える |
| 27 | 最終ハードニング（USB デバッグの無効化） | ◯ | ◎ | — | — | 全検証とエビデンスの回収を終えてから実行する。順序は `settings delete global adb_allowed_connection_time` → `settings put global development_settings_enabled 0` → 最後に `settings put global adb_enabled 0`。最後のコマンドで切断される見込みのため読み返しはできず、`-Verified $null` 相当とする。認可済み鍵の取り消しは端末の UI で行う。DO は `DISALLOW_DEBUGGING_FEATURES` で無効化を強制できる |

### ADB 単独で ◎ は 11 件、自作 DO を加えると 18 件

表を集計すると、(a) は ◎ 11・◯ 7・△ 3・✕ 6、(b) は ◎ 18・◯ 6・✕ 3 になります。(b) で ✕ として残るのは、前段（SUW の初回と USB デバッグ）と Google アカウントだけです（ホーム画面の配置は全方式で対象外です）。(c) では前段の手作業がなくなる代わりに、ADB で使えた settings の任意キーを使えなくなるため、端末名やユーザー補助が ✕ に変わります。なお、IMEI の扱いについては調査ノート間で見解に差がありました。fabriq 側の分析は「取得制限あり」を前提とし、ADB 側の分析は shell が `READ_PRIVILEGED_PHONE_STATE` を保持するため権限上は読めるとしています ([Shell manifest A13](https://github.com/aosp-mirror/platform_frameworks_base/blob/android-13.0.0_r1/packages/Shell/AndroidManifest.xml))。本書は AOSP ソースに基づく後者を採りますが、実用上の取得手段（`service call iphonesubinfo` の番号）が版や OEM によって変わり不安定であるため、表では ◯ としています。

### どの方式でも人の手に残る作業

どの方式を選んでも、次の作業は人の手に残ります ([石岡市 調達資料](https://www.city.ishioka.lg.jp/data/doc/1729155430_doc_26_0.pdf); [AMAPI issueCommand](https://developers.google.com/android/management/reference/rest/v1/enterprises.devices/issueCommand) ※)。具体的には、SIM の挿入と回線の開通（Android 15 以降の組織所有端末で AMAPI の `ADD_ESIM` を使える場合を除く）、保護フィルムとラベルの貼付、IMEI・シリアルとラベルの突合、最終状態の目視確認、QR フローでの初回ネットワーク接続、managed アカウントを使わない場合の Google アカウントのサインインです。

### PIN と最終ハードニングは必ず最後に置く

実行順序には 3 つの制約があります。まず、testharness はロック画面があると拒否されます。また、PIN を設定した端末は再起動後に初回ロック解除前（BFU）の状態になり、検証が失敗しうるため【推論】、PIN は最後に設定します。次に、`adb_enabled 0` を実行するとセッションが切れるため、最終ハードニングは最終ステップに置きます。最後に、`adb_allowed_connection_time 0` は Step 0 で設定し、ハードニングの直前に既定値へ戻します。

## fabriq からの移植：契約層はそのまま移り、実行層は「1 ホスト対 N 台」へ作り直す

fabriq は、対象 PC 上で 1 台につき 1 つのセッションを動かすプログラムですが、中核の契約は OS に依存しません。`New-ModuleResult` は Success／Error／Cancelled／Skipped／Partial の 5 値の Status と範囲付きの三値 `Verified` を返し、`New-BatchResult` は件数から Status を自動判定します。`Invoke-SafeCommand` は ModuleResult が返らなければ Error として記録する fail-closed の回収を行い、`Import-ModuleCsv` は読み込み失敗を `$null`、対象ゼロを空配列として区別します ([kernel/common.ps1](../kernel/common.ps1))。モジュールのテンプレートは、Step 1（CSV 読込）→ Step 2（前提確認）→ Step 3（dry-run 表示）→ Step 4（確認）→ Step 5（適用ループ）→ Step 5.5（読み返し検証）→ Step 6（集計して返却）の固定フローです ([dev/template/_template_script.ps1](../dev/template/_template_script.ps1))。`windows_feature_config` は、同じ判定関数を Step 5 の冪等性 skip と Step 5.5 の検証の両方で使っています ([windows_feature_config.ps1](../modules/standard/windows_feature_config/windows_feature_config.ps1))。このパターンは、Android の `settings get`・`pm list packages`・`dumpsys` にそのまま写せます。

一方、セッション状態の多くは、プロセス全体で 1 つしか持てないグローバル変数や環境変数として実装されています。`$env:SELECTED_*`、`$env:FABRIQ_SEGMENT`、`$global:AutoPilotMode`、`$global:FabriqEvidenceBasePath`、`$global:_LastModuleResult` がその例で、実行のたびに設定と解除を繰り返します ([kernel/main.ps1](../kernel/main.ps1); [kernel/KERNEL_API.md](../kernel/KERNEL_API.md))。このため、1 つのプロセスで N 台を並列に動かすことはサポートされません。

### 分類表：A はそのまま移植、B は作り直し、C は Windows 専用

| fabriq の要素 | 分類 | Android 版での扱い |
|---|---|---|
| `New-ModuleResult`／`New-BatchResult`、Verified の三値、無返却を Error とする規則 | A | そのまま移植する。端末×ステップのセルの色（OK／NG／未検証）に直結する |
| Step 1〜6＋5.5、`module.csv`、`*_list.csv`、Enabled／Segment／RequiredColumns | A | そのまま移植する。Segment が空欄のときの意味は、Guide.txt（全セグメント共通）と実装（厳密一致）で食い違っている ([dev/template/Guide.txt](../dev/template/Guide.txt))。新ツールでは最初に 1 つに決める |
| プロファイル CSV、`__GATE__`、ErrorMode（skip／retry、上限 5 回）、プロファイル別データオーバーレイ | A | `__GATE__` は端末ごとに評価する（NG の端末だけを止める）。オーバーレイは複数顧客のロットを扱ううえで必須に近い |
| `ENC:` 暗号（AES-256-CBC＋PBKDF2）、秘密値のマスク | A（要改修） | 固定ソルトは既知の弱点（TM t-0075）であるため、blob ごとの乱数ソルトに改める ([.tm/tasks.json](../.tm/tasks.json)) |
| 実行履歴 CSV、HTML チェックリスト（期待値と実測、OK／NG／Incomplete／CA）、manifest の前方互換ルール、log_uploader、JSONL テレメトリ | A | 「1 台 1 枚」と「ロット一覧」の 2 段構成にする ([kernel/EVIDENCE_MANIFEST.md](../kernel/EVIDENCE_MANIFEST.md)) |
| KERNEL_API.md、カーネルとモジュールの独立した SemVer、REQUIRES_KERNEL、CHANGELOG、設計ゲート、CI（exit code で判定）、TM 台帳 | A | 初版 0.1.0 から持ち込む |
| `SELECTED_*`・`$global:*`（プロセスに 1 つ） | B | 端末ごとにワーカプロセスを分け、`HOST_<column>` 相当の環境変数の契約を保つ |
| hostlist（`NewPCName` で照合、44 列） | B | 照合キーを `Serial` に変え、`AssetTag`・`DeviceName`・`WifiProfile`・`Segment`・`GoogleAccountPolicy` などの語彙に置き換える |
| セッション画面（作業者・パスフレーズ・ホスト選択） | B | 1 セッション＝1 ロットとする。ロット受付画面で `adb devices` → シリアル照合 → 一致しない端末を手動で割り当て |
| `Confirm-ModuleExecution`、ErrorMode 空欄時の MessageBox（タイムアウトなし） | B | 確認はロットで 1 回だけ行う。モーダル表示は、端末ごとの「要対応」セルに置き換える |
| FlexProfile（1 台×M 行） | B | 端末 N 台×ステップ M 列のマトリクス（電子チェックシート）にし、状態はシリアル単位で保存する ([flex_dashboard.ps1](../apps/fabriq_operator/lib/flex_dashboard.ps1)) |
| `__RESTART__`（Save-ResumeState・RunOnce・DPAPI・AutoLogon） | B | `adb reboot`＋起動完了を待つ同期ステップにする。resume_state は、ホスト自身のクラッシュへの備えとしてだけ残す |
| `Invoke-SafeCommandAsync`（`DefaultTimeoutSec=0`） | B | すべての adb 呼び出しに上限時間を設け、端末単位でキャンセルできるようにする ([kernel/json/async_config.json](../kernel/json/async_config.json)) |
| `Capture-ScreenEvidence`（DPI 依存） | B | `adb exec-out screencap -p` で撮影し、端末別のフォルダに保存する（DPI 由来の見切れ問題 t-0085／t-0101 は起きない） |
| ArtPulse、Execution Toolbar | B | ワーカのハートビートと `adb devices` の状態で置き換え、端末別の操作パネルを用意する |
| `Get-HardwareUniqueId`／`Test-AdminPrivilege`／`Test-FabriqProtectedPath` | B | それぞれ `ro.serialno`、`unauthorized` 状態の検出、`/sdcard`・`/data/local/tmp` 用のパスガードに置き換える |
| RunOnce・AutoLogon・WU ループ・Active Setup・HKCU・Zone.Identifier、および約 43 のモジュール（AD・GPO・BitLocker・レジストリ・Office・プリンタなど） | C | 対象外 |

### 端末ごとのワーカプロセスと、マトリクス型のチェックシート

並列化の方式としては、**端末ごとにワーカプロセスを分ける**ことが推奨されます。fabriq の環境変数による契約を壊さずに並列化でき、モジュール側から見た「1 プロセス＝1 台」の前提も保てるためです【推論】。

fabriq では、無人実行が止まる不具合がいくつも記録されています。上限なしの待機（t-0053、t-0054）、タイムアウトのない MessageBox（t-0055）、AutoPilot ガードのない `Read-Host`（t-0056）です ([.tm/tasks.json](../.tm/tasks.json))。1 対多モデルでは、こうしたハングの影響が N 倍になります。そのため、すべての外部呼び出し（adb・pm・am・ネットワーク待ち）に上限時間を設けること、無人実行中の対話をカーネルが一律に禁止して「要対応」状態に置き換えること、全行が無効のときや Segment が一致しないときは Error ではなく Skipped を返すことの 3 点を、カーネル側の設計で一律に封じる必要があります。`__RESTART__` は、ホストが再起動で落ちないため単純になります。ホストは「`adb reboot` → 再接続を待つ → `sys.boot_completed` をポーリング → 続行」を実行し、戻らない端末には `RebootTimeout` 遷移を設けます。fabriq_mac は再起動からの再開をしない設計でしたが、Android のホストモデルでは**再起動を跨いだ続行を安価に実現できます**【推論】。

### 自作 DO エージェントとのハイブリッド構成

tier (b) では、ホストが受付・配布・回収・一覧を担当し、エージェントは ADB では届かない操作だけを担当します。ホストからエージェントへの命令には、エクスポートした ContentProvider の `call()` を `adb shell content call --uri <URI> --method <METHOD> [--arg <ARG>]` で呼び出す方式が最も堅牢です。provider 側で `Binder.getCallingUid()` が shell（2000）か root（0）かを検査でき、戻り値の Bundle は `Result: Bundle[...]` として標準出力に出るため、ホストでパースして `New-ModuleResult` に写像できます ([Content.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/cmds/content/src/com/android/commands/content/Content.java))【確度: 高】。`am broadcast` は、送信元の UID を受信側で確実に取得できないため、補助的な用途に限ります ([BroadcastReceiver.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/core/java/android/content/BroadcastReceiver.java))。

証明書の一括導入のような長時間の処理は、`call()` でジョブ ID だけを返し、ホストが状態をポーリングする非同期パターンにすることが推奨されます。この形は、fabriq の async 中断点や再起動跨ぎ（再起動後に状態を再照会する）の考え方とも整合します【推論】。証明書や CSV のような大きな入力は `adb push` か `adb reverse` で渡しますが、エージェントが `/data/local/tmp` のファイルを読めるかは未検証です。

運用面で最も重いのは署名鍵です。署名が変わると通常の更新ができなくなり、QR の `SIGNATURE_CHECKSUM` も変わります。鍵を失った場合の逃げ道は、事前に用意した `transferOwnership`（API 28）による移譲か初期化しかありません ([API diff 28](https://developer.android.com/sdk/api_diff/28/changes/android.app.admin.DevicePolicyManager))。このため、鍵は fabriq の `kernel/` と同等の重要資産として永続的に厳重保管し、エージェントの versionCode とホスト側モジュールの互換表（REQUIRES_KERNEL 相当）を管理する必要があります【推論】。参照実装には Apache-2.0 の TestDPC が推奨されます ([android-testdpc](https://github.com/googlesamples/android-testdpc))。GPL-3.0 の OwnDroid や Dhizuku を取り込む場合は、ライセンス上の扱いに注意してください ([OwnDroid](https://github.com/BinTianqi/OwnDroid); [Dhizuku](https://github.com/iamr0s/Dhizuku))。

### 初版スコープ：約 12 モジュールと規模の目安

fabriq の 79 モジュール ([modules/](../modules/)) を分類すると、作り直せば Android に対応する概念を持つもの（B）が約 30、そのまま移せる汎用のもの（A。ログ送信・再試行ラッパー・試験用モジュールなど）が約 6、Windows 専用のもの（C）が約 43 になります【推論】。初版は次の 12 モジュールに絞ることが推奨されます。

| # | モジュール（案） | fabriq の手本 | 主な適用 | Step 5.5 の検証 |
|---|---|---|---|---|
| 1 | `device_precheck`（＋`__GATE__`） | `hostlist_precheck` | シリアル・機種・OS・パッチ・電池残量・ADB 認可・アカウント数・ユーザー数を照合 | 読み取り専用のゲート。不一致なら Error |
| 2 | `apk_install` | `app_config`／`winget_install` | `adb install -r -g`／`install-multiple`。versionCode が同じか新しければ skip | `pm list packages --show-versioncode` |
| 3 | `package_disable` | `storeapp_config`／`bloatware_remove` | `pm disable-user --user 0`／`pm uninstall -k --user 0` | `pm list packages -d`／`-u` |
| 4 | `settings_config` | `reg_hklm_config` | `settings put system/secure/global` | `settings get`（Test-…Match パターン） |
| 5 | `wifi_config` | `ssid_config` | `cmd wifi add-network`（PSK） | `cmd wifi list-networks`／`status` |
| 6 | `device_name_config` | `hostname_config` | `settings put global device_name` | 同じキーの `get` |
| 7 | `display_sound_config` | `power_config`／`brightness_config`／`volume_config` | 消灯・輝度・回転・`stay_on_while_plugged_in` | `settings get` |
| 8 | `time_config` | `time_sync_config` | `cmd alarm set-timezone`、自動検出の設定 | `getprop persist.sys.timezone` |
| 9 | `file_push_config` | `copyfile_config`／`file_delete` | `adb push`／`rm`（パスガードは確認ゲートの外側に置く） | 存在確認 |
| 10 | 再起動ステップ | `restart_config`／`__RESTART__` | `adb reboot` → `wait-for-device` | `sys.boot_completed`＋ヘルスチェック |
| 11 | `evidence_inventory` | `evidence_config` | `getprop`、パッケージ一覧、`settings list`、`dumpsys battery`、`screencap` | manifest の各セクションの status |
| 12 | `manual_step_guide` | `manual_kitting_assistant` | SUW、Google アカウント、CA 証明書の UI、OS 更新の案内と記録 | 作業者のチェック記録（`Verified=$null`） |

加えて、終端処理として**画面ロック PIN の設定と最終ハードニング**を初版から同梱することが推奨されます。どちらも、項目別表の #10（画面ロック）と #27（最終ハードニング）が示すとおり、実行順序に制約がある処理だからです。

規模の目安として、fabriq の各部の規模を次に示します ([kernel/common.ps1](../kernel/common.ps1); [kernel/main.ps1](../kernel/main.ps1); [kernel/KERNEL_API.md](../kernel/KERNEL_API.md); [modules/](../modules/); [tests/](../tests/); [CHANGELOG.md](../CHANGELOG.md))。

| 対象 | 規模 |
|---|---|
| カーネル | `common.ps1` が 5,500 行（関数 114）、`main.ps1` が 2,296 行（関数 9） |
| 公開 API | 関数 32、公開グローバル変数 6、特殊マーカー 6 |
| モジュール | 79 ディレクトリ（standard 63／extended 16）、スクリプト 104 本で計 36,548 行 |
| GUI アプリ | 計 12,947 行（operator が 4,065 行） |
| テスト | `It` ブロック 671 個（うちカーネル契約向けが 354 個） |
| CHANGELOG | 6,265 行 |

Android 版の初版は、カーネルの契約層（結果・CSV・暗号・マスク・プロファイル解析・ゲート・実行履歴・HTML、約 2〜3 千行相当）と 12 モジュールで構成します。規模は fabriq の 1〜2 割程度から始められる見込みです【推論】。

ただし、**端末ワーカ・ADB ラッパー・タイムアウト・再接続待ちという並列実行層は fabriq に前例のない新規部分で、最大の設計リスク**です【推論】。カーネル契約向けの既存テスト（Import-ModuleCsv 16 件、New-BatchResult 18 件、Resolve-ProfileModules 21 件、GateBarrier 15 件、ResumeState 29 件）は、そのまま受け入れ基準のチェックリストとして使えます。

### Mac 版との比較

| 観点 | Mac 版（fabriq_mac） | Android 版（本提案） |
|---|---|---|
| 実行位置 | 対象 Mac 上で bash を実行（1 対 1） | ホスト PC から ADB で N 台を操作（1 対多） |
| 再起動跨ぎ | 再開しない | `adb reboot` と起動完了待ちで、安価に続行できる |
| 状態の保存 | シリアル単位 | シリアル単位（同じ） |
| Studio／Evidence Manager との互換 | 追わない | 追わないことが推奨される（manifest の前方互換の考え方のみ採用） |
| GUI とコアの分離 | GUI が bash のコアを呼ぶ | GUI が CLI コアを呼ぶ二層構造が推奨される |

Android 版に固有の事情は 2 点あります。1 つは、ADB を使う方式では、SUW・開発者オプション・USB デバッグ・RSA 許可という前段の手作業が端末ごとに必ず残ることです。もう 1 つは、「MDM なし」の上限が二値ではないことです。ADB 単独では証明書・VPN・強制は不可能ですが、自作 DO で補完できます。ただし、その DO を「MDM」とみなすかどうかは顧客との合意事項になります。

## 運用制約：USB ドライバ、文字コード、ハブ構成、鍵とアカウントの管理が量産性を左右する

ホスト PC と ADB で量産する場合、ボトルネックは Android の API よりも、Windows ホストの環境、USB の物理構成、端末側のセキュリティ機能、アカウントと鍵の管理にあります。

| 区分 | 制約 | 影響 | 推奨対策 | 確度 |
|---|---|---|---|---|
| USB ドライバ | Google 端末には Google USB Driver が必要で、他社の端末は OEM が提供する ([win-usb](https://developer.android.com/studio/run/win-usb); [oem-usb](https://developer.android.com/studio/run/oem-usb))。公式ページには Windows 11 の手順も Windows Update への言及もない。`android_winusb.inf` を改変すると署名の警告が出る。シャープについては、旧メディアタブレット向けのページしか確認できなかった | 機種ごとに認識されないことがある | Google・Samsung・Sony のドライバを事前に導入する。AQUOS・arrows・京セラは、adbd の MS OS ディスクリプタによる inbox WinUSB での認識を期待値として、機種別に検証する ([adbd WinUSB commit](https://gerrit.omnirom.org/plugins/gitiles/android_system_core/+/b310da608b8292999cb30ed29bfb0bf4a9cb606d%5E%21)) | 中 |
| platform-tools の版 | 36.0.2 で、Windows の push／pull で末尾 1 文字が欠ける不具合と Samsung の検出不良が修正された。37.0.1（2026-07）では Windows の USB 実装が `libadbusb` に置き換わり、`ADB_USB_LEGACY=1` で旧実装に戻せる ([platform-tools release notes](https://developer.android.com/tools/releases/platform-tools)) | 版によって挙動が変わる | 36.0.2 以上に固定する。37.0.1 は台数を増やした状態で回帰試験を行う | 高 |
| ポート 5037 | すべての adb クライアントが 5037 番ポートを使う ([adb](https://developer.android.com/tools/adb))。他社のスイート（Apowermirror、Nox、BlueStacks など）が独自の adb を同梱しており、ポートを奪い合う ([Medium](https://medium.com/@liwp.stephen/solved-adb-server-version-doesnt-match-this-client-493cec01f032); [scrcpy FAQ](https://github.com/Genymobile/scrcpy/blob/master/FAQ.md)) | 版の不一致による切断 | 起動時に 5037 の所有プロセスを検出し、確認なしに停止しない。scrcpy などは `ADB` 環境変数で同じ adb を指すようにする | 高（仕様）／中（事例） |
| PowerShell 5.1 の文字コード | ネイティブ出力のデコードは `[Console]::OutputEncoding` に従い、日本語版 Windows では CP932 のままになる ([about_Character_Encoding](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_Character_Encoding); [マイナビ TECH+](https://news.mynavi.jp/techplus/article/techp5170/)) | 文字化けにより `-match` 判定が壊れる | adb を呼ぶ間は UTF-8 を明示するか、`System.Diagnostics.Process` の `StandardOutputEncoding` を使う。バイナリ出力（`screencap`）はパイプラインに通さない【推論】 | 中 |
| 台数・USB トポロジ | Intel の xHCI は総エンドポイント数 96 が一般的で、端末は約 5 エンドポイントを使う ([Cambrionix](https://docs.cambrionix.com/general/guides/understanding-and-managing-usb-endpoint-limitations))。STF は電源付きハブで約 28 台、Haswell 世代では 9〜12 台が限界、推奨ハブは USB 2.0 ([STF](https://github.com/DeviceFarmer/stf))。AstroFarm は 15 台 ([42Gears](https://www.42gears.com/wp-content/uploads/2026/01/AstroFarm-Recommendation-v2.pdf))。Tablet admin は最大 40 台で、USB 2.0 を推奨 ([tabletadmin](https://www.tabletadmin.com/android-kitting.html)) | 台数を増やすと間欠的な切断が起きる | 既定は 10〜20 台とする。40 台は、複数の xHCI コントローラ、USB 2.0 のセルフパワーハブ、ポート単位の電源制御で実証してから採用する | 中 |
| インストール速度 | Samsung の「USB 経由のアプリを検証」により、4.7 MB の APK に約 60 秒かかる（OFF にすると 2 秒未満） ([B4X](https://www.b4x.com/android/forum/threads/solved-b4a-installing-time-very-large.168106/post-1030317)) | スループットが落ちる | ポリシーが許す場合だけ OFF にする | 低 |
| Samsung Auto Blocker | USB 経由のコマンドを遮断する。One UI 6.1.1 で出荷された端末では既定で ON。One UI 9 beta の最大制限では USB 接続自体を遮断する ([Samsung](https://www.samsung.com/uk/support/mobile-devices/protect-your-galaxy-device-with-the-new-auto-blocker-feature/); [Android Authority](https://androidauthority.com/samsung-one-ui-9-auto-blocker-update-3666893/)) | ADB を使えない | 前段の手順に OFF を明記する | 中 |
| Advanced Protection | Android 16 では、提供元不明アプリの許可ができなくなり、ロック中は USB データが無効になる。開発者オプションを封鎖する計画も報じられている（未確認） ([Android Authority](https://www.androidauthority.com/android-16-usb-data-advanced-protection-3548018); [同](https://androidauthority.com/android-advanced-protection-mode-developer-options-3679725)) | ADB を使えない／将来のリスク | 有効化された端末は対象外にする | 中／低 |
| FRP | Android 15 から、SUW を回避しても FRP は解除されない ([Android Authority](https://www.androidauthority.com/android-15-factory-reset-protection-upgrades-3479431))。パスワード変更後 24〜72 時間は新しいパスワードが通らない ([Samsung NZ](https://www.samsung.com/nz/support/mobile-devices/what-is-google-frp/)) | 初期化で端末が使用不能になる | 既定ではアカウントを追加しない。アカウントが残った状態でのリカバリ初期化を禁止する | 中 |
| platform-tools の再配布 | SDK License の §3.4 は再配布を禁止し、§3.5 は OSS 部分を例外とする ([platform-tools](https://developer.android.com/tools/releases/platform-tools))。scrcpy は dl.google.com から取得して SHA-256 を照合している ([adb_windows.sh](https://github.com/Genymobile/scrcpy/blob/master/app/deps/adb_windows.sh))。WinGet の掲載情報は Apache-2.0 と表示しており、SDK License と矛盾する（信頼できないメタデータとみなす） ([wingetly](https://wingetly.io/apps/google/platform-tools)) | 同梱すると規約上のリスクがある | セットアップ時にダウンロードし、SHA-256 を照合する（法務確認が必要） | 中 |
| アプリの調達 | 参照機からの抽出は、著作権や Play の規約に抵触するおそれがある ([Protectstar](https://protectstar.com/de/blog/apk-files-what-you-need-to-know-are-they-harmful-or-illegal))。Play Integrity が UNLICENSED と判定すると、Play からのインストールを求められる ([Play Integrity](https://developer.android.com/google/play/integrity/remediation)) | 起動できない・更新が届かない | 再配布権のある APK に限定し、出所と SHA-256 を CSV に記録する | 中 |
| adbkey の共有 | `ADB_VENDOR_KEYS` で追加の鍵を指定できる ([adb man](https://man.archlinux.org/man/adb.1.txt))。鍵をコピーした PC は、端末から認可済みとみなされる ([O'Reilly](https://oreilly.com/library/view/practical-mobile-forensics/9781788839198/008a98ec-157d-42a2-960a-579e5b30cbfb.xhtml)) | 漏えいすると、USB デバッグが有効なままの全端末で shell を取られる | 拠点ごとに 1 本とし、ACL／DPAPI で保管する。案件ごとにローテーションし、出荷前に認可を取り消す | 中 |
| USB デバッグの残置 | STIG は無効化を要求する ([STIG V-56059](https://stigviewer.com/stig/samsung_android_with_knox_2.x/2015-01-05/finding/V-56059))。Singpass などは USB デバッグが有効だと動作を拒否する ([ask.gov.sg](https://ask.gov.sg/singpass/questions/cm8idm20i0057wd6ud5h2kq4a)) | 業務アプリが動かない／攻撃面が残る | 最終ハードニングを既定にする | 中 |
| 個人情報 | 携帯電話番号は、それ自体では個人識別符号とされていない ([アガルート資料](https://www.agaroot.jp/wp-content/uploads/2026/03/gyosei_tajima_kiso_kojin.pdf))。IMEI は端末に固有で、所有者が変わっても変わらない ([android.com](https://www.android.com/intl/ja_jp/articles/317)) | 利用者の氏名と結び付いた台帳やエビデンスは個人情報になる | マスク列、アクセス制御、保存期間を定める（法務確認が必要） | 低〜中 |
| テスト | エミュレータは、VM の中ではハードウェアアクセラレーションを使えない。AEHD は 2026-12-31 に終了する ([emulator acceleration](https://developer.android.com/studio/run/emulator-acceleration)) | fabriq の VM テストリグはそのまま流用できない | fake adb を使った PS 5.1 での単体テスト → 物理ホスト上のエミュレータ → 実機回帰、の 3 層で構成する | 高 |

### 台数の上限はハードウェアで決まる

公開されている adb 自体の資料には、USB 端末の台数上限は見当たりません。16 台が上限とする逸話は、Appium フォーラムの出典のない投稿で、STF や商用製品の実績と矛盾します ([Appium discuss](https://discuss.appium.io/t/maximum-how-many-mobile-phone-i-can-connect-via-usb-cable-over-adb/34491))。上限を決めているのは、xHCI のエンドポイント予算、ハブの品質、給電能力です。したがって、**USB の切断と再接続は例外ではなく通常の事象として扱い**、シリアルによる再接続の検出、端末単位の再試行、冪等なステップで吸収することが推奨されます【推論】。

### Wi-Fi ADB は量産の起点にはならない

Wi-Fi ADB は Android 11 以降でペアリングが必要で、そのペアリングは端末ごとの操作です。Android 17 と adb 37 で導入された「adb Wi-Fi 2.0」は、信頼済みネットワークへの自動接続を改善するものです ([adb](https://developer.android.com/tools/adb); [Android Developers Blog](https://android-developers.googleblog.com/2026/09/wireless-debugging-adb-wifi-2.html) ※)。いずれにしても、新品を何十台もまとめて立ち上げる起点にはなりません。

### 価値は人手の削減より、並列処理量とエビデンスの品質にある

経済性の観点では、国内のスマートフォンのキッティング委託相場は 1 台あたり約 ¥1,000〜3,000 です ([テクバン](https://biz.techvan.co.jp/tech-is/blog/device/002936.html))。ベンダーの公称値では、40 台構成で 1 台あたりの作業時間が 30 分から 3 分に短縮されています ([tabletadmin](https://www.tabletadmin.com/android-kitting.html))。人が行う前段の作業は残るため、本ツールの価値は、人手をなくすことよりも、並列処理量とエビデンスの品質にあります【推論】。

## ポリシー変更（2025〜2026 年）：自作 DPC と AMAPI 直結の道は狭まり、ADB 経路は残っている

| 変更 | 時期 | キッティングツールへの影響 | 確度 |
|---|---|---|---|
| AMAPI の Permissible Usage が、利用者を商用 EMM 開発者・Device Trust 提供者・OEM に限定。社内の自社アプリ専用ソリューションなどを禁止 ([AMAPI Permissible Usage](https://developers.google.com/android/management/permissible-usage); [AE community](https://www.androidenterprise.community/android-enterprise-general-discussions-3/android-management-api-for-internal-use-822)) | 現行 | 自社用や単一顧客用のツールが AMAPI を直接使うと、規約違反とされるリスクがある | 中 ※ |
| 未検証のプロジェクトの上限が 500 台に（2024-03）→ FRP・初期化ポリシーが使えなくなる（2025-05）→ **新規プロジェクトの既定クォータが 0 台に** ([Bayton 2024-03](https://bayton.org/blog/2024/03/amapi-permissible-usage/); [Bayton notes](https://bayton.org/notes/99/); [Bayton FAQ](https://bayton.org/android/android-enterprise-faq/amapi-usage-limits/); [AE community](https://www.androidenterprise.community/discussions/conversations/android-enterprise-partner-applicationquota-status/13543)) | 2024-03〜2025-11 | 承認がなければ 1 台も登録できない。承認の回答には 7 営業日以上から数週間かかる ([AE community](https://www.androidenterprise.community/android-enterprise-general-discussions-3/org-usage-limits-error-no-enrolled-devices-2661))。**日付に食い違いがある**: コミュニティ投稿は 2025-10-29 の文言更新とし、Bayton は 2025-11-01 の発効としている | 中 ※ |
| Play EMM API が custom DPC の新規登録を停止 ([Build a DPC](https://developer.android.com/work/dpc/build-dpc); [Play EMM API](https://developers.google.com/android/work/play/emm-api)) | 現行 | 自作 DPC では managed Google Play、managed アカウント、afw# を使えない | 高 |
| custom DPC 向けの Play EMM API の一部メソッドが 2025-09-30 に停止 ([Bayton](https://bayton.org/android/android-enterprise-provisioning-methods/); [Omnissa KB 93986](https://kb.omnissa.com/s/article/93986)) | 2025-09-30 | 既存の EMM ベンダーの DPC は動き続ける。custom DPC は延命状態と評されている ([Bayton](https://bayton.org/blog/2026/04/dpc-migration/)) | 中〜低 ※（Google の公式ページでは未確認） |
| Play Protect が、プロビジョニング中に承認済み DPC 以外をブロック（「App blocked to protect your device」） ([Google Help 16694822](https://support.google.com/work/android/answer/16694822); [Bayton 2025-12-24](https://bayton.org/blog/2025/12/the-dpc-allowlist/)) | 2025 年中に予告なく有効化 | tier (c) の成否を左右する。**証拠は混在している**: Test DPC や Android 14 でのブロック報告がある一方 ([AE #2909](https://www.androidenterprise.community/android-enterprise-general-discussions-3/play-protect-blocking-custom-dpc-app-during-6-tap-qr-provisioning-appeal-submitted-2909); [AE #14202](https://ypaoq37543.lithium.com/discussions/conversations/issue-play-protect-blocks-dpc-installation-during-qr-provisioning-on-android-14-/14202))、自作 DPC が 2026-10 に Pixel 6a でブロックされなかった報告もある ([muster #89](https://github.com/quadseven/muster/issues/89))。ADB の `dpm set-device-owner` 経路に適用されるかは不明（推論では対象外の可能性が高い） | 制度の存在: 中 ※／最小構成の自作 DPC がブロックされるか: 低／ADB 経路への適用: 不明 |
| Android 15 の FRP 強化（SUW を回避しても FRP が解除されず、アカウント追加・画面ロック・アプリ導入がブロックされる）。Android 16 では回避を検知したら再初期化する計画 ([Android Authority A15](https://www.androidauthority.com/android-15-factory-reset-protection-upgrades-3479431); [A16](https://www.androidauthority.com/android-16-factory-reset-protection-upgrades-3556859); [AE community](https://www.androidenterprise.community/blog/resources/enhanced-factory-reset-protection-in-android-15/9493)) | A15／A16 | Google アカウントを残したまま初期化すると致命的。testharness との相互作用は未検証。管理端末では、Android 15 以降の hard reset 後に EFRP が常に適用されるとの指摘がある ([Bayton](https://bayton.org/android/android-enterprise-faq/frp-vs-enterprise-frp/) ※) | A15: 中〜高／A16: 低（出荷されたか未確認） |
| 開発者検証制度が 2026-09-30 にブラジル・インドネシア・シンガポール・タイで開始。グローバル展開は 2027 年。**ADB は対象外**。未検証アプリ向けの「advanced flow」は 2026-08 から展開 ([Google Help](https://support.google.com/android-developer-console/answer/16561738); [Android Authority](https://www.androidauthority.com/google-android-advanced-flow-sideloading-rollout-begins-3700073/); [Mega Mobile Content](https://www.megamobilecontent.com/news/2026/10/01/android-developer-verification-live-brazil-indonesia-singapore-thailand/)) | 2026-09-30〜 | 2026-10 時点では、ADB による配布への影響は小さい。日本での施行日は記載がない。施行後も ADB の対象外扱いが続いているかは一次確認できていない | 中〜高 ※ |
| Android 17 の安定版リリース（Pixel は 2026-06-16、Samsung One UI 9 は 2026-09-16 から順次） ([Gizmochina](https://www.gizmochina.com/2026/06/09/heres-when-google-pixels-will-receive-android-17-update/); [9to5Google](https://9to5google.com/2026/09/16/samsung-android-17-one-ui-9-update-list/))。全アプリ向けの変更点に、adb・設定・インストール・ユーザー補助・Wi-Fi に関する記載はない ([Android 17 behavior changes](https://developer.android.com/about/versions/17/behavior-changes-all))。API 37 をターゲットにすると `ACCESS_LOCAL_NETWORK` が必要になり、Theft Detection Lock が既定で ON になる ([Bayton A17](https://bayton.org/android/android-17-enterprise-features/)) | 2026-06〜 | 日本メーカーの法人機は段階的に更新されるため、当面は Android 13〜17 の混在フリートを前提とする【推論】。Shell 権限と DPMS の Android 17 での差分は未確認。盗難対策との関係から、再キッティング時の挙動を検証する必要がある | リリース: 中（報道）／挙動変更: 高（公式）／コード差分: 未確認 ※ |
| そのほか: Android 14／15 の低 targetSdk のブロック、Android 16 での `SKIP_EDUCATION_SCREENS` の無視、AMAPI SDK 1.6.0 での APK 直接配布（2025-08）、platform-tools 37.0.1 の `libadbusb`、AEHD の終了（2026-12-31） ([Bayton A16](https://bayton.org/android/android-16-enterprise-features/); [platform-tools](https://developer.android.com/tools/releases/platform-tools)) | 2025〜2026 | 版ごとの能力表と、platform-tools の版固定で吸収する | 中〜高 |

### 変更の方向は一貫して締め付け

これらの変更は、一方向に締め付けが進んでいる点で共通しています。Google は、GMS 端末で Device Owner を使う用途を、AMAPI／Android Device Policy と、allowlist に載った検証済み DPC の 2 つの経路へ集約しています。どちらの経路も、単発の社内キッティングツールを想定していません ([Play EMM API](https://developers.google.com/android/work/play/emm-api) ※; [Bayton: MDM is dead](https://bayton.org/blog/2026/03/mdm-is-dead-long-live-ace/))。

### ADB 経路は「開発者向けの例外」として残っている

一方、ADB の経路は「開発者・パワーユーザー向けの例外」として残っています。開発者検証制度では対象外とされ、`dpm set-device-owner` は ManagedProvisioning を経由しません。ただし、Google は ADB による DO 設定を本番展開の手段とは位置付けていません ([Bayton FAQ](https://bayton.org/android/android-enterprise-faq/can-i-set-device-owner-without-factory-reset/) ※)。また、Advanced Protection による開発者オプション封鎖の計画も報じられています（公式確認なし） ([Android Authority](https://androidauthority.com/android-advanced-protection-mode-developer-options-3679725))。このため、**ADB＋DO の構成も「将来にわたって保証された経路」ではない**と認識しておく必要があります。

### 設計への示唆：適用経路を差し替えられる構造にする

今後も規制が締まる方向が続くと仮定して、設計に織り込むことが推奨されます。具体的には、(a) を基盤にし、(b) を任意で追加でき、将来は (d) の商用 EMM や OEM ツールへ適用経路を差し替えられる構造にします【推論】。

## 判断が必要な事項：方式、「MDM なし」の定義、初期化方式の 3 点が設計を分岐させる

| # | 判断事項 | 選択肢 | 推奨 | 理由 |
|---|---|---|---|---|
| 1 | 採用する方式 | (a) のみ／(a)＋(b)／(c)／(d) | (a) を基盤とする。証明書・EAP・強制が必要な案件に限って (b) を追加する。(c) は Play Protect の実機検証を通過した後に検討する。(d) は、顧客が EMM を持つ場合に限り「EMM の前段のオーケストレータ」として使う | (a) の上限、(b) で追加できる範囲、(c) と (d) のポリシーリスク |
| 2 | 「MDM なし」の定義 | サーバーやクラウドによる常駐管理がない／端末に管理アプリがない | 「常駐するクラウド管理がない」と定義する。自作 DO（組織管理端末になり、解除は原則初期化）を許容するかどうかは書面で合意する | DO 端末は組織管理端末として表示され、元に戻すには原則として初期化が必要 |
| 3 | 再キッティング（初期化）の方式 | 設定画面から初期化して SUW を完走／`cmd testharness enable`／DO の wipe | 標準は「Google アカウントを削除 → 設定画面から初期化 → SUW」。testharness は、対象機種で実機検証を通過するまで採用しない | testharness では、テストハーネスとしての判定、OTA の自動更新とアカウント同期の無効化、OEM 初期設定の未実施が起きる ([TestHarnessModeService.java](https://github.com/LineageOS/android_frameworks_base/blob/lineage-23.0/services/core/java/com/android/server/testharness/TestHarnessModeService.java)) |
| 4 | Google アカウントの方針 | 追加しない／会社の一時アカウント／利用者のアカウント | 既定は「追加しない」とし、hostlist に `GoogleAccountPolicy` 列を設ける | FRP 対策と、(b) の前提（アカウント 0） |
| 5 | 出荷時の USB デバッグ | OFF／保守用に残す | 既定は OFF（最終ハードニングを必須にする）。保守用に残す場合は、プロファイルに明示してリスクを記録する | STIG、業務アプリによる拒否、鍵共有のリスク |
| 6 | adbkey の共有範囲 | 全拠点で共通／拠点ごと／案件ごと | 拠点ごとに 1 本とし、案件ごとにローテーションし、出荷前に認可を取り消す | 漏えいしたときの被害範囲を限定する |
| 7 | platform-tools の入手方法 | 同梱／セットアップ時にダウンロード／AOSP からビルド | セットアップ時に dl.google.com から固定版を取得し、SHA-256 を照合して 3 ファイルだけ展開する | SDK License §3.4（法務確認が必要） |
| 8 | adb の版 | 36.0.2／37.0.1 | 36.0.2 以上を前提にする。37.0.1 は台数を増やした状態で検証し、`ADB_USB_LEGACY=1` を切り替え手段として文書化する | Windows の push／pull の不具合、`libadbusb` への移行 |
| 9 | 1 ホストあたりの台数 | 10〜20 台／20〜40 台 | 既定は 10〜20 台。40 台はハードウェア構成を実証した後 | STF、AstroFarm、Tablet admin の実績 |
| 10 | ホストの実装言語と GUI | PS 5.1＋WinForms／.NET exe | GUI が CLI コアを呼ぶ二層構造にする。第一候補は fabriq との継続性から PS 5.1 とし、配布物のコード署名費用を予算に入れる | fabriq の Defender 誤検知の前例（t-0059、t-0060） ([.tm/tasks.json](../.tm/tasks.json)) |
| 11 | fabriq の周辺ツールとの互換 | Studio・Evidence Manager と互換にする／独自にする | 互換は追わない。manifest の前方互換の考え方だけを採用する | fabriq_mac と同じ判断 |
| 12 | OS 更新の扱い | キッティング中に適用／Verify-only | 既定は Verify-only（パッチレベルと指紋を期待値と照合）。適用は任意の手作業モジュールにする | One UI 9 の更新は約 3.8 GB で、これが台数分かかる ([9to5Google](https://9to5google.com/2026/09/16/samsung-android-17-one-ui-9-update-list/)) |
| 13 | アプリの調達方法 | ベンダー提供の APK／参照機からの抽出／managed Play | 再配布権のある APK だけを使い、出所と SHA-256 を記録する。参照機からの抽出は社内アプリに限る | 著作権、Play Integrity |
| 14 | ロケール | ヘルパーで変更する／Verify-only | 日本語のまま使う案件では Verify-only にする。変更が必要な場合だけヘルパーを使う | hidden API への依存と、版ごとの回帰リスク |
| 15 | 自作 DPC の署名鍵と testOnly | testOnly ビルド／本番署名 | 鍵は永続的に厳重保管する。testOnly は検証用に限り、本番では `DISALLOW_DEBUGGING_FEATURES` と組み合わせる | testOnly の DO は、ADB で誰でも外せる |
| 16 | IMEI などの扱い | すべて保存／マスクする | 台帳とエビデンスにマスク列を設け、アクセス制御と保存期間を定める | 個人情報に当たる（法務確認が必要） |

## 実機検証が必要な項目：高リスク 5 項目が方式選定の前提を左右する

| # | 検証項目 | 確認する内容 | リスク | 失敗したときの代替 |
|---|---|---|---|---|
| 1 | `cmd testharness enable` を再キッティング経路として使えるか | Samsung・AQUOS・arrows・京セラでの実装有無（PersistentDataBlock に依存）、消去後に ADB が自動で再接続されるか、SUW スキップ後のキャリア初期設定・Play・OEM アプリの挙動、`isRunningInUserTestHarness()` による業務アプリの挙動差、FRP が掛からないか | 高 | 設定画面からの初期化＋SUW 完走 |
| 2 | 日本メーカーと Samsung での `dpm set-device-owner` | AQUOS・arrows・京セラ・Xperia・Galaxy で成功するか。プリインストールアプリやキャリアが作るアカウントで `ACCOUNTS_NOT_EMPTY` にならないか。Secure Folder などの追加ユーザーの扱い。参考として、Xiaomi・OPPO の失敗事例と、Samsung の「user limit reached」（文脈は不明確）が報告されている ([OwnDroid](https://github.com/BinTianqi/OwnDroid)) | 高 | 機種ごとに (b) の対象外とし、(a) に縮退する |
| 3 | 自作エージェントに対する Play Protect の挙動 | `adb install`＋`dpm set-device-owner` の経路でスキャンやブロックが起きるか。QR 経路で allowlist によるブロックが起きるか。宣言する権限を最小化したときの効果 | 高 | QR 経路を断念し、(b) だけにする |
| 4 | ハブ構成と台数スケール | 10／20／40 台での切断頻度、37.0.1 の `libadbusb` と `ADB_USB_LEGACY=1` の比較、並列インストールの所要時間 | 高 | 台数の上限を下げる、コントローラを増設する |
| 5 | Android 15 以降の FRP と初期化の関係 | testharness・設定画面からの初期化・DO の wipe のそれぞれで、FRP や EFRP が発動するか。Android 17 の Theft Detection Lock との関係 | 高 | アカウントを残さない運用を徹底する |
| 6 | Android 14 以降での APN の挿入 | `content insert --uri content://telephony/carriers` が成功するか。OEM が shell の `WRITE_APN_SETTINGS` を削っていないか。キャリア端末の APN 編集ロック、preferapn の選択 | 中 | DO の `addOverrideApn`、または手作業のガイド |
| 7 | 制限付き設定の下でのユーザー補助の有効化 | `settings put secure enabled_accessibility_services` が Android 13〜17（Android 15 の Enhanced Confirmation Mode を含む）で効くか | 中 | 手作業のガイド |
| 8 | 最終ハードニングの挙動 | `adb_enabled 0` を実行したときのホスト側の終了コード。`development_settings_enabled 0` を直接書いたときに開発者設定がリセットされるか。OEM が鍵を自動消去するか | 中 | UI で開発者オプションを OFF にするようガイドする |
| 9 | ロケール変更ヘルパー | `app_process` や `CHANGE_CONFIGURATION` を付与した APK が Android 14〜17 で動くか | 中 | Verify-only に切り替える |
| 10 | IMEI の取得 | `app_process` から `TelephonyManager.getImei()` を呼べるか、`service call iphonesubinfo` の番号は何か | 中 | `*#06#` 画面のスクリーンショット |
| 11 | `cmd wifi add-network` の機種差と導入版 | Android 11 以前で使えるか（導入版は一次資料で未確認）、OEM の Wi-Fi スタックによる差 | 中 | 手作業のガイド、または DO |
| 12 | Windows 11 の USB ドライバ | 各機種が inbox WinUSB で認識されるか | 中 | 機種別のドライバ表を整備する |
| 13 | Samsung 固有の機能 | Auto Blocker を OFF にする手順、「USB 経由のアプリを検証」を OFF にできるかとその所要時間 | 中 | 手順書に記載する |
| 14 | DO エージェントの IPC | `/data/local/tmp` のファイルを読めるか、`content call --extra` の構文、`adb reverse` の localhost 通信が Android 17 のローカルネットワーク権限の対象になるか | 中 | 渡し方を変更する |
| 15 | Android 17 の差分 | Shell マニフェスト、DPMS の ADB 前提条件、SettingsProvider の shell 免除が変わっていないか | 中 | 版ごとの能力表に反映する |
| 16 | 再起動後の安定化 | `sys.boot_completed` の後に PackageManager が安定するまでの時間、OEM の初回起動処理との競合 | 中 | ヘルスチェックの追加、待機時間の延長 |
| 17 | OEM 差の小さい項目 | `device_name` が BT や Wi-Fi Direct の名前に反映されるか、輝度の値域、Samsung などの省電力機能 | 低 | OEM ごとの CSV 値で吸収する |

## 次のステップ：一次情報の確定と 3 段階の PoC から着手する

| 順 | 作業 | 目的 | 完了条件 |
|---|---|---|---|
| 1 | 一次情報の再確認（developers.google.com、support.google.com、android.googlesource.com、docs.samsungknox.com、bayton.org を閲覧できる環境で実施） | ※ を付けた結論を確定させる | AMAPI Permissible Usage の原文、Google ヘルプ 16694822、FRP と EFRP のヘルプ、開発者検証の施行後の運用、AOSP の Android 17 の該当ソースを確認済み |
| 2 | 顧客との前提合意 | 設計の分岐を確定させる | 判断事項の #1〜#5 と #12 を書面で合意済み |
| 3 | 実機検証ラボの構築 | 機種×版の能力表を作る | Pixel・Galaxy・AQUOS・arrows・京セラの 5〜10 台と USB 2.0 のセルフパワーハブで、高リスクの 5 項目を完了 |
| 4 | PoC-1：1 台・ADB のみ | モジュール契約を検証する | `device_precheck`・`settings_config`・`apk_install` が Step 1〜6＋5.5 で動き、2 回目の実行が Skipped になる。fake adb の単体テストが PS 5.1 で exit code 0 |
| 5 | PoC-2：並列ワーカ | 新規の実行層を検証する | 10 台を並列に動かし、すべての adb 呼び出しにタイムアウトがあり、USB の抜き差しから復帰でき、端末×ステップのマトリクスが表示される |
| 6 | PoC-3：DO エージェントの最小構成 | tier (b) の価値を検証する | ADB 経路で DO 化し、`content call` で CA の導入・EAP Wi-Fi・ユーザー制限を適用して読み返せる |
| 7 | ガバナンスの初期化 | fabriq 流の保守性を初日から確保する | KERNEL_API 相当の文書 v0.1.0、乱数ソルトを使う暗号形式、CHANGELOG、CI（PS 5.1、exit code で判定）が揃っている |
| 8 | 初版（約 12 モジュール）の実装と現場試行 | 運用を評価する | 1 ロット（10〜20 台）で HTML チェックリストとスクリーンショットを納品形式で出力できる |

## 結論

Android 版で最も重要な発見は、**「MDM なし」という条件が、Mac 版のような二値ではなくなる**ことです。自作 DO エージェントは、サーバーを持たない端末上のポリシー制御アプリにすぎません。それでも、端末は組織管理端末になり、解除には原則として初期化が必要になります。したがって、ツールがどこまで自動化できるかを決めるのは技術ではなく、「その DO を MDM とみなすか」という契約上の定義です。Mac 版では「MDM なし」は固定の前提条件でしたが、Android 版では、この定義の交渉次第で ◎ の数が 11 件から 18 件まで動きます。顧客との最初の打ち合わせで決めるべきは機能一覧ではなく、この定義です。

もう 1 つの示唆は、エンジニアリング上のリスクが Android の API 側ではなくホスト側にあることです。Android で何ができるかは、AOSP のソースで高い確度で確定できます。一方、fabriq に前例がないのは、1 台のホストが N 台を同時に扱う並列実行層（タイムアウト、USB の切断と再接続、端末ごとの状態）です。さらにポリシーの締め付けが一方向に進んでいる以上、適用の手段（ADB・DO・EMM・OEM ツール）に依存しない ModuleResult 契約とエビデンス層を先に固め、適用経路を後から差し替えられる構造にしておくことが、fabriq の「保守性・堅牢性・柔軟性」を Android で再現する最短の道になります。
