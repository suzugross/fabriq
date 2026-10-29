# 自作 Device Owner (DPC) エージェントアプリ方式の実現性（tier (b) ADB+自作DO / tier (c) QR プロビジョニング）— 2026-10-29 時点

ラベル凡例: **[公式]** = developer.android.com / Google ヘルプ、**[AOSP]** = AOSP ソース（GitHub の aosp-mirror 経由で取得。スナップショットに API 36 のシンボル `setAutoTimePolicy` を含むので Android 16 以降のコード。正確な日付は未検証）、**[第三者]** = Bayton（Android Enterprise の専門家ブログ。bayton.org はこのセッションから到達不可のため、GitHub 上の同サイトのソース Markdown `jasonbayton/11ty` を読んだ）、**[コミュニティ]** = フォーラム / GitHub issue / OSS の README、**[推論]** = 本調査者の推論。

## Q1. `adb shell dpm set-device-owner`：Android 12〜17 での前提条件、OEM ごとの阻害要因、DO の解除方法

### Takeaway
AOSP のソースでは、**ADB(shell/root uid) からの set-device-owner はセットアップウィザード（SUW）完了後でも成功する**。条件は次の 4 つ: (1) 既存の DO/PO が無い、(2) 既定ユーザー以外のユーザー（ワークプロファイル、Secure Folder 系、アプリのクローン、プライベートスペースを含む）が存在しない、(3) アカウントが 1 つも無い（testOnly の admin で、全アカウントが「許可」フィーチャを宣言している場合のみ例外）、(4) 対象ユーザーが実行中。ADB 以外の経路（QR/NFC 等）は SUW 完了前に限られる。つまり tier (b) の運用は「初期化 → Google アカウントを追加せずに SUW を素通り → USB デバッグを有効化 → `adb install` → `dpm set-device-owner`」で成立する。解除は testOnly ビルドなら `dpm remove-active-admin` で可能。それ以外は factory reset か、アプリ自身が呼ぶ `clearDeviceOwnerApp()`（API 26 で deprecated）になる。

### Cited Findings
- **[AOSP] コマンド書式**: `set-device-owner [ --user <USER_ID> | current *EXPERIMENTAL* ] [ --device-owner-only ] [ --provisioning-context <CTX> ] <COMPONENT>`。help の説明は "Sets the given component as active admin, and its package as device owner." `--provisioning-context` はフラグ `provisioningContextParameter` が有効なビルドでのみ表示される — [DevicePolicyManagerServiceShellCommand.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/services/devicepolicy/java/com/android/server/devicepolicy/DevicePolicyManagerServiceShellCommand.java)
- **[AOSP] 処理の流れ**: `setActiveAdmin()` → `setDeviceOwner(component, userId, setProfileOwnerOnCurrentUserIfNecessary = !--device-owner-only)` → 失敗したら追加した admin を削除して例外 → 成功したら `setUserProvisioningState(STATE_USER_SETUP_FINALIZED)` を実行し、`Success: Device owner set to package ...` を出力 — 同上
- **[AOSP] 前提条件チェック** `checkDeviceOwnerProvisioningPreConditionLocked`: 既に DO がある → `STATUS_HAS_DEVICE_OWNER`。対象ユーザーに PO がある → `STATUS_USER_HAS_PROFILE_OWNER`。ユーザーが実行中でない → `STATUS_USER_NOT_RUNNING`。**isAdb の場合**: コメントは "If shell command runs after user setup completed check device status. Otherwise, OK."。SUW 完了済みなら次を確認する: テスト用でも既定でもないユーザーが 1 人（headless system user モードでは 2 人）を超える → `STATUS_NONSYSTEM_USER_EXISTS`、アカウントの非互換がある → `STATUS_ACCOUNTS_NOT_EMPTY`。**ADB 以外の場合**: 対象は user 0 でなければならない（`STATUS_NOT_SYSTEM_USER`）。コメントは "Only provision DO before setup wizard completes"（`STATUS_USER_SETUP_COMPLETED`） — [DevicePolicyManagerService.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/services/devicepolicy/java/com/android/server/devicepolicy/DevicePolicyManagerService.java)
- **[AOSP] エラーメッセージの文字列**（ホスト側でエラーを分類するのに使える）: "Trying to set the device owner (...), but device owner (...) is already set." / "...but the user already has a profile owner" / "Cannot set the device owner if the device is already set-up." / "Not allowed to set the device owner because there are already several users on the device." / "Not allowed to set the device owner because there are already some accounts on the device." / "...because this device has already paired."（watch 向け） / "Cannot provision an unsupported DPC into DO on a headless device" — 同上
- **[AOSP] アカウントの扱い**: `hasIncompatibleAccountsOrNonAdbNoLock` の javadoc は "If the user has no accounts, then return false. Otherwise, if the owner is unknown (== null), or is not test-only, then return true. Otherwise, if there's any account that does not have ..._ALLOWED, or does have ..._DISALLOWED, return true."。つまり**本番署名の（testOnly でない）DPC は、アカウントが 1 つでもあれば失敗する**（ログは "Non test-only owner can't be installed with existing accounts."）。判定に使うのは `ACCOUNT_FEATURE_DEVICE_OR_PROFILE_OWNER_ALLOWED/DISALLOWED` — 同上
- **[AOSP] DO 設定時の副作用**: `toggleBackupServiceActive(USER_SYSTEM, false)`（コメント "Shutting down backup manager service permanently."）。ADB 経由で設定した場合は `mAdminCanGrantSensorsPermissions = true`（センサー系の権限を admin が付与できる）。DO と managed profile の共存を防ぐため、managed profile の追加が制限される — 同上
- **[AOSP] `isAdb()` の判定** = `isShellUid(caller) || isRootUid(caller)`。`isShellUid` は `UserHandle.isSameApp(uid, Process.SHELL_UID)` — 同上
- **[AOSP] 解除**: `remove-active-admin` の help は "Disables an active admin, the admin must have declared android:testOnly in the application in its manifest. This will also remove device and profile owners." — [ShellCommand.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/services/devicepolicy/java/com/android/server/devicepolicy/DevicePolicyManagerServiceShellCommand.java)
- **[公式] `clearDeviceOwnerApp`** は API 26 で deprecated（API diff の Changed Methods に記載） — [API diff 26](https://developer.android.com/sdk/api_diff/26/changes/android.app.admin.DevicePolicyManager)
- **[第三者] Google の立場**: DO は "only possible during the initial device setup, before any accounts are added"。ADB コマンドは本番展開の手段ではなく開発・テスト用ツールと位置付けられている（検索スニペット） — [Bayton FAQ: can I set device owner without factory reset](https://bayton.org/android/android-enterprise-faq/can-i-set-device-owner-without-factory-reset/)
- **[コミュニティ] OEM 固有の失敗**（OwnDroid README）: **Xiaomi (MIUI/HyperOS)** は `MANAGE_DEVICE_ADMINS` の SecurityException が出る。開発者オプションの "USB debugging (Security settings)" を有効にするか root shell から実行する。**OPPO (ColorOS)** は "Unexpected @ProvisioningPreCondition" が出る。testkey 署名ビルドを使えという回避策（署名だけが違うビルド）が案内されている。**Samsung** は "user limit reached" で解決策なしとされる（ユーザー作成機能についての記述かどうか文脈が曖昧）。「複数ユーザー」エラーにはワークプロファイル、プライベートスペース、アプリクローンを削除して対処する。アカウントのエラーにはアカウントを削除するか、アカウントを持つアプリを凍結して対処する — [OwnDroid README](https://github.com/BinTianqi/OwnDroid)
- **[コミュニティ] "Trying to set device owner but device is already provisioned."** というエラーの報告がある（B4X フォーラム）。このメッセージは現行 AOSP の DPMS には存在しないので、旧版 Android の文言と見られる（どの版かは未特定） — [B4X forum](https://b4x.com/android/forum/threads/about-kiosk-not-allowed-to-set-the-device-owner.137700/)
- **[コミュニティ] Headwind MDM**: QR enrollment は Android 6 以下、Xiaomi、Huawei EMUI 5 以下では動かない（検索スニペット） — [Headwind quick start](https://h-mdm.com/quick-start)
- **[コミュニティ] Meraki**: あるベンダーが Android 13/14 の複数機種で DO の構成に問題があることを把握していた、という漠然とした報告のみ — [Meraki community](https://community.meraki.com/t5/Mobile-Device-Management/Unable-to-install-Meraki-in-Device-Owner-Mode-on-Android-phone/m-p/225790)

### Inferences
- [推論] **tier (b) の標準手順**: 初期化 → SUW では Wi-Fi もアカウントも入れずにスキップ → 開発者オプションと USB デバッグを有効化（ここだけは端末ごとの手作業として残る）→ `adb install agent.apk` → `adb shell dpm set-device-owner pkg/.AdminReceiver`。fabriq 流の Step 2（冪等性チェック）は `adb shell dpm list-owners` で既に DO かどうかを確認する形にできる（`list-owners` は AOSP に存在するコマンド）。
- [推論] **事前チェック（preflight）モジュールに入れるべき項目**: `dumpsys account` のアカウント数（キャリアのアカウント、Samsung アカウント、プリインストールアプリが自動作成するアカウントも ACCOUNTS_NOT_EMPTY の原因になる）、`pm list users` が 1 人だけか（Samsung Secure Folder、Dual Messenger、Xiaomi Second space やアプリの二重化も別ユーザーになりうる）。上記のエラー文字列で失敗理由を分類すれば ModuleResult の Error/Skipped に写像できる。
- [推論] **testOnly ビルドを本番に使う是非**: `remove-active-admin` で factory reset なしに DO を外せるので、再キッティングや検証には便利。一方で ADB にアクセスできる者は誰でも DO を剥がせる。`DISALLOW_DEBUGGING_FEATURES` で ADB を塞げばこのリスクは閉じられるが、ホストからの操作も同時にできなくなる。キッティングの完了時に「ADB を閉じる」最終ステップを置く設計が考えられる。
- [推論] QR 方式（SUW 前）と ADB 方式（SUW 後でも可）では前提条件の厳しさが異なる。ADB 方式は「SUW を素通りした端末」さえあれば成立する。

### Gaps
- Android 17 (API 37) で DPMS の ADB 前提条件が変わったかどうかは未確認（取得した mirror の日付が不明、android.googlesource.com / cs.android.com はこのセッションから到達不可）。
- 日本の OEM（SHARP / 京セラ / FCNT）、キャリア版端末（docomo / au / SoftBank）に固有の阻害要因について、信頼できる資料は見つからなかった。
- Samsung の Knox / One UI で ADB set-device-owner が成功するかについて、2025〜2026 年の一次報告は見つからなかった。OwnDroid の "user limit reached" も文脈が曖昧。
- OPPO の "Unexpected @ProvisioningPreCondition" が具体的に何を返しているのか（どの AOSP の状態コードに当たるのか）は未解明。

## Q2. QR コードプロビジョニング：提供状況、JSON フィールド、APK のホスティングと Wi-Fi、Android 12 以降の必須対応、NFC / afw#、Google のポリシー（自作 DPC への制限）

### Takeaway
QR プロビジョニングは Android 7 以降で使える。Android 9 以降は SUW の Welcome 画面を 6 回タップすると QR リーダーが内蔵で起動する。JSON は DevicePolicyManager の `EXTRA_PROVISIONING_*` キーをそのまま使う。Android 12 以降は DPC 側に `GET_PROVISIONING_MODE` と `ADMIN_POLICY_COMPLIANCE` の 2 つの Activity が**必須**で、無ければプロビジョニングに失敗し端末は factory reset される。**最大のリスクは 2025 年から Google Play Protect が「承認済み DPC の allowlist」をプロビジョニング中に強制していること**で、未承認の自作 DPC は「App blocked to protect your device / Harmful app blocked」で止まる可能性がある（実際に止まった報告と、止まらなかった報告の両方がある）。さらに Play EMM API への新規 custom DPC 登録は締め切られている（managed Google Play は使えない）。

### Cited Findings
- **[第三者] 提供バージョン**: QR は Android 7.0 から。Android 7.0〜8.x では 6 タップで QR リーダーアプリを先にダウンロードし、Android 9.0 以降はリーダーが内蔵で即起動する。Wi-Fi 情報は任意。含めた場合は平文で保存されるので、その QR は機密扱いにすべき。Android 13 以降は Wi-Fi を payload に含めないとネットワーク選択を求められる — [Bayton: provisioning methods](https://bayton.org/android/android-enterprise-provisioning-methods/)（[source md](https://github.com/jasonbayton/11ty/blob/main/_src/android/android-enterprise-provisioning-methods.md)）
- **[コミュニティ/公式サンプル] TestDPC README の最小 JSON**: `PROVISIONING_DEVICE_ADMIN_COMPONENT_NAME`、`PROVISIONING_DEVICE_ADMIN_SIGNATURE_CHECKSUM`、`PROVISIONING_DEVICE_ADMIN_PACKAGE_DOWNLOAD_LOCATION`（いずれも `android.app.extra.` 接頭辞付き）。QR は "Device Owner N+" が必要。configuring 画面で止まる場合は appspot.com に届いていない可能性があり、APK を自前でホストして QR を作り直すよう案内している — [android-testdpc](https://github.com/googlesamples/android-testdpc)
- **[AOSP] 現行の `EXTRA_PROVISIONING_*` の一覧**（QR/NFC で使う主なもの）: `DEVICE_ADMIN_COMPONENT_NAME`、`DEVICE_ADMIN_PACKAGE_NAME`、`DEVICE_ADMIN_PACKAGE_DOWNLOAD_LOCATION`、`DEVICE_ADMIN_PACKAGE_DOWNLOAD_COOKIE_HEADER`、`DEVICE_ADMIN_PACKAGE_CHECKSUM`、`DEVICE_ADMIN_SIGNATURE_CHECKSUM`、`DEVICE_ADMIN_MINIMUM_VERSION_CODE`、`ADMIN_EXTRAS_BUNDLE`、`LOCALE`、`TIME_ZONE`、`LOCAL_TIME`、`SKIP_ENCRYPTION`、`LEAVE_ALL_SYSTEM_APPS_ENABLED`、`USE_MOBILE_DATA`、`ALLOW_OFFLINE`、`SENSORS_PERMISSION_GRANT_OPT_OUT`、`WIFI_SSID / PASSWORD / SECURITY_TYPE / HIDDEN / PROXY_HOST / PROXY_PORT / PROXY_BYPASS / PAC_URL / EAP_METHOD / PHASE2_AUTH / CA_CERTIFICATE / USER_CERTIFICATE / IDENTITY / ANONYMOUS_IDENTITY / DOMAIN`、`ORGANIZATION_NAME`、`SUPPORT_URL`、`DISCLAIMERS`、`SKIP_EDUCATION_SCREENS`、`KEEP_SCREEN_ON` — [DevicePolicyManager.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/core/java/android/app/admin/DevicePolicyManager.java)
- **[AOSP] SIGNATURE_CHECKSUM と PACKAGE_CHECKSUM の違い**: `SIGNATURE_CHECKSUM` は "URL-safe base64 encoded SHA-256 checksum of any signature of the android package archive"。`PACKAGE_CHECKSUM` は "URL-safe base64 encoded SHA-256 hash of the file"。"Either this extra or ... must be present"。一致しなければ "an error will be shown to the user and the user will be asked to factory reset the device."。PACKAGE_CHECKSUM について "From Q, only SHA-256 hash is supported." — 同上
- **[AOSP] DOWNLOAD_LOCATION** は "A String extra holding a url ... When not provided it is assumed that the device admin package is already installed."（javadoc に HTTPS 必須という記述は無い） — 同上
- **[AOSP] WIFI_SECURITY_TYPE** の値は `NONE` / `WPA` / `WEP` / `EAP`。EAP 関連の extras（CA_CERTIFICATE / USER_CERTIFICATE / IDENTITY / DOMAIN 等）は API 29 で追加された（API diff 29 の Added Fields） — 同上 / [API diff 29](https://developer.android.com/sdk/api_diff/29/changes/android.app.admin.DevicePolicyManager)
- **[AOSP] USE_MOBILE_DATA**（API 33 で追加）: "For provisioning started via a QR code or an NFC tag, mobile data is always used for downloading the admin app"（WIFI_SSID が無く true の場合）。`ALLOW_OFFLINE` も API 33 で追加（javadoc は "whether offline provisioning should be used" とだけ書き、詳細は b/365955253 を参照） — 同上 / [API diff 33](https://developer.android.com/sdk/api_diff/33/changes/android.app.admin.DevicePolicyManager)
- **[AOSP] LOCALE** は "Format: xx_yy ... Use only for device owner provisioning ... can also be used for QR code provisioning."。プロビジョニング時の locale 設定はこの方法が唯一の公開手段（DPM API に setLocale は無い。Q4 参照） — [DevicePolicyManager.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/core/java/android/app/admin/DevicePolicyManager.java)
- **[AOSP] LEAVE_ALL_SYSTEM_APPS_ENABLED**: "skip the disabling of system apps during provisioning when set to true"。指定しないと DO プロビジョニングで不要なシステムアプリが無効化される — 同上
- **[公式] Android 12 の変更**: "`android.app.action.PROVISION_MANAGED_DEVICE` no longer works on Android 12. DPCs must implement activities with intent filters for the `ACTION_GET_PROVISIONING_MODE` and `ACTION_ADMIN_POLICY_COMPLIANCE` intent actions. Using `ACTION_PROVISION_MANAGED_DEVICE` to start provisioning causes the provisioning to fail." — [Android 12 enterprise](https://developer.android.com/work/versions/android-12)
- **[公式] API 履歴**: `ACTION_GET_PROVISIONING_MODE` と `ACTION_ADMIN_POLICY_COMPLIANCE` は API 29 で追加、API 31 で必須化。`ACTION_PROVISION_MANAGED_DEVICE` は API 31 で deprecated。`EXTRA_PROVISIONING_LOGO_URI` は API 33 で deprecated、`KEEP_SCREEN_ON` は API 34 で deprecated — [API diff 29](https://developer.android.com/sdk/api_diff/29/changes/android.app.admin.DevicePolicyManager) / [31](https://developer.android.com/sdk/api_diff/31/changes/android.app.admin.DevicePolicyManager) / [33](https://developer.android.com/sdk/api_diff/33/changes/android.app.admin.DevicePolicyManager) / [34](https://developer.android.com/sdk/api_diff/34/changes/android.app.admin.DevicePolicyManager)
- **[AOSP] GET_PROVISIONING_MODE の契約**: 渡される extras は `ADMIN_EXTRAS_BUNDLE`、`IMEI`、`SERIAL_NUMBER`、`ALLOWED_PROVISIONING_MODES`、`SENSORS_PERMISSION_GRANT_OPT_OUT`。返すのは `EXTRA_PROVISIONING_MODE` に `PROVISIONING_MODE_FULLY_MANAGED_DEVICE` か `MANAGED_PROFILE`。結果 intent には `LOCALE` / `LOCAL_TIME` / `TIME_ZONE` / `SKIP_ENCRYPTION` / `SENSORS_PERMISSION_GRANT_OPT_OUT` / `DISCLAIMERS` と追加の `ADMIN_EXTRAS_BUNDLE` を含めてよく、後者は ADMIN_POLICY_COMPLIANCE にマージされて渡る — [DevicePolicyManager.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/core/java/android/app/admin/DevicePolicyManager.java)
- **[コミュニティ] 2 つの Activity を欠いた場合の実例**: 自作 DPC "muster" を Pixel 6a で QR プロビジョニング（2026-08-19）したところ、APK のダウンロードまでは成功し、その後 "Something went wrong ... contact your IT admin" と Reset ボタンが表示されて factory reset された。原因は manifest に両 intent の Activity が無かったこと。修正後（PR #133、2026-10-15 マージ）は実機で DO に到達した（2026-10-19）。この件では Play Protect によるブロックは観測されなかった — [quadseven/muster #89](https://github.com/quadseven/muster/issues/89)
- **[公式] build-dpc ページ**: "Android Enterprise is no longer accepting new registrations for custom device policy controllers (DPCs) to the Google Play EMM API. To get your DPC approved by Android Enterprise for device management, visit the DPC approval support page." — [Build a DPC](https://developer.android.com/work/dpc/build-dpc)
- **[公式（スニペット）] DPC allowlist**: "Only DPCs verified and approved by Android Enterprise are permitted to install apps during enterprise device enrollment provisioning. The installation of any non-approved DPCs or app installation attempts not on the allowlist will be blocked during enrollment installation."。表示されるのは "App blocked to protect your device"（ページ本体は到達不可。検索スニペットのみ） — [Google Help 16694822: Approved Android Enterprise device policy controllers allowlist](https://support.google.com/work/android/answer/16694822)
- **[第三者] allowlist の導入経緯**: 2025 年中に事前告知なく有効化された。Play Protect が執行役で、未登録の DPC は悪意が無くても "Harmful app blocked" 扱いになる。SMS、通知、accessibility などのセンシティブな権限が問題視される。appeal は help の記事から行う。appeal の SLA は無く「数日〜数週間」かかる。審査中でも "continue" が出る場合があるが保証はない — [Bayton: The DPC allowlist (2025-12-24)](https://bayton.org/blog/2025/12/the-dpc-allowlist/) / [Bayton FAQ: Play Protect blocked my DPC](https://bayton.org/android/android-enterprise-faq/play-protect-blocked-my-dpc-why/)
- **[コミュニティ] ブロック報告**: Test DPC 自体が QR プロビジョニング中にブロックされたという投稿（2026-09-19 付とされるが日付は未検証）、"Play Protect blocks DPC installation during QR provisioning on Android 14" というスレッドなど（検索スニペット） — [AE Community #2909](https://www.androidenterprise.community/android-enterprise-general-discussions-3/play-protect-blocking-custom-dpc-app-during-6-tap-qr-provisioning-appeal-submitted-2909) / [AE Community #14202](https://ypaoq37543.lithium.com/discussions/conversations/issue-play-protect-blocks-dpc-installation-during-qr-provisioning-on-android-14-/14202)
- **[第三者] Play EMM API**: custom DPC 向けの一部メソッドは 2021 年 9 月に deprecated となり、turn-off は 2025-09-30。主要ベンダーの既存 custom DPC は動き続ける。Play EMM API が無いと managed Play による配信、managed account、承認、トラック管理は使えない — [Bayton: provisioning methods](https://bayton.org/android/android-enterprise-provisioning-methods/)。同じ日付は [Omnissa KB 93986](https://kb.omnissa.com/s/article/93986) にもある（スニペット）。Google の公式ページで turn-off 日付は確認できていない
- **[第三者] 「custom DPC は延命状態」という評価**: "the custom DPC itself for full-fat certified Android management is going the way of Device Admin"（2026-04-17、意見） — [Bayton: DPC migration](https://bayton.org/blog/2026/04/dpc-migration/)
- **[第三者] NFC**: Android 6.0 から。Android Beam は Android 10 で deprecated、Android 14 で完全削除。Android 10 以降はタグ型 NFC のみ。正式な deprecated 扱いではないが「niche」 — [Bayton: provisioning methods](https://bayton.org/android/android-enterprise-provisioning-methods/)
- **[第三者] afw#**: SUW のアカウント入力欄に `afw#<vendor>` を入れる方式（例: `afw#setup` = AMAPI）。Google にベンダー識別子が登録されている DPC 用で、自作 DPC では使えない（登録が締め切られているため） — 同上
- **[第三者] Android 16**: `EXTRA_PROVISIONING_SKIP_EDUCATION_SCREENS` は deprecated で無視される — [Bayton: Android 16 enterprise features](https://bayton.org/android/android-16-enterprise-features/)

### Inferences
- [推論] **tier (c) の成否は Play Protect の allowlist で決まる**。自作 DPC で Play Protect に止められなかった例（muster、Pixel 6a、2026-10）と、止められた例（Test DPC 等）が両方あり、判定は「センシティブ権限の有無」などのヒューリスティックで動いている可能性が高い。権限を最小にした DPC（accessibility、SMS、オーバーレイ等を宣言しない）にすれば通る確率は上がるが、**保証はなく、いつ変わってもおかしくない**。本番前に対象の全機種・全 OS 版で実機検証し、フォールバックとして tier (b) を残すのが妥当。
- [推論] allowlist の文言は「enrollment provisioning 中のインストール」なので、**ADB の `dpm set-device-owner`（ManagedProvisioning を経由しない）には適用されない可能性が高い**（未検証）。ただし `adb install` にも Play Protect の「USB 経由でインストールしたアプリのスキャン」が働く可能性はある。また GMS 非搭載端末（AOSP ベースの業務用端末など）には Play Protect 自体が無い。
- [推論] **SIGNATURE_CHECKSUM を使うべき**。署名証明書が変わらない限り APK を更新しても QR を作り直さなくてよい（PACKAGE_CHECKSUM はビルドごとに変わる）。
- [推論] APK のホスティングについて、javadoc に HTTPS 必須の記述は無いが、Android 9 以降のクリアテキスト制限が ManagedProvisioning のダウンロードにどう働くかは不明。HTTPS で、SUW 中の端末から到達できる URL（社内 LAN 上の HTTPS サーバー＋自己署名ではない証明書、または社外 CDN）を前提にするのが安全。COOKIE_HEADER で簡易認証を付けられる。
- [推論] QR のサイズは QR 規格の上限（バージョン 40 で約 2.9KB のバイナリ）に縛られる。Wi-Fi や EAP、admin extras を詰め込むと密度が上がり、SUW のカメラで読みにくくなる。admin extras には「サイト ID」と「設定取得先 URL」だけを入れ、本体の設定はダウンロードさせる設計が現実的。

### Gaps
- Google ヘルプ 16694822 の全文（適用されるプロビジョニング経路、GMS 非搭載端末や ADB 経路が含まれるか、appeal の基準）は本文を取得できなかった（support.google.com はこのセッションから到達不可）。
- DOWNLOAD_LOCATION に平文 HTTP が受け付けられるか、自己署名証明書が使えるかについて、一次情報は見つからなかった。
- `ALLOW_OFFLINE`（API 33）の正確な挙動（DPC をプリインストールしておく前提か）は javadoc からは判断できない。
- 2025-09-30 の turn-off が Google 公式に確認できるかは未検証（第三者ソースのみ）。

## Q3. admin extras bundle で端末ごと・サイトごとの設定を渡せるか／量販キッティングで端末別 QR をどう作るか

### Takeaway
渡せる。`EXTRA_PROVISIONING_ADMIN_EXTRAS_BUNDLE` は `PersistableBundle` として DPC に渡り、`GET_PROVISIONING_MODE` と `ADMIN_POLICY_COMPLIANCE` の両方の Activity で受け取れる。さらに GET_PROVISIONING_MODE には端末の **IMEI とシリアル番号** も渡される。このため「サイト単位の QR を 1 枚」用意し、端末側でシリアルを使って hostlist の行を引く設計が可能で、端末ごとに QR を作る必要はない。ただし QR の容量制約から、bundle には ID と URL 程度を入れるのが現実的。

### Cited Findings
- **[AOSP]** `EXTRA_PROVISIONING_ADMIN_EXTRAS_BUNDLE` は "A Parcelable extra of type PersistableBundle that is passed directly to the Device Policy Controller after provisioning."。NFC では "stringified java.util.Properties" の string プロパティが PersistableBundle に変換される — [DevicePolicyManager.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/core/java/android/app/admin/DevicePolicyManager.java)
- **[AOSP]** `ACTION_GET_PROVISIONING_MODE` の intent には `ADMIN_EXTRAS_BUNDLE`、`EXTRA_PROVISIONING_IMEI`、`EXTRA_PROVISIONING_SERIAL_NUMBER` が含まれうる。DPC が結果として返した ADMIN_EXTRAS_BUNDLE は、元の値とともに `ACTION_ADMIN_POLICY_COMPLIANCE` に渡される。`ACTION_ADMIN_POLICY_COMPLIANCE` の intent にも ADMIN_EXTRAS_BUNDLE が付く — 同上（IMEI / SERIAL_NUMBER は API 29 で追加: [API diff 29](https://developer.android.com/sdk/api_diff/29/changes/android.app.admin.DevicePolicyManager)）
- **[コミュニティ]** 実例: muster は QR の admin extras に `muster.server_url` を入れていた。修正版では `PolicyComplianceActivity` が bundle からサーバー URL を保存している（修正前は読まれず、「DO になったが登録先が無い」状態になった） — [quadseven/muster #89](https://github.com/quadseven/muster/issues/89)
- **[コミュニティ]** Headwind MDM は device ID を QR に埋め込んで登録を簡単にできる（既定は `h0001`）。端末の登録には IMEI を使う（検索スニペット） — [Headwind quick start](https://h-mdm.com/quick-start)
- **[コミュニティ]** MDM 製品の QR では、extras に MDM 固有のキー（"enrollmentId" 等）を入れるのが一般的（検索スニペット） — [AE Community: QR code provisioning](https://androidenterprise.community/t5/general-discussions/qr-code-provisioning/m-p/8996)
- **[第三者]** Knox Mobile Enrollment は CSV や IMEI/シリアルで端末を登録でき、KME が生成する QR でオフラインのプロビジョニングにも対応する（Samsung 限定） — [Bayton: provisioning methods](https://bayton.org/android/android-enterprise-provisioning-methods/)

### Inferences
- [推論] **fabriq 流の写像**: QR（サイト単位）に `{site_id, config_url, (任意) bootstrap_token}` を入れ、DPC は ADMIN_POLICY_COMPLIANCE でシリアルを使って `config_url` から hostlist の行と profile CSV を取得する。これで fabriq の `hostlist.csv` と profile モデルを再現できる。端末ごとの QR が要るのは、オフラインで個体ごとの値（ホスト名等）を直接渡したい場合に限られ、その場合は PowerShell で hostlist.csv から QR を一括生成する（QR 生成ライブラリを使う）。
- [推論] 試した範囲では、QR の JSON で bundle の中に入れ子のオブジェクトや非文字列型が使えるかを AOSP で確認できなかった（ManagedProvisioning のパーサーのソースを取得できなかった）。文字列値だけにしておくのが安全。

### Gaps
- QR の JSON パーサー（ManagedProvisioning）が admin extras でどの型を受け付けるか（int / boolean / 入れ子）は未検証。
- 量販キッティング業者が端末別 QR を実際にどう運用しているかについて、一次資料は見つからなかった。

## Q4. Device Owner で使える DevicePolicyManager の機能（キッティング関連・最小 API・Android 14〜17 の変化）と、素の ADB shell との比較

### Takeaway
DO にすると、**ADB shell では恒久化・強制できない「ポリシー」**が使えるようになる。具体的には CA 証明書やクライアント証明書の無言インストールと Wi-Fi 認証への付与、ユーザー制限（factory reset 禁止、USB ファイル転送禁止、提供元不明アプリ禁止、デバッグ禁止、Wi-Fi 設定禁止等）、企業 FRP、OTA ポリシー、APN override、Always-on VPN、ユーザー確認なしのアプリのインストール／アンインストール／無効化／非表示、ランタイム権限の固定、キオスク（lock task）、既定ダイヤラー／SMS アプリの設定、wipe、reboot、IMEI／シリアルの取得などである。一方 settings への書き込みは allowlist 制で、ADB の `settings put` より範囲が狭い。locale を設定する公開 API は無い。

### Cited Findings
- **[公式] API diff から確認した導入版**（DevicePolicyManager）:
  - **API 24**: `installKeyPair`（証明書チェーン付きのオーバーロード）、`reboot`、`setAlwaysOnVpnPackage`、`setDeviceOwnerLockScreenInfo`、`setOrganizationName`、`setPackagesSuspended`、`setSecurityLoggingEnabled` / `retrieveSecurityLogs`、`setShort/LongSupportMessage`、`requestBugreport`、`getWifiMacAddress`、`createAndManageUser` — [API diff 24](https://developer.android.com/sdk/api_diff/24/changes/android.app.admin.DevicePolicyManager)
  - **API 26**: `setNetworkLoggingEnabled` / `retrieveNetworkLogs`、`setDelegatedScopes`、`setResetPasswordToken` / `resetPasswordWithToken`、`setBackupServiceEnabled`、`setAffiliationIds`。deprecated になったのは `clearDeviceOwnerApp`、`clearProfileOwner`、`setCertInstallerPackage` — [API diff 26](https://developer.android.com/sdk/api_diff/26/changes/android.app.admin.DevicePolicyManager)
  - **API 28**: `setSystemSetting`、`setTime`、`setTimeZone`、`addOverrideApn` / `updateOverrideApn` / `removeOverrideApn` / `setOverrideApnsEnabled`、`installExistingPackage`、`setLockTaskFeatures`、`generateKeyPair`、`transferOwnership`、`setKeepUninstalledPackages`、`setMeteredDataDisabledPackages`、`clearApplicationUserData`、`wipeData`（新オーバーロード）、`installKeyPair`（flags 付きオーバーロード）。フィールドとして `ID_TYPE_IMEI/SERIAL`、`WIPE_EUICC` — [API diff 28](https://developer.android.com/sdk/api_diff/28/changes/android.app.admin.DevicePolicyManager)
  - **API 29**: `installSystemUpdate`、`setDefaultSmsApplication`、`setGlobalPrivateDnsModeSpecifiedHost/Opportunistic`、`setAlwaysOnVpnPackage`（lockdown allowlist 付き）、`getPasswordComplexity`。フィールドとして `WIPE_SILENTLY` — [API diff 29](https://developer.android.com/sdk/api_diff/29/changes/android.app.admin.DevicePolicyManager)
  - **API 30**: `setAutoTimeEnabled` / `setAutoTimeZoneEnabled`、`setFactoryResetProtectionPolicy`、`setLocationEnabled`、`setConfiguredNetworksLockdownState`、`setUserControlDisabledPackages`、`grantKeyPairToApp`、`setCommonCriteriaModeEnabled`。deprecated: `setAutoTimeRequired`、`resetPassword`、`setStorageEncryption` — [API diff 30](https://developer.android.com/sdk/api_diff/30/changes/android.app.admin.DevicePolicyManager)
  - **API 31**: `setRequiredPasswordComplexity`、`setUsbDataSignalingEnabled`、`grantKeyPairToWifiAuth`、`setPreferentialNetworkServiceEnabled`、`setNearbyAppStreamingPolicy`、`setOrganizationId` / `getEnrollmentSpecificId`。deprecated: `setPasswordQuality` と `setPasswordMinimum*` の全系統、`setOrganizationColor` — [API diff 31](https://developer.android.com/sdk/api_diff/31/changes/android.app.admin.DevicePolicyManager)
  - **API 33**: `setWifiSsidPolicy`、`setMinimumRequiredWifiSecurityLevel`、`setPreferentialNetworkServiceConfigs` — [API diff 33](https://developer.android.com/sdk/api_diff/33/changes/android.app.admin.DevicePolicyManager)
  - **API 34**: `setDefaultDialerApplication`、`wipeDevice`、`setMtePolicy`、`setCredentialManagerPolicy`、`addUserRestrictionGlobally`、`isStatusBarDisabled`、`setManagedSubscriptionsPolicy`。deprecated: `setCrossProfileCallerIdDisabled` / `setCrossProfileContactsSearchDisabled` / `setCrossProfileCalendarPackages` — [API diff 34](https://developer.android.com/sdk/api_diff/34/changes/android.app.admin.DevicePolicyManager)
  - **API 35**: `setContentProtectionPolicy`、`getSubscriptionIds` — [API diff 35](https://developer.android.com/sdk/api_diff/35/changes/android.app.admin.DevicePolicyManager)
  - **API 36**: `setAutoTimePolicy`、`setAutoTimeZonePolicy`、`setAppFunctionsPolicy` — [API diff 36](https://developer.android.com/sdk/api_diff/36/changes/android.app.admin.DevicePolicyManager)
- **[AOSP] settings の allowlist**（allowlist に無いキーは `SecurityException: "Permission denial: device owners cannot update ..."`。demo モードは例外）:
  - **Global**: `ADB_ENABLED`、`ADB_WIFI_ENABLED`、`AUTO_TIME`、`AUTO_TIME_ZONE`、`DATA_ROAMING`、`USB_MASS_STORAGE_ENABLED`、`WIFI_SLEEP_POLICY`、`STAY_ON_WHILE_PLUGGED_IN`、`WIFI_DEVICE_OWNER_CONFIGS_LOCKDOWN`、`PRIVATE_DNS_MODE`、`PRIVATE_DNS_SPECIFIER`。黙って無視されるのは `BLUETOOTH_ON`、`DEVELOPMENT_SETTINGS_ENABLED`、`MODE_RINGER`、`NETWORK_PREFERENCE`、`WIFI_ON`
  - **Secure**: `DEFAULT_INPUT_METHOD`、`SKIP_FIRST_USE_HINTS`、`INSTALL_NON_MARKET_APPS`、DO のみ `LOCATION_MODE`
  - **System**: `SCREEN_BRIGHTNESS`、`SCREEN_BRIGHTNESS_MODE`、`SCREEN_OFF_TIMEOUT`
  - 出典: [DevicePolicyManagerService.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/services/devicepolicy/java/com/android/server/devicepolicy/DevicePolicyManagerService.java)
- **[AOSP] `setTime` / `setTimeZone`** は DO または組織所有端末の PO だけが呼べる。`AUTO_TIME` / `AUTO_TIME_ZONE` が 1 のときは何もせず `false` を返す（"Don't allow set time when auto time is on."）。そのため先に `setAutoTimeEnabled(false)`（API 30）か `setAutoTimePolicy`（API 36）を呼ぶ必要がある — 同上
- **[AOSP] `setDefaultDialerApplication`** は DO か組織所有端末の PO だけ。telephony 機能が無い端末では何もしない。失敗すると `IllegalArgumentException("<pkg> cannot be set as the dialer")` — 同上
- **[AOSP] 無言インストール**: PackageInstallerSession のコメントは "Device owners and affiliated profile owners are allowed to silently install packages, so the permission check is waived if the installer is the device owner."（`noUserActionNecessary`） — [PackageInstallerSession.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/services/core/java/com/android/server/pm/PackageInstallerSession.java)
- **[AOSP] シリアル番号**: `Build.getSerial()` の javadoc は、READ_PRIVILEGED_PHONE_STATE、キャリア権限に加えて "If the calling app is the device owner of a fully-managed device, a profile owner of an organization-owned device, or their delegates" を許可条件として挙げる — [Build.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/core/java/android/os/Build.java)
- **[公式] Wi-Fi**: `WifiManager.addNetwork()` は API 29 で deprecated。"For applications targeting Q or above, this API will always fail and return -1. **Deprecation Exemptions: Device Owner (DO), Profile Owner (PO) and system apps.**" HTTP プロキシ付きの構成は "the calling app must be System, or be provisioned as the Profile or Device Owner"。`setWifiEnabled()` も DO/PO は例外扱いで、Android 13 以降は DO が `DISALLOW_CHANGE_WIFI_STATE` で DO/PO 以外による切替を禁止できる — [WifiManager reference](https://developer.android.com/reference/android/net/wifi/WifiManager)
- **[公式] Android 12**: `grantKeyPairToWifiAuth()` と `WifiEnterpriseConfig.getClientKeyPairAlias()` で、KeyChain の鍵を 802.1X に使える。`getCallerConfiguredNetworks()` は location 権限なしで使える。fully managed 端末の DPC は admin が提供したネットワークだけに制限できる。センサー権限は、管理する場合 SUW で明示メッセージが表示され、opt-out も可能 — [Android 12 enterprise](https://developer.android.com/work/versions/android-12)
- **[第三者] Android 16 (API 36、2025-06-10)**: `setAutoTimePolicy` / `setAutoTimeZonePolicy` / `setAppFunctionsPolicy` が追加された。DO と組織所有端末の PO は `NfcAdapter.enable()/disable()` が使える。ユーザー制限 `DISALLOW_THREAD_NETWORK` と `DISALLOW_CHANGE_NEAR_FIELD_COMMUNICATION_RADIO` が追加された。Advanced Protection は "cannot currently be enforced centrally ... through AMAPI or a custom DPC" — [Bayton: Android 16 enterprise features](https://bayton.org/android/android-16-enterprise-features/)
- **[第三者] Android 17 (API 37、Pixel 向けに 2026-06-16 リリース、他 OEM は 2026 年 Q3 後半から)**: `setNearbyAppStreamingPolicy` がエージェント型自動化（AI）を制御する窓口を兼ねる。`ACCESS_LOCAL_NETWORK` は `setPermissionGrantState` で事前付与できる。`setUsbDataSignalingEnabled(false)` は HID と USB4 の PCIe トンネリングも遮断する。ワーク／個人プロファイル間の localhost 通信は全アプリで遮断される。Advanced Protection を企業ポリシーで強制する機能は「2026 年後半」予定 — [Bayton: Android 17 enterprise features](https://bayton.org/android/android-17-enterprise-features/)
- **[第三者（スニペット）] Android 17 QPR2**: DPC がポリシーサイズの予算消費を確認できるようになる、COPE で `isDeviceManaged()` が true を返さなくなる（beta 情報） — [Bayton beta tracker](https://bayton.org/android/android-enterprise-beta-tracker/)

### Inferences
- [推論/既知だが今回未再検証] API 21〜23 から存在する DO の主要 API: `setGlobalSetting` / `setSecureSetting`、`installCaCert`、`setApplicationHidden`、`enableSystemApp`、`setUninstallBlocked`、`setLockTaskPackages`、`addUserRestriction`、`setScreenCaptureDisabled`、`setCameraDisabled`(API 14)、`setPermissionGrantState` / `setPermissionPolicy`(23)、`setStatusBarDisabled` / `setKeyguardDisabled` / `setSystemUpdatePolicy`(23)、`setMaximumTimeToLock`(8)。API 22/23 の diff ページは解析に失敗したので、導入版は reference ページで確認すること。
- [推論] **素の ADB と比べて DO で新たにできる主なこと**（キッティング観点）:
  - CA 証明書・クライアント証明書の無言インストールと、Wi-Fi 認証やアプリへの付与（ADB では UI 操作が必要）
  - EAP-TLS の Wi-Fi を構成して lockdown
  - ユーザー制限（`DISALLOW_FACTORY_RESET`、`DISALLOW_USB_FILE_TRANSFER`、`DISALLOW_INSTALL_UNKNOWN_SOURCES`、`DISALLOW_DEBUGGING_FEATURES`、`DISALLOW_CONFIG_WIFI`、`DISALLOW_MODIFY_ACCOUNTS` 等）
  - 企業 FRP（API 30）
  - OTA の延期・凍結（`setSystemUpdatePolicy`）と、ローカル OTA ファイルの適用（API 29）
  - APN override（API 28）
  - Always-on VPN の lockdown
  - キオスク（lock task）
  - アンインストールのブロック
  - ステータスバーとキーガードの無効化
  - 既定ダイヤラー／SMS（ADB でも `cmd role` 経由で可能な場合はある）
  - 端末ロック画面への所有者表示
  - IMEI／シリアルの取得（ADB でも `service call` 等で取れる場合があるが非公式）
  - ネットワークログ／セキュリティログ
  - 設定がユーザーに変更されない「強制」状態（ADB の `settings put` は値を書くだけで、後からユーザーが変更できる）
- [推論] 逆に **ADB の方が強い領域**もある: settings の任意キー（allowlist が無い）、`pm grant` / `pm disable-user`、`cmd` 系の各種サービス。そのため tier (b) は「ADB で設定値を書き、DO ではポリシー・証明書・制限・強制を担う」というハイブリッドが最も合理的。
- [推論] `setPasswordQuality` 系が API 31 で deprecated なので、新規実装では `setRequiredPasswordComplexity` を基本にする（数値で細かく指定するのが必須なら、deprecated API が DO で引き続き動くかを実機で確認する）。
- [推論] Wi-Fi の「Post-Apply Verification」は `getCallerConfiguredNetworks()`（API 31）で、DO 自身が追加したネットワークを location 権限なしに読み返せるので、fabriq の Step 5.5 に相当するものを実装しやすい。

### Gaps
- locale を DO から後で変更する公開 API は見つからなかった（プロビジョニング時の `EXTRA_PROVISIONING_LOCALE` だけ）。ADB 側でどう代替するかは本調査の範囲外。
- `TelephonyManager.getImei()` を DO が呼べる条件は、AOSP の該当 javadoc 本文を確認できなかった（Build.getSerial の条件のみ確認）。
- Android 17 で deprecated／削除された DPM API の公式な一覧（API diff 37）は取得していない。
- `setDefaultDialerApplication` / `setDefaultSmsApplication` の OEM 実装（Samsung 等）での成功率は不明。

## Q5. ホスト PC から DO エージェントへ命令し、結果を受け取る方法（セキュリティ含む）

### Takeaway
最も堅いのは **exported な ContentProvider の `call()`** を `adb shell content call` で叩く方式である。provider 側で `Binder.getCallingUid()` が shell（2000）か root（0）かを検査でき（DPMS の `isAdb()` と同じ判定）、戻り値の Bundle は stdout に `Result: Bundle[...]` として出るので、ホスト PowerShell でパースして ModuleResult に写像できる。`am broadcast` は手軽だが、受信側で送信元 UID を確実には取れない（`getSentFromUid()` は送信側が identity の共有を opt-in しないと `INVALID_UID`）。大きな設定は `adb push` で渡すか、`adb reverse` 経由の localhost HTTP で取得させる。

### Cited Findings
- **[AOSP] `content call` の書式**: `usage: adb shell content call --uri <URI> --method <METHOD> [--arg <ARG>] [--extra <BINDING> ...]`。実装は `provider.call(new AttributionSource(Binder.getCallingUid(), ...), authority, method, arg, extras)` で、結果を `System.out.println("Result: " + result)` で出力する — [Content.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/cmds/content/src/com/android/commands/content/Content.java)
- **[AOSP] `BroadcastReceiver.getSentFromUid()`**: "Returns the uid of the app that initially sent this broadcast ... or Process#INVALID_UID if the current receiver cannot access the identity of the broadcasting app" — [BroadcastReceiver.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/core/java/android/content/BroadcastReceiver.java)
- **[AOSP] DPMS の ADB 判定** `isAdb = isShellUid || isRootUid`（`Process.SHELL_UID`） — [DevicePolicyManagerService.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/services/devicepolicy/java/com/android/server/devicepolicy/DevicePolicyManagerService.java)
- **[第三者] Android 17**: API 37 をターゲットにするアプリは LAN へのアクセス（mDNS/SSDP、LAN 上の TCP/UDP、NSD）に `ACCESS_LOCAL_NETWORK` が必要。DO は `setPermissionGrantState` で事前付与できる。ワーク／個人プロファイル間の localhost 通信は遮断される（同一プロファイル内は影響なし） — [Bayton: Android 17 enterprise features](https://bayton.org/android/android-17-enterprise-features/)
- **[コミュニティ] 先行例**: Dhizuku は DO 権限を他アプリに共有する API を提供する（GPL-3.0、Android 6.0〜17 対応とされる）。DO 権限を IPC 越しに使わせる設計の参考になる — [Dhizuku](https://github.com/iamr0s/Dhizuku)

### Inferences
- [推論] **推奨する IPC 設計**:
  - 本線は `content://<pkg>.agent/` の `call(method, arg, extras)` とする。`method` = モジュール名（例: `wifi.apply` / `cert.install` / `restriction.set` / `verify.*`）、`arg` = JSON 文字列またはファイルパス。
  - provider の冒頭で `Binder.getCallingUid()` が {2000, 0} のどちらでもなければ `SecurityException` を投げる。
  - 戻りの Bundle に `status = Success|Skipped|Error|Partial`、`message`、`verified = true|false` を入れて、fabriq の `New-ModuleResult` 契約に写像する。
  - 冪等性チェック（Step 2）は `verify.*` メソッドで現在の状態を返させ、ホスト側で判定する。
- [推論] **大きな入力**（証明書、CSV、JSON）: `adb push` で端末の一時領域（例 `/data/local/tmp`）に置き、パスを `--arg` で渡す。エージェントがそのパスを読めるかは SELinux の文脈に依存し、未検証。代替案は、`adb reverse tcp:<port> tcp:<port>` でホストの HTTP を端末の localhost に見せ、エージェントに取得させる方式。Android 17 で API 37 をターゲットにした場合、localhost が local network permission の対象になるかは未確認なので、DO が自分自身に `ACCESS_LOCAL_NETWORK` を事前付与しておけば安全側になる。
- [推論] **長時間の処理**（APK の一括インストール、OTA 等）: `call()` はバインダーのタイムアウトに制約されるため、`call()` は受付だけしてジョブ ID を返し、ホストが `verify/status` をポーリングする非同期パターンにする。これは fabriq の async 中断点や `__RESTART__` 跨ぎの設計（再起動後に状態を再照会する）と相性がよい。
- [推論] **`am broadcast`**: 受信 receiver を `android:permission` で保護し、shell が保有する permission を指定する手はある。ただし確実な送信元認証にはならず、他アプリから叩かれるリスクが残るので、補助用途に限る。logcat のタグに結果を出すのはデバッグ用にとどめ、契約には使わない。
- [推論] **Wi-Fi ADB** (`adb tcpip` / ペアリング): 量産時に USB ハブを使わない選択肢になる。ただし `ADB_ENABLED` / `ADB_WIFI_ENABLED` は DO の global allowlist にあるので、キッティング完了時に DO が ADB を閉じられる（Q4 の allowlist 参照）。

### Gaps
- `/data/local/tmp` や `/sdcard/Android/data/<pkg>/` に adb で置いたファイルを DO アプリが読めるか（scoped storage / SELinux）は、Android 14〜17 で未検証。
- `content call` の `--extra` の型指定構文（`key:s:value` 等）の詳細は usage 本文で確認していない。
- `adb reverse` 経由の localhost 通信が Android 17 の local network permission の対象かどうかは不明。

## Q6. 自作 DPC の構築・保守コスト（SDK 要件、署名鍵、更新、再利用できる OSS）

### Takeaway
最小の DPC（DeviceAdminReceiver、`device_admin.xml`、Android 12 以降必須の 2 Activity、IPC 用 provider）は小規模だが、**保守コストの本体は外部要因の追従**にある。具体的には Android の年次の挙動変更（targetSdk の下限、権限、ローカルネットワーク）、Play Protect の allowlist、OEM 差分、署名鍵の永続管理である。再利用候補は TestDPC（Apache-2.0、Google 公式サンプル、API の実装例として最良）、Headwind MDM agent（Apache-2.0、サーバー前提）、OwnDroid と Dhizuku（GPL-3.0。社内ツールに取り込むならライセンスの扱いに注意）。

### Cited Findings
- **[コミュニティ/公式サンプル] TestDPC**: Apache-2.0、"designed to help EMMs, ISVs, and OEMs"。Android 5.0 以降をサポートし、v9.0.5 以降は Device Management Role Holder としても設定できる。Play ストアでも配布されている。アーカイブはされていない（commit 1,374、open issue 43） — [android-testdpc](https://github.com/googlesamples/android-testdpc)
- **[コミュニティ] Headwind MDM agent**: Apache-2.0、F-Droid で配布、star 366。デバッグ時は `dpm set-device-owner com.hmdm.launcher/.AdminReceiver`。目的は「サイレントインストールとサーバーへの情報送信」で、サーバーが前提。サーバー側は [hmdm-server](https://github.com/h-mdm/hmdm-server) — [hmdm-android](https://github.com/h-mdm/hmdm-android)
- **[コミュニティ] OwnDroid**: GPL-3.0、star 約 1.4k。DO の有効化は Shizuku、Dhizuku、root、ADB の 4 経路。機能はカメラ／スクショ禁止、USB 信号の無効化、lock task、wipe、Wi-Fi の追加・変更・削除、ネットワークログ、アプリの停止／非表示／アンインストールのブロック／権限付与、ユーザー制限、ユーザー管理、パスワードのリセットなど — [OwnDroid](https://github.com/BinTianqi/OwnDroid)
- **[コミュニティ] Dhizuku**: GPL-3.0、star 約 4.0k、アクティブ。Android 6.0〜17。アカウントがあると導入できない場合がある。Android 12 以降で Shizuku 経由の有効化は experimental — [Dhizuku](https://github.com/iamr0s/Dhizuku)
- **[第三者] targetSdk の下限**: Android 14 以降は targetSdk が 23 未満のアプリをインストールできない（`INSTALL_FAILED_DEPRECATED_SDK_VERSION`）。Android 15 では 24 未満に引き上げられたとする資料が多い（Android 16 では据え置き）。ADB では `adb install --bypass-low-target-sdk-block` で回避できる — [Bayton: Android 14 minimum SDK](https://bayton.org/android/android-14-minimum-sdk/) / [Bayton: minimum targetSdk matrix](https://bayton.org/android/android-minimum-targetsdk-matrix/)。この閾値は ROM 側の設定で、ビルドにより異なりうる — [F-Droid issue #2692](https://gitlab.com/fdroid/fdroidclient/-/issues/2692)
- **[第三者] Play の target API 要件**: Play で配布するアプリは 2026-08-31 から API 36 をターゲットにする必要があるが、"Privately distributed apps are not immediately affected"。API 37 の要件は 2027 年半ば（日付は TBC） — [Bayton: Android 16](https://bayton.org/android/android-16-enterprise-features/) / [Bayton: Android 17](https://bayton.org/android/android-17-enterprise-features/)
- **[公式]** DPC の実装は `DeviceAdminReceiver` のサブクラスが前提。DPC Support Library は Maven ではなく EMM Provider community で配布され、`play-services-auth` に依存する（managed Google Play アカウント用で、自作 DPC には不要） — [Build a DPC](https://developer.android.com/work/dpc/build-dpc)
- **[公式] API 28 の `transferOwnership`**: DO を別の admin コンポーネントへ移譲する（API diff 28 で追加） — [API diff 28](https://developer.android.com/sdk/api_diff/28/changes/android.app.admin.DevicePolicyManager)
- **[第三者] DPC 間の移行**: wipe なしで DPC を移行するプラットフォーム API は Android 9 で導入された。2 つの DPC 間での ownership 移譲を Pixel 9 Pro XL で PoC した例がある。AMAPI 側の DPC 移行機能（2024 年初頭）は同一 EMM 内に限られ、一方向 — [Bayton: DPC migration](https://bayton.org/blog/2026/04/dpc-migration/)
- **[AOSP] 自己更新の可能性**: DO はユーザー操作なしに PackageInstaller のセッションを完了できる（Q4 参照） — [PackageInstallerSession.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/services/core/java/com/android/server/pm/PackageInstallerSession.java)

### Inferences
- [推論] **最小構成**（Kotlin、Android Studio、minSdk 28〜31 程度、targetSdk は 35〜36）:
  - `AdminReceiver` と `device_admin.xml`（使うポリシーを宣言）
  - Android 12 以降必須の `ProvisioningModeActivity` と `PolicyComplianceActivity`（tier (c) のみ必要。ただし ADB 経由でも害はないので常に入れておく）
  - `AgentProvider`（`call()` で uid を検査）
  - ポリシーごとの小さな「モジュール」クラス群。fabriq の `modules/standard/<name>` に 1 対 1 で対応させる
  - 規模は数千行程度と見込まれる（未検証の見積もり）
- [推論] **署名鍵は一生ものになる**。署名が変わると Android の通常の更新ができず、QR の SIGNATURE_CHECKSUM も変わる。鍵を失った場合の脱出口は `transferOwnership`（新しいパッケージを旧署名でインストールできないので、実際には旧アプリから新アプリへ移譲する手順を事前に用意しておく必要がある）か、factory reset。鍵は HSM か厳重な保管庫に置き、fabriq の `kernel/` 相当の扱いにする。
- [推論] **更新経路**: (1) ホストから `adb install -r`（同じ署名） (2) エージェントが自分で PackageInstaller を使って更新（DO は無言インストール可） (3) ホストから QR を作り直す（新規端末のみ）。fabriq の VERSION / REQUIRES_KERNEL のように、エージェントの versionCode とホスト側モジュールの互換性マトリクスが必要になる（`call("agent.version")` で照会する）。
- [推論] **targetSdk の毎年の追従**: Play の要件は適用されないが、インストール時の下限（Android 14 で 23、Android 15 で 24）と、新しい targetSdk で導入される挙動変更（例: API 37 の local network permission）を受けて、少なくとも年 1 回はビルドと実機での回帰確認が必要になる。
- [推論] **ライセンス**: 社内配布のみなら GPL-3.0 のコード（OwnDroid / Dhizuku）を流用しても配布範囲は限られるが、取り込むなら派生物もソース開示義務の対象になりうる。TestDPC（Apache-2.0）を参照実装にするのが無難。
- [推論] **Play Protect 対策**: accessibility、SMS、オーバーレイ、通知リスナーなどの権限は宣言しない。QR 経路を使うなら、Google の DPC 承認／appeal プロセスを使うかどうかを判断する必要がある（自作の非 EMM DPC が承認されうるかは不明）。

### Gaps
- 自作の非商用（非 EMM）DPC が Google の DPC 承認／allowlist に申請でき、承認されうるかは不明（基準のページ本体に到達できなかった）。
- TestDPC の最終リリース日やコミット日はページから取得できなかった。メンテナンスの頻度は未確認。
- 「OpenDPC」「Flyve MDM」の 2026 年時点の状態は調べられなかった（Flyve MDM は過去に開発停止の情報があるが未検証）。Island / Shelter は work profile（PO）用で、DO を使うキッティング用途には直接の参考にならない（未検証）。
- 実装規模や工数の定量的な一次データは無い。

## Q7. ユーザー体験への影響（組織所有の表示、consumer 端末に戻せるか、Google アカウントと Play）

### Takeaway
DO 端末は「組織が管理する端末」となり、元に戻すには原則 factory reset が必要になる（testOnly 版を除く）。DO にするとバックアップサービスが停止され、managed profile は追加できなくなる。managed Google Play を使わない自作 DPC の場合、アプリ配布はサイドロード（DO による無言インストール）が前提になる。consumer の Google アカウントを追加できるかは `DISALLOW_MODIFY_ACCOUNTS` の運用次第と見られるが、一次情報では確認できていない。

### Cited Findings
- **[AOSP]** DO を設定するとバックアップサービスが停止する（"Shutting down backup manager service permanently."）。DO が設定された端末では managed profile の追加が制限される — [DevicePolicyManagerService.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/services/devicepolicy/java/com/android/server/devicepolicy/DevicePolicyManagerService.java)
- **[コミュニティ]** DO を設定すると `no_add_managed_profile` 制限が入る。DO とワークプロファイルの owner が同じアプリだと、再起動後に権限を失う — [OwnDroid README](https://github.com/BinTianqi/OwnDroid)
- **[公式]** Android 12 以降の fully managed 端末で、admin がセンサー権限を管理する場合、SUW でユーザーに明示メッセージが表示される — [Android 12 enterprise](https://developer.android.com/work/versions/android-12)
- **[AOSP]** QR／NFC による DO プロビジョニングは、`LEAVE_ALL_SYSTEM_APPS_ENABLED` を指定しないとシステムアプリを無効化する — [DevicePolicyManager.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/core/java/android/app/admin/DevicePolicyManager.java)。afw# 経路でも、ほとんどのシステムアプリが既定で無効化される — [Bayton: provisioning methods](https://bayton.org/android/android-enterprise-provisioning-methods/)
- **[第三者]** Play EMM API が無い custom DPC では、managed Play による配信、managed account、承認は使えない — 同上
- **[第三者] Android 17**: Theft Detection Lock と Remote Lock が既定で有効になる（新品、初期化後、アップグレード後のすべて） — [Bayton: Android 17](https://bayton.org/android/android-17-enterprise-features/)

### Inferences
- [推論] DO の表示（設定画面やクイック設定の「この端末は組織が管理しています」系の文言と、`setDeviceOwnerLockScreenInfo` / `setOrganizationName` による所有者表示）は、貸与・社用端末としてはむしろ望ましい。個人に払い下げる運用があるなら、払い下げ時の factory reset 手順（企業 FRP を設定しているなら解除も）を手順書に組み込む必要がある。
- [推論] 自作 DO 端末でも、制限をかけない限りユーザーは consumer の Google アカウントで Play ストアにサインインできると考えられる。業務用に固定するなら `DISALLOW_MODIFY_ACCOUNTS` を付け、必要なアプリを DO で事前に配布する。
- [推論] 企業 FRP（`setFactoryResetProtectionPolicy`）を設定すると盗難対策になる。その代わり、再キッティング時に指定した企業アカウントでのアンロックが必要になり、運用の負担が増える。Android 17 で既定有効になる盗難保護との組み合わせ（再キッティング時のロック）は検証が必要。

### Gaps
- 自作 DO 端末で consumer の Google アカウントの追加と Play ストアの利用が実際にできるかについて、2025〜2026 年の一次資料は見つからなかった。
- 「組織が管理しています」表示の正確な文言や表示場所（Android 14〜17）は一次資料で確認していない。
- Android 17 の Theft Detection Lock 既定有効化が、DO 端末のキッティングや再キッティングに与える影響は不明。
- 調査上の制約: bayton.org、support.google.com、androidenterprise.community、android.googlesource.com はこのセッションのネットワークポリシーで到達できなかった。Bayton は GitHub 上のソース Markdown、AOSP は GitHub の aosp-mirror、Google ヘルプとコミュニティは検索スニペットで代替した。
