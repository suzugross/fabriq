# ADB shell（uid 2000 / root なし / MDM・Device Owner なし）で Android に何をキッティングできるか — 2026-10-29 時点

> **凡例（確度ラベル）**
> - 【AOSP】= AOSP ソース（GitHub 上の aosp-mirror / LineageOS ミラー）を本調査で直接読んで確認。
>   - `aosp-mirror/platform_frameworks_base@main` は `Build.java` で `BAKLAVA = CUR_DEVELOPMENT`（＝Android 16 リリース前の開発スナップショット）([Build.java main](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/core/java/android/os/Build.java))。
>   - `LineageOS/*@lineage-23.0` は `BAKLAVA = 36`（＝Android 16 リリース版ベース）([Build.java lineage-23.0](https://github.com/LineageOS/android_frameworks_base/blob/lineage-23.0/core/java/android/os/Build.java))。Wi-Fi / Bluetooth / Permission(Role) / Connectivity / Telephony などの Mainline モジュールはこちらで確認。
>   - Android 13/14/15 の Shell 権限は `aosp-mirror` の `android-13.0.0_r1` / `android-14.0.0_r1` / `android-15.0.0_r1` タグで比較。
> - 【公式】= developer.android.com / Google ヘルプ / Samsung 公式ページ。
> - 【コミュ】= コミュニティ報告（XDA, Stack Overflow, GitHub, ブログ等）。実機差あり。
> - 【推論】= 上記からの推論・未検証。実機検証必須。
>
> **調査上の制約**: このセッションからは android.googlesource.com / cs.android.com / source.android.com / bayton.org / androidauthority.com への直接 fetch がブロックされたため、AOSP は GitHub ミラー経由、一部ニュースは検索スニペット経由で引用している（該当箇所に明記）。**Android 17 の AOSP ソースは未確認**。

---

## Q0. 総合ケイパビリティ・マトリクス（全項目サマリ）

### Takeaway
shell ユーザーは `WRITE_SECURE_SETTINGS`・`INSTALL/DELETE_PACKAGES`・`GRANT_RUNTIME_PERMISSIONS`・`MANAGE_APP_OPS_MODES`・`MANAGE_ROLE_HOLDERS`・`NETWORK_SETTINGS`・`SET_TIME_ZONE`・`CHANGE_CONFIGURATION`・`REBOOT`・`ENABLE_TEST_HARNESS_MODE` などを持つため、**設定値・アプリ・権限・既定アプリ・Wi-Fi(PSK系)・無線系トグル・画面ロック・タイムゾーン・棚卸し**の大半は「適用＋読み返し」の両方を CLI で組める。一方 **CA/クライアント証明書のサイレント導入、VPN プロファイル作成、EAP(802.1X) Wi-Fi、システムロケール直接変更（コマンド無し）、OTA 実行、壁紙設定（コマンド無し）**は shell 単体では不可またはヘルパー（APK/app_process）や UI 操作が必要。**前提の「開発者オプション＋USB デバッグ有効化＋RSA 許可」は端末ごとの手動タップが原則不可避**。

### Cited Findings

| # | キッティング項目 | Apply（適用） | Verify（読み返し） | 最小 Android | 制約・版差・OEM差 | 確度 |
|---|---|---|---|---|---|---|
| 1 | 画面消灯タイムアウト | `settings put system screen_off_timeout 600000` | `settings get system screen_off_timeout` | 4.2+ | shell は system テーブルの制限キー検査を免除 | 【AOSP】 |
| 2 | 輝度/自動輝度 | `settings put system screen_brightness 128` / `screen_brightness_mode 0` | 同 `get` | 4.2+ | OEM により輝度スケール（0-255 以外）差 | 【AOSP】(書込可否)＋【推論】(値域) |
| 3 | 自動回転 | `settings put system accelerometer_rotation 0` | 同 `get` | 4.2+ | — | 【AOSP】(書込可否) |
| 4 | 充電中スリープしない | `settings put global stay_on_while_plugged_in 7` | 同 `get` | 4.2+ | 開発者オプション項目 | 【AOSP】(testharness が同キー使用) |
| 5 | 端末名 | `settings put global device_name "PC-001"` | 同 `get` | 7.1+? | BT 名・Wi-Fi Direct 名に反映されるかは OEM 差 | 【推論】 |
| 6 | 自動日時/自動TZ | `cmd time_detector set_auto_detection_enabled false` / `cmd time_zone_detector set_auto_detection_enabled false`（または `settings put global auto_time 0` / `auto_time_zone 0`） | `cmd time_zone_detector is_auto_detection_enabled` / `settings get global auto_time_zone` | 12+（detector cmd） | サブコマンド文字列は端末で `help` 確認推奨 | 【AOSP】 |
| 7 | タイムゾーン | `cmd alarm set-timezone Asia/Tokyo` | `getprop persist.sys.timezone` / `date` | 9+?（cmd alarm） | 自動TZ有効だと上書きされうる→先に自動TZ off | 【AOSP】 |
| 8 | 時刻 | `cmd alarm set-time <epoch_ms>` | `date` | — | 自動時刻 on だと NTP で上書き | 【AOSP】 |
| 9 | NTP サーバ | `settings put global ntp_server ntp.nict.jp` | 同 `get` | — | キーは @hide/@Readable。反映経路は未検証 | 【AOSP】(キー)/【推論】(反映) |
| 10 | 12/24h | `settings put system time_12_24 24` | 同 `get` | 4.2+ | — | 【推論】 |
| 11 | フォント倍率 | `settings put system font_scale 1.15` | 同 `get` | 4.2+ | — | 【推論】 |
| 12 | 位置情報 ON/OFF | `cmd location set-location-enabled true` | `cmd location is-location-enabled` | 11+（推定） | 旧 `location_providers_allowed` は廃止 | 【AOSP】 |
| 13 | Wi-Fi ON/OFF | `cmd wifi set-wifi-enabled enabled` | `cmd wifi status` | 11+?（`svc wifi` は移行済） | — | 【AOSP】 |
| 14 | Wi-Fi 保存ネットワーク（open/OWE/WPA2/WPA3/WEP） | `cmd wifi add-network "<SSID>" wpa2 "<PSK>" [-h] [-m] [-d] [-r ...]` / `connect-network` | `cmd wifi list-networks` / `cmd wifi status` | 11+（コミュ） | **EAP/802.1X は非対応** | 【AOSP】 |
| 15 | スキャン常時可 | `cmd wifi set-scan-always-available disabled` | `settings get global wifi_scan_always_enabled` | — | — | 【AOSP】 |
| 16 | モバイルデータ | `cmd phone data enable\|disable`（旧: `svc data`） | `settings get global mobile_data` / `dumpsys telephony.registry` | — | DSDS はサブスクリプション別 | 【AOSP】(cmd phone) |
| 17 | 機内モード | `cmd connectivity airplane-mode enable\|disable` | `cmd connectivity airplane-mode`（引数なしで状態表示） | 11+? | — | 【AOSP】 |
| 18 | Bluetooth | `cmd bluetooth_manager enable\|disable`（`wait-for-state:STATE_ON` あり） | `settings get global bluetooth_on` / `dumpsys bluetooth_manager` | 13+?（`svc bluetooth` は移行済） | — | 【AOSP】 |
| 19 | NFC | `svc nfc enable\|disable` | `dumpsys nfc` | — | NFC 非搭載機は no-op | 【AOSP】 |
| 20 | APK 導入（単体/分割） | `adb install [-r -d -g] x.apk` / `adb install-multiple base.apk split_*.apk` / `pm install-create/-write/-commit` | `pm list packages --show-versioncode` / `dumpsys package <pkg>` | — | 14: targetSdk<23 ブロック、15: <24 ブロック（`--bypass-low-target-sdk-block`） | 【公式】【AOSP】 |
| 21 | 不要アプリ除去（可逆） | `pm uninstall -k --user 0 <pkg>` / `pm disable-user --user 0 <pkg>` | `pm list packages -u` / `pm list packages -d` | — | 戻し: `cmd package install-existing <pkg>` / `pm enable` | 【AOSP】【コミュ】 |
| 22 | アプリ hide | `pm hide` | — | — | **shell 不可**（MANAGE_USERS 必要・shell 非保持） | 【AOSP】 |
| 23 | アプリ suspend | `pm suspend --user 0 <pkg>` | `dumpsys package <pkg>`（suspended 表示） | 9+? | shell は SUSPEND_APPS 保持 | 【AOSP】 |
| 24 | ランタイム権限 | `pm grant <pkg> android.permission.POST_NOTIFICATIONS` / `pm revoke` | `dumpsys package <pkg>`（runtime permissions） | 6.0+ / 通知は 13+ | development 保護の特権（WRITE_SECURE_SETTINGS 等）もアプリへ grant 可 | 【AOSP】 |
| 25 | AppOps | `appops set <pkg> REQUEST_INSTALL_PACKAGES allow` / `SYSTEM_ALERT_WINDOW` / `MANAGE_EXTERNAL_STORAGE` / `RUN_ANY_IN_BACKGROUND` | `appops get <pkg> <OP>` | — | shell は MANAGE_APP_OPS_MODES 保持 | 【AOSP】 |
| 26 | 既定アプリ（ブラウザ/SMS/電話/ホーム） | `cmd role add-role-holder android.app.role.BROWSER <pkg>` / `pm set-home-activity <pkg>` | `cmd role get-role-holders android.app.role.BROWSER` | 10+ | アプリが role の資格要件を満たす必要 | 【AOSP】 |
| 27 | 電池最適化除外 | `dumpsys deviceidle whitelist +<pkg>` | `dumpsys deviceidle whitelist` | 6.0+ | OEM 独自の省電力（Samsung 等）は別管理 | 【AOSP】/【推論】(OEM) |
| 28 | ユーザー補助サービス | `settings put secure enabled_accessibility_services <pkg>/<svc>` + `settings put secure accessibility_enabled 1` | 同 `get` / `dumpsys accessibility` | — | 13+ の「制限付き設定」は UI 側の制約。ADB 経由での回避可否は要実機検証 | 【AOSP】(書込権限)/【コミュ】 |
| 29 | 通知リスナー | `cmd notification allow_listener <pkg>/<svc>` | `settings get secure enabled_notification_listeners` | — | — | 【AOSP】 |
| 30 | デバイス管理者（非 DO） | `dpm set-active-admin --user current <pkg>/<rcv>` | `dumpsys device_policy` | — | 削除コマンドは testOnly admin のみ | 【AOSP】 |
| 31 | IME | `ime enable <id>` / `ime set <id>` | `settings get secure default_input_method` / `enabled_input_methods` | — | — | 【AOSP】 |
| 32 | 画面ロック PIN/PW/パターン | `locksettings set-pin 1234` / `set-password` / `set-pattern` / `clear --old 1234` | `locksettings verify --old 1234` / `locksettings get-disabled` | 8.0+?（推定） | 設定済みなら `--old` 必須、管理ポリシー不適合は拒否 | 【AOSP】 |
| 33 | システムロケール | **専用コマンドなし**。ヘルパー（CHANGE_CONFIGURATION を `pm grant`）か shell-uid の app_process ヘルパー | `getprop persist.sys.locale` / `settings get system system_locales` | — | `settings put system system_locales` は「直接使うな」と AOSP コメント | 【AOSP】【コミュ】 |
| 34 | アプリ別ロケール | `cmd locale set-app-locales <pkg> --locales ja-JP` | `cmd locale get-app-locales <pkg>` | 13+ | — | 【AOSP】 |
| 35 | CA 証明書 | **不可（サイレント）**。ファイル push＋設定画面での手動導入 | `settings`では不可。UI or `dumpsys`（未検証） | — | Android 11+ で CA は Settings 経由のみ | 【AOSP】(KeyChain doc) |
| 36 | クライアント証明書（PKCS#12） | **サイレント不可**。push＋CertInstaller/設定 UI | UI | — | — | 【AOSP】/【推論】 |
| 37 | APN 追加 | `content insert --uri content://telephony/carriers --bind ...` | `content query --uri content://telephony/carriers` | **14+**（shell に WRITE_APN_SETTINGS 付与） | 13 以前は SecurityException。キャリア/OEM 制限は未検証 | 【AOSP】 |
| 38 | VPN プロファイル | **不可**（CONTROL_VPN 非保持） | — | — | VPN アプリ側の import 機能に依存 | 【AOSP】 |
| 39 | 壁紙 | **cmd なし**（`cmd wallpaper` は dim 系のみ）。ヘルパー APK（SET_WALLPAPER は normal 権限）等 | `dumpsys wallpaper` | — | — | 【AOSP】 |
| 40 | ホーム画面レイアウト | 不可（ランチャー固有） | — | — | — | 【推論】 |
| 41 | 棚卸し（シリアル/機種/OS/パッチ/指紋） | — | `getprop ro.serialno` / `ro.product.model` / `ro.build.version.release` / `ro.build.version.security_patch` / `ro.build.fingerprint`、`adb get-serialno` | — | — | 【推論】(一般知識) |
| 42 | IMEI/電話番号 | — | `service call iphonesubinfo <code> ...`（トランザクション番号は版依存） | — | shell は READ_PRIVILEGED_PHONE_STATE / READ_PHONE_NUMBERS 保持 | 【AOSP】(権限)/【コミュ】(コード番号) |
| 43 | エビデンス | — | `adb exec-out screencap -p > x.png` / `screenrecord --time-limit 180` / `uiautomator dump` | — | screenrecord 最大 180 秒・音声なし。FLAG_SECURE 画面は黒 | 【公式】/【コミュ】 |
| 44 | 再起動・起動完了待ち | `adb reboot` | `adb wait-for-device` → `getprop sys.boot_completed` = 1 | — | — | 【AOSP】(REBOOT 権限) |
| 45 | 初期化（再キッティング） | `cmd testharness enable`（全データ消去・ADB 鍵保持・SUW スキップ） | 再起動後 `getprop` / `settings get global device_provisioned` | 10+ | 画面ロック設定時は拒否。テストハーネスモードの副作用あり | 【公式】【AOSP】 |
| 46 | OTA 実行 | **不可**（UI。リカバリでの `adb sideload` は OEM 署名パッケージのみ） | `getprop ro.build.version.security_patch` | — | — | 【推論】 |
| 47 | 終盤ハードニング | `settings put global development_settings_enabled 0` → 最後に `settings put global adb_enabled 0` | 切断されるため **事前** に他検証を完了 | — | adb_enabled 0 で即切断（推定） | 【推論】 |

- 上表の「shell が持つ/持たない権限」の根拠: Shell の `AndroidManifest.xml`（main）に `WRITE_SETTINGS` / `WRITE_SECURE_SETTINGS` / `CHANGE_CONFIGURATION` / `INSTALL_PACKAGES` / `DELETE_PACKAGES` / `GRANT_RUNTIME_PERMISSIONS` / `INSTALL_GRANT_RUNTIME_PERMISSIONS` / `MANAGE_APP_OPS_MODES` / `CHANGE_COMPONENT_ENABLED_STATE` / `SUSPEND_APPS` / `MANAGE_ROLE_HOLDERS` / `MANAGE_DEFAULT_APPLICATIONS` / `SET_PREFERRED_APPLICATIONS` / `NETWORK_SETTINGS` / `CHANGE_WIFI_STATE` / `OVERRIDE_WIFI_CONFIG` / `READ_WIFI_CREDENTIAL` / `BLUETOOTH_PRIVILEGED` / `NFC` / `MODIFY_PHONE_STATE` / `READ_PRIVILEGED_PHONE_STATE` / `READ_PHONE_NUMBERS` / `LOCAL_MAC_ADDRESS` / `SET_TIME` / `SET_TIME_ZONE` / `MANAGE_TIME_AND_ZONE_DETECTION` / `SET_WALLPAPER` / `SET_WALLPAPER_COMPONENT` / `MANAGE_ACCESSIBILITY` / `MANAGE_NOTIFICATION_LISTENERS` / `MANAGE_DEVICE_ADMINS` / `MANAGE_PROFILE_AND_DEVICE_OWNERS` / `SET_AND_VERIFY_LOCKSCREEN_CREDENTIALS` / `REBOOT` / `DEVICE_POWER` / `DUMP` / `ENABLE_TEST_HARNESS_MODE` / `MANAGE_DEBUGGING` / `WRITE_APN_SETTINGS` などが列挙されている — [Shell AndroidManifest.xml (main)](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/packages/Shell/AndroidManifest.xml)
- 同マニフェストに **`MANAGE_USERS`、`MASTER_CLEAR`、`MANAGE_CA_CERTIFICATES`、`CONTROL_VPN` は存在しない**（main・Android 16(lineage-23.0) とも grep で 0 件）— [Shell AndroidManifest.xml (main)](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/packages/Shell/AndroidManifest.xml); [Shell AndroidManifest.xml (lineage-23.0 / Android 16)](https://github.com/LineageOS/android_frameworks_base/blob/lineage-23.0/packages/Shell/AndroidManifest.xml)
- バージョン比較（shell 保持の有無）: `WRITE_APN_SETTINGS` は A13 r1=無 / A14 r1=有 / A15 r1=有 / A16=有。`CHANGE_CONFIGURATION`・`SET_WALLPAPER`・`READ_PRIVILEGED_PHONE_STATE`・`MANAGE_ROLE_HOLDERS`・`SUSPEND_APPS`・`MANAGE_DEVICE_ADMINS`・`ENABLE_TEST_HARNESS_MODE`・`MANAGE_NOTIFICATION_LISTENERS`・`SET_TIME_ZONE`・`NETWORK_SETTINGS`・`LOCAL_MAC_ADDRESS`・`MANAGE_DEBUGGING` は A13〜A16 すべて有 — [A13](https://github.com/aosp-mirror/platform_frameworks_base/blob/android-13.0.0_r1/packages/Shell/AndroidManifest.xml); [A14](https://github.com/aosp-mirror/platform_frameworks_base/blob/android-14.0.0_r1/packages/Shell/AndroidManifest.xml); [A15](https://github.com/aosp-mirror/platform_frameworks_base/blob/android-15.0.0_r1/packages/Shell/AndroidManifest.xml); [A16](https://github.com/LineageOS/android_frameworks_base/blob/lineage-23.0/packages/Shell/AndroidManifest.xml)
- 主要権限の保護レベル（core/res/AndroidManifest.xml）: `CHANGE_CONFIGURATION`=signature|privileged|**development**|role、`WRITE_SECURE_SETTINGS`=signature|privileged|**development**|role|installer、`DUMP`/`READ_LOGS`=…|**development**、`PACKAGE_USAGE_STATS`=…|development|appop、`SET_WALLPAPER`=**normal**、`WRITE_APN_SETTINGS`/`MANAGE_CA_CERTIFICATES`/`CONTROL_VPN`=signature|privileged、`MASTER_CLEAR`=signature|privileged|role、`SYSTEM_ALERT_WINDOW`=…|appop|…|development、`MANAGE_EXTERNAL_STORAGE`=signature|appop|preinstalled、`REQUEST_INSTALL_PACKAGES`=signature|appop — [core AndroidManifest.xml](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/core/res/AndroidManifest.xml)

### Inferences
- `development` フラグ付き権限（WRITE_SECURE_SETTINGS / CHANGE_CONFIGURATION / DUMP / READ_LOGS 等）は `pm grant` でヘルパー APK に付与できるため、**「shell から直接はコマンドが無いが、shell 権限の範囲内の API はある」項目（システムロケール等）はヘルパー APK で埋められる**。壁紙は normal 権限なので任意アプリで可。
- `signature|privileged` 権限（CA 証明書・VPN・MASTER_CLEAR）は shell も保持しておらずヘルパーにも付与できない＝**ADB 単独 tier の本質的限界**。
- fabriq のモジュール設計では、各行の Apply/Verify をそのまま「Step 3 適用 / Step 5.5 検証」に写せる。`adb shell` は Android 7+（shell v2 プロトコル）でリモートの終了コードを返す（一般知識・未検証）ため、PowerShell 側は `$LASTEXITCODE` と出力文字列の両方で判定するのが安全。

### Gaps
- Android 17（API 37）の Shell マニフェスト差分は未確認（AOSP 17 ソースに到達できず）。
- 各 `cmd` サブコマンドの**導入バージョン**（wifi add-network、connectivity airplane-mode、bluetooth_manager 等）は git 履歴を辿れず、表中の「?」は未検証。
- OEM（特に日本メーカー）が Shell マニフェストやシェルコマンドを改変しているかは確認できず。

---

## Q1. 前提（開発者オプション・USB デバッグ・RSA 許可）は省略/削減できるか。鍵共有・ワイヤレスデバッグ・タイムアウト

### Takeaway
小売端末では **ビルド番号×7 → USB デバッグ ON → RSA「常に許可」** の手動操作は省略不可（ADB が動く前に ADB で設定できない鶏卵問題）。削減策は (a) **ホスト鍵 `adbkey`/`adbkey.pub` を全キッティング PC で共有**して RSA ダイアログを端末あたり 1 回にする、(b) 初回接続直後に `settings put global adb_allowed_connection_time 0` で **7 日無通信失効を止める**、(c) 再キッティング時は `cmd testharness enable` で **ADB 鍵を保持したまま初期化**する、の 3 点。ワイヤレスデバッグは端末ごとのペアリング操作が要るため大量一括には不向き（Android 17 で再接続は改善）。

### Cited Findings
- Android 4.2.2 以降、接続時に RSA 鍵を受け入れるかのダイアログが表示され、端末をアンロックしてダイアログを承認しない限り adb コマンドは実行できない — [Android Debug Bridge (developer.android.com)](https://developer.android.com/tools/adb)【公式】
- 「常に許可」した鍵は最終接続時刻とともに保存され、`ADB_ALLOWED_CONNECTION_TIME` 以内なら再接続時にユーザー操作不要。既定値 `DEFAULT_ADB_ALLOWED_CONNECTION_TIME = 604800000`（7 日）— [Settings.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/core/java/android/provider/Settings.java)【AOSP】
- `AdbKeyStore` は「一定期間の非アクティブ後に always-allow を失効させる」クラスで、`isKeyAuthorized()` は `allowedConnectionTime == 0 || now < lastConnectionTime + allowedConnectionTime` を返し、`filterOutOldKeys()` は `allowedTime == 0` なら何も削除しない（＝0 で無期限）— [AdbDebuggingManager.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/services/core/java/com/android/server/adb/AdbDebuggingManager.java)【AOSP】
- `adb_allowed_connection_time` は `@Readable` の Global 設定で、shell（uid 2000 < FIRST_APPLICATION_UID）は SettingsProvider の readable 制限チェックをスキップされる。shell は `WRITE_SECURE_SETTINGS` を保持 — [SettingsProvider.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/packages/SettingsProvider/src/com/android/providers/settings/SettingsProvider.java); [Shell manifest](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/packages/Shell/AndroidManifest.xml)【AOSP】
- テストハーネスモードの設定処理でも「AdbService 初期化前に ADB 鍵の TTL を無効化」として `ADB_ALLOWED_CONNECTION_TIME` に 0 を書いている（＝0 が「無期限」の正規の意味）— [TestHarnessModeService.java](https://github.com/LineageOS/android_frameworks_base/blob/lineage-23.0/services/core/java/com/android/server/testharness/TestHarnessModeService.java)【AOSP】
- adbd が認可に使うのは `/data/misc/adb/adb_keys`（ユーザー鍵）と `/adb_keys`（システムイメージ同梱鍵）— [AdbDebuggingManager.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/services/core/java/com/android/server/adb/AdbDebuggingManager.java)【AOSP】
- 「常に許可」は `~/.android/adbkey.pub` を端末の `/data/misc/adb/adb_keys` にコピーする動作 — [ArchWiki: ADB](https://wiki.archlinux.org/title/ADB)【コミュ】
- adbkey ファイルを別ワークステーションにコピーすると、端末はそれを既知の認可済みコンピュータとみなす（フォレンジック用途の記述）— [Practical Mobile Forensics (O'Reilly)](https://oreilly.com/library/view/practical-mobile-forensics/9781788839198/008a98ec-157d-42a2-960a-579e5b30cbfb.xhtml)【コミュ】
- ワイヤレスデバッグ: スマホは Android 11（API 30）以上、TV/Wear は 13 以上。「ペア設定コードによるペア設定」で表示された IP:port とコードを使い `adb pair ipaddr:port`。ペアリングは 1 回でよく、同一ネットワーク上で自動接続。Android 17 と adb 37.0.0 で「adb Wi-Fi 2.0」が導入され、ワイヤレスデバッグの信頼済みネットワークに接続すると端末がワークステーションへ自動接続する — [Android Debug Bridge](https://developer.android.com/tools/adb)【公式】
- AdbKeyStore はワイヤレスデバッグ用に「常に許可する信頼済みネットワーク（BSSID）」リストも保持 — [AdbDebuggingManager.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/services/core/java/com/android/server/adb/AdbDebuggingManager.java)【AOSP】
- Samsung **Auto Blocker** は USB ケーブル経由で悪意ある PC 等がコマンドを送れないようにする機能で、設定 > セキュリティとプライバシー > Auto Blocker にある — [Samsung UK: Auto Blocker](https://www.samsung.com/uk/support/mobile-devices/protect-your-galaxy-device-with-the-new-auto-blocker-feature/)【公式】。One UI 6.1.1 で出荷された端末では既定 ON（セットアップ時に OFF 可）、One UI 9 beta の「最大制限」モードでは USB 接続自体を完全ブロック — [Android Authority (検索スニペット経由)](https://androidauthority.com/samsung-one-ui-9-auto-blocker-update-3666893/)【コミュ/報道】
- Android 16 の Advanced Protection は「提供元不明アプリのインストール」権限付与を不可にし、ロック中の USB データ信号を無効化。2026 年 6 月報道では開発者オプション（ADB/ワイヤレスデバッグ）の封鎖を計画中とされるが、公式確認なし — [Android Authority: USB data](https://www.androidauthority.com/android-16-usb-data-advanced-protection-3548018); [Android Authority: developer options](https://androidauthority.com/android-advanced-protection-mode-developer-options-3679725)（いずれも検索スニペット経由）【報道】

### Inferences
- **鍵共有の運用**: Windows ホストでは `%USERPROFILE%\.android\adbkey` / `adbkey.pub` を全キッティング PC に同一配布すれば、端末側は 1 回の「常に許可」で全 PC から通る（鍵は秘密情報なので配布経路・保管は厳格に）。`ADB_VENDOR_KEYS` 環境変数で追加鍵パスを指定する方式もある（公式ページでの明記は未確認）。
- **最初の adb 接続直後の Step 0** として `settings put global adb_allowed_connection_time 0` を流すと、倉庫保管中の 7 日失効を防げる。ただし最終的な出荷前ハードニングで既定値（`settings delete global adb_allowed_connection_time` 等）に戻すのが安全。
- USB デバッグの ON 状態は `Settings.Global.ADB_ENABLED` に永続化されるので再起動で通常は維持される（一般知識）。OEM による自動 OFF（一定時間後に自動無効化等）の公式資料は見つからず。
- Samsung は **Auto Blocker を UI で OFF にするのが前提手順**（One UI 6.1.1+ 新品は既定 ON）。Advanced Protection を有効化した端末はキッティング対象外にする判断が妥当。
- ワイヤレスデバッグは「ペアリング＝端末ごとに UI 操作」なので USB ハブ並列より手数が多い。大量一括は USB を主、Wi-Fi は補助（ケーブル抜去後の検証など）に。
- 本当にゼロタッチ化するには Android Enterprise の QR/ゼロタッチ（＝DO/MDM）か OEM 独自（Samsung Knox Mobile Enrollment 等）が必要で、今回の「MDM なし」制約の外。

### Gaps
- 日本メーカー（AQUOS / arrows / Kyocera / Xperia）で USB デバッグの自動 OFF や Auto Blocker 相当機能があるかの資料は見つからず。
- `ADB_VENDOR_KEYS` と鍵共有について Google 公式の明文は今回取得できず（adb の man ページ未参照）。
- 「開発者向けオプション > adb 認証のタイムアウトを無効化」トグルの導入版（Android 11 と記憶）は一次資料で未確認。

---

## Q2. セットアップウィザード（SUW）前の ADB 利用・SUW スキップの可否とリスク

### Takeaway
新品は USB デバッグが OFF のため **SUW 中に ADB は原則使えない**（SUW をある程度進めて設定アプリに到達し、手動で開発者オプションを出す必要がある）。ADB が通れば `settings put global device_provisioned 1` / `settings put secure user_setup_complete 1` / SUW パッケージの `pm disable-user` は技術的に可能だが副作用が読めない。**公式に用意された「SUW スキップ付き初期化」は `cmd testharness enable`**（Android 10+、データ全消去・ADB 鍵保持・device_provisioned/user_setup_complete を自動設定・ロック画面無効化）で、再キッティング用途に最有力。

### Cited Findings
- `cmd testharness enable|restore`: 「端末の全データを消去しテストハーネスモードを有効化、保存済み ADB 鍵を保持し、計測テスト向けに設定を切替える」。既定で「Package Verifier 無効 / 充電中スリープしない有効 / OTA 更新無効 / アカウント自動同期無効」— [TestHarnessModeService.java](https://github.com/LineageOS/android_frameworks_base/blob/lineage-23.0/services/core/java/com/android/server/testharness/TestHarnessModeService.java)【AOSP】
- 同サービスは初回起動時に「データがあれば device を provisioned にして SUW をスキップ」として `DEVICE_PROVISIONED=1`・`USER_SETUP_COMPLETE=1` を書き、ロック画面を無効化、`persist` プロパティでテストハーネスを示す。起動完了後に ADB 鍵ファイルを復元し、`ADB_ALLOWED_CONNECTION_TIME=0`・`DEVELOPMENT_SETTINGS_ENABLED=1`・`PACKAGE_VERIFIER_INCLUDE_ADB=0`・`STAY_ON_WHILE_PLUGGED_IN=any`・`OTA_DISABLE_AUTOMATIC_UPDATE=1` を設定 — [TestHarnessModeService.java](https://github.com/LineageOS/android_frameworks_base/blob/lineage-23.0/services/core/java/com/android/server/testharness/TestHarnessModeService.java)【AOSP】
- 画面ロックがある場合は「Test Harness Mode cannot be enabled if there is a lock screen」で拒否（終了コード 2）。権限 `ENABLE_TEST_HARNESS_MODE` を要求し、消去は `Intent.ACTION_FACTORY_RESET`（外部ストレージも消去）を `android` パッケージへ送る — [TestHarnessModeService.java](https://github.com/LineageOS/android_frameworks_base/blob/lineage-23.0/services/core/java/com/android/server/testharness/TestHarnessModeService.java)【AOSP】
- 公式 adb ドキュメント: 「Android 10（API 29）以上のテスト端末は `testharness` adb シェルコマンドでファクトリーリセットできる」「`testharness` で復元時、現在のワークステーションでのデバッグを許可する RSA 鍵を永続領域へ自動バックアップ」— [Android Debug Bridge](https://developer.android.com/tools/adb)【公式】
- source.android.com のテストハーネス実装ガイドは、UI テストを壊す画面（キーボードや SUW）のスキップ等に限りカスタマイズ可と記述、日本語版はデータ消去を伴う旨明記 — [テストハーネス モードを実装する（source.android.com、検索スニペット経由）](https://source.android.com/compatibility/cts/harness?hl=ja)【公式】
- shell は `ENABLE_TEST_HARNESS_MODE` を A13〜A16 で保持 — [Shell manifest A13](https://github.com/aosp-mirror/platform_frameworks_base/blob/android-13.0.0_r1/packages/Shell/AndroidManifest.xml); [A16](https://github.com/LineageOS/android_frameworks_base/blob/lineage-23.0/packages/Shell/AndroidManifest.xml)【AOSP】
- AOSP のテスト用ツールには「SUW パッケージを disable し `user_setup_complete` 更新を最大 5 秒待ち、失敗時はプロパティ設定＋再起動にフォールバック」する実装がある（テスト用スクリプト）— [platform/tools/test/connectivity commit 7424c34（検索スニペット経由）](https://android.googlesource.com/platform/tools/test/connectivity/+/7424c34%5E%21/)【AOSP/テスト】
- Zebra の StageNow 文書は、Google が 2024 年に Zebra に SUW バイパス機能の削除を求め、その後部分的に復活したと記載（検索スニペット経由、未 fetch）— [Zebra techdocs: Wipe Device](https://techdocs.zebra.com/stagenow/latest/profiles/wipedevice/)【ベンダー文書】

### Inferences
- **`cmd testharness enable` はキッティングの「Restart-Resume」型ワークフローの起点に使える可能性が高い**（初期化→自動で ADB 再接続可→SUW スキップ→即モジュール投入）。ただし (1) 本来はテスト端末用で、`ActivityManager.isRunningInUserTestHarness()` が true になり**アプリ挙動が変わりうる**、(2) OTA 自動更新・アカウント同期が OFF にされる、(3) SUW をスキップするため Google アカウント・各種同意・OEM 初期設定（日本メーカーのキャリア初期設定等）が未実施になる、(4) テストハーネス状態から抜けるには再度の通常初期化が必要、という本番利用上の論点がある。**本番出荷端末に使うかは要判断・要実機検証**。
- `ACTION_FACTORY_RESET` を OS 内部から送る（＝設定アプリからの初期化と同種）ため、**FRP（Factory Reset Protection）は通常発動しない「信頼された初期化」扱いになる**と推定（未検証）。Google アカウントが残った状態でリカバリ初期化すると FRP が掛かるのとは対照的。
- 手動で `device_provisioned` / `user_setup_complete` を立てる方式は、SystemUI（通知・ホームボタン）や OEM SUW の後続処理が半端な状態になるリスクが指摘されがちだが一次資料はなく、**fabriq では採用せず testharness か通常 SUW 完走を推奨**。
- 新品の 1 台目は必ず「SUW 完走（または最小限）→手動で開発者オプション/USB デバッグ」を経るため、キッティング手順書には UI 操作パート（タップ数の最小化）を明記する必要がある。

### Gaps
- testharness 後の端末が Google Play / OEM アプリ（Samsung, 日本メーカー）で問題を起こすかの実例報告は見つからず。
- Samsung・日本メーカー端末で `cmd testharness` が実装/有効化されているか（`PersistentDataBlock` 依存）未確認。
- testharness 消去後に FRP が掛からないことの一次資料は未取得。

---

## Q3. `settings put system|secure|global` の可否と Android 12+ の制限

### Takeaway
**shell は system/secure/global のほぼ全キーを読み書き可能**。Android 12（S）で導入された「@hide キーの読取制限」は uid ≥ 10000 のアプリ向けで、shell（uid 2000）は対象外。system テーブルの「公開キー以外の変更禁止」も shell は免除。効く制限は (1) **ユーザー制限（DISALLOW_*）による拒否**、(2) **キーを書いても実体が追従しない設定**（システムロケール、旧 `install_non_market_apps`、旧 `location_providers_allowed` 等）。

### Cited Findings
- `enforceRestrictedSystemSettingsMutationForCallingPackage`: 「System/root/shell は任意の設定を変更できる」— SYSTEM_UID / SHELL_UID / ROOT_UID なら即 return — [SettingsProvider.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/packages/SettingsProvider/src/com/android/providers/settings/SettingsProvider.java)【AOSP】
- `enforceSettingReadable`: 呼出元 appId < `FIRST_APPLICATION_UID` なら即 return（＝shell 免除）。対象アプリには「From S+, settings keys annotated with @hide are restricted to system_server and system apps only, unless they are annotated with @Readable」の SecurityException — [SettingsProvider.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/packages/SettingsProvider/src/com/android/providers/settings/SettingsProvider.java)【AOSP】
- put/insert 経路で `isSettingRestrictedForUser(name, userId, value, callingUid)` を評価し、UserManager のユーザー制限（例: DISALLOW_SHARE_LOCATION → LOCATION_MODE、DISALLOW_DEBUGGING_FEATURES → ADB_ENABLED、DISALLOW_INSTALL_UNKNOWN_SOURCES → INSTALL_NON_MARKET_APPS）に応じて拒否/巻き戻し — [SettingsProvider.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/packages/SettingsProvider/src/com/android/providers/settings/SettingsProvider.java)【AOSP】
- `settings` コマンドの書式: `get/put/delete/reset/list [--user <USER_ID>|current] NAMESPACE KEY [VALUE] [TAG] [default] [overrideableByRestore]`、NAMESPACE は system/secure/global — [SettingsService.java (lineage-23.0)](https://github.com/LineageOS/android_frameworks_base/blob/lineage-23.0/packages/SettingsProvider/src/com/android/providers/settings/SettingsService.java)【AOSP】（※ Lineage は独自 `--lineage` を追加）
- DB アップグレード処理のコメント: 「Version 138: Settings.Secure#INSTALL_NON_MARKET_APPS is deprecated」「Version 193: remove obsolete LOCATION_PROVIDERS_ALLOWED settings」— [SettingsProvider.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/packages/SettingsProvider/src/com/android/providers/settings/SettingsProvider.java)【AOSP】
- `Settings.System.SYSTEM_LOCALES` は「直接使うな。更新は `LocalePicker#updateLocales` を使え」と注記された @hide/@Readable キー — [Settings.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/core/java/android/provider/Settings.java)【AOSP】
- `Settings.Global.NTP_SERVER`（`ntp_server`）は `time.android.com|ntp://time.other:123` のような複数指定書式を受け付ける @hide/@Readable キー — [Settings.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/core/java/android/provider/Settings.java)【AOSP】
- shell は `MANAGE_ACCESSIBILITY` と `WRITE_SECURE_SETTINGS` を保持 — [Shell manifest](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/packages/Shell/AndroidManifest.xml)【AOSP】
- Android 13 以降、サイドロードされたアプリはユーザーが手動で「制限付き設定を許可」するまでユーザー補助サービスとデバイス管理者を使えない（アプリ情報 > ⋮ > 制限付き設定を許可）— [Kaspersky: Restricted Settings in Android 13 and 14](https://kaspersky.com/blog/android-restricted-settings/49991); [CleanBrowsing](https://cleanbrowsing.org/support/mobile/disable-restricted-settings-android)【コミュ/ベンダー】

### Inferences
- 依頼にあった主要キーの扱い（推奨形）:
  - system: `screen_off_timeout`, `screen_brightness`, `screen_brightness_mode`, `accelerometer_rotation`, `time_12_24`, `font_scale`, `haptic_feedback_enabled`, `sound_effects_enabled`, `dtmf_tone`, `lockscreen_sounds_enabled` 等 → `settings put system` で可・読み返し可。
  - global: `stay_on_while_plugged_in`, `device_name`, `auto_time`, `auto_time_zone`, `wifi_scan_always_enabled`, `adb_enabled`, `development_settings_enabled`, `ntp_server`, `adb_allowed_connection_time` → 可。ただし `auto_time*` は Android 12+ では `cmd time_detector` / `cmd time_zone_detector` 側の API（MANAGE_TIME_AND_ZONE_DETECTION）を使う方が正規。
  - secure: `location_mode`（→ `cmd location` 推奨）, `default_input_method`/`enabled_input_methods`（→ `ime` 推奨）, `enabled_accessibility_services`, `accessibility_enabled`, `enabled_notification_listeners`（→ `cmd notification` 推奨）, `bluetooth_name`（OEM 依存、BT スタックが読むか未検証）。
  - **書いても効かない/非推奨**: `install_non_market_apps`（Android 8+ は per-app の REQUEST_INSTALL_PACKAGES → `appops`）、`location_providers_allowed`（廃止）、`system_locales`（実体はコンフィグ）、`bluetooth_on`/`wifi_on`/`airplane_mode_on` の直接書換（サービス側が追従しない→専用 cmd を使う）。
- **ロック画面系 secure キー（`lockscreen.*`, `lock_screen_owner_info` 等）は `locksettings` で扱うべき**で、設定 DB 直書きは不整合を招く（推論）。
- 冪等性: `settings get` → 期待値と一致なら skip → `put` → 再 `get` で一致確認、という fabriq の Test-…Match パターンがそのまま使える。
- 「制限付き設定」は UI フローの制限で、ADB から `settings put secure enabled_accessibility_services` を書く経路が Android 13–17 で依然有効かは**一次資料なし（要実機検証）**。shell の権限上は書込可能。

### Gaps
- OEM（Samsung/日本メーカー）の独自キー（例: Samsung の system テーブル独自項目）の一覧と挙動は未調査。
- Android 17 で SettingsProvider の shell 免除が変わっていないかは未確認。
- 「制限付き設定」（Android 15 の Enhanced Confirmation Mode 拡張を含む）が ADB 経由の secure 設定書込を検出・無効化するかは不明。

---

## Q4. ロケール（言語）変更

### Takeaway
**システムロケールを変える shell コマンドは無い**（`cmd locale` はアプリ別ロケール専用、Android 13+）。`settings put system system_locales` は AOSP 自身が「直接使うな」と注記。実用解は **(a) CHANGE_CONFIGURATION を `pm grant` したヘルパー APK**（protection level に development を含むので付与可能）、**(b) shell uid で動く app_process ヘルパー**（scrcpy/Shizuku と同じ方式）、**(c) UI 自動化**。

### Cited Findings
- `LocaleManagerShellCommand` のサブコマンドは `set-app-locales <PKG> [--user] [--locales en,en-US,hi] [--delegate]` / `get-app-locales` / `set-app-localeconfig` / `get-app-localeconfig` / `get-app-localeconfig-ignore-override` のみ — [LocaleManagerShellCommand.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/services/core/java/com/android/server/locales/LocaleManagerShellCommand.java)【AOSP】
- `SYSTEM_LOCALES` の注記「Do not use this value directory. To update system locale, use LocalePicker#updateLocales instead.」— [Settings.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/core/java/android/provider/Settings.java)【AOSP】
- `CHANGE_CONFIGURATION` = signature|privileged|development|role（development ＝ `pm grant` 可）。shell も保持 — [core AndroidManifest.xml](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/core/res/AndroidManifest.xml); [Shell manifest](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/packages/Shell/AndroidManifest.xml)【AOSP】
- `adb shell pm grant <package> android.permission.CHANGE_CONFIGURATION` をヘルパーに付与してロケールを切り替える OSS（Language Switcher Tile 等）がある — [AzimoLabs/Language-Switcher-Tile](https://github.com/AzimoLabs/Language-Switcher-Tile); [Changing Android 10+ device Locale programmatically](https://write.as/julioevm/changing-android-10-or-higher-device-locale-programatically)【コミュ】
- scrcpy サーバーは `adb shell CLASSPATH=/data/local/tmp/scrcpy-server.jar app_process / com.genymobile.scrcpy.Server ...` で起動され「shell として実行」され、shell に付与された特権を使う — [scrcpy doc/develop.md](https://github.com/Genymobile/scrcpy/blob/master/doc/develop.md)【コミュ/OSS】
- Shizuku は root または ADB でサーバープロセスを起動し、アプリにその binder を渡して ADB（shell）権限で system API を呼ばせる。ADB の権限は限定的で版により異なる（Shell マニフェスト参照）— [RikkaApps/Shizuku README](https://github.com/RikkaApps/Shizuku)【コミュ/OSS】

### Inferences
- fabriq-Android の `locale_config` モジュールは「小さな dex/jar を `/data/local/tmp` に push → `app_process` で shell uid 実行 → `LocalePicker.updateLocales()`（または `IActivityManager.updatePersistentConfiguration`）を呼ぶ」方式が、APK 常駐不要で最もクリーン（未検証）。Verify は `getprop persist.sys.locale` と `settings get system system_locales`。
- ヘルパー APK 方式は導入/削除の手間があるが実績が多い。いずれも hidden API 呼び出しになるため Android 版ごとに回帰テストが必要。
- 日本語端末を日本語のまま使うケースが多いなら、ロケールは「検証のみ（Verify only）」モジュールにして変更は例外扱いにする選択肢もある。

### Gaps
- Android 14–17 でのヘルパー方式の動作確認報告（一次）は今回見つからず（参照情報は 2017–2022 年が中心）。
- `persist.sys.locale` を shell が `setprop` できるか（SELinux）未確認（できない前提で設計すべき）。

---

## Q5. タイムゾーン・時刻・NTP

### Takeaway
`cmd alarm set-timezone <Olson>` / `set-time <epoch_ms>` が shell で使え（SET_TIME / SET_TIME_ZONE 保持）、Android 12+ は `cmd time_zone_detector` / `cmd time_detector` で自動検出の ON/OFF・手動提案・状態取得ができる。NTP サーバは `settings put global ntp_server` でキー自体は書けるが、反映経路は未検証（テスト用の `cmd network_time_update_service set_server_config_for_tests` は一時的）。

### Cited Findings
- AlarmManagerService のシェル: `set-time TIME`（エポック ms）、`set-timezone TZ`（Olson ID）、`get-config-version` — [AlarmManagerService.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/apex/jobscheduler/service/java/com/android/server/alarm/AlarmManagerService.java)【AOSP】
- TimeZoneDetectorShellCommand: 自動検出の取得/設定、テレフォニー/ジオ検出のサポート有無、手動 TZ 提案（`<manual suggestion opts>`）、`get-time-zone-state`、`confirm-time-zone --zone_id` 等（文字列は `TimeZoneDetector.SHELL_COMMAND_*` 定数）— [TimeZoneDetectorShellCommand.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/services/core/java/com/android/server/timezonedetector/TimeZoneDetectorShellCommand.java)【AOSP】
- TimeDetectorShellCommand: 自動時刻の取得/設定、手動/テレフォニー/ネットワーク/GNSS/外部時刻の提案、`get-time-state`、ネットワーク時刻の取得/クリア — [TimeDetectorShellCommand.java (lineage-23.0)](https://github.com/LineageOS/android_frameworks_base/blob/lineage-23.0/services/core/java/com/android/server/timedetector/TimeDetectorShellCommand.java)【AOSP】
- NetworkTimeUpdateService のシェル（サービス名 `network_time_update_service`）: `force_refresh`、`set_server_config_for_tests`、`reset_server_config_for_tests` — [NetworkTimeUpdateServiceShellCommand.java](https://github.com/LineageOS/android_frameworks_base/blob/lineage-23.0/services/core/java/com/android/server/timedetector/NetworkTimeUpdateServiceShellCommand.java)【AOSP】
- shell は `SET_TIME` / `SET_TIME_ZONE` / `MANAGE_TIME_AND_ZONE_DETECTION` / `SUGGEST_MANUAL_TIME` を保持 — [Shell manifest](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/packages/Shell/AndroidManifest.xml)【AOSP】

### Inferences
- 推奨手順: ① `cmd time_zone_detector set_auto_detection_enabled false`（必要時）→ ② `cmd alarm set-timezone Asia/Tokyo` → Verify `getprop persist.sys.timezone`。日本のフリートは通常「自動 TZ ON のまま」で問題ないため、**Verify-only（期待値 Asia/Tokyo）** を既定にする選択肢も。
- サブコマンド名（`is_auto_detection_enabled` 等のアンダースコア表記と記憶）は定数値を本調査で未確認のため、実装時に `adb shell cmd time_zone_detector help` で確定させる。
- `ntp_server` は社内 NTP 強制が必要な閉域網キッティングでのみ意味があり、反映（NetworkTimeUpdateService が再読込するか）を実機で `cmd network_time_update_service force_refresh` → `cmd time_detector get_network_time`（名称要確認）で検証すべき。

### Gaps
- `ntp_server` 書換の即時反映/再起動要否は未検証。
- Samsung / 日本メーカーが TZ 検出を独自実装していないか未確認。

---

## Q6. Wi-Fi（`cmd wifi`）

### Takeaway
`cmd wifi add-network` / `connect-network` は **open / owe / wpa2 / wpa3 / wep** のみ対応で **EAP（802.1X）は非対応**。オプションで hidden(-h) / metered(-m) / autojoin 無効(-d) / BSSID 固定(-b) / MAC ランダム化方式(-r) を指定でき、保存ネットワークとして登録される（＝通常の保存ネットワーク同様に永続・自動接続の対象と推定）。shell が使えるのは `NON_PRIVILEGED_COMMANDS` に列挙されたサブコマンドで、それ以外は root 専用。

### Cited Findings
- 非特権（shell 可）コマンド一覧 `NON_PRIVILEGED_COMMANDS`: add-suggestion, forget-network, connect-network, add-network, get-country-code, list-scan-results, list-networks, list-suggestions, remove-suggestion, remove-all-suggestions, set-scan-always-available, set-verbose-logging, set-wifi-enabled, set-passpoint-enabled, start-scan, status 等。それ以外は `Binder.getCallingUid() != ROOT_UID` で拒否 — [WifiShellCommand.java (lineage-23.0)](https://github.com/LineageOS/android_packages_modules_Wifi/blob/lineage-23.0/service/java/com/android/server/wifi/WifiShellCommand.java)【AOSP】
- `connect-network <ssid> open|owe|wpa2|wpa3|wep [<passphrase>] [-x] [-m] [-d] [-h] [-p] [-b <bssid>] [-r auto|none|persistent|non_persistent]`: 「指定パラメータで接続し保存ネットワークに追加」。`add-network`: 「保存ネットワークを追加/更新」。-m=従量制、-d=自動接続無効、-h=hidden、-p=private（非共有）、-x=SSID を16進で指定 — [WifiShellCommand.java](https://github.com/LineageOS/android_packages_modules_Wifi/blob/lineage-23.0/service/java/com/android/server/wifi/WifiShellCommand.java)【AOSP】
- 同ファイルに `"eap"` 文字列および `WifiEnterpriseConfig` の参照は 0 件（＝shell コマンドで EAP 設定不可）— [WifiShellCommand.java](https://github.com/LineageOS/android_packages_modules_Wifi/blob/lineage-23.0/service/java/com/android/server/wifi/WifiShellCommand.java)【AOSP】
- `list-networks`（保存ネットワーク一覧）、`forget-network <networkId>`、`status`（現在の Wi-Fi 状態）、`set-wifi-enabled enabled|disabled`、`set-scan-always-available enabled|disabled` — [WifiShellCommand.java](https://github.com/LineageOS/android_packages_modules_Wifi/blob/lineage-23.0/service/java/com/android/server/wifi/WifiShellCommand.java)【AOSP】
- `svc wifi` は `WifiShellCommand` へ移行済み（Svc.java のコメント）— [Svc.java](https://github.com/LineageOS/android_frameworks_base/blob/lineage-23.0/cmds/svc/src/com/android/commands/svc/Svc.java)【AOSP】
- shell は `NETWORK_SETTINGS` / `OVERRIDE_WIFI_CONFIG` / `READ_WIFI_CREDENTIAL` / `MANAGE_WIFI_NETWORK_SELECTION` / `CHANGE_WIFI_STATE` を保持 — [Shell manifest](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/packages/Shell/AndroidManifest.xml)【AOSP】
- Advanced Protection 有効時は WEP 接続がブロックされる — [Android Authority（検索スニペット経由）](https://www.androidauthority.com/android-16-advanced-protection-mode-3518368)【報道】

### Inferences
- 冪等性: `cmd wifi list-networks` の出力（Network Id / SSID / Security type）を SSID でパース → 既存なら skip（PSK の一致は読めないので「上書き更新したいなら forget→add」方針を CSV 列で選択）。Verify は `list-networks` 存在確認＋（電波環境があれば）`cmd wifi status` で接続 SSID 確認。
- **EAP-TLS/PEAP が必要な企業 Wi-Fi は ADB tier では不可**。shell は NETWORK_SETTINGS/OVERRIDE_WIFI_CONFIG を持つので、shell-uid の app_process ヘルパーから `WifiManager.addNetwork(WifiConfiguration+WifiEnterpriseConfig)` を呼べば理論上は可能だが、CA/クライアント証明書の受け渡しと Wi-Fi keystore への格納が絡み**未検証・高リスク**。
- `-r` でランダム MAC 方式を指定できるので、MAC 固定運用（社内 MAC 認証）がある場合は `-r none`（＝factory MAC）を CSV 列化できる。
- 導入版: `add-network`/`connect-network` は Android 11 以降との報告が一般的だが本調査で一次確認できず。

### Gaps
- `cmd wifi` で追加したネットワークの「作成者=shell」が他アプリからの管理・ユーザーによる編集に影響するかは未調査。
- 各サブコマンドの導入版（git 履歴）未確認。
- OEM（Samsung 等）の Wi-Fi スタック改変による挙動差は未調査。

---

## Q7. モバイルデータ・機内モード・Bluetooth・NFC・位置情報

### Takeaway
すべて shell のみで切替＋読み返し可能。新しめの端末では `svc wifi` / `svc bluetooth` は各モジュールの `cmd` に移行済みで、`svc` に残るのは power / usb / nfc / system-server。

### Cited Findings
- `cmd connectivity airplane-mode [enable|disable]`、引数なしで現在値（Settings.Global.AIRPLANE_MODE_ON）を `enabled/disabled` で表示 — [ConnectivityService.java (lineage-23.0)](https://github.com/LineageOS/android_packages_modules_Connectivity/blob/lineage-23.0/service/src/com/android/server/ConnectivityService.java)【AOSP】
- `cmd bluetooth_manager enable` / `disable` / `enableBle` / `disableBle` / `wait-for-state:<STATE>`（タイムアウトし得る旨注記）— [BluetoothShellCommand.java](https://github.com/LineageOS/android_packages_modules_Bluetooth/blob/lineage-23.0/service/src/com/android/server/bluetooth/BluetoothShellCommand.java)【AOSP】
- `svc` の登録コマンドは help / power / usb / nfc / system-server。「`svc wifi` は WifiShellCommand へ、`svc bluetooth` は BluetoothShellCommand へ移行」— [Svc.java](https://github.com/LineageOS/android_frameworks_base/blob/lineage-23.0/cmds/svc/src/com/android/commands/svc/Svc.java)【AOSP】
- `svc nfc [enable|disable]` は NfcAdapter.enable()/disable() を呼ぶ — [NfcCommand.java](https://github.com/LineageOS/android_frameworks_base/blob/lineage-23.0/cmds/svc/src/com/android/commands/svc/NfcCommand.java)【AOSP】
- Telephony のシェルに `data enable: enable mobile data connectivity` / `data disable` がある（`cmd phone data enable`）— [TelephonyShellCommand.java (lineage-23.0)](https://github.com/LineageOS/android_packages_services_Telephony/blob/lineage-23.0/src/com/android/phone/TelephonyShellCommand.java)【AOSP】
- `cmd location is-location-enabled [--user]` / `set-location-enabled true|false [--user]` — [LocationShellCommand.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/services/core/java/com/android/server/location/LocationShellCommand.java)【AOSP】
- shell は `MODIFY_PHONE_STATE` / `BLUETOOTH_PRIVILEGED` / `BLUETOOTH_CONNECT` / `NFC` / `CONNECTIVITY_INTERNAL` / `NETWORK_SETTINGS` を保持 — [Shell manifest](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/packages/Shell/AndroidManifest.xml)【AOSP】

### Inferences
- 旧端末互換のため、モジュールは「`cmd …` を試す → 未知コマンドなら `svc …` にフォールバック」の二段構えにする（`svc data enable` は旧来の手段として広く知られる）。
- Verify: 機内モード=`cmd connectivity airplane-mode`、BT=`settings get global bluetooth_on`（1/0）、NFC=`dumpsys nfc | grep mState`（推定）、位置=`cmd location is-location-enabled`、データ=`settings get global mobile_data`（DSDS は `mobile_data<subId>`、推定）。
- **USB 接続を切る操作（`svc usb` の機能切替）は ADB セッション断の原因になるため禁止リスト**に入れる。

### Gaps
- 各 cmd の導入版、`svc data` が Android 15–17 で残存しているかは未確認。
- 日本の SIM（docomo/au/SoftBank/楽天）や eSIM 端末で `cmd phone data` がサブスクリプションごとにどう効くかは未検証。

---

## Q8. アプリ（インストール/削除/無効化/権限/AppOps/既定アプリ/電池最適化/ユーザー補助/通知/デバイス管理者）

### Takeaway
アプリ配布と権限付与は ADB tier の得意領域。`adb install`（-r/-d/-g）・`install-multiple`・`pm install-create/write/commit`・`pm uninstall -k --user 0`（可逆 debloat）・`pm disable-user`・`pm suspend`・`pm grant/revoke`・`appops set`・`cmd role add-role-holder`・`pm set-home-activity`・`dumpsys deviceidle whitelist`・`cmd notification allow_listener`・`dpm set-active-admin`・`ime enable/set` がすべて shell で可能。**`pm hide` は shell 不可**。Android 14/15 で低 targetSdk APK のブロックがあり `--bypass-low-target-sdk-block` で回避。2026-09-30 開始の開発者検証制度でも ADB インストールは免除と Google が説明。

### Cited Findings
- `pm install [-rtfdg] [-i PACKAGE] [--user USER_ID|all|current] ...`: `-d` 「バージョンコードのダウングレードを許可（debuggable パッケージのみ）」、`-g` 「全ランタイム権限を付与」、`--install-reason 0..4`、`--restrict-permissions`、`--dont-kill` 等。`install-create/install-write/install-commit/install-abandon`、`install-existing [--user]`、`uninstall [-k] [--user USER_ID] [--versionCode]`（-k でデータ/キャッシュ保持）— [PackageManagerShellCommand.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/services/core/java/com/android/server/pm/PackageManagerShellCommand.java)【AOSP】
- `--bypass-low-target-sdk-block` は install オプション解析に存在（help には非掲載＝隠しオプション）— [PackageManagerShellCommand.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/services/core/java/com/android/server/pm/PackageManagerShellCommand.java)【AOSP】; Android 14 はハードコード値（既定 SDK 23）未満の targetSdk のインストールをブロックし `INSTALL_FAILED_DEPRECATED_SDK_VERSION`、回避は ADB の同フラグのみ — [Liliputing](https://liliputing.com/how-to-install-old-apps-on-android-14/); [F-Droid forum](https://forum.f-droid.org/t/force-install-old-targetsdk-apps-using-adb-root/28470)【コミュ】
- Android 15: targetSdkVersion 24 未満はインストール不可（`INSTALL_FAILED_DEPRECATED_SDK_VERSION: App package must target at least SDK version 24`）。アップグレード端末の既存アプリは残る。テスト用に `adb install --bypass-low-target-sdk-block FILENAME.apk` — [Android 15 behavior changes: all apps](https://developer.android.com/about/versions/15/behavior-changes-all)【公式】
- Android 16 の all-apps 変更ページには最小 targetSdk の追加引き上げ記述なし（16 KB ページ互換モード `android:pageSizeCompat` 等）— [Android 16 behavior changes: all apps](https://developer.android.com/about/versions/16/behavior-changes-all)【公式】
- Android 17 は 2026-06-16 に安定版リリース（June 2026 Pixel Drop と同時、Pixel 6 以降が対象、非 Pixel は OEM ごとに順次）。QPR2 は 2026-12 見込み（予測）— [Android Authority（検索スニペット経由）](https://www.androidauthority.com/android-17-release-date-3639790/); [Gizmochina](https://www.gizmochina.com/2026/06/09/heres-when-google-pixels-will-receive-android-17-update/)【報道】
- Android 17 の all-apps 変更: RAM ベースのアプリメモリ上限、SMS OTP 保護拡大、cross-profile loopback 遮断、バックグラウンド音声制限、keystore 鍵数上限等。adb/USB デバッグ/開発者オプション/Settings/インストール/ユーザー補助/ロケール/Wi-Fi に関する変更記載なし。テスト用 `am memory-limiter` コマンド追加 — [Android 17 behavior changes: all apps](https://developer.android.com/about/versions/17/behavior-changes-all)【公式】
- `pm hide/unhide` は存在するが、実装 `setApplicationHiddenSettingAsUser` に `@EnforcePermission(MANAGE_USERS)` — shell は MANAGE_USERS 非保持 — [PackageManagerService.java](https://github.com/LineageOS/android_frameworks_base/blob/lineage-23.0/services/core/java/com/android/server/pm/PackageManagerService.java); [Shell manifest](https://github.com/LineageOS/android_frameworks_base/blob/lineage-23.0/packages/Shell/AndroidManifest.xml)【AOSP】
- 同様に `pm set-user-restriction` の実装 `setUserRestriction` は `checkManageUsersPermission` で「You need MANAGE_USERS permission」— [UserManagerService.java](https://github.com/LineageOS/android_frameworks_base/blob/lineage-23.0/services/core/java/com/android/server/pm/UserManagerService.java)【AOSP】
- `pm suspend/unsuspend [--user] PKG...`、`disable-user`、`disable-until-used`、`default-state`、`grant/revoke [--user] [--all-permissions]`、`set-permission-flags`、`set-home-activity [--user] TARGET`（パッケージ名だけが意味を持つ）、`archive [--user] PKG`（APK とキャッシュを削除しユーザーデータ保持）/`request-unarchive` — [PackageManagerShellCommand.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/services/core/java/com/android/server/pm/PackageManagerShellCommand.java)【AOSP】
- `cmd role get-role-holders / add-role-holder [--user] ROLE PACKAGE [FLAGS] / remove-role-holder / clear-role-holders / set-bypassing-role-qualification` — [RoleShellCommand.java](https://github.com/LineageOS/android_packages_modules_Permission/blob/lineage-23.0/service/java/com/android/role/RoleShellCommand.java)【AOSP】
- `dumpsys deviceidle whitelist [package ...]`、`except-idle-whitelist`、`tempwhitelist` — [DeviceIdleController.java](https://github.com/LineageOS/android_frameworks_base/blob/lineage-23.0/apex/jobscheduler/service/java/com/android/server/DeviceIdleController.java)【AOSP】
- `cmd notification allow_listener COMPONENT [user_id]` / `disallow_listener` / `allow_assistant` / `set_dnd [on|priority|alarms|all|off]` — [NotificationShellCmd.java](https://github.com/LineageOS/android_frameworks_base/blob/lineage-23.0/services/core/java/com/android/server/notification/NotificationShellCmd.java)【AOSP】
- `dpm set-active-admin [--user <USER_ID>|current] <COMPONENT>`、`set-device-owner`、`set-profile-owner`、`remove-active-admin`（testOnly 宣言の admin のみ）— [DevicePolicyManagerServiceShellCommand.java](https://github.com/LineageOS/android_frameworks_base/blob/lineage-23.0/services/devicepolicy/java/com/android/server/devicepolicy/DevicePolicyManagerServiceShellCommand.java)【AOSP】
- IMMS シェル `ime list / enable / set / reset` — [InputMethodManagerService.java](https://github.com/LineageOS/android_frameworks_base/blob/lineage-23.0/services/core/java/com/android/server/inputmethod/InputMethodManagerService.java)【AOSP】
- shell は `INSTALL_PACKAGES` / `DELETE_PACKAGES` / `GRANT_RUNTIME_PERMISSIONS` / `MANAGE_APP_OPS_MODES` / `UPDATE_APP_OPS_STATS` / `SUSPEND_APPS` / `MANAGE_ROLE_HOLDERS` / `MANAGE_DEVICE_ADMINS` / `MANAGE_NOTIFICATION_LISTENERS` / `DEVICE_POWER` / `CHANGE_COMPONENT_ENABLED_STATE` を保持 — [Shell manifest](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/packages/Shell/AndroidManifest.xml)【AOSP】
- 公式 adb ドキュメント: `install-multiple` は Play Console から取得した端末向け複数 APK のインストール用。`pm install` オプション `-r -t -d -g --fastdeploy --incremental` — [Android Debug Bridge](https://developer.android.com/tools/adb)【公式】
- 開発者検証: Google のヘルプは「開発者・パワーユーザーは引き続き ADB で未検証アプリを自分の端末にビルド/テスト/インストールできる」と説明。2026-09-30 からブラジル・インドネシア・シンガポール・タイで施行、認定 Android 端末で登録済み開発者のアプリが必要に — [Understanding Android developer verification (Google)](https://support.google.com/android-developer-console/answer/16561738); [Foresight Mobile](https://foresightmobile.com/blog/android-developer-verification-2026)【公式/コミュ】
- Xperia 1 IV/V（docomo）で `adb shell pm uninstall -k --user 0 <pkg>` によりキャリア初期アプリを削除、`cmd package install-existing <pkg>` で復元できるとの日本語事例 — [kunkoku.jp](https://kunkoku.jp/xperia1vso51d-2.html)【コミュ】
- UAD-ng は不要システムアプリ除去ツールで「自己責任」の警告付き、Wiki に「OEM の奇妙な挙動」ページ — [UAD-ng README](https://github.com/Universal-Debloater-Alliance/universal-android-debloater-next-generation)【コミュ/OSS】
- Android 13 の「制限付き設定」によりサイドロードアプリのユーザー補助/デバイス管理者は手動許可が必要 — [Kaspersky](https://kaspersky.com/blog/android-restricted-settings/49991)【ベンダー】

### Inferences
- **debloat は `pm uninstall -k --user 0` より `pm disable-user --user 0` を既定**にし（可逆・OTA 後も一貫）、削除はリスト管理された対象のみ。戻しは `cmd package install-existing` / `pm enable`。重要システムアプリ（SUW、Phone、SystemUI、OEM のキャリア設定アプリ）は禁止リスト化。
- `-d` ダウングレードは help 上「debuggable のみ」なので本番 APK のロールバックには使えない前提。
- **既定アプリ**: Android 10+ は `cmd role add-role-holder android.app.role.{BROWSER,SMS,DIALER,HOME}` が正規。Verify は `get-role-holders`。対象アプリが role 要件（intent-filter 等）を満たさないと失敗する。
- **権限**: `pm grant` は runtime 権限と development 権限のみ。`SYSTEM_ALERT_WINDOW`/`MANAGE_EXTERNAL_STORAGE`/`REQUEST_INSTALL_PACKAGES`/`WRITE_SETTINGS`/`PACKAGE_USAGE_STATS` は appop 系なので `appops set <pkg> <OP> allow`。Verify は `appops get` と `dumpsys package <pkg>`。
- **ユーザー補助/通知リスナー/デバイス管理者** を ADB から有効化した場合に Android 13+ の制限付き設定が阻害しないかは要実機検証（shell 権限上は可能）。`adb install` 経由のインストールが「セッション型インストーラ扱い」になるかで結果が変わる可能性（推論）。
- 開発者検証は日本は初期施行国に含まれず、かつ ADB インストールは免除と説明されているため、2026-10 時点では ADB 配布への影響は小さいと推定。ただし将来の運用変更に注意。

### Gaps
- `pm suspend` を shell が行った場合の「suspending package=com.android.shell」の UI 表示・解除挙動は未検証。
- Samsung / 日本メーカー固有の電池最適化（Samsung「スリープ中のアプリ」等）を ADB で制御する方法は未調査。
- 2026-09-30 施行後に ADB 免除が実際に維持されているかの一次確認は未取得（検索結果は 2026 年中盤の記事中心）。

---

## Q9. 画面ロック（`locksettings`）

### Takeaway
`locksettings set-pin / set-password / set-pattern / clear / verify / get-disabled / set-disabled` が shell で使える（`SET_AND_VERIFY_LOCKSCREEN_CREDENTIALS` 保持）。資格情報が既にある場合は `--old <CREDENTIAL>` 必須、管理ポリシー（複雑度）を満たさない新資格情報は拒否。Verify は `locksettings verify --old <PIN>`（成功/失敗）と `get-disabled`（None かどうか）。

### Cited Findings
- ヘルプ: 「secure なロック画面が設定済みの場合、ほとんどのコマンドに `--old <CREDENTIAL>` が必要」。`get-disabled`（None なら true）、`set-disabled <true|false>`（secure 時は即時効果なし＝Swipe/None の切替のみ）、`set-pattern [--old] [--user] <PATTERN>`（1-9 の数字列）、`set-pin`、`set-password`、`clear [--old]`、`verify [--old]`、`require-strong-auth` — [LockSettingsShellCommand.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/services/core/java/com/android/server/locksettings/LockSettingsShellCommand.java)【AOSP】
- `hasSecureLockScreen()` が false の端末では help/get-disabled/set-disabled 以外は「The device does not support lock screen - ignoring the command.」で -1 — [LockSettingsShellCommand.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/services/core/java/com/android/server/locksettings/LockSettingsShellCommand.java)【AOSP】
- `checkCredential()`: secure かつ `--old` 未指定なら「User has a lock credential, but old credential was not provided」で失敗。統合チャレンジの管理プロファイルは「Profile uses unified challenge」で拒否。`isNewCredentialSufficient()` は `getRequestedPasswordMetrics/Complexity` に照らし「New credential doesn't satisfy admin policies」で拒否 — [LockSettingsShellCommand.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/services/core/java/com/android/server/locksettings/LockSettingsShellCommand.java)【AOSP】
- shell は `SET_AND_VERIFY_LOCKSCREEN_CREDENTIALS` を保持 — [Shell manifest](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/packages/Shell/AndroidManifest.xml)【AOSP】

### Inferences
- **testharness 初期化と相性が悪い**（ロック画面があると testharness 不可）ため、キッティング順序は「各種設定 → 最後に PIN 設定」。再キッティング時は先に `locksettings clear --old <PIN>`。
- PIN を CSV に平文で持つのはセキュリティ上問題。fabriq の既存方針（秘密情報の扱い）に合わせ、端末ごと PIN の受け渡し・ログマスクが必要。
- `locksettings` の終了コードは成功 0 / 失敗 -1（255 として見える可能性）なので、出力文字列（"Lock credential cleared" 等）と併用判定。
- 指紋/顔認証の登録は ADB 不可（UI 必須、推論）。

### Gaps
- `locksettings` の導入版（Android 8.0 と推定）と、Samsung One UI / 日本メーカーでの動作差（独自ロック画面の互換性）は未確認。
- `dumpsys lock_settings` で資格情報種別（PIN/パターン/パスワード）を読み取れるかは未確認。

---

## Q10. 証明書（CA / PKCS#12）

### Takeaway
**Android 11 以降、CA 証明書は「設定」からしかインストールできない**（アプリの `KeyChain.createInstallIntent()` も CA には使えない）。shell も `MANAGE_CA_CERTIFICATES`（signature|privileged）を持たないため **ADB でのサイレント導入は不可**。最小タップ手順は「`adb push` で Download に置く → 設定の該当画面を `am start` で開く → ファイル選択」を UI 自動化で補助する形。

### Cited Findings
- KeyChain の Javadoc: 「Android R（11）以降、このメソッドが返す Intent は CA 証明書のインストールには使えない。CA 証明書は Settings からのみインストールできるため、アプリは CA 証明書ファイルをユーザーに提供すべき（例: MediaStore で Downloads に書き出す）」。`ACTION_INSTALL = "android.credentials.INSTALL"`、`com.android.certinstaller.CertInstallerMain` を起動 — [KeyChain.java](https://github.com/LineageOS/android_frameworks_base/blob/lineage-23.0/keystore/java/android/security/KeyChain.java)【AOSP】
- `MANAGE_CA_CERTIFICATES` は signature|privileged、shell マニフェストに無し — [core AndroidManifest.xml](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/core/res/AndroidManifest.xml); [Shell manifest](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/packages/Shell/AndroidManifest.xml)【AOSP】

### Inferences
- 手順案: `adb push corp-ca.crt /sdcard/Download/` → `adb shell am start -a android.settings.SECURITY_SETTINGS`（「暗号化と認証情報」への到達は OEM 差）→ 「証明書のインストール > CA 証明書」→ 警告確認 → ファイル選択。uiautomator で半自動化可能だが**ロケール/OEM 依存**。
- PKCS#12（クライアント証明書）は `android.credentials.INSTALL` Intent に byte[] extra を渡す必要があり `am start` では渡せないため、結局ファイル選択 UI になる（推論）。ヘルパー APK から `KeyChain.createInstallIntent()` に PKCS#12 を渡せば「名前入力＋確認」程度まで短縮できる可能性。
- Verify: ユーザー CA は `/data/misc/user/0/cacerts-added/` にあるが shell から読めるかは未確認。UI（「ユーザー」タブ）のスクリーンショットをエビデンスとする方式が現実的。
- 証明書が必須の環境（EAP-TLS Wi-Fi、社内 TLS インスペクション）は **ADB tier の対象外＝MDM が必要**と明示すべき。

### Gaps
- Android 14–17 で CA 導入画面へ直接飛べる公開 Settings アクションの有無は未確認。
- OEM ごとの「暗号化と認証情報」メニュー位置の差は未調査。

---

## Q11. APN / VPN

### Takeaway
**APN は Android 14 以降なら shell で `content insert --uri content://telephony/carriers` できる見込み**（shell に `WRITE_APN_SETTINGS` が付与されたのは A14 から。A13 以前は SecurityException）。**VPN プロファイル作成は不可**（`CONTROL_VPN` 非保持）。

### Cited Findings
- TelephonyProvider の `checkPermissionForApnTable()`: `WRITE_APN_SETTINGS` を持てば許可、無ければキャリア特権を確認し、どちらも無ければ「No permission to access APN settings」。insert（`insertSingleRow`）は `checkPermission(match)` を通る — [TelephonyProvider.java](https://github.com/aosp-mirror/platform_packages_providers_TelephonyProvider/blob/main/src/com/android/providers/telephony/TelephonyProvider.java)【AOSP】
- Shell マニフェストの `WRITE_APN_SETTINGS` は「Permission required for CTS test - CtsTelephonyProviderTestCases」コメント付きで、A13 r1 には無く A14 r1 以降に存在 — [Shell manifest A13](https://github.com/aosp-mirror/platform_frameworks_base/blob/android-13.0.0_r1/packages/Shell/AndroidManifest.xml); [A14](https://github.com/aosp-mirror/platform_frameworks_base/blob/android-14.0.0_r1/packages/Shell/AndroidManifest.xml); [main](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/packages/Shell/AndroidManifest.xml)【AOSP】
- TelephonyProvider には DPC 管理 APN の強制（`carriers/enforce_managed`、「If enforced, return DPC records only」）がある — [TelephonyProvider.java](https://github.com/aosp-mirror/platform_packages_providers_TelephonyProvider/blob/main/src/com/android/providers/telephony/TelephonyProvider.java)【AOSP】
- `CONTROL_VPN` は signature|privileged、shell マニフェストに無し — [core AndroidManifest.xml](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/core/res/AndroidManifest.xml)【AOSP】
- Vpn.java は always-on VPN の設定を `Settings.Secure.ALWAYS_ON_VPN_APP` / `ALWAYS_ON_VPN_LOCKDOWN` に保存し、そこから読み出す — [Vpn.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/services/core/java/com/android/server/connectivity/Vpn.java)【AOSP】

### Inferences
- APN 例（未検証）: `content insert --uri content://telephony/carriers --bind name:s:CorpAPN --bind apn:s:corp.example --bind mcc:s:440 --bind mnc:s:10 --bind numeric:s:44010 --bind type:s:default,supl --bind protocol:s:IPV4V6`。Verify: `content query --uri content://telephony/carriers --where "apn='corp.example'"`。既定 APN 選択（preferapn）の設定方法は未検証。
- Shell の WRITE_APN_SETTINGS は CTS 目的で付与されたもので、**OEM が削っていない保証はない**。日本のキャリア端末はキャリア設定で APN 編集を制限している場合があり実機確認必須。
- always-on VPN は `settings put secure always_on_vpn_app <pkg>` で値自体は書けるが、VPN アプリの同意（VpnService.prepare）や読込タイミング（再起動要否）が絡むため**非推奨・未検証**。VPN は各 VPN アプリ側のインポート機能（設定ファイル push＋Intent）に依存する設計に。

### Gaps
- Android 14–17 実機で shell からの APN insert 成功報告は見つからず（一次の動作確認なし）。
- 日本 MVNO/キャリア端末での APN 編集ロックの挙動不明。

---

## Q12. 壁紙・ホーム画面レイアウト・着信音

### Takeaway
**壁紙を設定する shell コマンドは無い**（`cmd wallpaper` は dim 量の操作のみ）。`SET_WALLPAPER` は normal 権限なので小さなヘルパー APK で設定可能、または shell-uid app_process ヘルパー。**ホーム画面レイアウトはランチャー固有で ADB 標準手段なし**。着信音は `settings put system ringtone <content-uri>` で理論上可能だが未検証。

### Cited Findings
- `cmd wallpaper` のサブコマンドは `set-dim-amount`、`dim-with-uid`、`get-dim-amount`、help のみ — [WallpaperManagerShellCommand.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/services/core/java/com/android/server/wallpaper/WallpaperManagerShellCommand.java)【AOSP】
- `SET_WALLPAPER` の保護レベルは normal、shell も `SET_WALLPAPER` / `SET_WALLPAPER_COMPONENT` を保持 — [core AndroidManifest.xml](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/core/res/AndroidManifest.xml); [Shell manifest](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/packages/Shell/AndroidManifest.xml)【AOSP】

### Inferences
- 壁紙: ヘルパー APK（`WallpaperManager.setStream()`、`FLAG_SYSTEM|FLAG_LOCK`）を `am start`/`am broadcast` で起動し画像パスを extra で渡す方式が単純。Verify は `dumpsys wallpaper`（ファイルサイズ/コンポーネント）程度で、画像同一性の厳密検証は難しい→スクリーンショットをエビデンスに。
- 着信音: 音源を push → MediaStore 登録（`am broadcast` の MEDIA_SCANNER は廃止済のため `content insert` か UI）→ `settings put system ringtone content://media/...`。手数が多く OEM 差も大きいので優先度低。
- ランチャーレイアウト: AOSP Launcher3 / Pixel Launcher / One UI Home などで保存形式が異なり、ADB からの公式インポート手段は無い。MDM 無しなら「ウィジェット/アイコン配置は手作業 or 対象外」と割り切る。

### Gaps
- Samsung One UI Home / 日本メーカー独自ホームのレイアウト事前定義手段は未調査。

---

## Q13. 棚卸し・エビデンス（シリアル/IMEI/MAC/電池/パッケージ/スクショ/UI ダンプ）

### Takeaway
シリアル・機種・OS 版・セキュリティパッチ・ビルド指紋は `getprop` で取得可能。**IMEI/電話番号は shell が `READ_PRIVILEGED_PHONE_STATE` / `READ_PHONE_NUMBERS` を持つため権限的には読めるが、`service call iphonesubinfo` のトランザクション番号は版/OEM 依存で不安定**。エビデンスは `screencap`（PNG）・`screenrecord`（最大 180 秒・音声なし）・`uiautomator dump`。

### Cited Findings
- shell は `READ_PRIVILEGED_PHONE_STATE` / `READ_PRECISE_PHONE_STATE` / `READ_PHONE_NUMBERS` / `LOCAL_MAC_ADDRESS` / `DUMP` を保持（A13〜A16 で READ_PRIVILEGED_PHONE_STATE・LOCAL_MAC_ADDRESS 有）— [Shell manifest](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/packages/Shell/AndroidManifest.xml); [A13](https://github.com/aosp-mirror/platform_frameworks_base/blob/android-13.0.0_r1/packages/Shell/AndroidManifest.xml)【AOSP】
- `READ_PRIVILEGED_PHONE_STATE` は signature|privileged|role（一般アプリには付与不可）— [core AndroidManifest.xml](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/core/res/AndroidManifest.xml)【AOSP】
- `service call iphonesubinfo` のインデックスは端末により 14/15 等で異なり、Android 11+ ではパッケージ名を引数に渡す形の報告。シェルから権限エラーになった報告や、root 必須との古い回答もあり — [Reddit r/AutomateUser（ミラー）](https://redlib.ssps.io/r/AutomateUser/comments/1in8mc8/get_self_phone_number); [SO ミラー: IMEI via adb](https://windows-hexerror.linestarve.com/q/so49644219-how-to-get-imei-number-from-adb-command-of-higer-level-api-devices)【コミュ・低信頼】
- `screencap`: `adb exec-out screencap -p > screen.png` で標準出力へ。`screenrecord`: 既定かつ最大 180 秒（`--time-limit`）、音声は記録されない、録画中の回転は非対応 — [Android Debug Bridge](https://developer.android.com/tools/adb)【公式】
- Android 15 の画面共有保護（通知内容・パスワード入力の秘匿、OTP を含む通知から起動した Activity の秘匿）。既定のシステム画面録画は対象外 — [Android 15 behavior changes: all apps](https://developer.android.com/about/versions/15/behavior-changes-all)【公式】
- `pm list packages` は `--show-versioncode` / `--apex-only` 等のオプションを持つ — [PackageManagerShellCommand.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/services/core/java/com/android/server/pm/PackageManagerShellCommand.java)【AOSP】
- `cmd wifi add-network/connect-network` の `-r auto|none|persistent|non_persistent` で MAC ランダム化方式を指定できる — [WifiShellCommand.java](https://github.com/LineageOS/android_packages_modules_Wifi/blob/lineage-23.0/service/java/com/android/server/wifi/WifiShellCommand.java)【AOSP】

### Inferences
- 推奨インベントリ取得（全て Verify-only モジュール）: `adb get-serialno`、`getprop ro.serialno`、`ro.product.manufacturer`/`ro.product.model`/`ro.product.device`、`ro.build.version.release`/`ro.build.version.sdk`、`ro.build.version.security_patch`、`ro.build.fingerprint`、`pm list packages --show-versioncode`、`dumpsys battery`（level/health/temperature）、`cat /proc/meminfo`、`df /data`。
- **IMEI は shell-uid の app_process ヘルパーから `TelephonyManager.getImei(slot)` を呼ぶのが最も堅牢**（shell が READ_PRIVILEGED_PHONE_STATE を持つため。scrcpy と同じ起動方式）— 未検証だが `service call` の番号依存を避けられる。代替は `*#06#` を `am start -a android.intent.action.DIAL -d tel:*%2306%23` で表示しスクショ。
- Wi-Fi MAC は接続ネットワークごとにランダム化されるため、資産台帳用の「factory MAC」と接続時 MAC を区別して記録する必要。`cmd wifi status` / `ip addr show wlan0` は現在 MAC（ランダム化後）を返すと推定。
- FLAG_SECURE ウィンドウ（銀行アプリ等・一部設定画面）は `screencap` で黒く写る（一般に知られる挙動、未検証）。

### Gaps
- Android 14+ のバッテリー健全性（サイクル数/SoH）を `dumpsys battery` で読めるかは未確認（BATTERY_STATS 系の公開状況要調査）。
- Android 13–17 の各版での `iphonesubinfo` トランザクション番号表は信頼できる一次資料なし。

---

## Q14. コマンドが無い設定の UI 自動化フォールバック

### Takeaway
`input tap/swipe/text/keyevent`、`am start -a android.settings.*` による設定画面の直接起動、`uiautomator dump` による UI ツリー取得で「半自動化」は可能。ホスト側ライブラリは **uiautomator2（Python, 端末側は jar を push する方式）** と **Appium UiAutomator2 ドライバ（端末に Server APK と io.appium.settings を導入）**。ロケール・OEM・版依存が強く、FLAG_SECURE 画面はスクショ不可。

### Cited Findings
- uiautomator2: 端末側で UiAutomator ベースの HTTP サービスを動かし、Python クライアントが HTTP で操作。端末側は jar（android-uiautomator-server-jar）で動作、Android 4.4+ / Python 3.8+ — [openatx/uiautomator2 README](https://github.com/openatx/uiautomator2)【OSS】
- Appium UiAutomator2 ドライバ: 端末に UiAutomator2 Server をインストール（`skipServerInstallation` 等の capability）、通知取得等で Appium Settings ヘルパー（`io.appium.settings`）を使用し、通知アクセスは手動許可が必要。`autoGrantPermissions` は targetSdk 23+ 前提 — [appium-uiautomator2-driver README](https://github.com/appium/appium-uiautomator2-driver)【OSS】
- scrcpy は shell として `app_process` で jar を実行（画面キャプチャには shell の特権が必要）— [scrcpy doc/develop.md](https://github.com/Genymobile/scrcpy/blob/master/doc/develop.md)【OSS】
- `am start -a android.intent.action.VIEW` 等、am で Activity 起動・ブロードキャスト等が可能 — [Android Debug Bridge](https://developer.android.com/tools/adb)【公式】
- Android 16 の Intent リダイレクト防御強化 — [Android 16 behavior changes: all apps](https://developer.android.com/about/versions/16/behavior-changes-all)【公式】

### Inferences
- fabriq 流にするなら「UI 自動化モジュール」は (1) `am start` で目的画面へ直行 → (2) `uiautomator dump /sdcard/x.xml` → `adb pull` → PowerShell で XML を XPath 検索（`resource-id` 優先、`text` は ja/en 両辞書）→ (3) `input tap x y` → (4) 再 dump で状態確認、という純 ADB 実装で APK 不要にできる（uiautomator コマンドは標準搭載）。
- 外部ライブラリを使う場合、uiautomator2 は Python ランタイム、Appium は Node.js＋Java＋Server APK が必要で、PowerShell 5.1 中心の fabriq とは依存が増える。**まず純 ADB（input + uiautomator dump）で足りるかを PoC すべき**。
- 依存: ロケール（文言）、OEM（Settings の構成・resource-id）、Android 版、画面サイズ/回転、ダークモード。CSV に「OEM×版ごとのセレクタ」を持つ設計が必要で保守コストが高い→**UI 自動化は最後の手段**と位置付ける。

### Gaps
- 日本メーカー端末の Settings resource-id の安定性は未調査。
- `uiautomator dump` が Android 16/17 で遅延・失敗する既知問題の有無は未調査。

---

## Q15. OS 更新（OTA）・Google Play システムアップデート

### Takeaway
**root なしで ADB から OTA の実行を指示する手段は見当たらない**（設定アプリの UI 操作が必要）。`adb sideload` はリカバリ（ADB sideload モード）で OEM 署名済み OTA パッケージに限られる。検証（パッチレベル・ビルド指紋の読み取り）は ADB で可能。

### Cited Findings
- テストハーネスモードは OTA 自動更新を無効化する（`OTA_DISABLE_AUTOMATIC_UPDATE=1`）— [TestHarnessModeService.java](https://github.com/LineageOS/android_frameworks_base/blob/lineage-23.0/services/core/java/com/android/server/testharness/TestHarnessModeService.java)【AOSP】
- shell は `REBOOT` を保持（リカバリへの再起動 `adb reboot recovery`/`sideload` 自体は可能）— [Shell manifest](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/packages/Shell/AndroidManifest.xml)【AOSP】

### Inferences
- 出荷前の OS 最新化は「Settings > システム > ソフトウェアアップデート」を `am start` で開いて手動実行（または UI 自動化）し、ADB では `getprop ro.build.version.security_patch` と `ro.build.fingerprint` を**期待値（CSV）と照合する Verify-only モジュール**とするのが現実的。
- Pixel は Google 公式に「OTA イメージ（adb sideload 用）」を配布しており、リカバリでの `adb sideload` が可能（一般知識・本調査で未取得）。Samsung/日本メーカーは原則不可（OEM 署名パッケージ非公開）。
- Google Play システムアップデート（Mainline）のバージョンは `pm list packages --apex-only --show-versioncode` 等で棚卸し可能と推定。実行は UI。

### Gaps
- `cmd` 経由で SystemUpdate/GMS のアップデートチェックを起動する非公開手段の有無は未調査（root 不要なものがあるか不明）。
- Pixel OTA sideload の公式ページ（developers.google.com/android/ota）は今回取得していない。

---

## Q16. 初期化（再キッティング）と FRP

### Takeaway
通説「ADB では初期化不可」は**不正確**。`am broadcast -a android.intent.action.FACTORY_RESET` は shell が `MASTER_CLEAR` を持たないため不可だが、**`cmd testharness enable`（Android 10+）は shell で全データ消去を実行できる**（ADB 鍵保持・SUW スキップ付き）。ただし画面ロック設定済みだと拒否、結果はテストハーネスモード端末になる。

### Cited Findings
- `MASTER_CLEAR` = signature|privileged|role、shell マニフェストに無し — [core AndroidManifest.xml](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/core/res/AndroidManifest.xml); [Shell manifest](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/packages/Shell/AndroidManifest.xml)【AOSP】
- testharness は `ENABLE_TEST_HARNESS_MODE` を確認し、ロック画面があれば拒否、ADB 鍵を PersistentDataBlock に保存して `ACTION_FACTORY_RESET`（EXTRA_WIPE_EXTERNAL_STORAGE=true）を送る — [TestHarnessModeService.java](https://github.com/LineageOS/android_frameworks_base/blob/lineage-23.0/services/core/java/com/android/server/testharness/TestHarnessModeService.java)【AOSP】
- 公式: Android 10+ のテスト端末は `testharness` でファクトリーリセット可能 — [Android Debug Bridge](https://developer.android.com/tools/adb)【公式】

### Inferences
- 再キッティングの手順案: `locksettings clear --old <PIN>` → Google アカウント削除（UI、FRP 回避のため推奨）→ `cmd testharness enable` → `adb wait-for-device` → `getprop sys.boot_completed` → 即モジュール投入。
- FRP: OS 内部からの初期化（設定画面からの初期化と同系統）なので FRP は通常掛からないと推定。それでも**Google アカウントは事前に削除**しておくのが安全（未検証）。
- testharness を使わない通常初期化（SUW 完走が必要）とのトレードオフを、fabriq の profile で「再キッティング方式」として選択可能にする設計が妥当。
- `PersistentDataBlock`（FRP パーティション）実装が無い/改変された OEM 端末では testharness が「Failed to enable Test Harness Mode」になりうる（ソース上の失敗経路から推論）。

### Gaps
- Samsung / 日本メーカーで testharness が使えるかの報告なし。
- Android 15 で強化された FRP と testharness の相互作用の一次資料なし。

---

## Q17. 再起動・起動完了検出（restart-resume 相当）

### Takeaway
`adb reboot` → `adb wait-for-device` → `getprop sys.boot_completed` が `1` になるまでポーリング、が定石。fabriq の `__RESTART__` 相当は「ホスト側で状態 JSON に次ステップを記録 → 再起動 → 起動完了待ち → 再開」で実装できる（端末側の常駐不要）。

### Cited Findings
- shell は `REBOOT` を保持 — [Shell manifest](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/packages/Shell/AndroidManifest.xml)【AOSP】
- testharness は初回起動の早い段階（システム起動完了前）で provisioned 設定を行い、起動完了後に ADB 鍵復元・設定変更を行う（＝起動完了イベント後でないと揃わない設定がある）— [TestHarnessModeService.java](https://github.com/LineageOS/android_frameworks_base/blob/lineage-23.0/services/core/java/com/android/server/testharness/TestHarnessModeService.java)【AOSP】
- Android 15: アプリが停止状態になると pending intent がキャンセルされウィジェット無効化、ユーザー操作で停止状態を解除すると `ACTION_BOOT_COMPLETED` が配信される — [Android 15 behavior changes: all apps](https://developer.android.com/about/versions/15/behavior-changes-all)【公式】

### Inferences
- `sys.boot_completed=1` 後も PackageManager/Settings が安定するまで数秒〜数十秒かかることがあるため、追加で `pm path android` 成功や `cmd package list packages` 応答をヘルスチェックにする。
- PIN 設定済み端末は再起動後「初回ロック解除前（BFU）」状態で、認証情報暗号化ストレージ（CE）が未解錠。ADB 自体は動くがアプリデータ依存の検証は失敗しうる→ PIN はキッティング最後に設定する理由がここにもある（推論）。
- 並列多台数では `adb -s <serial> wait-for-device` を端末ごとに走らせ、タイムアウトを設けて「再起動後に戻らない端末」を Error 状態に落とす（fabriq の状態機械に `RebootTimeout` 遷移を追加）。

### Gaps
- OEM ごとの起動完了後の「初回起動時処理」（Samsung のアプリ最適化、キャリアアプリ更新）が ADB 操作と競合する事例は未調査。

---

## Q18. 終盤ハードニング（USB デバッグ/開発者オプション無効化）

### Takeaway
`settings put global development_settings_enabled 0` と `settings put global adb_enabled 0` は shell で書けるが、**`adb_enabled 0` は adbd を止めるため実行した瞬間にセッションが切れる**（推定）。よって「全 Verify・エビデンス取得・ログ回収を完了 → `adb_allowed_connection_time` を既定に戻す → 開発者オプション OFF → 最後に adb_enabled 0」の順にし、**adb_enabled の読み返しは不可能＝実行の終了をもって成功とみなす（-Verified=$null 相当）**。

### Cited Findings
- `DISALLOW_DEBUGGING_FEATURES` ユーザー制限の変化時に SettingsProvider が `ADB_ENABLED` を更新する（ユーザー制限で ADB 設定が強制されうる）— [SettingsProvider.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/packages/SettingsProvider/src/com/android/providers/settings/SettingsProvider.java)【AOSP】
- `ADB_ALLOWED_CONNECTION_TIME` の既定は 7 日、0 は無期限 — [Settings.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/core/java/android/provider/Settings.java); [AdbDebuggingManager.java](https://github.com/aosp-mirror/platform_frameworks_base/blob/main/services/core/java/com/android/server/adb/AdbDebuggingManager.java)【AOSP】
- 公式 adb ドキュメントは「adb デバッグの承認を取り消す」で全ペアリング済みワークステーションを解除できると記載 — [Android Debug Bridge](https://developer.android.com/tools/adb)【公式】

### Inferences
- 推奨順序: ① 全モジュールの Verify・スクショ・`uiautomator dump`・ログを pull → ② `settings delete global adb_allowed_connection_time`（既定 7 日へ）→ ③ `settings put global development_settings_enabled 0` → ④ `settings put global adb_enabled 0`（ここで切断）。
- 設定アプリから開発者オプションを OFF にした場合は各種開発者設定もリセットされるが、`settings put` 直書きではリセット処理が走らない可能性（未検証）。厳密にやるなら最後に UI で OFF。
- 認可済み鍵（`/data/misc/adb/adb_keys`）は shell から削除できないため、鍵を残したくない場合は端末 UI の「USB デバッグの承認を取り消す」を手動実行（または adb_enabled 0 で実質無効化）。
- 一度 adb_enabled 0 にすると以後のリモート再操作は再度の手動有効化が必要＝「出荷後のメンテ用に ADB を残すか」は運用判断として profile で選択可能に。

### Gaps
- `adb_enabled 0` 実行時に adb がリモート終了コードを返す前に切断されるか（ホスト側の戻り値の扱い）は未検証。
- 一部 OEM が開発者オプション OFF 時に USB デバッグ鍵も自動消去するかは未確認。
