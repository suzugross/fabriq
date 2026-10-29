# Practical, operational and legal constraints of a host-PC (Windows + ADB over USB) Android bulk-kitting tool — ecosystem & market practice (as of 2026-10-29, Japan focus)

> Method note (read first): research date 2026-10-29. In this research session many hosts were **blocked for direct page fetch** (www.fcnt.com, developer.samsung.com, docs.cambrionix.com, discuss.appium.io, bayton.org, support.google.com, learn.microsoft.com, android-developers.googleblog.com, www.docomo.ne.jp, pc.watch.impress.co.jp, houjin.bellpark.co.jp, biz.techvan.co.jp; GitHub REST API also gated). Pages on developer.android.com and github.com were read directly. Findings from blocked hosts come from **search-engine snippets/summaries** and are marked "(snippet)". Treat those as lower confidence than directly read pages.
> Labels used: **[Official]** = Google/Microsoft/OEM primary doc; **[Vendor]** = vendor marketing / product page; **[Community]** = forums, blogs, READMEs of OSS projects reporting experience; **[Unknown]** = not found.

---

## Q1. Windows host prerequisites (USB drivers, AV/corporate policy, PowerShell 5.1 encoding, port 5037 conflicts, adb server stability)

### Takeaway
A Windows host needs (a) a working ADB USB driver per OEM — Google USB Driver for Pixel, OEM packages for Samsung/Sony/etc., with modern adbd exposing Microsoft OS descriptors so the inbox WinUSB driver *can* bind automatically, but this is not documented as guaranteed — (b) a single, pinned adb version owning TCP 5037 (other phone suites/emulators bundle their own adb and fight for the port), and (c) explicit UTF-8 decoding of adb output in PowerShell 5.1 on Japanese Windows (CP932 default ⇒ mojibake). Platform-tools changed its Windows USB backend in 37.0.1 (July 2026), so version pinning and regression testing matter.

### Cited Findings
**USB drivers**
- [Official] Google: "The Google USB Driver is required to perform adb debugging on Windows with Google devices"; Windows drivers for all other devices are provided by the hardware manufacturer (OEM USB drivers page). Download requires accepting the Android SDK License Agreement; current file is `usb_driver_r13-windows.zip`, license text dated 2026-04-28, page updated 2026-04-30. Also installable via SDK Manager into `android_sdk\extras\google\usb_driver\`. — [Get the Google USB Driver](https://developer.android.com/studio/run/win-usb)
- [Official] Google's OEM driver list links: Samsung `developer.samsung.com/galaxy/others/android-usb-driver-for-windows`, Sharp `k-tai.sharp.co.jp/support/`, FCNT `fcnt.com/support/develop/#anc-03`, Kyocera `kyoceramobile.com/support/drivers/`, Sony `developer.sonymobile.com/downloads/drivers/`. Page gives install steps only for Windows 10/8.1/7 (nothing for Windows 11) and **does not mention Windows Update at all**. Editing `android_winusb.inf` (e.g., to add other VID/PIDs) "will lead to security warnings when you install or upgrade the driver" (i.e., breaks the driver signature). — [Install OEM USB drivers](https://developer.android.com/studio/run/oem-usb)
- [Official/OEM] Sharp's only official "ADB USB driver" download page found is for old **media tablets** (EB-W51GJ/EB-WX1GJ, A01SH, EB-W700G), distributed as `usb_driver_SHARP.zip` behind a license agreement (use in Japan, free of charge). No AQUOS-smartphone ADB driver page was found. — [Sharp soft_dl_menu](https://jp.sharp/support/mediatablet/soft_dl_menu.html), [Sharp USB driver agreement](https://jp.sharp/support/mediatablet/doc/usb_driver_agreement.html)
- [Official/OEM] Sony's support article (last modified 2026) still points to `developer.sony.com/develop/drivers/` and a manual "Add legacy hardware" install of the ADB Interface driver; pages are titled for Windows 7/10; no statement that Windows 10/11 need no driver. (snippet) — [Sony support SX278901](https://www.sony.ca/en/electronics/support/articles/SX278901)
- [Official source code] AOSP adbd provides **Microsoft OS descriptors** (Extended Compat ID + DeviceInterfaceGUID) so that "the inbox WinUSB driver [gets] automatically installed" instead of requiring a VID-specific INF. — [adbd: Update DeviceInterfaceGUID for WinUSB (commit)](https://gerrit.omnirom.org/plugins/gitiles/android_system_core/+/b310da608b8292999cb30ed29bfb0bf4a9cb606d%5E%21); explanation in [Fabien Sanglard – Android Windows driver](https://fabiensanglard.net/android_windows_driver). A secondary (Korean) summary claims this covers all CTS-compliant Android 11+ devices — unverified. — [GeekNews summary](https://news.hada.io/topic?id=12647)
- [Community] Windows Update as a driver source is described as "possible but unreliable"; guides recommend manual Device Manager → "Android ADB Interface" or OEM packages. A Pixel 7 user reported Windows Update did not fix detection. — [WindowsReport ADB driver guide](https://windowsreport.com/windows-unable-install-adb-interface/), [GrapheneOS forum](https://discuss.grapheneos.org/d/2753-trouble-installing-google-usb-drivers-w-fastboot-onto-windows-10)
- FCNT (arrows) and Kyocera official Japanese driver pages could not be read (fcnt.com blocked; Kyocera JP page not found by search). — see Gaps.

**platform-tools versions / Windows-specific changes** [Official] — [SDK Platform Tools release notes](https://developer.android.com/tools/releases/platform-tools)
- Versions: 37.0.1 (July 2026), 37.0.0 (Feb 2026), 36.0.2 (Sep 2025), 36.0.0 (Apr 2025), 35.0.2 (Jul 2024)…
- 37.0.1: "**Windows USB: Introduce `libadbusb` to replace `libusb`.** `libadbusb` can be disabled by setting the `ADB_USB_LEGACY` environment variable to `1`." Also lazy-loads `wlanapi.dll` "to allow adb to run on Windows Server that doesn't include `wlanapi.dll`"; openscreen mDNS backend deleted (`ADB_MDNS_OPENSCREEN` no effect); new `ADB_TRACE` log levels.
- 36.0.2: "Fix bug where **last character of a file was truncated when pulled or pushed [Windows only]**"; "Fix bug where Samsung devices were not properly detected"; legacy USB backend made default on Linux "due to instability reports (device disconnection, interface read failure…)".
- 36.0.0: re-written libusb backend; "Libusb USB backend hot-plug now supports for Windows".
- Release notes never mention a device-count limit, many-device behavior, or output encoding.

**Antivirus / corporate policy**
- [Community] No report found of Defender flagging Google's official `adb.exe`; the documented pattern is wrapper executables being flagged: a debloat toolkit's packaged `run_toolkit.exe` "was flagged as a trojan by Windows Defender and VirusTotal. This was a false positive caused by how the exe was packaged". — [android-tv-debloat-toolkit v1.2 release note](https://newreleases.io/project/github/seun-novodev/android-tv-debloat-toolkit/release/v1.2)
- [Community] AppLocker/WDAC: publisher-based rules can reject binaries whose signatures fail stricter Authenticode checks; WDAC matches `OriginalFileName` so renaming does not evade deny rules (secondary). No source confirming whether platform-tools `adb.exe` is Authenticode-signed by Google LLC. — [Microsoft Security Advisory 2915720](https://technet.microsoft.com/library/security/2915720), [appcontrol.ai on WDAC block rules](https://www.appcontrol.ai/post/signed-trusted-abused-making-sense-of-wdac-s-recommended-block-rules)

**PowerShell 5.1 + adb encoding**
- [Official] `$OutputEncoding` governs what PowerShell sends **to** native programs; decoding of a native program's stdout uses `[Console]::OutputEncoding` (snippet). — [about_Character_Encoding](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_Character_Encoding)
- [Community] On Japanese Windows the console code page typically stays CP932 (Shift-JIS) even when `$OutputEncoding` is UTF-8, producing mojibake for UTF-8 output; fix is `[Console]::OutputEncoding = [Text.Encoding]::UTF8` (and optionally `chcp 65001`). — [Mynavi TECH+ (PowerShell 文字化け)](https://news.mynavi.jp/techplus/article/techp5170/)

**Port 5037 / version conflicts**
- [Official] "When the server starts, it binds to local TCP port 5037… All adb clients use port 5037 to communicate with the adb server." — [Android Debug Bridge (adb)](https://developer.android.com/tools/adb)
- [Community] Other software ships its own adb and grabs 5037: Apowermirror started "the adb with another version of adb" ([Medium](https://medium.com/@liwp.stephen/solved-adb-server-version-doesnt-match-this-client-493cec01f032)); Wandoujia's kadb/qadb, HBuilder's adb vs Nox `nox_adb.exe` ([CSDN](https://damodev.csdn.net/6a4357ac10ee7a33f28485be.html)); BlueStacks `hd-adb.exe` listed as a 5037 holder ([adbkonnekt on PyPI](https://pypi.org/project/adbkonnekt/0.12)). Diagnosis: `netstat -ano | findstr 5037`.
- [Community/OSS] scrcpy FAQ: "adb server version doesn't match" means several adb versions run at once — "use the same adb version everywhere", or point scrcpy at one binary via the `ADB` environment variable. — [scrcpy FAQ](https://github.com/Genymobile/scrcpy/blob/master/FAQ.md)
- [Official source] A device shows `offline` when the server sees the device but cannot connect to adbd. — [adb dev overview](https://android.googlesource.com/platform/packages/modules/adb/+/HEAD/docs/dev/overview.md)

### Inferences
- Bundle-or-require one pinned platform-tools (≥36.0.2 because of the Windows push/pull truncation bug; test 37.0.1's new `libadbusb` backend at scale and keep `ADB_USB_LEGACY=1` as a documented fallback switch).
- At tool start: detect other adb servers on 5037 (Get-NetTCPConnection -LocalPort 5037 → owning process path), refuse/ask before killing; optionally run the tool's own server on a non-default port (`ADB_SERVER_SOCKET` / `-P`) to coexist with vendor suites — needs verification by the ADB-capabilities researcher (not in the official page text read here).
- In PowerShell 5.1, don't parse adb text through the default pipeline: set `[Console]::OutputEncoding` to UTF-8 around adb calls or use `System.Diagnostics.Process` with `StandardOutputEncoding=UTF8`; binary output (e.g., `exec-out screencap -p`) must bypass the PowerShell pipeline entirely (PS 5.1 re-encodes native stdout as text) — the latter is general PowerShell knowledge, not sourced here.
- Since fabriq's own launcher .exe has had Defender false positives, avoid wrapping adb/scrcpy in a self-made packed exe; ship plain signed binaries + .ps1, and document Defender/ASR/WDAC allow-listing by hash/publisher.
- Driver strategy for a kitting PC: pre-install Google USB Driver + Samsung driver + Sony driver; treat "no driver" (inbox WinUSB via MS OS descriptors) as the expected path for recent AQUOS/arrows/Kyocera but verify per model in a device-compatibility matrix.

### Gaps
- FCNT (arrows) and Kyocera Japan current ADB driver availability/Windows 11 support — fcnt.com blocked; Kyocera JP page not found.
- Whether Windows Update reliably auto-installs ADB drivers for Samsung/Pixel — no official Microsoft/Google statement found.
- Whether Google's platform-tools `adb.exe` is Authenticode-signed (useful for WDAC publisher rules) — not found; check with `Get-AuthenticodeSignature` on a real copy.
- No official data on adb server stability vs number of devices.

---

## Q2. Scale: devices per host (USB topology/xHCI limits, hubs, cables, throughput, adb limits, Wi-Fi ADB, device-farm experience)

### Takeaway
The binding constraint is **xHCI endpoint/resource budget per host controller, hub quality and power**, not adb. Device-farm experience converges on **~15–30 devices per host** (STF: ~28 on powered hubs; 42Gears AstroFarm: 15 Android per machine; Japanese commercial kitting tools: 20–40 per PC, with USB 2.0 recommended). Wi-Fi ADB is not a bootstrap path for fresh devices (needs Wi-Fi + on-device pairing first) but became more reliable with Android 17 + adb 37 ("adb Wi-Fi 2.0").

### Cited Findings
**USB controller / endpoints**
- [Vendor] Endpoints, not device count, are the limit; "a typical mobile device might enumerate as 5 endpoints"; "common implementations of the XHCI controllers impose their own limit on the total number of endpoints to 96" (Intel); AMD controllers reported up to 254 endpoints. Symptoms: Windows "run out of USB resources" or intermittent failures without clear error. Mitigations: lower-endpoint hubs / USB 2.0 paths for devices not needing 3.0 speed. (snippet) — [Cambrionix: USB endpoint limitations](https://docs.cambrionix.com/general/guides/understanding-and-managing-usb-endpoint-limitations)
- [Vendor] Windows "Not enough USB controller resources" error KB. (snippet) — [Plugable KB](https://kb.plugable.com/en_US/general-support/not-enough-usb-controller-resources-error); Intel community thread on xHCI endpoint limits (one reply disputes how commonly 96 is hit) — [Intel Community](https://community.intel.com/t5/Embedded-Intel-Core-Processors/Hardware-limitations-on-USB-endpoints-XHCI/m-p/264571)
- [Patent] ADB requires the host to configure a USB interface opening two pipes by default; a single Android device may need up to 16 pipes depending on enabled functions. (snippet) — [US12147375B2](https://patents.google.com/patent/US12147375)

**Device-farm / tool experience**
- [Community/OSS] DeviceFarmer/STF README: a recommended build "can handle about **28 devices on powered hubs**, plus about 10 more on motherboard ports" (motherboard ports "usually not recommended for stability reasons"); "Intel Haswell processors may run into USB host controller limits at around **9 to 12 devices**"; budget AMD builds "can sometimes completely lose the USB PCIE extension cards"; best hub = Plugable **USB 2.0** 7-port with 60 W adapter (1.5 A/port BC); "The more recent USB 3.0 version has proven unreliable"; "most powered USB 3.0 hubs it tested occasionally disconnected entirely"; some hubs cannot charge + transfer simultaneously → charging cycles; "**On Windows you're on your own**". Supports Android 2.3.3–15 (SDK 35). — [DeviceFarmer/stf](https://github.com/DeviceFarmer/stf)
- [Vendor] 42Gears AstroFarm guideline (Jan 2026): up to **15 Android devices per contributor machine** (5 iOS); 16-port hub with per-port switches; "based on real-world experience… not official specifications". (snippet) — [AstroFarm Recommendation v2 (PDF)](https://www.42gears.com/wp-content/uploads/2026/01/AstroFarm-Recommendation-v2.pdf)
- [Community] STF single-host deployment noted as fine "when scaling to 14 devices is sufficient"; a Chinese guide says STF can control 160+ devices (multi-host) with good hubs/wireless. — [GitHub topics: stf](https://github.com/topics/stf), [TesterHome](https://testerhome.com/topics/2988?locale=en)
- [Vendor, JP] 東信電気 "Tablet admin for キッティング": Windows software, USB-connected, **up to 40 devices at once**, "1台あたり30分→3分" in a 40-unit USB-hub case (environment-dependent); management PC Windows 10, Android 6.0+, **USB 2.0 recommended**. (snippet) — [tabletadmin.com Android kitting](https://www.tabletadmin.com/android-kitting.html), [東信電気 product page](https://www.toshin-et.co.jp/service_product/tabletadmin/index.php)
- [Vendor, JP] PMT "キッティング支援システム": up to **20 devices simultaneously**, image-judgement to prevent errors. (snippet) — [PMT](https://pmt-inc.jp/kitting.html)
- [Community] Appium forum anecdote of an "adb limit of 16 phones"; a Russian Q&A claims only 3 simultaneous devices — both unsourced anecdotes, contradicted by STF/commercial numbers above. — [Appium discuss](https://discuss.appium.io/t/maximum-how-many-mobile-phone-i-can-connect-via-usb-cable-over-adb/34491), [Habr Q&A](https://qna.habr.com/q/368083)

**adb's own limits**
- [Official] adb scans odd ports 5555–5585 for emulators — "the range used by the first 16 emulators". — [adb docs](https://developer.android.com/tools/adb). `ADB_LOCAL_TRANSPORT_MAX_PORT` changes the scan range ([adb man page](https://man.archlinux.org/man/adb.1.txt)); an AOSP-derived commit raised the emulator count cap from 16 to 64 ([commit](https://gerrit.omnirom.org/plugins/gitiles/android_system_core/+/094fd86b1423fbb4f4c421447ddf6d10a5027c7b%5E%21)). These limits concern **emulators**, not USB devices; no documented USB-device cap was found.

**Hubs / power / cables**
- [Vendor, JP] サンワサプライ USB-2HCS20: USB 2.0 20-port, sync + charge, **self-powered only** (no bus power), up to 2.4 A/port charging, **1.5 A while syncing/PC-connected**. — [Sanwa flyer PDF](https://cdn.sanwa.co.jp/support/chirashi/hansoku/PDF/PDF190822_04.pdf)
- [Vendor] Wholesale 20-port USB+OTG "screen casting / synchronous operation" hubs ~US$134–141 per box (Global Sources listing). — [Global Sources](https://www.globalsources.com/Multifunction-USB/Synchronous-Operation-USB-Hub-1230022305p.htm)
- [Community, JP] Charge-only USB-C cables are sold alongside data cables; tell-tales are "充電専用" labels and absence of the USB logo on packaging. — [Mynavi compass-551](https://news.mynavi.jp/article/compass-551/), [ITmedia PC USER](https://www.itmedia.co.jp/pcuser/articles/2503/07/news106.html)

**Throughput**
- [Community] Install time often dominated by device-side verification, not USB transfer: a 4.7 MB APK took ~60 s on a Samsung phone; unchecking Samsung's "Check applications with USB" (Play Protect "verify apps over USB") cut it to <2 s. — [B4X forum](https://www.b4x.com/android/forum/threads/solved-b4a-installing-time-very-large.168106/post-1030317)
- No reliable USB 2.0 vs 3.x adb throughput benchmark for parallel installs was found.

**Wi-Fi ADB**
- [Official] Wireless debugging requires Android 11+ (phones), same network, one-time pairing by QR or pairing code; Android 10 and lower need USB first then `adb tcpip 5555`. "Android 17, alongside adb 37.0.0 introduces **adb Wi-Fi 2.0**… the device will automatically connect to the workstation when the device connects to a wireless debugging trusted network." — [adb docs](https://developer.android.com/tools/adb)
- [Official] Google blog (Sep 2026): new mDNS stack replaced Bonjour and legacy mDNS "so that your wireless devices more reliably stay connected". (snippet) — [Android Developers Blog: adb Wi-Fi 2.0](https://android-developers.googleblog.com/2026/09/wireless-debugging-adb-wifi-2.html)
- [Community] Auto-discovery needs multicast/mDNS between clients (guest/AP-isolation Wi-Fi breaks it); the connection port changes every wireless-debugging session; discovery historically disappeared after network change/PC restart/device shutdown. — [apk-extractor-adb wireless guide](https://www.mintlify.com/KilzaNiko/apk-extractor-adb/guides/wireless-debugging), [remio summary](https://www.remio.ai/post/google-adb-wi-fi-2-0-makes-wireless-android-debugging-more-reliable-but-compatib)

### Inferences
- Design target: **10–20 devices per Windows host as the default batch**, 20–40 as a stretch with (a) several independent xHCI controllers (add-in PCIe USB cards), (b) USB 2.0 self-powered sync-and-charge hubs (USB 2.0 hubs have their own TT and are what STF and Tablet admin recommend), (c) per-port power switching for recovery.
- The tool must treat USB flapping as normal: reconnect detection by serial, per-device retry, idempotent steps — matches fabriq's idempotency principles.
- APK distribution bottleneck is likely device-side verification/dexopt, not bandwidth; pre-disable "verify apps over USB" only if policy allows, otherwise budget minutes per large APK.
- Wi-Fi ADB is useful only for *later* phases (post-setup, Android 17+ trusted networks) or for scrcpy remote viewing; bootstrapping 40 fresh devices still needs USB.

### Gaps
- No Microsoft-official per-controller endpoint numbers for current Intel/AMD chipsets (Cambrionix/Intel forum figures only).
- No published benchmark of parallel `adb install` throughput on Windows with N devices.
- Industrial kitting hub vendors (Cambrionix, Acroname) pages could not be read (blocked).

---

## Q3. Licensing: redistributing adb.exe / AdbWinApi.dll, AOSP builds, scrcpy, OSS libraries, QR libraries

### Takeaway
Google's prebuilt platform-tools and Google USB Driver are under the **Android SDK License Agreement**, whose §3.4 forbids redistribution "except to the extent required by applicable third party licenses", while §3.5 carves out open-source components. adb's source is Apache-2.0, and well-known Apache-2.0 projects (scrcpy, QtScrcpy) **do ship Google's prebuilt adb.exe/AdbWinApi.dll/AdbWinUsbApi.dll** in their Windows zips — a widely tolerated practice but not an explicit grant. The legally clean options are: have the installer download platform-tools from dl.google.com with SHA-256 verification (exactly what scrcpy's build script does), or build adb from AOSP source under Apache-2.0. OSS deps of interest are permissive (QRCoder MIT; AdvancedSharpAdbClient, scrcpy, Shizuku, Headwind MDM Apache-2.0) except UAD-ng (GPL-3.0) and ADB AppControl (proprietary).

### Cited Findings
- [Official] SDK License (shown on platform-tools/USB-driver download): §3.1 "limited, worldwide, royalty-free, non-assignable, non-exclusive, and non-sublicensable license to use the SDK solely to develop applications for compatible implementations of Android"; §3.4 "Except to the extent required by applicable third party licenses, you may not copy (except for backup purposes), modify, adapt, **redistribute**, decompile… the SDK or any part of the SDK"; §3.5 "Use, reproduction and distribution of components of the SDK licensed under an open source software license are governed solely by the terms of that open source software license"; §3.8 do not remove proprietary notices. — [Platform-tools page / license](https://developer.android.com/tools/releases/platform-tools), [Google USB Driver page](https://developer.android.com/studio/run/win-usb)
- [Official source] adb module `Android.bp` declares `SPDX-license-identifier-Apache-2.0`. (snippet) — [AOSP adb Android.bp (android17-release)](https://android.googlesource.com/platform/packages/modules/adb/+/refs/heads/android17-release/Android.bp)
- [Community] Debian discussion: the SDK terms "seem to be specific to the SDK package distributed by Google" and do not prevent downstream packagers from distributing (source-built) adb under Apache-2.0; Replicant/Trisquel note Google's binaries "come with a proprietary license". — [Debian bug #792670](https://lists.alioth.debian.org/pipermail/pkg-java-maintainers/2015-July/052132.html), [Trisquel forum](https://trisquel.info/en/forum/what-license-android-sdk-platform-tools-how-free-are-they)
- [Conflict] A WinGet listing labels Google's Windows platform-tools zip "Apache-2.0", contradicting the SDK EULA — treat as unreliable metadata. — [wingetly platform-tools](https://wingetly.io/apps/google/platform-tools)
- [OSS precedent] scrcpy's `app/deps/adb_windows.sh` downloads `https://dl.google.com/android/repository/platform-tools_r37.0.1-win.zip`, verifies SHA-256 `45f4d631…1a99`, extracts only `adb.exe`, `AdbWinApi.dll`, `AdbWinUsbApi.dll` and puts them in the release ("Google only provides x86 Windows platform-tools"). — [scrcpy adb_windows.sh](https://github.com/Genymobile/scrcpy/blob/master/app/deps/adb_windows.sh)
- [OSS] scrcpy: Apache-2.0, v5.0.1 shown; features include virtual display (`--new-display`), camera mirroring (Android 12+), audio (Android 11+), OTG/HID mode not requiring USB debugging; multi-device via `-s/--serial` or `ANDROID_SERIAL`; `ADB` env var selects the adb binary. — [scrcpy README](https://github.com/Genymobile/scrcpy), [connection.md](https://github.com/Genymobile/scrcpy/blob/master/doc/connection.md), [FAQ](https://github.com/Genymobile/scrcpy/blob/master/FAQ.md)
- [OSS] QtScrcpy: Apache-2.0 ("same license as scrcpy"); group control ("control all your phones at the same time"); Windows release archives "with all the dependencies (including ADB)"; 632 open issues. — [QtScrcpy](https://github.com/barry-ran/QtScrcpy)
- [OSS] QRCoder: **MIT**; supports .NET Framework 3.5+ / .NET Standard 1.3+ / .NET 5+; `PngByteQRCode`, `SvgQRCode`, `BitmapByteQRCode`, `Base64QRCode` need no System.Drawing; maintained by Shane32 since 2025. — [QRCoder](https://github.com/codebude/QRCoder)
- [OSS] AdvancedSharpAdbClient: **Apache-2.0**; .NET Framework 2.0–4.8, .NET Standard 1.3–2.1, .NET 6/8/10; "uses the `adb.exe` server process as an intermediate" (so still needs adb.exe); NuGet 3.6.16 seen. — [AdvancedSharpAdbClient](https://github.com/SharpAdb/AdvancedSharpAdbClient), [NuGet](https://www.nuget.org/packages/AdvancedSharpAdbClient)
- [OSS] Shizuku: code Apache 2.0 with restrictions — don't reuse its icons, the name "Shizuku", app ID `moe.shizuku.privileged.api`, or its permission namespace; server started via root or ADB. — [Shizuku](https://github.com/RikkaApps/Shizuku)
- [OSS] UAD-ng (Universal Android Debloater Next Generation): **GPL-3.0**, Rust GUI using ADB for non-rooted devices, 9.6k stars. — [UAD-ng](https://github.com/Universal-Debloater-Alliance/universal-android-debloater-next-generation)
- [OSS] Headwind MDM server: **Apache-2.0** community edition (QR-code enrollment); Enterprise edition (kiosk/COSU, restrictions) is commercial. — [hmdm-server](https://github.com/h-mdm/hmdm-server)
- [Proprietary] ADB AppControl: freemium/proprietary; terms say commercial use requires the author's permission/license; decompiling prohibited; batch install/operations are PRO/Extended (paid) features. — [ADB AppControl terms](https://adbappcontrol.com/?go=terms), [AlternativeTo](https://alternativeto.net/software/adb-appcontrol/about), [ADB AppControl PRO](https://adbappcontrol.com/en/mobile/pro)

### Inferences
- For fabriq-style folder distribution with `LICENSES/` + `THIRD_PARTY_NOTICES.md`: safest is **not bundling** Google binaries; ship a `setup` step that downloads the pinned platform-tools zip from dl.google.com, verifies SHA-256, and extracts only the 3 files (copying scrcpy's approach), with the user accepting Google's terms. If bundling anyway, follow the scrcpy precedent but record the risk (SDK §3.4) and keep Google's `NOTICE.txt`.
- scrcpy itself can be bundled (Apache-2.0: include LICENSE + NOTICE), but its zip already contains Google's adb — the same SDK-license question re-appears; prefer pointing scrcpy at the tool's adb via `ADB` env var to keep one adb version (also avoids 5037 conflicts).
- QRCoder (MIT, net35+) is directly usable from PowerShell 5.1 via `Add-Type -Path QRCoder.dll` (PngByteQRCode avoids System.Drawing issues) — inference; test on 5.1.
- Avoid embedding UAD-ng code/lists without GPL-3.0 compliance review; ADB AppControl is not reusable commercially.

### Gaps
- Whether Google considers prebuilt `adb.exe`/`AdbWinApi.dll` "components licensed under an open source software license" (§3.5) — no Google statement found; needs legal review.
- No maintained third-party **Windows** build of AOSP adb (source-built, Apache-2.0) was identified.
- Licenses/maintenance of ppadb (pure-python-adb), adbkit, Fastboot Enhance, "adb-multi" style tools, and STF's exact license name (README read did not state it) — not verified in this session.

---

## Q4. Remaining human/manual steps per device, typical time, and Japanese kitting-service prices

### Takeaway
ADB can only start after a human has powered the device, passed (or partially passed) the setup wizard, enabled Developer options + USB debugging, and approved the RSA prompt; SIM insertion, film, labels and visual checks also stay manual. Japanese smartphone kitting is priced roughly **¥1,000–3,000/device** for basic work (packages up to ¥4,400–8,800 with MDM), and the closest automation benchmark is Tablet admin's vendor claim of **30 → 3 minutes/device** for 40 devices.

### Cited Findings
**Task breakdown**
- [Official, municipal] A Japanese city's smartphone procurement spec lists SIMカード挿入, 初期設定, 動作確認, 液晶保護フィルム装着, MDM によるアプリ制限・アップデート許可設定. (snippet) — [石岡市 資料 PDF](https://www.city.ishioka.lg.jp/data/doc/1729155430_doc_26_0.pdf)
- [Official] USB debugging requires on-device acceptance of the RSA key ("cannot be executed unless you're able to unlock the device and acknowledge the dialog"). — [adb docs](https://developer.android.com/tools/adb)
- [Vendor] Developer options are unlocked by tapping Build number ≥7 times; menu paths vary by model/OS. QR-code enrollment needs no developer mode/USB debugging. — [TCL FAQ](https://www.tcl.com/uk/en/support-mobile/faq/7033), [ArborXR help](https://help.arborxr.com/en/articles/8437122-developer-mode-usb-debugging-on-android-tablets-phones)
- [Official, carrier R&D] NTT docomo Technical Journal Vol.25 No.4 (Jan 2018): growing setting items make Android kitting take "場合によっては1台当り数時間", whole deliveries can take months; docomo built an internal PC tool "Androidキッティングツール（AKT）" (mentions a UI Automator method among two approaches). (snippet) — [docomo Technical Journal PDF](https://www.docomo.ne.jp/binary/pdf/corporate/technology/rd/technical_journal/bn/vol25_4/vol25_4_007jp.pdf)
- [Vendor] SoftBank blog: manual kitting "1台あたり数時間かかるのが一般的" (mostly PC context). — [SoftBank ビジネスブログ](https://www.softbank.jp/biz/blog/business/articles/202403/what-is-kitting/)
- [Vendor] Tablet admin for キッティング: RPA that automates the taps/swipes; 40-unit case 30 min → 3 min per device (environment-dependent). — [tabletadmin.com](https://www.tabletadmin.com/android-kitting.html); earlier releases claimed "約10分の1の作業時間" — [@Press 2020](https://www.atpress.ne.jp/news/213787)

**Prices (JPY; retrieved 2026-10-29; publication dates of the listings were not verified — several are months to ~1.5 years old)**
- [Vendor] ベルパーク (法人): "すぐ使えるパッケージ" (SIM insertion, film etc.) **¥4,400/台**; MDM-setting plans **¥6,600 / ¥8,800/台**; system-side initial setup **¥110,000** separate. (snippet) — [ベルパーク キッティングサービス](https://houjin.bellpark.co.jp/services/kitting)
- [Vendor/aggregator] Comparison articles: smartphone **¥1,000–2,000/台**, tablet **¥1,500–3,000/台**, MDM/app bulk distribution as options; one provider "基本初期設定 ¥2,480〜(税抜)" from 1 unit; Seino Information Service quoted as iPhone ¥1,550〜/台, iPad ¥2,900〜/台 (attribution via aggregator, unverified). (snippet) — [テクバン Tech-is](https://biz.techvan.co.jp/tech-is/blog/device/002936.html), [GIZUMO](https://gizumo-inc.jp/media/kitting-service-cost/), [クロスネットワーク](https://www.xnetwork.jp/contents/kitting-cost)
- [Marketplace] Individual on coconala: Android kitting "1台2,000円〜". — [coconala](https://coconala.com/services/3104658)
- [Vendor] エクスモバイル (rental phones): phone-book registration option ¥220〜/台 (tax incl.). — [エクスモバイル キッティング](https://www.xmobiles.jp/pc/special/kitting.php)
- [Vendor] Kitting price structure = base + options (asset-system registration, security software, reports, travel/holiday surcharges). — [IT トレンド](https://it-trend.jp/kitting/article/717-9136), [BOXIL](https://boxil.jp/mag/a3004/)

### Inferences
- Rough per-device manual baseline for an Android smartphone with SIM+Wi-Fi+several apps+settings+evidence: **20–40 minutes** (consistent with Tablet admin's "30 min" baseline); a USB/ADB tool can plausibly bring the *operator-attended* time to ~3–8 minutes/device (unbox, SIM, power, wizard to home screen, dev options + USB debugging, RSA tap, cable-in, final visual check), the rest running in parallel.
- At ¥1,000–3,000 per device of market price, the economic value is in throughput (40/batch) and evidence quality (screenshots + HTML checklist), not in eliminating the human front-end.
- The setup-wizard → enable USB debugging front-end is the dominant fixed cost; only DO/QR/zero-touch (other researchers) remove it.

### Gaps
- No Japanese provider publishes a per-device **minutes** breakdown for Android smartphones; docomo's AKT numbers (devices in parallel, time saved) not readable (PDF blocked).
- Carrier (docomo/KDDI/SoftBank) business-kitting price lists not found publicly; Ricoh/Otsuka/Canon MJ smartphone kitting prices not found.

---

## Q5. FRP and Google-account considerations

### Takeaway
Consumer FRP engages on an **untrusted reset** (recovery/remote wipe) when a Google account is on the device; Android 15 moved enforcement out of the setup wizard (bypassing the wizard no longer disables FRP, and blocks adding accounts/screen lock/apps). For ADB-based kitting without Device Owner, the safe practice is **no personal Google account at all**, or a company-controlled account removed before shipment; with managed devices, EFRP lists the company accounts allowed to unlock. Password changes trigger a 24–72 h re-login lockout window.

### Cited Findings
- [Official] Google: on fully managed devices a reset from device settings does not trigger FRP; on company-owned devices with a work profile FRP is generally triggered if a Google account is present at reset; with Android Management API EMMs, "FRP will only be initiated if the EFRP policy is set"; EFRP specifies which Google Accounts can activate a reset device. (snippet) — [Enable enterprise factory reset protection](https://support.google.com/work/android/answer/14549362)
- [Community expert] Bayton: EFRP is always enforced after a hard reset on managed devices running **Android 15+**, regardless of OEM unlock; configure EFRP before problems arise; recovery is hard if the enrolment account was deleted. (snippet) — [bayton.org FRP vs EFRP](https://bayton.org/android/android-enterprise-faq/frp-vs-enterprise-frp/)
- [Official community] Android Enterprise community: "in the past, the setup wizard took care of FRP, but it could be circumvented"; Android 15 moved enforcement deeper into the system. — [Enhanced FRP in Android 15](https://www.androidenterprise.community/blog/resources/enhanced-factory-reset-protection-in-android-15/9493)
- [Press] Android 15: bypassing setup wizard no longer deactivates FRP; blocked from adding a Google account, setting a screen lock, or installing apps until ownership is verified. Android 16 (planned, May 2025 report): force another factory reset if bypass detected. — [Android Authority A15](https://www.androidauthority.com/android-15-factory-reset-protection-upgrades-3479431), [Android Authority A16](https://www.androidauthority.com/android-16-factory-reset-protection-upgrades-3556859)
- [Vendor/Press] FRP is triggered by recovery/remote wipe; removing the Google account from the device removes FRP ("never perform an untrusted reset on a device which has an unmanaged Google account" — Zebra); after a password change, the new password may not work on the device for 24 h (72 h on older phones); wrong entries restart the timer (Samsung). — [Zebra developer blog](https://developer.zebra.com/blog/factory-reset-protection), [Android Central](https://androidcentral.com/factory-reset-protection-what-you-need-know), [Samsung NZ](https://www.samsung.com/nz/support/mobile-devices/what-is-google-frp/)
- [Vendor] EFRP via MDM: Sophos needs internal Google IDs of FRP accounts; ManageEngine/Applivery (`frpAdminEmails`) equivalents; requires Device Owner (Seqrite). — [Sophos](https://docs.sophos.com/central/Mobile/help/en-us/AdminHelp/EnrollDevices/AndroidEnterprise/FRP/), [ManageEngine](https://www.manageengine.com/mobile-device-management/help/profile_management/android/android_enterprise_factory_reset_protection.html), [Applivery](https://docs.applivery.com/en/device-management/android/policies/factory-reset-protection.md)

### Inferences
- Tool rule: per-device CSV must declare `GoogleAccountPolicy` = none / company-temp / user; default **none**. If an account is added for Play installs, the tool should remove it (and verify via `dumpsys account` or similar — ADB researcher to confirm) **before** any reset and before shipment; log this in the HTML checklist.
- Never use recovery-mode wipes in the kitting flow while any account may be present; "factory-state hardening" should be a settings-initiated reset after account removal.
- Do not create "kitting staff" personal Gmail accounts at scale — besides FRP risk, Google may flag bulk-created accounts (not sourced; general caution).

### Gaps
- Exact current Google wording (support.google.com blocked); Android 16/17 FRP "second reset" behavior shipped or not — not confirmed.
- How specific Japanese kitting vendors handle Google accounts (company accounts vs none vs managed Google Play accounts) — no public procedure found; Optim's Android kitting manual exists but content not readable. — [OPTiM Android_Kitting.pdf](https://biz3.optim.co.jp/dist/Manual/ja/Android_Kitting.pdf)

---

## Q6. App sourcing without Google Play (APK sources, extraction, split APKs, updates, Play Integrity, sideloading policy)

### Takeaway
`adb install` of vendor-provided APKs remains viable: Google's new **developer verification** (live 2026-09-30 in Brazil/Indonesia/Singapore/Thailand, global 2027) explicitly leaves the **ADB workflow unchanged**. Extracting Play-installed apps from a reference device (`pm path` + `adb pull`) is technically easy but redistribution may infringe copyright/Play terms, split-APK apps need all splits, and apps using Play Integrity licensing will report UNLICENSED and may force "Install from Play". Updates of sideloaded apps won't come from Play unless the user/account becomes licensed.

### Cited Findings
- [Official/Press] Developer verification enforcement began **2026-09-30** in Brazil, Indonesia, Singapore, Thailand; global rollout 2027; initial enforcement covers participating stores; "ADB remains exempt"; Google's guide says the ADB workflow "stays the same"; a new "advanced flow" for unverified apps (restart + 24 h wait + warnings) began rolling out Aug 2026. Japan's date not stated (presumably 2027 global phase). — [Android Developer Console Help](https://support.google.com/android-developer-console/answer/16561738?hl=en), [Android Authority](https://www.androidauthority.com/google-android-advanced-flow-sideloading-rollout-begins-3700073/), [Gadget Hacks](https://android.gadgethacks.com/news/google-android-sideloading-restrictions-sept-30-explained/), [Mega Mobile Content](https://www.megamobilecontent.com/news/2026/10/01/android-developer-verification-live-brazil-indonesia-singapore-thailand/)
- [Community] Extraction: `pm list packages -f` / `pm path <pkg>` then `adb pull`; base.apk plus config splits (ABI, language, density) — all splits needed. — [Codemia](https://codemia.io/knowledge-hub/path/how_do_i_get_an_apk_file_from_an_android_device)
- [Community/legal commentary] Extracting and distributing APKs "may infringe on the copyrights of the app developer"; Google Play Terms license is limited to uses explicitly authorized. — [Protectstar](https://protectstar.com/de/blog/apk-files-what-you-need-to-know-are-they-harmful-or-illegal), [ConductAtlas: Google Play Terms](https://conductatlas.com/platform/google-play-store/google-play-terms/license-scope-and-restrictions/)
- [Official] Play Integrity `appLicensingVerdict` UNLICENSED = "they didn't install or buy the app from Google Play"; developers can show a GET_LICENSED dialog that adds the app to the user's Play library so "Google Play can deliver app updates". Press: dialog can be unskippable with "Install from Play". — [Play Integrity remediation](https://developer.android.com/google/play/integrity/remediation), [Android Authority](https://androidauthority.com/play-integrity-sideloading-detection-3480639)

### Inferences
- Preferred sources in order: (1) vendor/in-house APKs delivered with redistribution rights (record source + SHA-256 in CSV); (2) apps the customer licenses via managed Google Play / EMM (other researchers); (3) reference-device extraction **only for in-house or explicitly licensed apps** — never as a general Play-app redistribution mechanism.
- CSV schema should support multi-file app entries (`install-multiple`) and a per-app flag "requires Play licensing/Integrity" so operators know a Play account step is needed.
- Post-kitting updates for sideloaded apps must be planned (MDM, vendor updater, or re-kitting); the tool can implement "update if versionCode lower" idempotently.

### Gaps
- Japan-specific developer-verification timing and whether enterprise/EMM-pushed APKs are exempted in 2027 — not found.
- Play Integrity behavior for apps installed via managed Google Play private apps vs ADB — not researched (other researchers).

---

## Q7. Security considerations (USB debugging after kitting, shared adbkey, secrets, evidence privacy)

### Takeaway
Leaving USB debugging/developer options on is a recognized risk (STIG baselines require it off; some banking/government apps refuse to run with it on), so final hardening should disable it and revoke authorizations. A shared adb key across kitting PCs is supported (`ADB_VENDOR_KEYS`) but makes that key a fleet-wide credential. In Japan, mobile phone numbers are not 個人識別符号 per se, but IMEI/phone numbers tied to employee names in kitting ledgers/evidence are personal information to be protected.

### Cited Findings
- [Vendor] Debugging should be "a temporary elevation of privilege… never a persistent configuration state on corporate hardware"; MDMs offer allow/block USB debugging policies. — [Scalefusion](https://blog.scalefusion.com/usb-debugging-on-android/), [Hexnode](https://www.hexnode.com/blogs/usb-debugging-in-android-devices-and-why-disable-it/)
- [Official baseline] DISA STIG (Samsung Knox): checks USB debugging disabled and cannot be re-enabled; "disabling Developer Modes will also automatically disable USB Debugging and Mock Locations". — [STIG V-56059](https://stigviewer.com/stig/samsung_android_with_knox_2.x/2015-01-05/finding/V-56059)
- [Official, apps] A bank warns USB debugging "could pose a security risk to your banking information"; Singapore's Singpass app refuses to run with debugging on. — [Standard Chartered](https://www.sc.com/gh/security-warning-1004/), [ask.gov.sg Singpass](https://ask.gov.sg/singpass/questions/cm8idm20i0057wd6ud5h2kq4a)
- [Official source/man] `ADB_VENDOR_KEYS`: colon-separated paths (or directories) of extra RSA keys used for authorization; keys deduplicated; device trusts keys listed in `adb_keys`/`adb_user_keys`. Chromium infra used it to push a host key onto unauthorized devices. — [adb man page](https://man.archlinux.org/man/adb.1.txt), [AOSP: How ADB manages keys](https://android.googlesource.com/platform/packages/modules/adb/+/fdbd8a5f/docs/dev/keystore.md), [Chromium code review](https://codereview.chromium.org/979013002)
- [Commentary, JP] 携帯電話番号・クレジットカード番号 "それ自体は個人識別符号とされませんでした" (契約形態・運用が多様なため). — [アガルート 資料 PDF](https://www.agaroot.jp/wp-content/uploads/2026/03/gyosei_tajima_kiso_kojin.pdf); 個人関連情報 scope commentary — [西村あさひ ニューズレター](https://www.nishimura.com/sites/default/files/newsletters/file/data_protection_260818_ja.pdf); 総務省/PPC 電気通信事業における個人情報等保護ガイドライン解説 (updated May 2026) — [総務省 PDF](https://www.soumu.go.jp/main_content/001078650.pdf)
- [Official, Google JP] IMEI is assigned to the device and does not change when owner/phone number changes. — [android.com IMEI](https://www.android.com/intl/ja_jp/articles/317)

### Inferences
- Final step "factory-state hardening": disable USB debugging and developer options, revoke USB debugging authorizations, verify (read back) and record in the checklist — fits fabriq Step 5.5 verification. Exact `settings put global adb_enabled 0` / `development_settings_enabled 0` behavior per OEM should be verified by the ADB researcher; disabling adb over adb terminates the session, so it must be the last step.
- Shared key: keep one adbkey per kitting site, stored with OS-level ACL/DPAPI; rotate per project; revoke on devices at the end — otherwise any leaked key grants shell to every unrevoked device still in debugging mode.
- Secrets: reuse fabriq's AES-encrypted CSV values for Wi-Fi PSK/cert passwords; avoid passing secrets on adb command lines visible in logs/screenshots; screenshots of Wi-Fi screens must not show passwords.
- Evidence (screenshots/HTML) containing IMEI + 電話番号 + 利用者氏名 constitutes 個人情報 when linkable to individuals; apply access control, retention limits and optional masking columns. (Legal interpretation — needs confirmation.)

### Gaps
- No PPC/official statement found on whether IMEI alone is 個人識別符号 or 個人関連情報.
- No official Android doc read on revoking authorizations programmatically.

---

## Q8. Existing tools, commercial staging tools, and Japanese kitting workflows

### Takeaway
The closest Japanese analog is **東信電気 Tablet admin for キッティング** (Windows, USB, RPA tap/swipe automation, ≤40 devices, subscription), plus PMT's 20-device system and docomo's internal AKT. OSS building blocks are mature (scrcpy/QtScrcpy for viewing/group control, AdvancedSharpAdbClient, UAD-ng lists, Shizuku, STF for farm architecture lessons). Global "staging" tools (Zebra StageNow, SOTI/MobiControl QR, 42Gears) are barcode/QR + MDM oriented and mostly OEM-specific (Zebra).

### Cited Findings
- [Vendor, JP] Tablet admin for キッティング (東信電気): Windows software; Android 6.0+; Windows 10 mgmt PC; USB 2.0 recommended; up to 40 devices; plans: limited-count monthly, unlimited monthly, prepaid (1 year); price "オープン" (TEKWIND SHOP shows a ¥999,999 placeholder); subscription minimum term described differently by vendor vs reseller; sister product "Tablet admin PRO" ended sales end of April 2025 while "for Kitting" continues. (snippet) — [東信電気](https://www.toshin-et.co.jp/service_product/tabletadmin/index.php), [tabletadmin.com](https://www.tabletadmin.com/android-kitting.html), [@Press 2022](https://www.atpress.ne.jp/news/319120), [TEKWIND](https://tablet.tekwind.co.jp/application/tabletadmin/), [TEKWIND SHOP](https://www.tekwind.shop/shopdetail/000000002168/)
- [Vendor, JP] PMT キッティング支援システム: up to 20 simultaneous, image-judgement error prevention; "Certino2" 20 simultaneous / 200 with multicast (snippet). — [PMT](https://pmt-inc.jp/kitting.html)
- [Official, JP carrier] docomo AKT (Androidキッティングツール) internal PC tool (2018). — [docomo Technical Journal](https://www.docomo.ne.jp/binary/pdf/corporate/technology/rd/technical_journal/bn/vol25_4/vol25_4_007jp.pdf)
- [Vendor] Zebra StageNow: stage "a handful or thousands of Android devices with a quick scan of a barcode or tap on an NFC tag" (Zebra devices). SOTI support recommends StageNow for Zebra enrollment; older "MobiControl Stage" programmer largely replaced by the built-in QR generator (forum). 42Gears SureMDM integrates with StageNow. — [Zebra StageNow](https://prod-www.zebra.com/us/en/software/mobile-computer-software/stagenow.html), [SOTI Pulse thread](https://pulse.soti.net/community/thread/94976786-d69a-4fb6-abfa-a8240082b42a), [42Gears](https://www.42gears.com/?p=113125)
- [OSS] STF/DeviceFarmer: active but "largely funded by individual team members and their unpaid free time, leading to slow progress"; Linux-oriented ("On Windows you're on your own"). — [DeviceFarmer/stf](https://github.com/DeviceFarmer/stf)
- [OSS] Other building blocks and licenses: see Q3 (scrcpy, QtScrcpy, AdvancedSharpAdbClient, QRCoder, Shizuku, UAD-ng, Headwind MDM, ADB AppControl).
- [Survey] GIGA School 2nd phase (MM総研, Jun–Jul 2025 survey, 621万台): ChromeOS 60%, iPadOS 31%, Windows 10%; average unit price ChromeOS ¥54,000, Windows ¥55,000, iPadOS ¥57,000 — Android tablets are not a GIGA category. — [helentech (MM総研)](https://helentech.jp/news-71040/), [マイナビ ニュース](https://news.mynavi.jp/article/newsinsight-320/)

### Inferences
- Differentiators vs Tablet admin: CSV/module-driven per-serial configuration, evidence (screenshots + HTML checklist), OSS/transparent, folder distribution, no subscription — but RPA tap/swipe automation per OEM UI (what Tablet admin sells) is the expensive part to maintain; prefer `settings`/`cmd`/`pm`/`dpm` shell APIs over UI automation where possible.
- GIGA school practices are Chromebook/iPad zero-touch based and not directly transferable to Android ADB kitting.

### Gaps
- Esper provisioning and 42Gears bulk enrollment details/pricing not retrieved.
- "adb-multi"/"multi-adb" style OSS repos and "Android Platform Tools GUI" wrappers — maintenance/licence not verified.
- docomo/KDDI/SoftBank, Ricoh, Otsuka, Canon MJ standard Android kitting workflows — no public detailed procedures found.

---

## Q9. Android version / OEM landscape in Japan (2026) and update-policy implications

### Takeaway
As of Oct 2026, **Android 17 is released** (Pixel stable 2026-06-16; Samsung One UI 9 rolling out since 2026-09-16), while Japanese OEM business models will mostly be on Android 15/16 with staggered updates. Android ≈ 51–54% of Japanese users; AQUOS is the largest Android brand, then Xperia/Pixel; Kyocera is business-only (DIGNO/TORQUE/DuraForce); FCNT's arrows continues. Kitting must handle a mixed fleet and decide whether to apply OS/security updates during kitting (multi-GB downloads).

### Cited Findings
- [Press] Android 17 stable released 2026-06-16 with the June Pixel Drop. — [Gizmochina](https://www.gizmochina.com/2026/06/09/heres-when-google-pixels-will-receive-android-17-update/), [Android Authority](https://androidauthority.com/android-17-release-date-3639790)
- [Press] Samsung One UI 9 (Android 17) stable from 2026-09-16 (S26, Korea), S25/Fold7/Flip7 in US by early Oct, S24/Fold6/Flip6 from 2026-10-06 (Korea); full update ~**3.8 GB** for non-beta users; One UI 9.5 expected early 2027. — [9to5Google](https://9to5google.com/2026/09/16/samsung-android-17-one-ui-9-update-list/), [Android Authority S24](https://androidauthority.com/samsung-one-ui-9-galaxy-s24-fold-6-flip-6-rollout-3719569), [Nokiamob](https://nokiamob.net/2026/10/05/samsung-rolls-out-one-ui-9-and-android-17-to-galaxy-z-flip-7-and-fold-7-globally/)
- [Survey] MMD研究所 (Feb 2026): main OS Android 50.8% vs iPhone 49.0%; among Android: AQUOS 25.9%, Xperia 16.3%, Pixel 16.2%. A later (Aug) survey: Android 54.1%. — [Web担当者Forum](https://webtan.impress.co.jp/n/2026/04/01/52414)
- [POS data] BCN rankings Apr–Jul 2026: Pixel 10a, Galaxy S26, AQUOS wish5 / sense10, arrows We2 in top 10. Counterpoint Q2 2026: Pixel 10a #3 in Japan after iPhone 17 / 17e. — [BCN 2026-04-23](https://www.bcnretail.com/topics_detail10/20260423_620118.html), [BCN 2026-07-05](https://www.bcnretail.com/research/detail/20260705_640086.html), [helentech](https://helentech.jp/news-android-share-pixel-10a-japan-q2-2026-90585/)
- [Press] Kyocera exited consumer phones (announced 2023-05-15, by Mar 2025) but continues business phones and TORQUE; 2026: DIGNO SX5 KC-S306 (docomo, 2026-03-27, business), DuraForce EX2 (Sep 2026), TORQUE G07 (au, announced 2026-02-25). (snippet) — [ITmedia 2023](https://www.itmedia.co.jp/mobile/amp/2305/28/news047.html), [ITmedia 法人スマホ](https://www.itmedia.co.jp/keywords/business_sp.html), [Impress 法林岳之](https://k-tai.watch.impress.co.jp/docs/column/mobile_catchup/1507953.html)
- [Press] Sharp AQUOS: from AQUOS R8 (2023) "OSバージョンアップ最大3回・セキュリティ更新最大5年" (previously 2 / 3 years); extended to sense series from sense8; older business sense5G pages show "最大2回". — [SMHN R8](https://smhn.info/202305-aquos-r8-os-update), [SMHN sense8](https://smhn.info/202310-sharp-aquos-sense-8-update), [Sharp sense5G business](https://jp.sharp/k-tai/aquos-sense5g/business/)
- [Official] STF lists support through Android 15 (SDK 35) — an example of OSS lagging the newest Android. — [DeviceFarmer/stf](https://github.com/DeviceFarmer/stf)

### Inferences
- Tool must support at least Android 13–17 in 2026–2027 fleets (business devices often ship 1–2 versions behind); keep an OEM×Android-version capability matrix (driver, settings keys, wizard behavior).
- OS update during kitting: multi-GB OTA per device (e.g., 3.8 GB One UI 9) × 40 devices saturates Wi-Fi; treat "OS update" as an optional module with explicit operator decision, ideally via carrier/OEM FOTA after delivery; record build fingerprint + security patch level in evidence.

### Gaps
- No Japan **business** (法人) device share data found; Pixel/Galaxy 7-year policy and arrows/Kyocera business update commitments not verified in this session.
- AQUOS/arrows/Xperia Android 17 schedules not found.

---

## Q10. Testing strategy (emulator/CI vs real devices; comparison with fabriq VM tests)

### Takeaway
The Android Emulator **cannot run hardware-accelerated inside another VM** (Google), so fabriq's VM-based test pattern does not transfer directly; emulator tests need a bare-metal Windows host with **WHPX** (AEHD sunsets 2026-12-31) or a Linux CI runner. Use `google_apis` (debuggable, `adb root`) images for logic tests; Play images are production builds without root. Emulators do not reproduce OEM drivers, hubs, setup wizards or OEM settings, so a small real-device lab remains mandatory.

### Cited Findings
- [Official] "You can't run a VM-accelerated emulator inside another VM, such as a VM hosted by VirtualBox, VMWare, or Docker. You must run a VM-accelerated emulator directly on your host computer." WHPX recommended (Windows 10 1803+); "Android Emulator hypervisor driver will be **sunset on December 31, 2026**"; AEHD requires Hyper-V features disabled. — [Configure hardware acceleration](https://developer.android.com/studio/run/emulator-acceleration)
- [Official, Microsoft] If Windows runs in a VM, nested virtualization must be enabled; forum reports of emulator boot taking hours in Windows VMs (ESXi/Azure). (snippet) — [MS .NET MAUI hardware acceleration](https://learn.microsoft.com/dotnet/maui/android/emulator/hardware-acceleration), [MS Q&A](https://learn.microsoft.com/en-us/answers/questions/1681555/running-android-studio-(and-specifically-an-androi)
- [Community] `google_apis` images allow `adb root`; `google_apis_playstore` are production builds ("adbd cannot run as root in production builds"). android-emulator-runner `target`: default / google_apis / playstore. — [8kSec](https://8ksec.io/rooting-an-android-emulator-for-mobile-security-testing), [ReactiveCircus/android-emulator-runner](https://github.com/ReactiveCircus/android-emulator-runner)
- [Official] Up to 16 emulators are auto-discovered by default (ports 5555–5585). — [adb docs](https://developer.android.com/tools/adb)

### Inferences
- Test pyramid: (1) PowerShell unit tests with a **fake adb** (recorded stdout fixtures incl. Japanese UTF-8 strings, offline/unauthorized states) — runs in fabriq's existing `run_tests.ps1` on PS 5.1; (2) emulator integration tests (google_apis, several API levels, ≤16 instances) on a bare-metal Windows box with WHPX or Linux CI; (3) monthly real-device regression on a 5–10 device matrix (Pixel, Galaxy, AQUOS, arrows, Kyocera rugged) through the actual hub topology.
- Note the conflict: enabling Hyper-V/WHPX on the kitting/test PC may interfere with other virtualization tools; keep the test host separate from fabriq's Hyper-V VM host if needed.

### Gaps
- Whether Hyper-V nested virtualization + WHPX gives usable emulator performance for CI on Windows — only anecdotes (negative).
- Behavioral differences of emulator vs OEM devices for `settings`/`dpm`/setup-wizard keys — not researched here (ADB researcher scope).
