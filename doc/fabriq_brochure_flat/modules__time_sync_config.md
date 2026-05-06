# time_sync_config (Standard)

**カテゴリ**: System
**メニュー名**: Time Sync
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 実装あり（同期ソース確認、最大 5 回リトライ）
**サブスクリプト**: なし

## 目的
Windows Time サービス (W32Time) に NTP サーバを設定し、時刻同期をトリガーするモジュール。
`w32tm /config /manualpeerlist` で複数サーバを構成し、`/resync /force` で同期を実行した上で、
同期ソースが「Local CMOS Clock」等のローカル系から外れたかを `w32tm /query /status` で
確認する。CSV の Enabled=1 が 0 件のときは「同期のみモード」として、既存の NTP 設定を
保ったまま `/resync /force` のみを実行する 2 モード設計。

## 入力 (CSV)
`time_sync_list.csv`
- `Enabled`: 1=NTP サーバとして使用 / 0=リスト表示のみ（同期のみモードの判定にも使用）
- `NtpServer`: NTP サーバアドレス（例 `time.windows.com`, `ntp.nict.jp`）
- `Description`: 説明
- `Segment`: Segment フィルタ（任意）

デフォルト同梱: time.windows.com (Enabled=1) / time.google.com (Enabled=0) /
ntp.nict.jp (Enabled=0)。

## 主要ステップ
1. `time_sync_list.csv` 全行読み込み（Enabled フィルタなし、表示と判定の両方に使用）
2. W32Time サービス存在確認
3. ドライラン表示:
   - W32Time 現状 (Status / StartType)
   - 各 NTP サーバへの ICMP 疎通 (`[REACHABLE]` / `[UNREACHABLE]`)
4. 実行確認（AutoPilot 自動 Y）
5. 適用:
   - 5-1. W32Time を Start + StartupType=Automatic
   - 5-2. ICMP 確認（表示のみ、ブロック環境でも継続）
   - 5-3. Enabled=1 の行があれば `w32tm /config /manualpeerlist:"server,0x9 ..." /syncfromflags:manual /reliable:yes /update`
     （`,0x9` = SpecialPollInterval + Client）→ 3 秒待機
   - 5-4. 同期実行（最大 5 回リトライ、3 秒間隔）: `w32tm /resync /force` →
     `w32tm /query /status` でソース確認 → ローカル系以外なら成功
   - 5-5. 最終同期状態を表示
6. Status 判定: ソース切替成功なら Success、5 回失敗なら Partial

## 注意点・運用メモ
- 管理者権限必須
- 標準モード実行時は既存 NTP 設定を上書き
- Hyper-V VM では Time Synchronization Integration Service（ホスト同期）が優先されることがあり、
  w32tm 設定が無視される
- ドメイン参加 PC は通常 DC が NTP ソースになる。意図的に外部 NTP を使う場合のみ本モジュール利用
- ICMP ブロック環境ではドライラン疎通表示が `[UNREACHABLE]` になるが、UDP/123 が通っていれば
  実際の同期は成功する（ICMP は情報表示のみで処理継続）
- 起動直後の同期はサービス初期化遅延で 1-2 回リトライしてから成功するパターンが多い

## 検証
Step 5.4 が事実上の Post-Apply Verification:
- `w32tm /query /status` を再実行して `Source` フィールドを取得
- ソースが「Local CMOS Clock」「LOCL」「Free-running System Clock」のいずれでもなければ
  外部 NTP に切り替わったと判定して成功
- 最大 5 回 (3 秒間隔) リトライして切替えなければ Partial
  （NTP 経路未疎通や起動直後の一時的遅延で発生しうるが、設定自体は正しく適用されているため
  致命的失敗扱いにはしない設計判断）

`-Verified` 引数の渡し方は実装上 Status と連動するため、Success の場合に Verified=true 相当の
履歴が残る。Partial は途中までの同期試行を記録としては残す。
