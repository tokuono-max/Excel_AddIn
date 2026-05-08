# 共通仕様：JSON 定義（画面設定）

**共通仕様は次の 2 種類に分かれる。本ドキュメントは「JSON の共通仕様」。機能面（シートプロパティ・実行権復帰等）は docs/共通仕様_機能.md を参照すること。**

**対象**: 機能別画面設定ファイル（`config/ui_*.json`）の共通ルール・キー名・構造  
**作成日**: 2026-03-06  
**目的**: 各機能の JSON 定義を共通仕様としてまとめ、キー名・階層・値の意味を統一する。

---

## 1. 参照関係

| 文書 | 内容 |
|------|------|
| **本ドキュメント** | JSON の共通仕様（キー名・構造・ルール）の一覧。 |
| **docs/共通仕様_機能.md** | 機能面の共通仕様（シートプロパティ・VBA 実行権復帰・共通モジュール変更時のデグレ防止）。 |
| **config/README_ui_config.md** | 設定ファイルの配置・方針・ファイル名・読込タイミング・エラー時挙動。 |
| **各機能仕様書** | `docs/svc_csv_ld_仕様書.md` 等の「設定・config」で機能ごとの利用キーを記載。 |

---

## 2. ファイル名・配置

| 項目 | 内容 |
|------|------|
| **命名規則** | `ui_<機能キー小文字>.json`（例: ui_csv_mg.json, ui_csv_ld.json, ui_csv_sv.json, ui_csv_sp.json） |
| **配置** | プロジェクトルートの `config/` フォルダ。 |
| **読込** | `core.core_cst.get_ui_config_from_file_required(feature_key)`。機能キーは §7 の機能別 JSON 一覧に準ずる（csv_mg, csv_ld, csv_sv, csv_sp, hd_nr, undo, dupli, row_dl, col_dl, dt_ymd, dt_hm, trm_ex, help 等）。 |
| **読込タイミング** | 各画面表示時（create_dialog 等の呼び出し時）。core_cst がファイル mtime でキャッシュするため、保存後に次回表示で反映。 |
| **除外キー** | ルートの `_header` と `_separator` は実行時に設定として使わない（メタ情報・区切り用）。 |

---

## 3. 共通トップレベル構造

| キー | 説明 | 必須 |
|------|------|------|
| **_header** | メタ情報（File, Function, Created, Updated, Setting_details）。実行時は参照しない。 | 任意 |
| **_separator** | 区切り用。実行時は参照しない。 | 任意 |
| **WINDOW** | ウィンドウ共通の既定。サイズ・位置・タスクバー・最小化/最大化等。 | 機能による |
| **MAIN** | メイン画面（結合・分割・読込・保存のメインダイアログ）用。TITLE, WINDOW, 説明文, ボタン等。 | 機能による |
| **SCREENS** | サブ画面（進捗・完了・ワーニング・重複確認等）の設定。PROGRESS, DONE, WARNING, DUPLICATE 等。 | 機能による |
| **COMMON** | （csv_mg 等）全画面で共有する設定。WINDOW の既定・ICON ルール等。 | 機能による |

---

## 4. WINDOW 共通キー

各機能の WINDOW（トップレベルまたは MAIN.WINDOW / SCREENS.*.WINDOW）で共通に使うキー。

| キー | 型 | 説明 | 例 |
|------|-----|------|-----|
| **DEFAULT_WIDTH** | number | 初期幅（ピクセル）。0 は未指定。**適用時、ボタン等の配置に必要な最小幅より小さくはならない**（実装で max(設定値, minimumSizeHint().width()) とする）。 | 560 |
| **DEFAULT_HEIGHT** | number | 初期高さ（ピクセル）。0 は未指定。 | 420 |
| **RESIZABLE** | boolean | ユーザーがリサイズできるか。 | true / false |
| **CENTER_ON_EXCEL** | boolean | Excel ウィンドウを基準に中央配置するか。 | true / false |
| **SHOW_IN_TASKBAR** | boolean | タスクバーにアイコンを表示するか。false で Excel の子として非表示にすることが多い。 | true / false |
| **SHOW_MINIMIZE** | boolean | 最小化ボタンを表示するか。 | true / false |
| **SHOW_MAXIMIZE** | boolean | 最大化ボタンを表示するか。 | true / false |
| **SHOW_CLOSE_BUTTON** | boolean | ×閉じるボタンを表示するか。進捗画面では false にすることが多い。 | true / false |
| **STARTUP_POSITION** | string | 起動位置。center / center_on_excel / remember_last 等。 | "center" |
| **TOPMOST** | boolean | `WindowStaysOnTopHint` で最前面ヒント。他アプリの上に出やすい。 | true / false |
| **EXCEL_FRONT_FOLLOW** | boolean | Excel が前景のときにダイアログを前面へ追従するか（前景フック）。 | true / false |

- マージ順（csv_mg 等）: トップレベル WINDOW → COMMON.WINDOW → MAIN.WINDOW → SCREENS.*.WINDOW（画面ごとで上書き）。
- 機能により WINDOW に上記以外のキー（例: csv_mg の STORAGE_KEY）を追加してよい。未記載のキーは実装・機能仕様で定義する。

---

## 5. MAIN 共通キー

| キー | 型 | 説明 | 例 |
|------|-----|------|-----|
| **TITLE** | string | ウィンドウタイトル。 | "ファイル分割" |
| **DESC** / **DESCRIPTION** | string | 説明文（機能概要）。改行 `\n`・タブ `\t`（4文字空白）・文末 `\n`（下に1行空欄）は有効。§6.1 参照。 | "アクティブシートの…" |
| **DESC_VISIBLE** | boolean | 説明文を表示するか。 | true / false |
| **WINDOW** | object | メイン画面用の WINDOW 設定。上記 WINDOW キーと同じ構造。 | { "RESIZABLE": true } |
| **FILTER** | string | ファイルダイアログの種別フィルタ（csv_ld / csv_sv）。 | "CSVファイル (*.csv);;…" |
| **TABLE** | object | テーブル設定。TOOLTIP, COLUMNS（key, title, width の配列）等。 | 機能仕様参照 |
| **DIALOG_BUTTONS** | object | ボタン配置。LEFT / RIGHT 配列（id, label, tooltip）や OK / CANCEL 等。 | 機能仕様参照 |
| **DIALOGS** | object | サブダイアログ（FOLDER, ADD 等）の TITLE, FILTER, DEFAULT_BASE_FILENAME 等。FOLDER の USE_NATIVE: true で OS 標準（左にツリー）、false で Qt（右にファイル一覧）。 | 機能仕様参照 |

---

## 6. SCREENS 共通パターン

サブ画面（進捗・完了・ワーニング等）で共通に使うキー。

| 画面キー | 主なキー | 説明 |
|----------|----------|------|
| **PROGRESS** | TITLE, TOOLTIP, WINDOW | 進捗画面。WINDOW で SHOW_CLOSE_BUTTON: false 等。 |
| **DONE** | TITLE, MSG_HEADER, MSG_COUNT_PREFIX, ICON, ICON_SIZE, LIST_BG_SAME_AS_WINDOW, BTN_OK, BTN_OK_TOOLTIP, WINDOW | 完了通知。ICON_SIZE で S/M/L または数値。リスト背景を画面色にする場合は LIST_BG_SAME_AS_WINDOW: true。 |
| **WARNING** | TITLE, MSG / MESSAGE, ICON, ICON_SIZE, BTN_OK, BTN_OK_TOOLTIP, WINDOW | ワーニング通知。MSG は省略可（本文をコードで渡す場合あり）。MSG 指定時は改行 `\n` または `\\n`。ICON 未定義でアイコン非表示。ICON_SIZE で S/M/L または数値。 |
| **DUPLICATE** | （csv_mg）重複確認画面。 | 機能仕様参照 |
| **DATA_SHORTAGE** | TITLE, MSG, **SHORTAGE_CELL_BG_RGB**, ICON, ICON_SIZE, BTN_OK, WINDOW | （hd_nr）データ不足通知。SHORTAGE_CELL_BG_RGB は不足セルのシート上ハイライト色を [R, G, B] で指定（例: [255, 255, 224]）。実装は xlwings の Range.color に RGB タプルで設定する。 |

### 6.1 表示文字列の改行・タブ・文末改行（共通仕様）

- **文中の `\n`**: 有効。改行として表示する。JSON では `\n`（1文字の改行）またはリテラル `\\n`（バックスラッシュ＋n）のどちらでも指定可能。実装側で `_normalize_message_newlines` により `\\n` を改行に変換する。
- **文中の `\t`**: 有効。**タブは 4 文字分の空白**として表示する。実装側で `\t` を半角スペース 4 文字に変換する。
- **文末の `\n`**: 有効。**末尾で改行し、その下に 1 行分の空欄**を入れて表示する。説明文（DESC）の下にリスト等がある場合、DESC とリストの間に 1 行分の余白をレイアウトで確保する（`addSpacing` 等）。文字列の前後からは空白・タブのみ strip し、改行（`\n`）は strip しない。
- 対象キー: **DESC**, **DESCRIPTION**, **MSG**, **MESSAGE**, **MSG_HEADER** およびラジオ等の補助説明文（desc）など、設定から表示する文字列全般でこの共通仕様を適用する。

### 6.2 アイコン（ICON・ICON_SIZE）

- **ICON 未設定・空**: アイコン非表示。
- **ICON 設定値あり**: その標準アイコンを表示。**種類**は Information / Info, Warning / Warn, Critical / Error, Question のいずれか（大文字小文字は区別しない）。
- **ICON_SIZE**: アイコンの**大きさ**。`S`（小＝16px）, `M`（中＝24px）, `L`（大＝32px）のいずれか。数値（12〜48）でピクセル指定も可能。未設定時は M（24px）相当。

### 6.3 機能別に許容されるキー・画面

以下のキーは §6 の表にないが、機能ごとの拡張として許容する。トップレベル・SCREENS いずれも、本共通仕様の WINDOW / TITLE / ICON 等のキー名に揃える。

| 種別 | キー・画面 | 説明 |
|------|------------|------|
| **SCREENS** | **UNDO_DONE**, **UNDO_FAILED** | （ui_undo）復元成功・復元不可時の通知。DETAIL_TEXT で本文の固定部分を指定し、コードから渡す detail_text と結合して表示する。 |
| **SCREENS** | **CHOICE**, **NO_TARGET** | （ui_trm_ex）削除種別選択ダイアログ、削除対象なし時の通知。 |
| **SCREENS** | **HELP** | （ui_help）操作マニュアル表示窓。TITLE, BTN_CLOSE, WINDOW 等。 |
| **SCREENS** | **REPORT** | （ui_dupli）重複検出レポート画面。§7 で言及済み。 |
| **トップレベル** | **MESSAGES** | 機能別の文言（STATUS_DONE, ERROR_PREFIX 等）。§6.1 の改行ルールは「表示する文字列全般」に含め得る。 |
| **トップレベル** | **HIGHLIGHT** | （ui_dupli）着色設定。USE_CORE_ERR_BG, RGB 等。 |
| **表示キー** | **DETAIL_TEXT** | （ui_undo 等）本文の固定部分。§6.1 の改行・タブルールを適用する。 |

---

## 7. 機能別 JSON ファイル一覧

| ファイル | 機能 | 主なトップレベルキー |
|----------|------|----------------------|
| **ui_csv_mg.json** | CSV結合 | WINDOW, COMMON, MAIN（SCREENS: PROGRESS, DONE, DUPLICATE） |
| **ui_csv_ld.json** | CSV読込 | WINDOW, MAIN（TITLE, FILTER）, SCREENS.PROGRESS |
| **ui_csv_sv.json** | CSV保存 | WINDOW, MAIN（TITLE, FILTER）, SCREENS.PROGRESS |
| **ui_csv_sp.json** | CSV分割 | WINDOW, MAIN（TITLE, DESC, TABLE, DIALOG_BUTTONS, DIALOGS.FOLDER）, SCREENS（PROGRESS, DONE, WARNING） |
| **ui_hd_nr.json** | 行整形（ヘッダブロック横結合） | WINDOW, MAIN（TITLE, DESC, DIALOG_BUTTONS）, SCREENS（WARNING, PROGRESS, DONE, DATA_SHORTAGE） |
| **ui_undo.json** | 元に戻す（復元不可時） | SCREENS.UNDO_FAILED 等。 |
| **ui_dupli.json** | 重複チェック | MESSAGES, HIGHLIGHT, SCREENS（PROGRESS, DONE, REPORT） |
| **ui_row_dl.json** | 空白行削除 | MESSAGES, SCREENS（PROGRESS, DONE） |
| **ui_col_dl.json** | 空白列削除 | MESSAGES, SCREENS（PROGRESS, DONE） |
| **ui_dt_ymd.json** | 日付変換 YYYY/MM/DD | MESSAGES, SCREENS（PROGRESS, DONE, WARNING） |
| **ui_dt_hm.json** | 日付・時刻変換 YYYY/MM/DD HH:MM | MESSAGES, SCREENS（PROGRESS, DONE, WARNING） |
| **ui_trm_ex.json** | 文頭・文末トリム | MESSAGES, SCREENS（CHOICE, DONE, NO_TARGET） |
| **ui_help.json** | 操作マニュアル（Info.txt 表示） | MESSAGES, SCREENS.HELP |
| **ui_data_agg.json** | データ集約・クレンジング | WINDOW, MAIN, MESSAGES, SCREENS（PROGRESS, DONE, WARNING, STEP_POPUP, **SCENARIO_EDIT** ほか）。**シナリオ編集の詳細キー（DETAIL_CELL / DETAIL_NAME 等）の意味の正本は「データ集約ツール要求定義書.md」§12**（実体は `config/ui_data_agg.json`）。 |

読込対象の機能キーは上記一覧に準ずる（csv_mg, csv_ld, csv_sv, csv_sp, hd_nr, undo, dupli, row_dl, col_dl, dt_ymd, dt_hm, trm_ex, help, data_agg 等）。

---

## 8. 新規機能の JSON を追加するとき

1. **ファイル名**: `config/ui_<機能キー小文字>.json` とする。
2. **読込**: `core_cst.get_ui_config_from_file_required("<機能キー>")` で読む。`_header` / `_separator` は設定として使わない。
3. **キー名**: 本共通仕様の WINDOW / MAIN / SCREENS のキー名に揃える（TITLE, DEFAULT_WIDTH, ICON, BTN_OK 等）。
4. **詳細**: 機能仕様書の「設定・config」に、その機能で利用するキーと意味を記載する。**データ集約（data_agg）**は要求定義書 **§12** を正とする。

---

## 9. 画面の共有化（共通コンポーネント）

ワーニング・完了通知・フォルダ選択・ファイル選択は共通モジュールで提供し、機能ごとの JSON で見た目を制御する。

| 画面種別 | 共通モジュール | 説明 |
|----------|----------------|------|
| **ワーニング** | `ui_qt.ui_common` | `create_warning_dialog(req, parent_hwnd, warning_cfg)`。warning_cfg は SCREENS.WARNING 相当（TITLE, MSG, ICON, BTN_OK, WINDOW）。 |
| **完了通知** | `ui_qt.ui_common` | `create_done_dialog(req, parent_hwnd, parent_widget, done_cfg)`。done_cfg 未指定時は CSV_MG 用の既定を使用。 |
| **フォルダ選択** | `ui_qt.ui_fld` | `show_folder_dialog(parent, title, initial_dir, config)`。USE_NATIVE で OS 標準 / Qt カスタム（左ツリー＋右一覧）を切替。 |
| **ファイル選択（開く/保存）** | `ui_qt.ui_fil` | `show_open_file_dialog(...)` / `show_save_file_dialog(...)`。標準 OS ダイアログのラッパ。 |

**適用対象（現時点）**: csv_ld, csv_sv, csv_sp, dupli, row_dl, col_dl, dt_ymd, dt_hm に適用。**csv_mg は今回の共有化から除外**し、詳細動作確認が終わるまで従来実装のままとする。適用後も各機能の JSON（SCREENS.WARNING, SCREENS.DONE, MAIN.DIALOGS.FOLDER, MAIN.TITLE/FILTER 等）で文言・アイコン・サイズを定義する。各 JSON の `_header.Shared_UI`（任意）に、当該機能で利用する共通モジュールの説明を記載できる。

---

## 10. 参照

- 設定の配置・方針・エラー時: **config/README_ui_config.md**
- 機能面の共通仕様・共通モジュール変更時: **docs/共通仕様_機能.md**（同文書内で 共通モジュール変更時_デグレ防止.md を参照）
- 各機能の設定利用: **docs/svc_csv_ld_仕様書.md**, **docs/svc_csv_sv_仕様書.md**, **docs/svc_csv_sp_仕様書.md** および CSV結合の仕様
