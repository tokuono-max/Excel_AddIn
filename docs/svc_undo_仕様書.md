# svc_undo 機能仕様書（元に戻す）

**対象**: 元に戻す（Undo）— キャッシュに保存したシート状態の復元および復元用スナップショットの保存  
**作成日**: 2026-03-09  
**目的**: 機能概要・用語・保存先・復元フロー・共通事項・呼び出し元を定義する。

---

## 1. 機能概要

**元に戻す（Undo）** は、破壊的処理（CSV 読込・行整形・トリム等）の**直前**に保存したシートの有効データを、Pickle キャッシュから読み出し、シート (1,1) から上書きして復元する機能である。

- **復元（exec_undo）**: VBA のリボン等から「元に戻す」が実行されたとき、対象ブック・シートに対応するキャッシュを検索し、見つかれば `data` をシートに書戻す。成功後は当該キャッシュエントリを削除する（1 世代のみ）。
- **スナップショット保存（save_undo_snapshot）**: 破壊的処理を行う各機能が、ユーザーが確認ダイアログで OK した時点で呼び出す。シートの UsedRange を 1 回読込み、キャッシュに保存する。これにより後から「元に戻す」で復元可能になる。

本モジュールは **docs/共通仕様_機能.md** の「5. 元に戻す（Undo）の共通仕様」に準拠し、**core_xlc / core_stat / core_sys / core_w32** を使用する。**復元処理中**は `ensure_ui_server` により **ui_qt.ui_undo** の進捗ダイアログ（`config/ui_undo.json` の `SCREENS.PROGRESS`、キャンセルなし）を表示する。完了／失敗のモーダルは従来どおり。

---

## 2. 語句定義

| 用語 | 定義 |
|------|------|
| **キャッシュ** | core_sys.CacheManager が管理する Pickle 永続化。実体は Windows 一時フォルダ（%TEMP%）内の core_cst.CACHE_FILE_NAME（例: header_converter_cache.pkl）の 1 ファイル。キーごとにペイロードを格納。 |
| **Undo キー** | ブック・シートを一意に識別するキャッシュキー。`{hwnd}_{pid}_{ブック名}_{シート名}` をファイル名禁止文字（`\ / : * ? " < > \|`）で `_` に置換した文字列。 |
| **データペイロード** | 通常の Undo 用。`{"data": list_2d, "book_name": str, "sheet_name": str}`。`data` はシート (1,1) から復元する 2 次元リスト。 |
| **構造ペイロード** | ヘッダ集約等の専用復元用。`num_rows` キーを持つペイロード。hc_hd_rs 等の別モジュールに復元を委譲する。本モジュールでは委譲先が無い場合はエラー表示。 |

---

## 3. 共通事項

| 項目 | 内容 |
|------|------|
| **共通仕様** | docs/共通仕様_機能.md の「5. 元に戻す（Undo）の共通仕様」に従う。 |
| **ステータス報告** | core_stat.set_status_info で HC_STATUS_INFO にメッセージを保存。VBA の RestoreStatBar 等で参照される。キー名は直書きせず core_stat の API を使用。 |
| **ログ** | 処理の起点で core_log により `[UNDO] exec_undo start ...` / `[UNDO] save_undo_snapshot start ...` を出力。エラー・警告・保存成功時もログに残す。 |
| **Excel 操作** | exec_undo 実行中は Interactive / ScreenUpdating を False にし、完了・異常時は finally で True に復帰。復元後に Excel を前面に表示（core_w32.bring_to_front）。 |
| **進捗 UI** | キャッシュ読込後、データ復元・構造復元のいずれも **IPC 進捗**（`progress/progress_undo_*.pkl` をポーリング）を表示。完了後に進捗を閉じ、`UNDO_DONE`／`UNDO_FAILED` を表示。 |

---

## 4. 保存先とキー・ペイロード

| 項目 | 内容 |
|------|------|
| **保存先** | core_sys.CacheManager。実体パスは `tempfile.gettempdir()` + core_cst.CACHE_FILE_NAME。 |
| **キー生成** | `_make_undo_key(hwnd, wb_name, sh_name)`。hwnd は Excel ウィンドウハンドル、pid は os.getpid()、wb_name / sh_name はブック名・シート名。 |
| **データペイロード** | save_undo_snapshot が保存する形式。`data`: シート UsedRange を (1,1) 起点で読んだ 2 次元リスト。exec_undo は**表示停止→クリア（ClearContents 優先）→値設定→余白クリア→解除**の順で実行。復元前に旧 UsedRange 寸法を取得し、書込後に保存データより大きい範囲の余白をクリアして有効データ領域の拡大を防ぐ。 |
| **構造ペイロード** | `num_rows` キーを持つペイロード。exec_undo は hc_hd_rs.restore_header_logic に委譲（モジュールが存在する場合）。ImportError または実行時例外の場合はステータスにエラーを表示。 |

---

## 5. save_undo_snapshot（スナップショット保存）

### 5.1 役割

破壊的処理の**前**に、呼び出し元（csv_ld, hd_nr 等）が「ユーザーが確認で OK した時点」で実行する。シートの有効データを 1 回で読込み、キャッシュに保存する。

### 5.2 引数

| 引数 | 型 | 説明 |
|------|-----|------|
| book | Any | 対象 Excel ブック（xlwings Book） |
| sheet_id | str | シートの GUID（HC_GUID_B64）。空の場合はアクティブシートを対象とする |
| target_hwnd | Optional[int] | Excel ウィンドウハンドル。キー生成に使用 |
| excel_hwnd | Optional[int] | target_hwnd の代替。キー生成に使用 |

### 5.3 処理概要

1. book / sheet を取得。sheet は _get_sheet(book, sheet_id) で解決（GUID 検索またはアクティブ）。
2. CacheManager が利用できない場合は False を返す。
3. シートの UsedRange を取得。無い場合は空リストをペイロードの `data` とする。
4. 有る場合は `ptr_s.range((1, 1), (last_row, ncols)).value` で一括読込し、2 次元リストに正規化（1 行の場合は [list] に包む）。
5. ペイロード `{"data": list_2d, "book_name": wb_name, "sheet_name": sh_name}` を _make_undo_key で生成したキーで CacheManager.save する。
6. 成功時 True、失敗時（シート取得失敗・読取失敗・保存失敗）False を返す。

### 5.4 戻り値

- **True**: 保存成功。
- **False**: 保存失敗（book が None、シート未検出、CacheManager 不可、UsedRange 読取失敗、CacheManager.save 失敗のいずれか）。

---

## 6. exec_undo（復元執行）

### 6.1 呼び出し元

- **svc_server**: action `"undo"` で、book / sheet_id / target_hwnd（excel_hwnd）を渡して exec_undo を呼ぶ。
- **hc_main**: undo_last_action は exec_undo のエイリアス。新方式では svc_server 経由で book 渡しとなる想定。

### 6.2 引数

| 引数 | 型 | 説明 |
|------|-----|------|
| book | Any | 対象 Excel ブック |
| sheet_id | str | シートの GUID。空の場合はアクティブシート |
| target_hwnd | Optional[int] | Excel ウィンドウハンドル |
| excel_hwnd | Optional[int] | target_hwnd の代替 |

### 6.3 処理フロー

1. **起点ログ**: `[UNDO] exec_undo start book=... sheet_id=...` を出力。
2. **ブック・シート取得**: book が None の場合は警告ログで終了。ptr_s = _get_sheet(book, sheet_id)。シートが無い場合はステータスに「シートを特定できませんでした」を設定して終了。
3. **キー生成**: _make_undo_key(hwnd, wb_name, sh_name) で str_undo_key を生成。
4. **キャッシュ読込**: CacheManager.load(str_undo_key)。None の場合はステータスに「元に戻すための物理キャッシュ情報が見つかりませんでした。」を設定して終了。
5. **ペイロード種別判定**: `"num_rows" in dict_undo_payload` で構造ペイロードかどうかを判定。
6. **構造ペイロードの場合**:  
   - hc_hd_rs.restore_header_logic(target_hwnd=hwnd) を実行。成功時はステータス "UI"。  
   - ImportError の場合は「構造復元モジュールが利用できません。」をステータスに設定。  
   - その他例外の場合は「ヘッダ復元失敗 Detail: ...」をステータスに設定。
7. **データペイロードの場合**:  
   - list_data = dict_undo_payload.get("data", [])。空の場合は何もせず終了。  
   - **表示停止**: Interactive / ScreenUpdating を False に設定（高速化のため最初に実施）。  
   - **クリア**: 現状の UsedRange 寸法（old_rows, old_cols）を取得後、ClearContents（または value=None）でクリア。  
   - **値設定**: core_xlc.write_chunk(ptr_s, 1, 1, list_data) で (1,1) から書込。  
   - **余白クリア**: 保存データ寸法（saved_rows, saved_cols）より大きい範囲（右側・下側）をクリアし、有効データ領域の拡大を防止。  
   - UsedRange の列の AutoFit を試行（失敗時は無視）。  
   - ステータスに「加工直前の物理状態へ正常に復元されました。」を設定。  
   - finally で Interactive / ScreenUpdating を True に復帰。
8. **キャッシュ削除**: 上記いずれかの復元が行われた後、CacheManager.delete(str_undo_key) で当該キーを削除。
9. **例外時**: 復元処理中の例外はログに記録し、ステータスに「ERROR: 元に戻す処理中に例外。 Detail: ...」を設定。delete は成功時のみ実行されるため、例外時はキャッシュは残る。
10. **finally**: Interactive / ScreenUpdating の復帰、bring_to_front(hwnd)、core_stat.get_status_info でステータスを取得し Excel の StatusBar に反映。

### 6.4 画面・UI（復元不可時）

- **復元できない場合**（キャッシュなし・シート未検出・キャッシュモジュール不可・構造復元モジュール不可・データ空・処理中例外）は、既存のステータス通知（core_stat.set_status_info）に加え、**ui_common.DoneDialog** で「復元できません」旨を表示する。
- **「物理キャッシュが見つからない」の主因**: 破壊的処理（例: hd_nr 行整形）で **save_undo_snapshot を呼んでいない**と、キャッシュが作成されず exec_undo で復元できない。共通仕様どおり、**ユーザーが確認で OK した直後・シート変更の直前**に save_undo_snapshot(book, sheet_id=..., target_hwnd=..., excel_hwnd=...) を呼ぶこと。hd_nr は 2.2.0 でこの呼び出しを追加済み。
- 実装: ensure_ui_server() のあと、module=ui_qt.ui_common・action=undo_failed・req_dict.detail_text=メッセージ で req_*.pkl を投入し、res_*.pkl の完了を最大 60 秒待つ。ui_common.create_dialog は action=undo_failed で **config/ui_undo.json** の **SCREENS.UNDO_FAILED** を読み、DoneDialog の表示内容（TITLE, MSG_HEADER, WINDOW, BTN_OK, ICON 等）を設定する。
  - 本文は、まず ui_undo.json の `SCREENS.UNDO_FAILED.DETAIL_TEXT` を先頭に表示し、その後に svc_undo 側から渡されたエラー詳細（req_dict.detail_text）を空行を挟んで追記する。DETAIL_TEXT 未設定時は、従来どおり req_dict.detail_text のみを表示する。
- 通常の復元成功時は専用の Qt 画面は出さず、ステータスバーおよびログで通知する。

---

## 7. ログ出力

| タイミング | 内容 |
|------------|------|
| exec_undo 開始 | `[UNDO] exec_undo start book=... sheet_id=...` |
| save_undo_snapshot 開始 | `[UNDO] save_undo_snapshot start book=... sheet_id=...` |
| 保存成功 | `[UNDO] save_undo_snapshot saved key=... rows=...` |
| キャッシュなし | `[UNDO] 元に戻すための物理キャッシュ情報が見つかりませんでした。` |
| 構造復元モジュールなし | `[UNDO] structure undo: hc_hd_rs not available` |
| その他エラー | logger.warning / logger.exception でメッセージと例外を出力 |

---

## 8. 設定・モジュール・参照

| 項目 | 内容 |
|------|------|
| **サービス** | svc.svc_undo（exec_undo, save_undo_snapshot, undo_last_action） |
| **内部** | _make_undo_key, _get_sheet |
| **依存 core** | core_xlc（find_sheet_by_guid, write_chunk）, core_stat（set_status_info, get_status_info）, core_sys（CacheManager）, core_w32（bring_to_front）, core_log |
| **設定ファイル** | config/ui_undo.json（復元不可画面 SCREENS.UNDO_FAILED: TITLE, MSG_HEADER, DETAIL_TEXT, WINDOW, BTN_OK, ICON 等） |
| **共通仕様** | docs/共通仕様_機能.md 第 5 節 |
| **キャッシュ定義** | core/core_sys.py（CacheManager）, core/core_cst.py（CACHE_FILE_NAME） |

---

## 9. 破壊的処理側の責務

CSV 読込・行整形・トリム等、シート内容を上書きまたは削除する機能では、**ユーザーが確認ダイアログで OK した直後**に、次のいずれかでスナップショットを保存すること。

- `from svc.svc_undo import save_undo_snapshot`  
- `save_undo_snapshot(book, sheet_id=..., target_hwnd=..., excel_hwnd=...)`

戻り値が False の場合でも処理を続行してよいが、その場合は「元に戻す」で復元できない可能性がある。必要に応じてログやステータスで通知する。
