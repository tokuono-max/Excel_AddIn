# リボンから Python（hc_main.invoke）までの呼び出し経路

Excel アドインのリボン操作から、`hc_main.invoke` を経由して Python 側が動くまでの**メソッド／モジュール単位**の流れをまとめた仕様です。

**対象バージョン目安**: `hc_main` 1.11.0 以降、`VBA\Main.bas` 2.1.0 以降（リボンは `Main.RibbonCallback_hc_main` のみ、各 button の **tag 必須**。Python 公開入口は `invoke` のみ）。

---

## 1. 全体像（1 行の要約）

```
customUI（xlam 内） → Main.RibbonCallback_hc_main → Main.RibbonInvokeFromControl
  → ExcelUtil.GetSheetIdSafe / control.Tag（action）→ Main.RunPythonSafe → RunPython（xlwings）
  → hc_main.invoke → _invoke_impl_* →（分岐）_call_svc_server または svc.* 直接呼び出し
```

---

## 2. リボン（customUI）側


| 項目             | 内容                                                                                        |
| -------------- | ----------------------------------------------------------------------------------------- |
| **onAction**   | **すべてのカスタム button で同一**: `Main.RibbonCallback_hc_main`（標準モジュール名 `Main` を付与）               |
| **tag**        | **必須**。`hc_main.invoke` の **action** 文字列と完全一致（例: `load_csv`, `merge_csv`, `run_data_agg`）。**例外**: `check_for_updates` は **svc ブリッジを使わず** `RunPython` で `core.packaged_update.check_for_updates_interactive` を直接呼ぶ（`VBA\Main.bas`）。 |
| **リポジトリ上のソース** | `CSV_Tool_xml.txt`（xlam の `customUI*.xml` へ手動マージ想定）                                       |


`tag` が空の場合、VBA はログを出して処理せず終了する。`CSV_Tool_xml.txt` を修正すること。

### 2.1 待機 UserForm（WaitForm）

- **表示**: `Main.RibbonInvokeFromControl` 内、`RunPythonSafe` の直前に `HC_WaitForm.BeginWaitForRibbon(control.ID, act)`。ボタン **id ごと**に **WaitForm を出すか**（`ShowWaitForm`）と **日本語表示名**（`DisplayName`）を `HC_WaitForm` 内の **1 か所の `Select Case`**（`RibbonWaitInfo`）で定義する（`CSV_Tool_xml.txt` の各 `button id` と一致させる）。`BeginWaitForRibbon` は `ShowWaitForm` が False のときは UserForm を出さず即 return する。表示名だけが必要な場合は `RibbonDisplayNameFromControlId`（内部で `RibbonWaitInfo` と同じ対応表を参照）。**位置**: `WaitForm` の `UserForm_Activate` で `Application.Left/Top/Width/Height` を使い **Excel メインウィンドウの中央**に配置（`StartUpPosition = 0` 手動）。
- **未登録 id**: `RibbonWaitInfo` の `Case Else` で `ShowWaitForm = True`、`DisplayName = "処理"`（リボンに無い id が来た場合も従来どおり待機表示）。
- **閉じる**: Python から `Application.Run "HC_WaitForm.NotifyUiReady"`（`core.core_cursor.notify_wait_form_ready` または既存の `notify_ui_ready` 内で同時実行）。**× ボタンは `UserForm_QueryClose` で禁止**。
- **タイムアウト**: 30 秒で `HC_WaitForm.WaitFormTimeout`（OnTime）。
- **エラー時**: `Main` の `RunPythonSafe` / `RibbonInvokeFromControl` の `ErrorHandler` で `NotifyUiReady`。

---

## 3. VBA 呼び出しチェーン（メソッド名）

```
[Excel リボン ボタン押下]
    ↓
Main.RibbonCallback_hc_main(ByVal control As Object)     … Public、リボン onAction の唯一の入口
    ↓
Main.RibbonInvokeFromControl(ByVal control As Object)     … Private、実処理の集約
    │
    ├─ ExcelUtil.GetSheetIdSafe(ActiveSheet)  →  sheet_id（sId）
    ├─ act = Trim$(control.Tag)   … 空ならエラーログして終了（tag 必須）
    ├─ HC_WaitForm.BeginWaitForRibbon(control.ID, act)
    └─ Main.RunPythonSafe(act, sId)
           ↓
Main.RunPythonSafe(ByVal methodName As String, ByVal sId As String)   … Public
    │
    ├─ methodName = 上記 act（= hc_main.invoke の action）
    ├─ PyEscSq(methodName), PyEscSq(sId) で Python 単一引用符用エスケープ
    ├─ sCmd = "import hc_main; hc_main.invoke(action='…', target_hwnd=…, sheet_id='…')"
    └─ RunPython sCmd          … xlwings が提供する VBA の RunPython（参照元は環境依存）
           ↓
[別プロセスの Python が sCmd を 1 行実行]
```

### 補助プロシージャ


| 名前                        | 公開      | 役割                                                          |
| ------------------------- | ------- | ----------------------------------------------------------- |
| `PyEscSq`                 | Private | `action` / `sheet_id` に `\` や `'` が含まれても RunPython 文字列を壊さない |
| `RibbonInvokeFromControl` | Private | `sId` 取得、`control.Tag` から `act` を取得、`RunPythonSafe` 呼び出し    |


---

## 4. Python 入口（hc_main）


| 名前                                                                | 役割                                                                                                  |
| ----------------------------------------------------------------- | --------------------------------------------------------------------------------------------------- |
| `hc_main.invoke(action, target_hwnd=..., sheet_id=..., **kwargs)` | VBA および他 Python からの**単一受け口**。許可された `action` のみ `_INVOKE_HANDLER_MAP` でディスパッチ（`getattr` による任意実行はしない） |
| `_INVOKE_HANDLER_MAP`                                             | `action` 文字列 → `_invoke_impl_`* の対応表（単一ソース）                                                         |
| `INVOKE_ACTIONS`                                                  | 上記マップのキー集合（許可 action の列挙に利用可能）                                                                      |


`load_csv` 等の旧トップレベル関数は **提供しない**（1.11.0 以降）。

---

## 5. invoke 後の分岐（2 系統）

`_invoke_impl_`* の実装は、次のどちらか（または両方のパターンが混在）です。

### 5.1 常駐プロセスへ IPC: `_call_svc_server`

`ensure_svc_server()` のあと、依頼を pickle で書き出し、**svc_server** 側が `action`（**内部名**）に応じて処理します。

**代表例（VBA の action → 内部 action の対応）**


| invoke の action（リボン tag と同一） | 主な内部 action（`_call_svc_server` 第 1 引数） |
| ---------------------------- | -------------------------------------- |
| `load_csv`                   | `csv_ld`（`ACTION_CSV_LD`）              |
| `save_csv`                   | `csv_sv`                               |
| `merge_csv`                  | `csv_mg`                               |
| `split_csv`                  | `csv_sp`                               |
| `trim_spaces`                | `trm_ex`                               |
| `normalize_header`           | `hd_nr`                                |
| `insert_shuka_header`        | `hd_in`                                |
| `undo_last_action`           | `undo`                                 |
| `show_help`                  | `help`                                 |
| `run_data_agg`               | `data_agg`（`payload` は kwargs 経由で渡しうる） |


内部 action の許可集合は `hc_main._ALLOWED_SVC_ACTIONS`（`svc/svc_server.py` の `_ACTION_MAP` と同期が前提）。

### 5.2 同一 Python プロセスで `svc` パッケージを直接 import

次の invoke action は、`**_call_svc_server` を経由せず**、対応モジュールの関数を直接呼びます。


| invoke の action       | Python モジュール     | 呼び出す関数                |
| --------------------- | ---------------- | --------------------- |
| `check_duplicates`    | `svc.svc_dupli`  | `check_duplicates`    |
| `delete_empty_rows`   | `svc.svc_row_dl` | `delete_empty_rows`   |
| `delete_empty_cols`   | `svc.svc_col_dl` | `delete_empty_cols`   |
| `convert_date_ymd`    | `svc.svc_dt_ymd` | `convert_date_ymd`    |
| `convert_date_ymd_hm` | `svc.svc_dt_hm`  | `convert_date_ymd_hm` |


---

## 6. 具体例: 「ファイル結合」ボタン（merge_csv）

```
リボン button（onAction="Main.RibbonCallback_hc_main", tag="merge_csv"）
  → Main.RibbonCallback_hc_main
  → Main.RibbonInvokeFromControl
       act = "merge_csv", sId = GetSheetIdSafe(...)
  → Main.RunPythonSafe("merge_csv", sId)
  → RunPython("import hc_main; hc_main.invoke(action='merge_csv', target_hwnd=…, sheet_id='…')")
  → hc_main.invoke(...)
  → hc_main._invoke_impl_merge_csv
  → hc_main._call_svc_server("csv_mg", book_ptr, sheet_id)
  → （IPC）svc.svc_server 側で csv_mg 処理
```

---

## 7. RunPython 実行後の VBA（参考）

`RunPython` が戻ったあと、`RunPythonSafe` 内で次を実行します。

- `Main.CheckAndNotifyVBA(sId)` … シートプロパティ `HC_NOTIFY_RETV` 等の通知
- `HC_Bridge.RestoreStatBar` … ステータスバー同期

詳細は `docs\VBA_Python_実行権とシートプロパティ_解析.md` などを参照。

---

## 8. 関連ファイル


| 種別                 | パス                                                                |
| ------------------ | ----------------------------------------------------------------- |
| Python 司令塔         | `hc_main.py`（`invoke`, `_INVOKE_HANDLER_MAP`, `_call_svc_server`） |
| VBA ゲートウェイ         | `VBA\Main.bas`                                                    |
| リボン XML ソース（リポジトリ） | `CSV_Tool.xlam.txt`                                               |
| 取り込み先（運用）          | `excel_addin\CSV_Tool.xlam` 内 customUI                            |


---

## 9. 変更履歴（本ドキュメント）


| 日付         | 内容                                                                                            |
| ---------- | --------------------------------------------------------------------------------------------- |
| 2026-04-06 | 2.1.0 / hc_main 1.11.0 整合: onAction・tag 必須、`ActionForRibbonControlId`・`Call*`・Python ラッパ削除を反映 |
| 2026-04-05 | 初版（invoke 一元化・単一リボンコールバック前提で整理）                                                               |


