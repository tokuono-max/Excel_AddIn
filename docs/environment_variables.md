# CSV Tool：ログファイルと環境変数

本書は **どのログがどこにあり、何のために使うか** を中心に整理し、それに紐づく **環境変数** を同じ流れで説明します。ログ行の書式や出力内容の細部は扱いません。

---

## 1. ログの保存場所（基準パス）

| 項目 | 内容 |
|------|------|
| 共有ログのフォルダ | Windows の **`%TEMP%\csv_tool\`**（VBA・Python ともここを既定とする） |
| パス例 | `C:\Users\<ユーザー名>\AppData\Local\Temp\csv_tool\` |
| IPC ルートを変えた場合 | 環境変数 **`HC_IPC_ROOT`**（別名 **`HC_QT_IPC_DIR`**）で pickle や制御フラグのルートを変えられます。**上記 3 本の共有ログ**（`hc_csv*.log`）は通常 **引き続き `%TEMP%\csv_tool\`** に出力されます。 |
| 子プロセス用の別フォルダ | IPC ルート配下の **`logs\`** に、起動時のブートログ（ファイル名にタイムスタンプが付く個別ファイル）が作成されることがあります。既定 IPC なら **`%TEMP%\csv_tool\logs\`** です。 |

---

## 2. 共有ログファイル（3 種類）

VBA（`HC_Log`）と Python（`core_log`）が **同じファイル名** で追記します。用途に応じて **どのファイルを開くか** を選びます。

| ファイル名 | フルパス（既定） | 用途（いつ使うか） | 環境変数（関連） |
|------------|------------------|-------------------|------------------|
| **`hc_csv.log`** | `%TEMP%\csv_tool\hc_csv.log` | **運用ログ**。通常の動作確認、ユーザー操作に伴う記録、エラー報告の一次情報。 | **常時**出力。ON/OFF 用の環境変数はなし。 |
| **`hc_csv_diag.log`** | `%TEMP%\csv_tool\hc_csv_diag.log` | **診断ログ**。不具合の深掘り、内部経路の追跡、データ集約まわりの調査など。**普段は無効**でよい。 | **`HC_LOG_DIAG=1`** で有効。別名 **`HC_DEBUG=1`** でも有効（後方互換）。データ集約用の **`HC_DIAG_DATA_AGG_*`**（別名 **`DATA_AGG_*`**）のいずれかを有効にした場合も、診断ログファイルへの出力が有効になります。CSV 分割・同名確認の HWND 調査用 **`HC_CSV_SP_CONFLICT_HWND_DIAG=1`** のみでも有効（**`[CONFLICT_HWND_DIAG]`**）。**`HC_CSV_SP_CONFLICT_HWND_DIAG_TREE=1`** を併用すると Excel 配下の子孫 HWND スナップショット **`[CONFLICT_EXCEL_DESC]`**（ログ肥大に注意）。**`HC_UI_FG_DIAG=1`** のみでも有効（**`[UI_FG]`**・フォアグラウンド HWND／Excel・ダイアログの PID・クラス名・`GW_OWNER`・`SetForegroundWindow` 成否など。`ui_common` の `prepare_dialog`／`ensure_front`／行・列削除完了ダイアログ）。**`HC_UI_WINDOW_CAPTION_DIAG=1`** のみでも有効（**`[UI_CAPTION]`**・`apply_window_config` の最小化/最大化の設定解釈・Qt ウィンドウフラグ・Win32 `GWL_STYLE` と min/max ボックス bit の遅延サンプル・`set_window_style_remove_min_max` 発火）。いずれも `core_env.diag_log_file_enabled()` と整合。 |
| **`hc_csv_perf.log`** | `%TEMP%\csv_tool\hc_csv_perf.log` | **計測ログ**。起動や処理の所要時間の比較・改善用。**普段は無効**でよい。 | **`HC_LOG_PERF=1`** のときのみ出力。CSV 読込の区間定義・ログ対応は **`docs/csv_ld_perf_measurement.md`**。 |

**使い分けの目安**

- 現場サポート・通常運用 → **`hc_csv.log`**
- 再現調査・詳細トレース → **`hc_csv_diag.log`**（上記環境変数を立ててから再現）
- 速度改善・ボトルネック把握 → **`hc_csv_perf.log`**（`HC_LOG_PERF=1`）

**データ集約（`svc_data_agg` ↔ `ui_server`）の相関**

- **運用ログ（`hc_csv.log`）**: `[DATA_AGG] ui_ipc ok|fail|skip`（req ファイル名・sheet_id・hwnd）、`[UI_DISPATCH] source_req=...`、`[UI_DATA_AGG] create_dialog ok ... elapsed_ms=...`、`[DATA_AGG_UI] create_dialog enter ...`（UI プロセス pid）。
- **診断ログ（`hc_csv_diag.log`）**: 上記と同じ経路の詳細（`[DATA_AGG_TRACE]`、`[UI_TRACE] data_agg ... wall_perf_s=...`）。**`HC_LOG_DIAG=1`** またはデータ集約用診断フラグ等で診断ファイルが有効なときのみ。

`source_req=`（例: `req_data_agg_main_*.pkl`）を揃えて読むと、**svc 側送出 → ui_server 取り込み → `create_dialog`** の繋がりを追えます。

**CSV 読込（`svc_csv_ld` ↔ `ui_server`）の相関**

- **運用ログ（`hc_csv.log`）**: `[CSV_LD]` の `phase=`（`enter` / `after_ensure_ui_server` / `ready_ui` / `result_ok` / `result_cancel` / `do_load_enter` / `jit_import_done` / `load_csv_flow_done`）、`ui_ipc ok req=...`（`req_*.pkl` 名）、`[UI_DISPATCH]` / `[UI_CSV_LD] create_dialog ok source_req=... elapsed_ms=...`。
- **診断ログ（`hc_csv_diag.log`）**: `[CSV_LD_TRACE]`（診断有効時）、`[UI_TRACE] csv_ld create_dialog ok ...`。**`HC_LOG_DIAG=1`** 等で診断ファイルが有効なときのみ後者の詳細行。

`svc` 側の `req=` と `ui_server` 側の `source_req=` を同じファイル名で突き合わせると、**依頼 → ダイアログ生成**の遅延を切り分けやすくなります。

**CSV 保存（`svc_csv_sv` ↔ `ui_server`）の相関**

- **運用ログ（`hc_csv.log`）**: `[CSV_SV]` の `phase=`（`enter` / `sheet_valid_check` / `after_ensure_ui_server` / `ready_ui` / `result_ok` / `result_cancel` / `do_save_enter` / `read_matrix_done` / `write_csv_done` / `save_csv_flow_done` / 無データ警告は `no_valid_warn_flow_done`）。**`result_ok` / `result_cancel` には `dialog_wait_ms=`**（`ui_ipc ok` 直後〜結果 pickle 検知まで＝ダイアログ操作待ち＋ポーリング、機械処理と切り分け）。`ui_ipc ok`（通常 `req=...`、無データ分岐は `kind=no_valid_warn req=...`）、`[UI_DISPATCH]` / `[UI_CSV_SV] create_dialog ok source_req=... elapsed_ms=...`。
- **診断ログ（`hc_csv_diag.log`）**: `[CSV_SV_TRACE]`、`[UI_TRACE] csv_sv create_dialog ok ...`（診断有効時のみ後者の詳細行）。

`req=` と `source_req=` を揃えて、**ファイル保存ダイアログ・警告ダイアログ・進捗**のいずれの経路でも UI 側の `create_dialog` コストを追えます。

**CSV 結合（`svc_csv_mg` ↔ `ui_server`）の相関**

- **運用ログ（`hc_csv.log`）**: `[CSV_MG]` の `phase=`（`enter` / `after_ensure_ui_server` / `ready_ui` / `result_ok` / `result_cancel` / `merge_prep_done` / `merge_excel_write_done` / `merge_csv_flow_done`）、`ui_ipc ok req=...`、`[UI_DISPATCH]` / `[UI_CSV_MG] create_dialog ok source_req=... elapsed_ms=...`。
- **診断ログ（`hc_csv_diag.log`）**: `[CSV_MG_TRACE]`、`[UI_TRACE] csv_mg create_dialog ok ...`（診断有効時のみ後者の詳細行）。

**CSV 分割（`svc_csv_sp` ↔ `ui_server`）の相関**

- **運用ログ（`hc_csv.log`）**: `[CSV_SP]` の `phase=`（`enter` / `after_ensure_ui_server` / `result_ok` / `result_cancel` / `before_save_sleep` / `split_save_loop_done` / `split_csv_flow_done`）、警告経路は `ui_ipc ok kind=sp_warn` と `warn_flow_done`、`[UI_DISPATCH]` / `[UI_CSV_SP] create_dialog ok source_req=... elapsed_ms=...`。
- **診断ログ（`hc_csv_diag.log`）**: `[CSV_SP_TRACE]`、`[UI_TRACE] csv_sp create_dialog ok ...`（診断有効時のみ後者の詳細行）。同名確認のライフサイクル調査用: **`[CONFLICT_LIFECYCLE]`**（`ui_csv_sp`・`finish_enter` / `result_pickle_ok|fail` / `before_accept` / `after_accept_ok`）、**`[UI_TRACE] csv_sp_conflict lifecycle_end`**（`exec_plus_teardown_ms=`・`rc=`。同名確認は `ui_server` で後処理が軽いため、分割メインより短く出やすい）。**`HC_CSV_SP_CONFLICT_HWND_DIAG=1`** のとき **`[CONFLICT_HWND_DIAG]`**（`ui_server`・`winId` / `IsWindow` / `IsWindowVisible` / `GW_OWNER` / `rect` / **`GetClassName`（`cls=`）** / **`GetWindowText`（`text=`）** / **`GetParent`（`parent_hwnd=`）** / **`GetAncestor(..., GA_ROOT)`（`root_hwnd=`）**）。**`HC_CSV_SP_CONFLICT_HWND_DIAG_TREE=1`** 併用で **`[CONFLICT_EXCEL_DESC]`**（Excel 子孫の `hwnd,cls,vis,is_win` を最大 48 件まで）。上記フラグ単独でも `hc_csv_diag.log` が有効化される）。

**空白列削除（`svc_col_dl`）・空白行削除（`svc_row_dl`）**

- **運用ログ（`hc_csv.log`）**: 従来どおり `[COL_DL]` / `[ROW_DL]`（開始・使用範囲・完了など）。
- **計測ログ（`hc_csv_perf.log`）**: **`HC_LOG_PERF=1`** のときのみ **`[COL_DL_PERF]`** / **`[ROW_DL_PERF]`**。`phase=`（`enter` / `after_context` / `after_used_range` / `after_progress_ui_submit` / `after_matrix_read` / `after_scan_blank_*` / `after_delete_done` / `flow_end` など）と **`cumulative_ms=`**（`delete_empty_cols` / `delete_empty_rows` 開始からの経過）。
- **診断ログ（`hc_csv_diag.log`）**: **`HC_LOG_DIAG=1`**（等）のとき **`[COL_DL_TRACE]`** / **`[ROW_DL_TRACE]`**（上記 phase と同趣旨の区間）。**`HC_UI_FG_DIAG=1`** のとき **`[UI_FG]`**（完了ダイアログの表示タイミング・前面化の切り分け）。

**出荷ヘッダ挿入（`svc_hd_in`）・行整形ヘッダ（`svc_hd_nr`）**

- **運用ログ（`hc_csv.log`）**: `[HD_IN]` / `[HD_NR]`。
- **計測ログ（`hc_csv_perf.log`）**: **`HC_LOG_PERF=1`** のとき **`[HD_IN_PERF]`** / **`[HD_NR_PERF]`**。`phase=`（例: `enter` / `after_resolve` / `after_insert_ok` / `after_user_confirm` / `after_data_matrix_read` / `after_reshape` / `after_sheet_write_and_fit` / `flow_end`）と **`cumulative_ms=`**。
- **診断ログ（`hc_csv_diag.log`）**: 有効時 **`[HD_IN_TRACE]`** / **`[HD_NR_TRACE]`**。

**重複チェック（`svc_dupli`）**

- **運用ログ（`hc_csv.log`）**: `[DUPLI]`。
- **計測ログ（`hc_csv_perf.log`）**: **`HC_LOG_PERF=1`** のとき **`[DUPLI_PERF]`**。`phase=`（`enter` / `after_context` / `early_no_valid_range` / `after_intersection` / `after_progress_ui_submit` / `after_matrix_read` / `abort_matrix_read_failed` / `after_analyze` / `early_no_duplicates` / `after_duplicate_highlight` / `after_report_submit` / `flow_end` など）と **`cumulative_ms=`**。
- **診断ログ（`hc_csv_diag.log`）**: 有効時 **`[DUPLI_TRACE]`**。

**日付変換 YYYY/MM/DD（`svc_dt_ymd`）・日付時刻変換（`svc_dt_hm`）**

- **運用ログ（`hc_csv.log`）**: `[DT_YMD]` / `[DT_HM]`。
- **計測ログ（`hc_csv_perf.log`）**: **`HC_LOG_PERF=1`** のとき **`[DT_YMD_PERF]`** / **`[DT_HM_PERF]`**。`phase=`（`enter` / `after_context` / `early_no_selection` / `after_selection` / `early_empty_selection` / `after_progress_ui_submit` / `after_matrix_read` / `abort_matrix_read_failed` / `after_analyze` / `early_warning_not_date` / `after_write_chunk` / `after_done_ui` / `flow_end` など）と **`cumulative_ms=`**。
- **診断ログ（`hc_csv_diag.log`）**: 有効時 **`[DT_YMD_TRACE]`** / **`[DT_HM_TRACE]`**。

**文頭・文末トリム（`svc_trm_ex`）**

- **運用ログ（`hc_csv.log`）**: `[TRM_EX]`。
- **計測ログ（`hc_csv_perf.log`）**: **`HC_LOG_PERF=1`** のとき **`[TRM_EX_PERF]`**。`phase=`（`enter` / `after_context` / `after_range_resolve` / `early_no_data` / `abort_used_range_failed` / `after_matrix_read` / `abort_matrix_read_failed` / `abort_matrix_invalid` / `after_scan_counts` / `early_no_target` / `after_choice_ui_submit` / `user_cancel` / `after_apply_choice` / `after_undo_snapshot_attempt` / `after_write_chunk` / `abort_write_failed` / `after_done_ui` / `except_flow` / `flow_end` など）と **`cumulative_ms=`**。
- **診断ログ（`hc_csv_diag.log`）**: 有効時 **`[TRM_EX_TRACE]`**。

**元に戻す（`svc_undo` の `exec_undo`）**

- **運用ログ（`hc_csv.log`）**: `[UNDO]`。
- **計測ログ（`hc_csv_perf.log`）**: **`HC_LOG_PERF=1`** のとき **`[UNDO_PERF]`**。`phase=`（`enter` / `after_payload_load` / `branch_structure` / `after_structure_restore` / `abort_structure_import` / `abort_structure_error` / `branch_data` / `abort_empty_data` / `after_data_restore` / `before_cache_delete` / `after_cache_delete` / `after_done_ui` / `abort_exception` / `abort_no_book` / `abort_no_sheet` / `abort_no_app` / `abort_no_hsys` / `abort_no_cache` / `flow_end` など）と **`cumulative_ms=`**。
- **診断ログ（`hc_csv_diag.log`）**: 有効時 **`[UNDO_TRACE]`**。

---

### 2.1 VBA リボン待機（WaitForm）と `core.ribbon_invoke.invoke`（`excel_session.invoke_action`）

| 項目 | 内容 |
|------|------|
| **表示** | `Main.RibbonInvokeFromControl` 内、`RunPythonSafe` 直前に `HC_WaitForm.BeginWaitForRibbon`（`VBA\Main.bas` / `HC_WaitForm.bas`）。 |
| **閉じる（同期）** | `core.ribbon_invoke.invoke` の **`finally`** で、`core.ribbon_public_to_svc.RIBBON_INVOKE_FINALLY_NOTIFY_WAITFORM` に含まれる action のとき **`notify_wait_form_ready()`**（`ribbon_invoke` 内 `_INVOKE_NOTIFY_WAITFORM_ACTIONS` と同一集合）。短寿命 RunPython 終了前に VBA の `HC_WaitForm.NotifyUiReady` を実行。**`RIBBON_ACTIONS_READY_UI_CLOSES_WAITFORM`**（`load_csv` / `save_csv` / `merge_csv` / `split_csv`）は含めず、READY_UI 等で `notify_ui_ready` 側が砂時計・WaitForm を扱う。 |
| **閉じる（異常・VBA）** | `RunPythonSafe` / `RibbonInvokeFromControl` / `RibbonCallback_hc_main` の **ErrorHandler** でも `NotifyUiReady`。 |
| **未知の action** | `invoke` がマップに無い action で **`ValueError`** を投げる前に、**ベストエフォートで `notify_wait_form_ready()`**（リボンと `tag` の不整合時の取りこぼし防止）。 |

---

## 3. IPC 配下のブートログ（補助）

| 保存場所（既定） | 用途 |
|------------------|------|
| **`%TEMP%\csv_tool\logs\`**（`HC_IPC_ROOT` を変えた場合は **`%HC_IPC_ROOT%\logs\`**） | `svc_server` / UI サーバー / ブリッジ等を **別プロセスで起動した直後**の標準出力・標準エラーなどがファイルに分かれることがあります。**プロセスが立ち上がらない・すぐ落ちる**ときの補助資料として参照します。 |

共有の `hc_csv.log` とは役割が異なり、**起動失敗の切り分け**向けです。

---

## 4. 環境変数一覧（ログ以外）

ログと直接は無関係ですが、アドイン全体で参照します。値の詳細は各モジュールの実装を参照してください。

### 4.1 IPC・パス

| 正規名 | 別名 | 意味 |
|--------|------|------|
| `HC_IPC_ROOT` | `HC_QT_IPC_DIR` | IPC（pickle・制御フラグ等）のルート。未設定時は `%TEMP%\csv_tool`。 |
| `HC_PROJECT_ROOT` | — | 子プロセス起動時に内部で設定されることがある。 |
| `HC_INSTALL_ROOT` | — | 配布ツリーのルート（例: `CSV_Tool`）。`core.runtime_layout` が参照。短寿命 `xlwings_short_runner` はここを `sys.path` 先頭にし `chdir` する。 |
| `HC_PACKAGED_DEPLOYMENT` | — | `1` / `true` / `yes` で「配布（EXE）モード」を意味する。`svc_host` の子プロセス起動が **`app\bin\hc_main.exe`** 等（**`hc_svc_server.exe` / `hc_ui_server.exe`** は同じ **`app\bin\`**）に切り替わる条件の一部。VBA `xlwings.bas` も `RunPython` の短寿命経路判定に参照（`USE_PACKAGED_RUNPYTHON` と併用可）。配布時、`svc_host` / `ui_data_agg` が **`hc_*.exe` を起動するとき**子プロセスの **`PATH` 先頭**に **`%HC_INSTALL_ROOT%\app\bin`** と **`%HC_INSTALL_ROOT%`**（インストールルート）を足す（**`runtime_layout.env_with_packaged_dll_search_path`**）。 |
| `HC_DEPLOY_ROOT` | — | **共有側の配布ルート**（`catalog.json` は通常 **`%HC_DEPLOY_ROOT%\catalog.json`**）。**`installer\CSV_Tool_Setup.iss`** が **`SHAREPAYLOAD` の親フォルダ**をインストール時に `HKCU\Environment` へ書く。詳細は **`docs\インストールと運用（利用者・運用向け）.md` §1.5**。薄いインストーラ EXE のビルド手順は **`docs\インストーラ化（開発者向け）.md` §2**（**`installer\build_csv_tool_setup.bat`**）。 |
| `HC_CATALOG_PATH` | — | **`catalog.json` のフルパス**（任意）。設定時は **`HC_DEPLOY_ROOT`** より優先（`core.packaged_update`）。 |
| `HC_UPDATE_CHECK_AT_STARTUP` | — | **`0` / `false`** で、アドイン起動直後の **版通知（`catalog.json` 照合）**を行わない。未設定は **有効**（ただし **配布モード**かつカタログ解決可能なときだけ実際に読む）。 |

**版確認ログ（更新チェック）** … 配布モードで `core.packaged_update` が **`%HC_INSTALL_ROOT%\logs\hc_update.log`** に追記する（共有に届かない・比較結果など）。運用ログの `hc_csv.log` とは別ファイル。

**`config\` の JSON（UI・warmup 等）**: **`core_cst.resolve_config_file_path`** は、**`HC_INSTALL_ROOT`** が有効かつ **`<HC_INSTALL_ROOT>\config\<ファイル名>` が存在する** とき **そのパスを最優先**し、無ければ **各 EXE バンドル（または開発時はリポジトリルート）の `config\`** にフォールバックする（`docs/Exe化（開発者向け）.md` **セクション 2.7**）。

**運用の目安**: **配布PC**ではインストーラが上記を設定する。**開発PC**で **cmd にだけ `HC_*` を載せ替えてから** Excel を手動起動する場合は `docs/Exe化（開発者向け）.md` **セクション 2.6**（`tools\dev\start_excel_*.bat` は **Excel を起動しない**）を参照。

#### 4.1.1 IPC 起動時スイープ（滞留ファイルの TTL 削除）

常駐プロセス（`ui_server` / `bridge_runner` / `svc_server`）が **単一インスタンス mutex 取得成功直後**に、古いキュー・結果・起動ガードを削除する。目的・フォルダ対応表は **`docs/IPC_TEMP_CLEANUP.md`** を参照。**`bridge_runner` のみ** `bridge_requests` の `*.json` を起動時に **全削除**してから TTL スイープ（補助）を行う。

| 変数名 | 既定（秒） | 意味 |
|--------|------------|------|
| `HC_IPC_DISABLE_STARTUP_SWEEP` | — | `1` でスイープをすべて無効化（トラブルシュート用）。 |
| `HC_IPC_SWEEP_QUEUE_TTL_SEC` | `86400`（24h） | `requests` / `bridge_requests` / `svc_requests`。 |
| `HC_IPC_SWEEP_SVC_RESULTS_TTL_SEC` | `3600`（1h） | `svc_results`（`hc_main` の古い `svc_res` 掃除と整合）。 |
| `HC_IPC_SWEEP_STARTING_FLAG_TTL_SEC` | `600`（10m） | `control` 直下の `*_starting.flag` のみ（`shutdown.flag` / `svc_shutdown.flag` は対象外）。 |

### 4.2 データ集約（診断ログを有効にしうるフラグ群）

正規名は **`HC_DIAG_DATA_AGG_*`**。従来の **`DATA_AGG_*`** も読みます。いずれかを有効にすると診断ログファイルが使われる場合があります（詳細は `core_env`）。

| 正規名 | 別名 |
|--------|------|
| `HC_DIAG_DATA_AGG_NAMES` | `DATA_AGG_NAME_PATH_DIAG` |
| `HC_DIAG_DATA_AGG_NAMES_MAX_ROWS` | `DATA_AGG_NAME_PATH_DIAG_MAX_ROWS` |
| `HC_DIAG_DATA_AGG_NAMES_COL` | `DATA_AGG_NAME_PATH_DIAG_COL` |
| `HC_DIAG_DATA_AGG_BATCH_TIMING` | `DATA_AGG_COMPUTE_BATCH_TIMING` |
| `HC_DIAG_DATA_AGG_FILE_TIMING` | `DATA_AGG_PER_FILE_TIMING` |
| `HC_DIAG_DATA_AGG_MASTER_PREFETCH` | `DATA_AGG_MASTER_OFF_PREFETCH` |
| `HC_DIAG_DATA_AGG_JOIN` | `DATA_AGG_JOIN_DUMP`（`1` のみ有効） |
| `HC_DIAG_DATA_AGG_JOIN_COL` | `DATA_AGG_JOIN_DUMP_COL`（省略時は全列。指定時は代入先列名に部分一致する項目だけ詳細ログ） |
| `HC_DIAG_DATA_AGG_JOIN_MAX_SLICES` | `DATA_AGG_JOIN_DUMP_MAX_SLICES`（スライスごとの比較値ログ本数上限） |
| `HC_DIAG_DATA_AGG_JOIN_MAX_ROWS` | `DATA_AGG_JOIN_DUMP_MAX_ROWS`（ファイル単位 post_merge の行プレビュー本数） |

**結合キー調査（`HC_DIAG_DATA_AGG_JOIN=1`）**  
`hc_csv_diag.log` に **`[DATA_AGG_JOIN_DUMP]`** を出します。各ファイルの結合書き込み前後で、`phase=enter` / `slice` / `done` / `skip`（理由付き）および **`phase=post_merge`**（結合定義がある列の先頭行プレビュー）が記録されます。`HC_LOG_DIAG=1` と併用してください。MAC RMT など特定列に絞る場合は **`HC_DIAG_DATA_AGG_JOIN_COL=MAC RMT`**（列名の部分一致・大小無視）。

**本番一括の性能（`hc_csv.log` にも出力）**

| 変数名 | 既定 | 意味 |
|--------|------|------|
| `DATA_AGG_FILE_PARALLEL_WORKERS` | `auto` | 入力ファイル並列読込。`0` で逐次。`auto` は `min(8, CPU, ファイル数)`。本番一括開始時に `[DATA_AGG] batch extract parallel …` が出る。 |
| `DATA_AGG_BATCH_FILE_PATH_FILTER` | 有効 | `0` で無効。`file_pattern` 付き cell 項目の OR で走査結果を絞る（マスタプレビューと同ロジック）。 |
| `DATA_AGG_PER_FILE_TIMING` / `HC_DIAG_DATA_AGG_FILE_TIMING` | オフ | `1` でファイル別 open/read/merge ms を診断ログへ。 |
| `DATA_AGG_COMPUTE_BATCH_TIMING` / `HC_DIAG_DATA_AGG_BATCH_TIMING` | 本番は常時1行 | 本番一括完了時 `[DATA_AGG] compute_batch_timing …`（extract/merge/total_ms）を `hc_csv.log` に出力。 |

**縦反復抽出の上限・打ち切り**

| 変数名 | 既定 | 意味 |
|--------|------|------|
| `HC_DATA_AGG_EXTRACT_ABSOLUTE_MAX` | `999999` | 「N件」指定および「空白まで」の絶対上限。UI の取得件数スピンもこの範囲。 |
| `HC_DATA_AGG_EXTRACT_TRUNC_POLICY` | `warn`（未設定時） | 読取上限に達し未読データがあると判定したときの方針。`warn` … ログのみで続行（本番一括・デバッグの既定）。`abort` … 結合前に処理中断（`DataAggExtractTruncated`）。上限の確認はシナリオ／マスタデバッグで行い、本番一括は止めない思想。 |

**互換（コード側の既定・環境変数なし）**: シナリオで `repeat_max` 未設定かつ「空白まで」でもない場合、縦反復の上限は従来どおり **9999** 件。

### 4.3 svc / ブリッジ・Excel 連携

| 変数名 | 既定（概要） | 意味（概要） |
|--------|----------------|----------------|
| `HC_SVC_TIMEOUT_SEC` | `180` | `svc_server` 応答待ちの既定秒数。 |
| `HC_SVC_TIMEOUT_<ACTION>_SEC` | — | アクション別の上書き。 |
| `HC_RETURN_EARLY` | `1` | 依頼書き出し後に早期 return するか。 |
| `HC_RETURN_EARLY_WAIT_SEC` | `1.0` | 早期 return 前の sleep 秒数（`HC_RETURN_EARLY=1` のときのみ）。短縮例: `0.5`。 |
| `HC_SVC_IDLE_POLL_SEC` | `0.1` | アイドルポーリング間隔（秒）。 |
| `HC_SVC_WARMUP_ACTIONS` | — | ウォームアップで先行ロードする action。 |
| `HC_MAIN_POLL_SEC` | `0.05` | 常駐 **`hc_main.py`** のポーリング間隔（秒）。**読取**: `core_env.hc_main_poll_sec()`。**互換**: 未設定なら **`HC_BRIDGE_POLL_SEC`**。 |
| `HC_MAIN_MIN_FILE_AGE_SEC` | `0.05` | `bridge_requests` の `.json` を読む前の最小経過秒（書き込み途中の読取り緩和）。**読取**: `core_env.hc_main_min_file_age_sec()`。**互換**: **`HC_BRIDGE_MIN_FILE_AGE_SEC`**。 |
| `HC_MAIN_BAD_FILE_MAX_POLLS` | `100` | 同一ファイルの JSON 解釈または転送がこのポーリング回数を超えたら削除。**読取**: `core_env.hc_main_bad_file_max_polls()`。**互換**: **`HC_BRIDGE_BAD_FILE_MAX_POLLS`**（積算秒の目安は `HC_MAIN_POLL_SEC` または互換の `HC_BRIDGE_POLL_SEC`）。 |
| `HC_PROGRESS_WINDOW_STARTUP_WAIT_SEC` | `1.0`（秒） | 進捗ウィンドウ起動待ち（CSV 読込等）。読取は **`core_env.progress_window_startup_wait_sec()`**。 |
| `HC_EXCEL_HWND` | — | 実行中に Python 側が設定（子プロセスが HWND を参照）。**`core_env.set_excel_hwnd_for_spawn(hwnd)`** / 定数 **`core_env.ENV_EXCEL_HWND`**。 |

**svc_server ウォームアップ（B+）**: 既定は **`config/svc_warmup.json`** の `warmup_actions`。B+ 常駐のため **初回 spawn 時のみ** `_run_warmup` が走る。上記 `HC_SVC_WARMUP_ACTIONS` は JSON が無いときのフォールバック。COM 常駐・再起動方針は **`docs/svc_com_session.md`**。

**試験: `HC_RETURN_EARLY_WAIT_SEC`（待機短縮の効果・取りこぼし確認）**

1. **`HC_LOG_PERF=1`** を付けたうえで Excel を起動し（環境変数は起動前に設定）、リボンから同一操作（例: `run_data_agg`）を実行する。
2. **`%TEMP%\csv_tool\hc_csv.log`** で `hc_main.ribbon_invoke` の次を比較する。
   - `call_svc phase=return_early_wait_ready` … `wait_sec_effective`（実効秒）・`env_defined` / `env_value`（変数の有無と生文字列）
   - `call_svc phase=after_return_early_sleep` … `actual_sleep_ms`（`perf_counter` 実測。`wait_sec_effective×1000` に近いこと）
   - 直後の `call_svc phase=after_return_early` の `cumulative_ms`（待機短縮分だけ減るか）
3. **`%TEMP%\csv_tool\hc_csv_perf.log`** の `after_xlwings_runpython` / `elapsed_since_click_ms` と突き合わせ、VBA に戻るまでの壁時計がどれだけ変わるか見る。
4. 値は **`1.0` → `0.5` → `0.25`** など段階的に変え、**Excel 再起動のたびに**試す（プロセスは起動時の環境を読む）。
5. **`0` や極小**で `svc_server` が依頼を拾えない症状がないか、実操作で確認する。

### 4.4 その他

| 変数名 | 意味 |
|--------|------|
| `CSV_TOOL_SPAWN_ID` | プロセス_spawn_識別用。 |
| `TEMP` | Windows 一時フォルダ。ログ・IPC 既定の基準。 |

---

## 5. VBA での設定の注意

- ユーザー環境変数として設定すると、Excel 起動プロセスに引き継がれます。
- **`HC_LOG_DIAG`** / **`HC_LOG_PERF`** は Python の **`core_env`** と同名で解釈を揃えています。
- VBA ソースは **Shift-JIS（CP932）** で保存する運用とし、ログ**ファイル**は Python と共有するため **UTF-8** のままです。
- **`%TEMP%\\csv_tool\\bridge_requests\\` の `req_*.json`**（リボン経由の依頼）は **UTF-8** で出力されます（常駐 **`hc_main.py`** は **utf-8-sig → utf-8 → cp932** の順で解釈し、旧 ANSI 出力も読み取り可能）。

---

## 6. 設定例（Windows）

Excel を**いったん終了**してから変数を設定し、**新しく Excel を起動**すると VBA の `Environ` および Excel から起動する Python に反映されます（開いたままの Excel では読み直されません）。

### 6.1 グラフィカル（恒久的・ユーザー単位）

1. **設定** → **システム** → **バージョン情報**（または **システムの詳細情報**）→ **システムの詳細設定** → **環境変数(N)...**
2. **ユーザー環境変数**で **新規(N)...**
3. 変数名・値を入力（下表など）。
4. OK で閉じ、Excel を再起動。

| 目的 | 変数名 | 値の例 |
|------|--------|--------|
| 診断ログを出す | `HC_LOG_DIAG` | `1` |
| 分割・同名確認の HWND 診断のみ（`hc_csv_diag.log`） | `HC_CSV_SP_CONFLICT_HWND_DIAG` | `1` |
| 上に加え Excel 配下 HWND 列挙（`[CONFLICT_EXCEL_DESC]`、ログ多め） | `HC_CSV_SP_CONFLICT_HWND_DIAG_TREE` | `1`（`HC_CSV_SP_CONFLICT_HWND_DIAG=1` 必須） |
| UI 前面・Z 順の切り分け（`hc_csv_diag.log`、`[UI_FG]`） | `HC_UI_FG_DIAG` | `1` |
| タイトルバー最小化/最大化・Win32 スタイルの切り分け（`hc_csv_diag.log`、`[UI_CAPTION]`。`gwl_exstyle` / `ws_ex_toolwindow` を含む） | `HC_UI_WINDOW_CAPTION_DIAG` | `1` |
| 旧方式のタスクバー抑止（`WS_EX_TOOLWINDOW` を付与。タイトルバーの最小／最大化ボタンが消えることがある） | `HC_USE_WS_EX_TOOLWINDOW_FOR_TASKBAR` | `1` |
| 上記を含め `WS_EX_TOOLWINDOW` を絶対に付けない（`USE=1` より優先） | `HC_SKIP_WS_EX_TOOLWINDOW` | `1` |
| ui_server を CMD から `python.exe` で起動してログを見る（`FreeConsole` を抑止） | `HC_UI_KEEP_CONSOLE` | `1` |
| 計測ログを出す | `HC_LOG_PERF` | `1` |
| 従来どおりデバッグ相当 | `HC_DEBUG` | `1` |
| IPC を別フォルダに固定 | `HC_IPC_ROOT` | `D:\work\csv_tool_ipc`（実在パス推奨） |
| svc 応答を長く待つ | `HC_SVC_TIMEOUT_SEC` | `300` |
| 早期 return をやめる | `HC_RETURN_EARLY` | `0` |
| 早期 return 直前の待機を短くする | `HC_RETURN_EARLY_WAIT_SEC` | `0.5` |

削除するときは一覧で変数を選び **削除(D)**。

### 6.2 PowerShell（そのセッションだけ）

Excel を**同じ PowerShell から**起動したいときに使います。

```powershell
$env:HC_LOG_DIAG = "1"
$env:HC_LOG_PERF = "1"
& "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
```

### 6.3 PowerShell（ユーザーに恒久的に書き込む）

```powershell
[Environment]::SetEnvironmentVariable("HC_LOG_DIAG", "1", "User")
[Environment]::SetEnvironmentVariable("HC_LOG_PERF", "1", "User")
```

解除例:

```powershell
[Environment]::SetEnvironmentVariable("HC_LOG_DIAG", $null, "User")
```

### 6.4 コマンドプロンプト（そのセッションだけ）

`set` は**いま開いている cmd** にだけ効き、**同じウィンドウから起動した** Excel などに引き継がれます。等号の前後に空白を入れないでください。

```cmd
set HC_LOG_DIAG=1
set HC_LOG_PERF=1
set HC_DIAG_DATA_AGG_NAMES=1
start "" "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
```

値にスペースが含まれる場合:

```cmd
set "HC_IPC_ROOT=D:\work\csv_tool ipc"
```

### 6.5 コマンドプロンプト（setx でユーザー環境変数・恒久的）

`setx` は**新しく開いたプロセス**から見えます。**いまの cmd には反映されません**。Excel は**再起動**が必要です。

```cmd
setx HC_LOG_DIAG 1
setx HC_LOG_PERF 1
setx HC_RETURN_EARLY 0
```

```cmd
setx HC_IPC_ROOT "D:\work\csv_tool"
```

注意: `setx` の値は**約 1024 文字**まで。複数回続ける場合は数秒空けると安全です。

**変数の削除**は `setx` ではできません。GUI で削除するか、例:

```cmd
reg delete "HKCU\Environment" /v HC_LOG_DIAG /f
```

（削除後も Excel は一度終了してから起動し直してください。）

### 6.6 組み合わせ例

| 目的 | 変数名 | 値 |
|------|--------|-----|
| 診断＋計測＋データ集約の名前パス調査 | `HC_LOG_DIAG` / `HC_LOG_PERF` / `HC_DIAG_DATA_AGG_NAMES` | 各 `1` |
| IPC を RAM ディスク等へ | `HC_IPC_ROOT` | 例: `R:\csv_tool` |
| トラブルシュートで結果まで Python が待つ | `HC_RETURN_EARLY` | `0` |

### 6.7 Qt ウィンドウ・タイトルバーとタスクバー（`WS_EX_TOOLWINDOW`）

**背景**  
Excel をオーナーとする Qt ダイアログに拡張スタイル **`WS_EX_TOOLWINDOW`** を付けると、**`GWL_STYLE` 上は** 最小化・最大化ビットが有効でも、**タイトルバーに最小／最大化ボタンが描画されない**ことがある（システムメニューでは利用可）。

**既定（現在）**  
`core_w32.apply_taskbar_hiding_extended_style` は **既定では何もしない**（`WS_EX_TOOLWINDOW` を付けない）。タスクバーとの関係は主に **GW_OWNER** と **`SHOW_IN_TASKBAR`** に寄せる。副作用として、**環境によってはタスクバーに Qt ウィンドウのボタンが出やすい**場合があります。

| 手段 | 内容 |
|------|------|
| **`HC_USE_WS_EX_TOOLWINDOW_FOR_TASKBAR=1`** | 従来方式を有効化（`WS_EX_TOOLWINDOW` + `WS_EX_APPWINDOW` 解除）。タスクバー抑止を強める一方、**タイトルバー min/max 非表示**や **Windows 11 で閉じるボタンの見た目が変わる**等のリスクあり。 |
| **`HC_SKIP_WS_EX_TOOLWINDOW=1`** | `WS_EX_TOOLWINDOW` を**付けない**（**`USE=1` より優先**）。 |
| **`WINDOW.USE_WS_EX_TOOLWINDOW_FOR_TASKBAR`**（各 `config/ui_*.json` の `WINDOW` 内・任意） | **`true`** の画面だけ `WS_EX_TOOLWINDOW` をオプトイン。**キー未指定**の画面はグローバルな `HC_USE_...` のみが効く。`false` を明示すると、その画面では環境変数が `1` でも付与しない。 |

**診断**  
`HC_UI_WINDOW_CAPTION_DIAG=1` のとき、`hc_csv_diag.log` の **`[UI_CAPTION]`** に `gwl_exstyle`（符号なし 8 桁 hex）と **`ws_ex_toolwindow`**（`WS_EX_TOOLWINDOW` の有無）が出る。

**未実装（将来案）**  
**`ITaskbarList`** 等でタスクバーから隠しつつ `WS_EX_TOOLWINDOW` を使わない方式は、コスト・互換調査が必要なため未着手。

**手動回帰の目安**  
1. **データ集約メイン**：タイトルバーに最小化・最大化・閉じる。  
2. **データ集約デバッグ**：同上。  
3. **`SHOW_IN_TASKBAR: false` の代表**（例：CSV 読込の進捗ダイアログ）：タイトルバー操作とタスクバー表示の許容確認。

---

## 7. Python API（参照）

- **`core.core_env`** … 真偽・文字列取得、**`ipc_dir_raw()`**（`HC_IPC_ROOT` / `HC_QT_IPC_DIR`）、**`progress_window_startup_wait_sec()`**、**`set_excel_hwnd_for_spawn`**、診断ファイル要否。
- **`core.ipc_cleanup`** … IPC ルート配下の **起動時スイープ**（TTL および `bridge_requests` の全 JSON 削除を含む `run_ui_server_startup_sweeps` 等）。**`docs/IPC_TEMP_CLEANUP.md`**。
- **`core.core_log`** … `get_logger`（運用）、`get_diag_logger`、`get_perf_logger`、データ集約向け診断は **`hc_csv_diag.log` に統合**。

---

## 8. 本ドキュメントの変更履歴

| 日付 | 内容 |
|------|------|
| 2026-04-16 | §4.1: **`HC_INSTALL_ROOT`** と **`core_cst.resolve_config_file_path`**（`config\` JSON の優先順）の説明を追記。`docs/Exe化（開発者向け）.md` セクション 2.7 へ相互参照。 |
| 2026-04-10 | §4.1.1・§7: `bridge_runner` 起動時に `bridge_requests` の `*.json` を全削除してから TTL 補助スイープする旨を追記。 |
| 2026-04-10 | `docs/csv_ld_perf_measurement.md`（`load_csv` 計測）。`hc_csv_perf.log` 行に CSV 読込区間の参照を追記。 |
| 2026-04-10 | §4.1.1 IPC 起動時スイープ環境変数。`core.ipc_cleanup`・`docs/IPC_TEMP_CLEANUP.md` 参照。 |
| 2026-04-06 | `svc_csv_sv`: `dialog_wait_ms` ログ、`HC_EXCEL_HWND` を `core_env.set_excel_hwnd_for_spawn` に統一。 |
| 2026-04-06 | `core_env` に `progress_window_startup_wait_sec` / `set_excel_hwnd_for_spawn` / `ENV_EXCEL_HWND`。`svc_csv_ld`・`ui_qt.ipc_file` の環境参照を `core_env` 経由に統一。 |
| 2026-04-06 | `svc_trm_ex` / `svc_undo`（`exec_undo`）の perf・診断タグ（`[TRM_EX_PERF]` / `[UNDO_PERF]` 等）。 |
| 2026-04-06 | `svc_dupli` / `svc_dt_ymd` / `svc_dt_hm` の perf・診断タグ（`[DUPLI_PERF]` 等）。 |
| 2026-04-06 | `svc_col_dl` / `svc_row_dl` / `svc_hd_in` / `svc_hd_nr` の perf・診断タグ。短寿命 invoke（当時 `hc_main.invoke`）と VBA WaitForm の対応（上記 2.1）。 |
| 2026-04-13 | §6.7: `WS_EX_TOOLWINDOW` 既定オフ・`HC_USE_WS_EX_TOOLWINDOW_FOR_TASKBAR` / `HC_SKIP_WS_EX_TOOLWINDOW` / JSON `USE_WS_EX_TOOLWINDOW_FOR_TASKBAR`・手動回帰目安。§6.1 表を更新。 |
| 2026-04-13 | `HC_UI_WINDOW_CAPTION_DIAG`・`[UI_CAPTION]`（`apply_window_config` のタイトルバー／GWL_STYLE 診断）。§2 表・§6.1 表を更新。 |
| 2026-04-11 | 短寿命司令塔を **`core.ribbon_invoke`** に統一（旧ルート `hc_invoke.py` 削除）。§2.1 のモジュール名を更新。 |
| 2026-04-06 | ログファイル視点に再構成（保存先、`hc_csv` 3 種＋IPC 配下ブートログ、用途と環境変数の対応）。ログ以外の環境変数は第 4 章に集約。設定例（GUI / PowerShell / cmd）は継承。 |
