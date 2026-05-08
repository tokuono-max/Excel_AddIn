# ui_common 仕様書（表示共通）

**対象**: ui_qt/ui_common.py — Qt UI サーバ側の表示共通処理  
**作成日**: 2026-03-09  
**目的**: ウィンドウ設定適用・前面化・中央配置・共通ダイアログ・設定読込の役割と API を定義する。

---

## 1. 機能概要

**ui_common** は、機能非依存の「表示共通」を集約するモジュールである。次の責務を持つ。

- **ウィンドウフラグ・スタイル**: JSON（WINDOW）に基づく TOPMOST・最小化/最大化・タスクバー・リサイズ・サイズ・位置の適用。
- **Excel 連携**: Excel HWND をオーナーにした親子関係の設定、Excel 前面化のうえでのダイアログ前面化、Excel ウィンドウ中央への配置、DPI/マルチモニタ対策。
- **設定読込**: UI_COMMON と機能別設定のマージ（get_ui_config / get_ui_config2）、深い階層のマージ（_deep_merge）。
- **共通ダイアログ**: ワーニング（WarningDialog）、進捗（ProgressDialog）、完了通知（DoneDialog）の生成と、メッセージの改行・タブ正規化（_normalize_message_newlines）。
- **Excel 実質モーダル**: ダイアログ表示中の Excel 操作ロック/解除（enable_excel_window：子 HWND 列挙＋EnableWindow）。※ 推奨は ui_win.enable_excel_window（有効化時にルート含む）。
- **終了・ジオメトリ**: シャットダウンフラグ監視・親プロセス死活監視（UiShutdownGuard）、ウィンドウ位置・サイズの保存/復元（save_geometry / restore_geometry）。

進捗・完了は「進捗を閉じてから完了通知を表示」するシーケンスを標準とし、進捗 100% 表示後は 1 秒経過で進捗を閉じたあと完了通知を表示する。

---

## 2. 依存関係

| 対象 | 内容 |
|------|------|
| **core** | core_log, core_cst（UI_COMMON, get_ui_config_from_file_required）, core_w32（set_owner, bring_to_front, get_root_window, enum_child_windows, enable_windows, set_window_style_remove_min_max 等） |
| **ui_qt** | ipc_file（write_pickle, read_pickle, get_control_dir, get_shutdown_flag_path） |
| **PySide6** | QtCore, QtWidgets（QDialog, QMessageBox, QLabel, QProgressBar, QListWidget 等） |

設定の解釈（WINDOW キー・TOPMOST・CENTER_ON_EXCEL 等）は本モジュールの責務。core_w32 は Win32 API の実行のみ行い、設定は読まない。

---

## 3. 公開 API（主要）

### 3.1 ウィンドウ設定・配置

| 関数 | 説明 |
|------|------|
| **apply_window_config(w, ui_cfg, parent_hwnd, screen_key)** | UI_COMMON および画面固有の WINDOW を統合し、フラグ・サイズ・位置・オーナー・前面化を適用。DONE/PROGRESS はオーナー/前面化をスキップし、各ダイアログの showEvent で中央配置等を行う。 |
| **apply_common_window_flags(w)** | cst.UI_COMMON.WINDOW の TOPMOST のみ適用。 |
| **apply_common_window_style(w, parent_hwnd)** | 共通フラグ適用＋CENTER_ON_EXCEL 時は center_on_excel を実行。 |
| **center_on_excel(w, parent_hwnd, rect_override)** | Excel ウィンドウ矩形の中央にウィジェットを配置。rect_override 指定時は get_excel_rect を使わずその矩形を使用。DPI 対策のため _with_thread_dpi_physical 内で実行。 |
| **center_on_rect(w, rect)** | 指定矩形 (左,上,右,下) の中央にウィジェットを移動。HWND 取得可能時は SetWindowPos で物理ピクセル配置。 |
| **get_excel_rect(parent_hwnd)** | Excel HWND の GetWindowRect 結果を (left, top, right, bottom) で返す。 |
| **ensure_front(w, parent_hwnd)** | 先に Excel を前面化し、続けてウィジェットを raise/activate/SetWindowPos(HWND_TOP)/SetForegroundWindow。GW_OWNER が 0 のとき Win32 `set_owner` を再適用（`_hc_show_taskbar` 時はスキップ）。SFW 失敗時は短い遅延で最大 2 回再試行。 |
| **done_dialog_show_event_on_excel** | showEvent 用。CENTER_ON_EXCEL・Excel 無効化に加え、**TOPMOST / ALWAYS_IN_FRONT_OF_EXCEL** または **EXCEL_FRONT_FOLLOW（旧 FRONT_FOLLOW）** のとき `ensure_front` と 80/200ms の再試行（ヘルプ等 TOPMOST なしでも Excel 手前化）。 |
| **ensure_owner_and_front(w, owner_hwnd)** | _set_owner_hwnd のあと ensure_front。 |
| **ensure_dialog_front_of_excel(w, parent_hwnd, rect_override)** | オーナーをルートに設定・Excel 前面化・center_on_excel・raise/activate/nudge。完了通知の前面化補強用。 |
| **_set_owner_hwnd(w, owner_hwnd)** | ウィジェットを Excel ルート HWND の子に紐付け（Qt setTransientParent ＋ Win32 set_owner）。タスクバー非表示は _hc_show_taskbar が False のときのみ実施。 |

### 3.2 Excel 操作ロック

| 関数 | 説明 |
|------|------|
| **enable_excel_window(hwnd, enabled)** | 指定 HWND の子を再帰列挙し、それらに対して EnableWindow を実行。Z オーダー連動を保つためルートは変更しない実装。※ 有効化時にシート操作を復帰させるには ui_win.enable_excel_window（ルート含む）の利用を推奨。 |

### 3.3 設定読込

| 関数 | 説明 |
|------|------|
| **get_ui_config(screen_key)** | cst.UI_COMMON と cst.UI_&lt;screen_key&gt; を _deep_merge して返す。 |
| **get_ui_config2(feature_key, screen_key)** | 機能キー（CSV_MG 等）に対応する設定を取得。CSV_MG は get_ui_config_from_file_required でファイルから読込。それ以外は cst.UI_SCREENS から取得。COMMON と screen_key をマージして返す。 |
| **apply_tooltip_if_set(widget, cfg, key)** | cfg の TOOLTIP があれば _normalize_message_newlines して setToolTip。 |
| **_normalize_message_newlines(text)** | \\n → 改行、\\t およびタブ → 4 文字空白。共通仕様に従う。 |

### 3.4 共通ダイアログ

| 関数・クラス | 説明 |
|--------------|------|
| **create_warning_dialog(req, parent_hwnd, warning_cfg)** | WarningDialog を生成。req.message があればそれを使用、なければ warning_cfg.MSG。TITLE, ICON, BTN_OK, WINDOW を適用。 |
| **create_progress_dialog(req_dict, parent_hwnd, parent_widget, progress_cfg)** | ProgressDialog を生成し、モデルレス一覧に登録。progress_cfg 未指定時は _get_progress_config()（CSV_MG 基準）。 |
| **create_done_dialog(req, parent_hwnd, parent_widget, done_cfg)** | DoneDialog を生成。done_cfg 未指定時は _get_done_config()。 |
| **WarningDialog** | ワーニング用 QDialog。表示時 Excel ロック、閉じ時ロック解除。CENTER_ON_EXCEL 時は透明表示→中央配置→不透明化。 |
| **ProgressDialog** | 進捗用 QDialog（モデルレス）。progress_path の Pickle をポーリングし、RUN/DONE/OVER_LIMIT に応じて表示更新。DONE かつ show_done_dialog 時は進捗を閉じたあと create_done_dialog で完了通知を表示。 |
| **DoneDialog** | 完了通知用 QDialog。items / detail_text を表示。showEvent で ensure_front を常に呼び Excel 前面に表示。OK/close で enable_excel_window(True)。 |

### 3.5 ジオメトリ・監視

| 関数・クラス | 説明 |
|--------------|------|
| **save_geometry(screen_key, w)** | ウィンドウの geometry を ipc_file の control/geometry 配下に Pickle 保存。 |
| **restore_geometry(screen_key, w)** | 保存済み位置・サイズを復元。成功時 True。 |
| **UiShutdownGuard** | 終了フラグ・親/Excel プロセス死活をポーリング。検知時は on_shutdown または reject/close でダイアログを閉じる。 |

### 3.6 前面追従（オプション）

| 関数 | 説明 |
|------|------|
| **start_front_follow(dialog, parent_hwnd)** | Excel が前面になったときにダイアログを ensure_front するフックを開始。WINDOW.EXCEL_FRONT_FOLLOW が True のとき apply_window_config から呼ばれる（旧キー FRONT_FOLLOW も互換読取）。ウィジェット `destroyed` では `stop_front_follow_if_matches(そのウィジェット)` のみ（進捗→プレビュー移行後に進捗破棄で追従を止めない）。 |
| **stop_front_follow()** | 前面追従を停止。 |
| **stop_front_follow_if_matches(dialog)** | `_front_follow_dialog` が `dialog` と同一インスタンスのときだけ `stop_front_follow`。進捗終了がプレビューより遅れた場合にプレビュー側のフックを止めない。ヘルプ（`ui_help._HelpDialog`）は閉じる／`closeEvent` 先頭で呼び、`enable_excel_window` より前に置く（閉鎖直後の `not_visible_try_restore` が `show()` して窓を蘇生するのを防ぐ）。 |
| **teardown_feature_ui_shared_state(...)** | 機能終了時の共有状態片付け。`stop_front_follow`（`stop_front_follow_match_widget` 指定時は上記条件付き停止）、任意で `enable_excel_window(True)`、任意で `_remove_from_modeless`。二重呼び可。`_close_all_modeless` は含めない。`ProgressDialog` は `match_widget=self`、他機能は未指定で無条件停止。 |

**診断（症状の切り分け）:** **A（Excel 操作ロック／Win32 有効化）** … **`HC_UI_EXCEL_LOCK_DIAG=1`** で `enable_excel_window` 適用時に `hc_csv_tool.diag.ui_excel_lock` へ `[UI_EXCEL_LOCK] sym_type=A` を出力。**B（前面・Z 順）** … **`HC_UI_FG_DIAG=1`** で `[UI_FG]`（`hc_csv_tool.diag.ui_fg`）。

**診断（EXCEL_FRONT_FOLLOW）:** `hc_csv_tool.diag.front_follow` に `ensure_gen_bump`・`schedule_ensure_deferred`・`scheduled_ensure_skip`（`gen_stale` / `dialog_mismatch` / `invalid_shiboken` / `not_visible`）・`scheduled_ensure_run`。`gen_stale` は `stop` 後に積まれていた遅延 `ensure_front` を捨てたとき。メイン `hc_csv.log` には `scheduled_ensure_skip gen_stale` を INFO で残す。`handle_fg` / `cooldown_skip` / `ensure_front_snap` の前景 PID に **`fg_exe`**（`core_w32.get_process_image_path_for_diag`：実行パス短縮）を付与。`schedule_ensure_deferred` / `start_front_follow ok` / `ensure_front enter` / `scheduled_ensure_run` に **`dlg_id`**（`id(widget)`）。`not_visible` スキップ時は **`winId`・`is_hidden`・`window_state`・`elapsed_since_schedule_ms`** を付与し、**`not_visible_try_restore`**（`setVisible`/`show`・最小化解除・`raise_`）のあと可視なら即 `ensure_front`。それでも不可視なら **config/ui_window_timing.json の EXCEL_FRONT_FOLLOW** に従い遅延再試行し、失敗時 `not_visible_exhausted`。`ensure_front_snap` は `ensure_front` 段階ごとに前景・親ルート・`dlg_owner`・（最終段のみ）`sfw_ok` を出す（`HC_UI_FG_DIAG` 不要）。

**TOPMOST と EXCEL_FRONT_FOLLOW（apply_window_config）:** `WINDOW.TOPMOST` または `ALWAYS_IN_FRONT_OF_EXCEL` が true のときは `WindowStaysOnTopHint` のみ。`EXCEL_FRONT_FOLLOW` は **TOPMOST 系が false のときだけ** `start_front_follow` を開始する（両方 true でも FOLLOW は付かない＝表示は TOPMOST 優先）。`SHOW_IN_TASKBAR` は `_hc_show_taskbar` のみで `_set_owner_hwnd` と連動し、本判定では変更しない。

| TOPMOST 系 | EXCEL_FRONT_FOLLOW | 表示方針（apply_window_config） |
|------------|---------------------|----------------------------------|
| true | false / true | `StaysOnTopHint` のみ（FOLLOW なし） |
| false | false | 追従フックなし |
| false | true | `start_front_follow` |

---

## 4. 画面キーと apply_window_config の挙動

- **DONE / PROGRESS**: オーナー設定・ensure_front のタイマーをスキップ（_skip_owner_front）。中央配置は各ダイアログの showEvent 内で実施。
- **その他**: WINDOW に従いフラグ・サイズ・CENTER_ON_EXCEL・remember_last を適用。CENTER_ON_EXCEL 時は 0/50/100/150/300/350 ms でオーナー・前面化、400 ms で再中央配置。**TOPMOST 系が false かつ** EXCEL_FRONT_FOLLOW 時のみ `start_front_follow` を開始。

---

## 5. 参照

| 項目 | 内容 |
|------|------|
| **モジュール** | ui_qt/ui_common.py |
| **共通仕様** | docs/共通仕様_JSON定義.md（WINDOW キー）、docs/共通仕様_機能.md |
| **デグレ防止** | docs/共通モジュール変更時_デグレ防止.md（ui_common 変更時の確認対象：csv_ld, csv_sv, csv_sp, csv_mg の画面・進捗・完了・ワーニング） |
