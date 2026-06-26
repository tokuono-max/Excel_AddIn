# -*- coding: utf-8 -*-
"""
Python: 3.12+
Module: ui_qt/ui_data_agg.py
Created: 2026-03-18
Updated: 2026-06-22
Version: 0.4.48
Purpose:
  データ集約ツールの UI。メイン画面・対象ファイル一覧（別画面）・シナリオ編集・デバッグ（ui_data_agg_debug）・ステップ実行ポップ・進捗・完了を担当する。
  設定は config/ui_data_agg.json。create_dialog は ui_server から呼ばれる。
History (latest 3):
  - 0.4.48 (2026-06-22) showEvent: 同期 _pulse を廃止し _schedule_excel_unlock_pulse_chain に統合（QTimer(0)+90/200/450ms・二重 guid_scan 防止）。create_dialog の重複 deferred pulse を削除。
  - 0.4.47 (2026-06-22) create_dialog(main): WaitForm 合図を prepare 直後に移動。起動時の同期 _pulse を廃止し show 後 QTimer(0) で非同期実行（COM 競合・WaitForm 待ち短縮。guid_scan は維持）。
  - 0.4.46 (2026-06-22) create_dialog(main): 起動区間のフェーズ計測プローブ（DATA_AGG_TRACE phase=… elapsed_ms/step_ms）。_pulse 内 COM サブ区間も診断時に出力。
  - 0.4.45 (2026-06-03) 基準フォルダ走査: 拡張子チェックに .xlsm を追加（.xls/.xlsx と同形式）。
  - 0.4.44 (2026-05-03) EXCEL_LOCK false 時: _pulse_excel_unlock_if_excel_lock_off で Win32 解除に加え CommandBars 有効化・Interactive=True。showEvent で 0/130/400ms の再パルス（ensure_front 直後の無効化取りこぼし緩和）。
  - 0.4.43 (2026-05-03) EXCEL_LOCK false 時: メイン create_dialog（prepare 後）と showEvent で enable_excel_window(True)＋メニューロック解除を明示（子 HWND 無効の取り残しでリボンが効かない事象の緩和）。
  - 0.4.42 (2026-05-03) メイン showEvent: EXCEL_MENU_BAR_LOCK_ON_SHOW は want_excel_child_hwnd_lock_while_modal（WINDOW.EXCEL_LOCK）が true のときのみ適用。ルート共通の EXCEL_LOCK false を開いた直前から反映。
  - 0.4.41 (2026-05-03) merge_screen_cfg_window_from_root で PROGRESS／DONE の WINDOW にルート＋MAIN を反映。完了 showEvent・フォルダ選択の enable_excel_window を want_excel_child_hwnd_lock_while_modal 連動。進捗 req.excel_lock を同一判定で設定。
  - 0.4.40 (2026-05-03) _DataAggDoneDialog: OK を accept 直結から変更し、enable_excel_window(True)＋focus_excel_after_modal_close を OK 押下で明示（closeEvent 非経路のロック残り対策・症状 A）。
  - 0.4.39 (2026-05-02) メイン teardown: モードレス hide では destroyed が来ないため stop_front_follow() をベストエフォートで実行（EXCEL_FRONT_FOLLOW 残留・ゴースト対策）。[DATA_AGG_MAIN_LIFE] stop_front_follow_called。
  - 0.4.38 (2026-05-02) メイン: ゴースト調査用 [DATA_AGG_MAIN_LIFE]（teardown/reject/close/hide/destroyed）と destroyed 接続。ui_common は MODELESS_REMOVE / stop_front_follow を hc_csv.log に出力。
  - 0.4.37 (2026-05-02) メイン: _keep_modeless(..., exclude_from_bulk_close=True) で共通完了の _close_all_modeless から除外（ゴースト／誤閉防止）。
  - 0.4.36 (2026-05-02) メイン項目表: 横ヘッダ左寄せ（setDefaultAlignment）。config の MAIN.WINDOW に EXCEL_FRONT_FOLLOW 追加と整合。
  - 0.4.35 (2026-04-19) 配布 EXE は ``HC_INSTALL_ROOT\\app\\bin\\`` に統一。packaged_app_exe・フォールバックパスを更新。PATH 補強は ``app\\bin`` + インストールルート。
  - 0.4.34 (2026-04-19) 一括実行で short_runner EXE を使うとき、子 env の PATH 先頭に ``app\\shared``・``app`` を追加（runtime_layout.env_with_packaged_dll_search_path）。
  - 0.4.33 (2026-04-16) 一括実行の子プロセス: Nuitka 配布時は `hc_xlwings_short_runner.exe --script-file=...` を使用（`hc_ui_server.exe -c` は WinError 2 になり得る）。`cwd` は `HC_INSTALL_ROOT` が取れるときインストールルートへ。
  - 0.4.32 (2026-04-14) シナリオ Excel 出力を固定列（項目名〜行のルール）の表形式に変更。build_scenario_definition_sheet_matrix_with_headers を使用。
  - 0.4.31 (2026-04-13) apply_window_config: メインは _window_cfg を {"WINDOW": ...} で渡す（SHOW_MINIMIZE 等が効く）。シナリオ編集はルート WINDOW と SCREENS.SCENARIO_EDIT.WINDOW をマージして同関数でタイトルバー・タスクバー方針を統一。
  - 0.4.30 (2026-04-13) create_dialog(main): show 前に install_ribbon_startup_wait_dismiss_on_first_show（VBA WaitForm 解除の初回 Show を取りこぼさない）。デバッグ起動を try/except し失敗時は QMessageBox＋logger.exception。
  - 0.4.29 (2026-04-13) メイン・シナリオ編集: WA_AlwaysShowToolTips と全主要ウィジェットのツールチップ（MAIN.UI の TOOLTIP_*、SCENARIO_EDIT／DETAIL の TIP_*／scenario_layout の既定）。
  - 0.4.28 (2026-04-13) シナリオ編集: SCENARIO_EDIT の MSGBOX_TITLE を DETAIL_NAME に継承。登録検証・種別混在のメッセージを JSON 化（MSG_REGISTER_VALIDATE_PREFIX 等）。レイアウト側のフォールバック文言を JSON 実体に整合。
  - 0.4.27 (2026-04-13) メイン: ui_data_agg.json の DESC_VISIBLE・BTN_SEARCH_RUN・LABEL_MASTER／LABEL_MASTER_ACTIVE_SHEET を反映。走査・項目読込・シナリオ読込／保存・フォルダ選択のダイアログ文言を MAIN.UI の新キーで上書き可能に。
  - 0.4.26 (2026-04-10) ちらつき抑制: メイン・シナリオ編集の遅延オーナー／前面を最小限に（再センタ多段は廃止、prepare の中央寄せに任せる）。デバッグ showEvent の二重 ensure_front を 1 回に。
  - 0.4.25 (2026-04-10) モードレスメイン: 閉じるボタンは reject→hide のみで closeEvent を通さないため、reject でも closeEvent と同じ Excel ロック解除・modeless 除去を実行（× と同等）。
  - 0.4.24 (2026-04-10) メイン: DATA_AGG_MAIN は apply_window_config で遅延オーナーが付かないため、初回 showEvent で ui_common と同系の QTimer オーナー・前面・再センタを追加。リボン抑止に EXCEL_LOCK_INTERACTIVE（Application.Interactive）を任意で併用（MAIN.WINDOW、既定 true）。
  - 0.4.23 (2026-04-10) メイン・DEBUG の SHOW_IN_TASKBAR をオフにし Excel オーナー紐づけを有効化（前面・マスタデバッグの Excel 背面を改善）。リボン抑止は成功時のみ applied フラグ＋遅延再試行。
  - 0.4.22 (2026-04-10) 段階1: シナリオ編集・デバッグ子ダイアログをメインと同系統に（exec 前 prepare_dialog_excel_center、シナリオ編集 showEvent で遅延オーナー・前面・再センタ）。
  - 0.4.21 (2026-04-10) シナリオ編集: 空のシナリオ名の既定表示を行番号ではなく「既存名と重ならない連番」（上から順に割当）に変更。挿入でシナリオ1が重複しない。
  - 0.4.20 (2026-04-10) シナリオ編集: `_summary_table_tooltip` を `_DataAggMainWindow` 専用と誤参照していたデグレを修正（`_data_agg_summary_table_tooltip` 共用）。
  - 0.4.19 (2026-04-10) シナリオ編集診断: `_refresh_sources_table` に `refresh_enter` / `refresh_leave` / `refresh_abort`（例外時 traceback）。blockSignals を try/finally で確実に解除。
  - 0.4.18 (2026-04-10) シナリオ編集診断: `sync_enter` / `after_refresh`（追加ボタン経路）/ `sync_abort`（同期途中の未捕捉例外）。
  - 0.4.17 (2026-04-10) シナリオ編集ダイアログ: 診断ログ `[DATA_AGG_SCENARIO_EDIT]`（hc_csv_diag.log、診断有効時）。load 失敗時も右ペイン有効化を継続。
  - 0.4.16 (2026-04-09) IPC の excel_rect を center_on_excel の rect_override に反映（メイン遅延再センタ・完了・ステップポップ）。進捗は req 引き継ぎで既存 ProgressDialog と整合。
  - 0.4.15 (2026-04-06) シナリオ出力: タイトル行・/区切り要約・列オートフィット。メイン項目表の右クリック。メイン表示中は Excel リボン抑止（ベストエフォート）。
  - 0.4.14 (2026-04-06) 一括後はデータ集約シートを再アクティブ化。メイン WINDOW マージで Excel 中央。ソース表の右クリック。シナリオ定義の Excel 出力。
  - 0.4.13 (2026-04-06) 一括実行・デバッグ: いずれかのマスタ項目に取得ソースがあれば有効。子プロセスに core_env.ipc_dir_raw() で HC_IPC_ROOT を明示。
  - 0.4.12 (2026-04-06) すべてクリア: 基準フォルダ入力と検出ファイル一覧も空にする。
  - 0.4.11 (2026-04-06) 一括実行: シナリオファイル有無ではなく、全項目に取得ソース登録があるかで有効化。
  - 0.4.10 (2026-04-06) Excel タブ背景を兄弟タブの Base（白系）に揃え、Window グレーとの差を解消。
  - 0.4.9 (2026-04-06) Excel タブ: QGroupBox 等の背景をタブと同色に統一。一括実行はシナリオファイル読込済みのときのみ有効。
  - 0.4.8 (2026-04-06) create_dialog 入口に運用ログ・診断トレース（ui_server source_req との相関用）。
  - 0.4.7 (2026-04-05) クリアボタンをシナリオ名行右寄せ。シナリオクリア後の保存はファイル名空。項目読込前にすべてクリア。
  - 0.4.6 (2026-04-05) シナリオ編集: 0件から追加時の右ペイン有効化。メイン: すべてクリア／シナリオクリア。
  - 0.4.5 (2026-04-04) シナリオ名ラベル、キー不一致／指定行 UI 削除、イベントログに書込み方式・出力シート名列。
  - 0.4.4 (2026-04-04) Excel タブ: 新規シート「シート名入力」、シナリオ名_連番は読込ファイル名ベース。
  - 0.4.3 (2026-04-04) 一括実行はシナリオ JSON を上書きせず一時スナップショットを子プロセスへ渡す。
  - 0.4.2 (2026-04-04) メイン画面に Excel タブ（出力先・ジャンプ・並べ替え）。文言・ツールチップは MAIN.UI、QGroupBox でグループ化。
  - 0.4.1 (2026-04-01) 一括進捗: 表示中メイン画面を親とし中央表示（Excel 中央への依存を回避）。
  - 0.4.0 (2026-03-25) デバッグボタン・SCREENS.DEBUG 連携（ui_data_agg_debug.py）。シナリオ編集のステップ実行から同一デバッグを起動。
  - 0.3.0 (2026-03-18) メイン画面レイアウト変更: 対象ファイル一覧を別画面に、項目縦列+シナリオ編集+要約、項目読込3通り（CSV/シート/直接編集）。
  - 0.2.0 (2026-03-18) Phase A/B: ボタンハンドラ接続、項目一覧・シナリオ要約・走査条件 UI 追加。
  - 0.1.0 (2026-03-18) Phase4: create_dialog（main / progress / done / step_popup）とメイン画面の骨子。
"""
from __future__ import annotations

import copy
import json
import os
import subprocess
import sys
import tempfile
import time
import traceback
import uuid
from pathlib import Path
from typing import Any, Callable

from PySide6.QtCore import QItemSelectionModel, QObject, QPoint, QThread, QTimer, Qt, Signal, Slot
from PySide6.QtGui import (
    QAction,
    QColor,
    QCloseEvent,
    QHideEvent,
    QPalette,
    QResizeEvent,
    QShowEvent,
)
from PySide6.QtWidgets import (
    QAbstractButton,
    QAbstractItemView,
    QButtonGroup,
    QCheckBox,
    QComboBox,
    QDialog,
    QFileDialog,
    QFormLayout,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QInputDialog,
    QLabel,
    QLineEdit,
    QListWidget,
    QMenu,
    QMessageBox,
    QPushButton,
    QRadioButton,
    QScrollArea,
    QSizePolicy,
    QSpinBox,
    QSplitter,
    QStackedWidget,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from core import core_env
from ui_qt.ui_common import (
    _normalize_message_newlines,
    excel_rect_tuple_from_req as _excel_rect_tuple_from_req,
    show_done_notice,
    show_error_notice,
    show_info_notice,
    show_warning_notice,
)
from ui_qt.ui_data_agg_scenario_layout import (
    build_scenario_detail_cell_scroll,
    build_scenario_detail_name_scroll,
)
from ui_qt.ipc_file import (
    delete_batch_done_notify,
    write_pickle,
    get_last_folder,
    set_last_folder,
    try_read_batch_done_notify,
)
from svc.data_agg_source_ui import ensure_source_ui_block, source_ui_block
from svc.data_agg_name_extract_summary import (
    fmt_ne_length_mode as _fmt_ne_length_mode,
    fmt_ne_start_mode as _fmt_ne_start_mode,
    fmt_ne_write_mode as _fmt_ne_write_mode,
    ja_search_cond_static as _ja_search_cond_static,
    ja_search_target_static as _ja_search_target_static,
    name_extract_full_detail_lines,
    name_extract_setting_lines,
)
from svc.data_agg_cell_coordinate_summary import cell_coordinate_full_detail_lines
from svc.data_agg_source_list_display import (
    scenario_source_kind_label_and_summary,
    scenario_source_tooltip_plain,
)

from core.core_log import get_data_agg_diag_logger, get_logger

logger = get_logger(__name__)
_data_agg_ui_diag = get_data_agg_diag_logger()


def folder_scan_paths_from_state(state: dict[str, Any]) -> list[str]:
    """走査条件 dict から scan_folder を実行し、パス文字列リストを返す（UI スレッド外可）。"""
    from svc import svc_data_agg_scan as scan_mod  # noqa: WPS433

    sp = str(state.get("start_path") or "").strip() or "."
    exts = list(state.get("extensions") or [])
    if not exts:
        return []
    paths = scan_mod.scan_folder(
        sp,
        recursive=bool(state.get("recursive")),
        extensions=tuple(exts),
        keyword=str(state.get("keyword") or ""),
    )
    return [str(p) for p in paths]


def should_apply_folder_scan_result(generation: int, current_generation: int) -> bool:
    """走査世代が最新のときだけ UI へ結果を反映する。"""
    return int(generation) == int(current_generation)


class _FolderScanWorker(QObject):
    """フォルダ走査をバックグラウンドスレッドで実行する。"""

    finished = Signal(int, object)
    failed = Signal(int, str)

    def __init__(self, generation: int, state: dict[str, Any]) -> None:
        super().__init__()
        self._generation = int(generation)
        self._state = dict(state)

    @Slot()
    def run(self) -> None:
        try:
            paths = folder_scan_paths_from_state(self._state)
            self.finished.emit(self._generation, paths)
        except Exception as ex:
            self.failed.emit(self._generation, str(ex))


def _log_scenario_edit_diag(fmt: str, *args: Any) -> None:
    """シナリオ編集の診断行（hc_csv_diag.log、診断ファイル有効時のみ）。"""
    try:
        _data_agg_ui_diag.info("[DATA_AGG_SCENARIO_EDIT] " + fmt, *args)
    except Exception:
        pass


def _log_data_agg_main_lifecycle(w: QWidget, where: str, extra: str = "") -> None:
    """メイン画面の閉じる経路調査（hc_csv.log / hc_csv_diag）。"""
    try:
        wid = int(w.winId())
    except Exception:
        wid = 0
    try:
        vis = bool(w.isVisible())
    except Exception:
        vis = False
    sid = str(getattr(w, "_sheet_id", "") or "")
    ph = int(getattr(w, "_parent_hwnd", 0) or 0)
    tail = (extra or "").strip()
    if tail:
        tail = " " + tail
    msg = (
        "[DATA_AGG_MAIN_LIFE] %s sheet_id=%s parent_hwnd=%s winId=%s visible=%s%s"
        % (where, sid, ph, wid, vis, tail)
    )
    try:
        logger.info(msg)
    except Exception:
        pass
    try:
        _data_agg_ui_diag.info(msg)
    except Exception:
        pass


def _log_data_agg_create_dialog_phase(
    phase: str,
    *,
    t0: float,
    t_prev: float,
    parent_hwnd: int = 0,
    extra: str = "",
) -> float:
    """create_dialog(main) 起動区間の計測プローブ（hc_csv_diag.log・HC_LOG_DIAG=1 等）。"""
    now = time.perf_counter()
    tail = (" " + str(extra).strip()) if str(extra or "").strip() else ""
    try:
        _data_agg_ui_diag.info(
            "[DATA_AGG_TRACE] create_dialog phase=%s parent_hwnd=%s "
            "elapsed_ms=%d step_ms=%d wall_perf_s=%.6f%s",
            str(phase or "").strip() or "?",
            int(parent_hwnd or 0),
            int((now - t0) * 1000),
            int((now - t_prev) * 1000),
            now,
            tail,
        )
    except Exception:
        pass
    return now


def _data_agg_summary_table_tooltip(display: str) -> str:
    """要約表ツールチップ: 区切り「 | 」を改行して複数行表示（メイン項目表・シナリオ編集一覧で共用）。"""
    return (display or "").replace(" | ", "\n")


__version__ = "0.4.48"

# メイン項目表: 連携参照行・結合参照行の背景（連携優先で灰）
_ROW_BG_LINK = QColor("#E0E0E0")
_ROW_BG_JOIN = QColor("#F5F0E6")
_ROW_FG_LINKED = QColor("#666666")

# メイン項目名列: 連携/結合参照のリネーム追跡用（直前の確定マスタ名）
_ITEM_MASTER_NAME_ROLE = int(Qt.ItemDataRole.UserRole)

_BTN_EDIT_ENABLED = (
    "QPushButton { background-color: #2196F3; color: white; font-weight: bold; "
    "border-radius: 4px; padding: 4px 12px; border: 1px solid #1976D2; } "
    "QPushButton:hover { background-color: #1976D2; } "
    "QPushButton:pressed { background-color: #0D47A1; }"
)
_BTN_EDIT_LINKED_DISABLED = (
    "QPushButton { background-color: #E0E0E0; color: #888888; font-weight: bold; "
    "border-radius: 4px; padding: 4px 12px; border: 1px solid #CCCCCC; }"
)

# Excel タブ: コンボ・入力の縦幅・最大幅（コンパクト表示）
_EXCEL_CTRL_MAX_H = 26
_EXCEL_COMBO_WRITE_MODE_W = 200
_EXCEL_COMBO_SHEET_RULE_W = 320
_EXCEL_EDIT_ANCHOR_W = 88
_EXCEL_COMBO_SORT_ITEM_W = 220
_EXCEL_COMBO_SORT_ORDER_W = 88
_EXCEL_SORT_BTN_MAX_H = 26


def _batch_active_path(sheet_id: str, ipc_root: Path) -> Path:
    d = ipc_root / "progress"
    d.mkdir(parents=True, exist_ok=True)
    sid = str(sheet_id or "").strip() or "default"
    return d / ("data_agg_batch_active_%s.pkl" % sid)


def _excel_compact_control(w: QWidget, max_width: int | None = None) -> None:
    """Excel オプション用のコンボ・入力をコンパクトにする。"""
    w.setMaximumHeight(_EXCEL_CTRL_MAX_H)
    if max_width is not None:
        w.setMaximumWidth(max_width)


def _get_cfg() -> dict[str, Any]:
    """
    データ集約用の画面設定を config/ui_data_agg.json から読み込む。
    読込失敗時は UiConfigLoadError が発生する。
    """
    from core import core_cst as cst

    return cst.get_ui_config_from_file_required("data_agg")


def _ui_disp_str(block: dict[str, Any], key: str, default: str) -> str:
    """UI ブロック（JSON）由来の表示文字列。\\n / リテラル \\\\n を改行に。"""
    return _normalize_message_newlines(str((block or {}).get(key) or default).strip())


def _global_detail_name_cfg() -> dict[str, Any]:
    """SCREENS.SCENARIO_EDIT.DETAIL_NAME（名前から取得フォーム文言）。メイン要約など編集ダイアログ外から参照。"""
    try:
        root = _get_cfg()
    except Exception:
        return {}
    se = (root.get("SCREENS") or {}).get("SCENARIO_EDIT") or {}
    d = se.get("DETAIL_NAME")
    return d if isinstance(d, dict) else {}


def _global_detail_cell_cfg() -> dict[str, Any]:
    """SCREENS.SCENARIO_EDIT.DETAIL_CELL（セル座標フォーム文言）。メイン要約ツールチップ等で参照。"""
    try:
        root = _get_cfg()
    except Exception:
        return {}
    se = (root.get("SCREENS") or {}).get("SCENARIO_EDIT") or {}
    d = se.get("DETAIL_CELL")
    return d if isinstance(d, dict) else {}


def _scenario_lineage_bucket(stype: str) -> str:
    """シナリオソース type をセル座標系 / 名前・パス系のいずれかに分類する。"""
    t = (stype or "cell").strip().lower()
    if t == "cell":
        return "cell"
    return "path"


def _find_data_agg_main_window(sheet_id: str, excel_hwnd: int) -> QWidget | None:
    """
    表示中のデータ集約メイン画面を返す。一括進捗を Excel ではなく当該画面の中央に置くため。
    sheet_id が IPC で「_」になった場合は Excel HWND が一致するメイン画面を優先する。
    """
    from PySide6.QtWidgets import QApplication

    app = QApplication.instance()
    if not isinstance(app, QApplication):
        return None
    candidates: list[QWidget] = []
    for w in app.topLevelWidgets():
        if not w.isVisible():
            continue
        if type(w).__name__ != "_DataAggMainWindow":
            continue
        candidates.append(w)
    if not candidates:
        return None
    sid = str(sheet_id or "").strip()
    if sid and sid != "_":
        for w in candidates:
            if str(getattr(w, "_sheet_id", "") or "").strip() == sid:
                return w
    eh = int(excel_hwnd or 0)
    if eh:
        same_excel = [
            w
            for w in candidates
            if int(getattr(w, "_parent_hwnd", 0) or 0) == eh
        ]
        if same_excel:
            return same_excel[0]
    if len(candidates) == 1:
        return candidates[0]
    return None


def _data_agg_excel_parent_hwnd_rect(
    from_widget: QWidget | None,
) -> tuple[int, tuple[int, int, int, int] | None]:
    """
    親チェーン上の _DataAggMainWindow から Excel HWND と IPC の excel_rect を取得する。
    子ダイアログの Win32 オーナー・中央配置をメインと揃えるために使う。
    """
    w: QWidget | None = from_widget
    while w is not None:
        ph = int(getattr(w, "_parent_hwnd", 0) or 0)
        if ph:
            req = getattr(w, "_req", None)
            rect = (
                _excel_rect_tuple_from_req(req)
                if isinstance(req, dict)
                else None
            )
            return ph, rect
        w = w.parentWidget()
    return 0, None


def _data_agg_prepare_subdialog_excel_center(
    dlg: QWidget,
    from_widget: QWidget | None,
    window_cfg: dict[str, Any] | None = None,
) -> None:
    """exec/show 直前に、Excel 中央・オーナー HWND・前面化をメイン画面と同様に 1 回適用する。"""
    ph, rect = _data_agg_excel_parent_hwnd_rect(from_widget)
    if not ph:
        return
    try:
        from ui_qt.ui_common import prepare_dialog_excel_center_before_show

        prepare_dialog_excel_center_before_show(dlg, ph, rect, window_cfg)
    except Exception:
        pass


def _data_agg_warn_debug_open_failed(
    parent: QWidget | None, title: str, exc: BaseException
) -> None:
    """デバッグダイアログ生成・表示の失敗をユーザーへ通知し、運用ログにスタックを残す。"""
    try:
        logger.exception("[DATA_AGG_UI] debug dialog open failed title=%s", title)
    except Exception:
        pass
    try:
        show_error_notice(
            parent,
            title,
            "デバッグ画面を開けませんでした。\n\n%s" % exc,
        )
    except Exception:
        pass


class _DataAggMainWindow(QDialog):
    """
    データ集約ツールのメイン画面。
    項目一覧（縦）と各項目ごとのシナリオ要約を表示する。フォルダ指定・走査条件・シナリオ保存/読込の入口。
    """

    def __init__(
        self,
        req: dict[str, Any],
        parent_hwnd: int,
        sheet_id: str,
        main_cfg: dict[str, Any],
        window_cfg: dict[str, Any],
    ) -> None:
        super().__init__()
        try:
            self.setAttribute(Qt.WidgetAttribute.WA_AlwaysShowToolTips, True)
        except Exception:
            pass
        self._req = req or {}
        self._parent_hwnd = int(parent_hwnd or 0)
        self._sheet_id = str(sheet_id or "")
        self._main_cfg = main_cfg or {}
        self._window_cfg = window_cfg or {}
        self._scenario: dict[str, Any] = {}
        self._scenario_path: str = ""
        self._scenario_dirty: bool = False
        self._suppress_scenario_dirty: bool = False
        self._scenario_save_empty_filename: bool = False
        self._batch_poll_timer: QTimer | None = None
        self._batch_poll_deadline: float = 0.0
        self._batch_poll_run_id: str = ""
        self._excel_menu_bar_lock_applied: bool = False
        self._excel_menu_lock_app: Any = None
        self._excel_lock_interactive_prev: bool | None = None
        self._excel_deferred_owner_front_scheduled: bool = False
        self._messages = _get_cfg().get("MESSAGES") or {}
        self._ui = (self._main_cfg.get("UI") or {})
        _u = lambda k, d: _ui_disp_str(self._ui, k, d)

        title = _normalize_message_newlines(
            str(self._main_cfg.get("TITLE") or "データ集約ツール").strip()
        )
        self.setWindowTitle(title)
        layout = QVBoxLayout(self)
        desc_visible = (
            self._main_cfg.get("DESC_VISIBLE")
            if isinstance(self._main_cfg.get("DESC_VISIBLE"), bool)
            else True
        )
        if desc_visible:
            desc_raw = str(
                self._main_cfg.get("DESC")
                or self._main_cfg.get("DESCRIPTION")
                or ""
            ).strip(" \t\r")
            if not desc_raw:
                desc_raw = (
                    "項目（列）ごとに取得ソース・抽出ルールを定義し、マスターへ統合します。"
                )
            desc = _normalize_message_newlines(desc_raw)
            lbl = QLabel(desc)
            lbl.setWordWrap(True)
            lbl.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
            layout.addWidget(lbl)
            self._main_set_tip(
                lbl,
                "TOOLTIP_DESC_MAIN",
                "データ集約の概要です。項目ごとの取得ルールと一括実行の入口です。",
            )
            if desc.endswith("\n"):
                try:
                    fm = lbl.fontMetrics()
                    layout.addSpacing(max(4, fm.lineSpacing()))
                except Exception:
                    layout.addSpacing(8)
        # 項目読込（アクティブシート1行目 / CSVファイル → 取得ボタンで実行）
        grp_item_read = QGroupBox(_u("GROUP_ITEM_READ", "項目読込"))
        self._main_set_tip(
            grp_item_read,
            "TOOLTIP_GROUP_ITEM_READ",
            "マスタ項目名を Excel シート先頭行または CSV から読み込む設定です。",
        )
        read_layout = QHBoxLayout(grp_item_read)
        self._item_read_group = QButtonGroup(self)
        self._rad_sheet = QRadioButton(_u("RADIO_SHEET", "アクティブシート1行目"))
        self._rad_csv = QRadioButton(_u("RADIO_CSV", "CSVファイル"))
        self._main_set_tip(
            self._rad_sheet,
            "TOOLTIP_ITEM_READ_SHEET",
            "アクティブブックのシート先頭行を項目名として読み込みます。",
        )
        self._main_set_tip(
            self._rad_csv,
            "TOOLTIP_ITEM_READ_CSV",
            "項目定義 CSV を選んで項目名を読み込みます。",
        )
        self._rad_sheet.setChecked(True)
        self._item_read_group.addButton(self._rad_sheet)
        self._item_read_group.addButton(self._rad_csv)
        read_layout.addWidget(self._rad_sheet)
        read_layout.addWidget(self._rad_csv)
        self._btn_item_load = QPushButton(_u("BTN_LOAD_ITEMS", "取得"))
        self._btn_item_load.clicked.connect(self._on_item_load)
        self._btn_item_load.setAutoDefault(False)
        self._btn_item_load.setDefault(False)
        self._main_set_tip(
            self._btn_item_load,
            "TOOLTIP_BTN_LOAD_ITEMS",
            "選択した読込方式で項目一覧を取得・更新します。",
        )
        read_layout.addWidget(self._btn_item_load)
        read_layout.addStretch(1)
        layout.addWidget(grp_item_read)
        # 項目一覧・シナリオ要約（タブ切り替え）
        grp_items = QGroupBox(_u("GROUP_ITEMS", "項目一覧・シナリオ要約"))
        self._main_set_tip(
            grp_items,
            "TOOLTIP_GROUP_ITEMS",
            "項目・シナリオ要約・基準フォルダ・Excel オプションをまとめた領域です。",
        )
        items_layout = QVBoxLayout(grp_items)
        hint_master = str(self._ui.get("LABEL_MASTER") or "").strip()
        if hint_master:
            lbl_master_hint = QLabel(hint_master)
            lbl_master_hint.setStyleSheet("color: #555; font-size: 11px;")
            items_layout.addWidget(lbl_master_hint)
            self._main_set_tip(
                lbl_master_hint,
                "TOOLTIP_LBL_MASTER_LINE",
                "マスタ（出力先）に関する注記です。",
            )
        self._lbl_scenario_display_name = QLabel()
        self._lbl_scenario_display_name.setStyleSheet(
            "color: #1565C0; font-weight: bold; font-size: 12px;"
        )
        row_scenario_header = QHBoxLayout()
        row_scenario_header.setContentsMargins(0, 0, 0, 0)
        row_scenario_header.setSpacing(8)
        row_scenario_header.addWidget(self._lbl_scenario_display_name, 0)
        row_scenario_header.addStretch(1)
        self._main_set_tip(
            self._lbl_scenario_display_name,
            "TOOLTIP_SCENARIO_DISPLAY_NAME",
            "現在読み込んでいるシナリオファイル名の表示です。",
        )
        btn_clear_all = QPushButton(_u("BTN_SCENARIO_CLEAR_ALL", "すべてクリア"))
        btn_clear_all.setToolTip(
            _u(
                "TOOLTIP_SCENARIO_CLEAR_ALL",
                "マスタ項目一覧とシナリオ定義を空にし、未読込状態に戻します。シナリオ保存は無効のままです。",
            )
        )
        btn_clear_all.setAutoDefault(False)
        btn_clear_all.setDefault(False)
        btn_clear_all.clicked.connect(self._on_scenario_clear_all)
        btn_clear_sc = QPushButton(_u("BTN_SCENARIO_CLEAR_SOURCES", "シナリオクリア"))
        btn_clear_sc.setToolTip(
            _u(
                "TOOLTIP_SCENARIO_CLEAR_SOURCES",
                "項目名は残し、各項目に登録した取得シナリオのみ削除します。保存が必要な変更としてシナリオ保存を有効にします。",
            )
        )
        btn_clear_sc.setAutoDefault(False)
        btn_clear_sc.setDefault(False)
        btn_clear_sc.clicked.connect(self._on_scenario_clear_sources_only)
        row_scenario_header.addWidget(btn_clear_all, 0)
        row_scenario_header.addWidget(btn_clear_sc, 0)
        items_layout.addLayout(row_scenario_header)
        tab_widget = QTabWidget()
        # タブ1: 項目・シナリオ
        tab_items = QWidget()
        tab_items_layout = QVBoxLayout(tab_items)
        items_row = QHBoxLayout()
        self._item_table = QTableWidget()
        self._item_table.setColumnCount(3)
        self._item_table.setHorizontalHeaderLabels([
            _u("TABLE_HEADER_NAME", "項目名（マスター）"),
            _u("TABLE_HEADER_EDIT", "シナリオ"),
            _u("TABLE_HEADER_SUMMARY", "シナリオ要約"),
        ])
        self._item_table.setMinimumHeight(300)
        self._item_table.verticalHeader().setDefaultSectionSize(24)
        _ith = self._item_table.horizontalHeader()
        _ith.setDefaultAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        _ith.setStretchLastSection(False)
        _ith.setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        _ith.setSectionResizeMode(1, QHeaderView.ResizeMode.Fixed)
        _ith.setSectionResizeMode(2, QHeaderView.ResizeMode.Fixed)
        self._item_table.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAsNeeded
        )
        _ww = getattr(self._item_table, "setWordWrap", None)
        if callable(_ww):
            _ww(False)
        self._item_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self._item_table.setSelectionMode(QTableWidget.SelectionMode.ExtendedSelection)
        self._item_table.setAlternatingRowColors(False)
        self._item_table.setStyleSheet(
            "QTableWidget::item:selected { background-color: #B0C4DE; } "
            "QTableWidget::item:selected:!active { background-color: #C8D4E0; }"
        )
        items_row.addWidget(self._item_table, 1)
        self._item_table.cellDoubleClicked.connect(self._on_item_table_double_clicked)
        self._item_table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self._item_table.customContextMenuRequested.connect(
            self._on_item_table_context_menu
        )
        self._main_set_tip(
            self._item_table,
            "TOOLTIP_ITEM_TABLE",
            "項目名・シナリオ編集・要約を表示する一覧です。項目名列のダブルクリックで編集し、シナリオ編集はボタンから開きます。",
        )
        self._main_set_tip(
            self._item_table.horizontalHeader(),
            "TOOLTIP_ITEM_TABLE_HEADER",
            "項目名・シナリオ・要約の各列見出しです。",
        )
        self._main_set_tip(
            self._item_table.verticalHeader(),
            "TOOLTIP_ITEM_TABLE_VERTICAL",
            "項目行の番号です。",
        )
        tab_items_layout.addLayout(items_row)
        move_row = QHBoxLayout()
        move_row.addStretch(0)
        btn_up = QPushButton(_u("BTN_MOVE_UP", "▲ 上へ"))
        btn_up.setToolTip(_u("TOOLTIP_MOVE_UP", "選択行を1行上に移動（連続・歯抜け選択可）"))
        btn_up.clicked.connect(self._on_move_items_up)
        btn_down = QPushButton(_u("BTN_MOVE_DOWN", "▼ 下へ"))
        btn_down.setToolTip(_u("TOOLTIP_MOVE_DOWN", "選択行を1行下に移動（連続・歯抜け選択可）"))
        btn_down.clicked.connect(self._on_move_items_down)
        move_row.addWidget(btn_up)
        move_row.addWidget(btn_down)
        move_row.addStretch(1)
        tab_items_layout.addLayout(move_row)
        self._lbl_item_total = QLabel()
        self._lbl_item_total.setStyleSheet("color: #555; font-size: 11px;")
        tab_items_layout.addWidget(self._lbl_item_total)
        self._main_set_tip(
            self._lbl_item_total,
            "TOOLTIP_LBL_ITEM_TOTAL",
            "読み込んだマスタ項目の件数表示です。",
        )
        # タブ2: フォルダ/ファイル・検出一覧（ウィジェットは後で addTab 順のみ入替）
        tab_scan = QWidget()
        tab_scan_layout = QVBoxLayout(tab_scan)
        grp_scan = QGroupBox(_u("GROUP_SCAN", "フォルダ/ファイル条件"))
        self._main_set_tip(
            grp_scan,
            "TOOLTIP_GROUP_SCAN",
            "一括実行・デバッグで参照する基準フォルダとファイル検索条件です。",
        )
        scan_layout = QVBoxLayout(grp_scan)
        row_path = QHBoxLayout()
        lbl_base_folder = QLabel(_u("LABEL_BASE_FOLDER", "基準フォルダ") + ":")
        self._main_set_tip(
            lbl_base_folder,
            "TOOLTIP_LABEL_BASE_FOLDER",
            "走査の起点となるフォルダです。",
        )
        row_path.addWidget(lbl_base_folder)
        self._edit_start_path = QLineEdit()
        self._main_set_tip(
            self._edit_start_path,
            "TOOLTIP_EDIT_START_PATH",
            "基準フォルダのパスです。フォルダ選択で設定できます。",
        )
        row_path.addWidget(self._edit_start_path, 1)
        btn_folder = QPushButton(_u("BTN_FOLDER", "フォルダ選択"))
        self._main_set_tip(
            btn_folder,
            "TOOLTIP_BTN_FOLDER",
            "基準フォルダをダイアログで選びます。",
        )
        btn_folder.setAutoDefault(False)
        btn_folder.setDefault(False)
        btn_folder.clicked.connect(self._on_folder_select)
        row_path.addWidget(btn_folder)
        scan_layout.addLayout(row_path)
        row_opts = QHBoxLayout()
        self._chk_recursive = QCheckBox(_u("LABEL_SUBFOLDER", "サブフォルダ含む"))
        self._chk_recursive.setChecked(False)
        self._main_set_tip(
            self._chk_recursive,
            "TOOLTIP_CHK_SUBFOLDER",
            "オンにするとサブフォルダも再帰的に走査します。",
        )
        row_opts.addWidget(self._chk_recursive)
        self._lbl_keyword = QLabel(_u("LABEL_KEYWORD", "キーワード") + ":")
        self._edit_keyword = QLineEdit()
        self._lbl_keyword.setVisible(False)
        self._edit_keyword.setVisible(False)
        self._main_set_tip(
            self._lbl_keyword,
            "TOOLTIP_LABEL_KEYWORD",
            "（将来用）ファイル名などのキーワード条件のラベルです。",
        )
        self._main_set_tip(
            self._edit_keyword,
            "TOOLTIP_EDIT_KEYWORD",
            "（将来用）キーワード入力欄です。",
        )
        row_opts.addWidget(self._lbl_keyword)
        row_opts.addWidget(self._edit_keyword)
        lbl_ext = QLabel(_u("LABEL_EXTENSIONS", "拡張子") + ":")
        self._main_set_tip(
            lbl_ext,
            "TOOLTIP_LABEL_EXTENSIONS",
            "検索対象とする拡張子です。",
        )
        row_opts.addWidget(lbl_ext)
        self._chk_ext_xls = QCheckBox(_u("CHK_EXT_XLS", ".xls"))
        self._chk_ext_xlsx = QCheckBox(_u("CHK_EXT_XLSX", ".xlsx"))
        self._chk_ext_xlsm = QCheckBox(_u("CHK_EXT_XLSM", ".xlsm"))
        self._chk_ext_csv = QCheckBox(_u("CHK_EXT_CSV", ".csv"))
        self._main_set_tip(
            self._chk_ext_xls,
            "TOOLTIP_CHK_EXT_XLS",
            ".xls を検索対象に含めます。",
        )
        self._main_set_tip(
            self._chk_ext_xlsx,
            "TOOLTIP_CHK_EXT_XLSX",
            ".xlsx を検索対象に含めます。",
        )
        self._main_set_tip(
            self._chk_ext_xlsm,
            "TOOLTIP_CHK_EXT_XLSM",
            ".xlsm を検索対象に含めます。",
        )
        self._main_set_tip(
            self._chk_ext_csv,
            "TOOLTIP_CHK_EXT_CSV",
            ".csv を検索対象に含めます。",
        )
        self._chk_ext_xls.setChecked(True)
        self._chk_ext_xlsx.setChecked(True)
        self._chk_ext_xlsm.setChecked(True)
        self._chk_ext_csv.setChecked(True)
        row_opts.addWidget(self._chk_ext_xls)
        row_opts.addWidget(self._chk_ext_xlsx)
        row_opts.addWidget(self._chk_ext_xlsm)
        row_opts.addWidget(self._chk_ext_csv)
        row_opts.addStretch(1)
        scan_layout.addLayout(row_opts)
        row_scan_run = QHBoxLayout()
        row_scan_run.addStretch(1)
        self._btn_scan_run = QPushButton(_u("BTN_SEARCH_RUN", "検索実行"))
        self._btn_scan_run.setAutoDefault(False)
        self._btn_scan_run.setDefault(False)
        self._btn_scan_run.clicked.connect(lambda: self._on_scan(False))
        btn_scan_run = self._btn_scan_run
        self._main_set_tip(
            btn_scan_run,
            "TOOLTIP_BTN_SEARCH_RUN",
            "条件に従いフォルダを走査し、下の一覧を更新します。",
        )
        row_scan_run.addWidget(btn_scan_run)
        scan_layout.addLayout(row_scan_run)
        tab_scan_layout.addWidget(grp_scan)
        grp_files = QGroupBox(_u("GROUP_DETECTED_FILES", "検出ファイル一覧"))
        self._main_set_tip(
            grp_files,
            "TOOLTIP_GROUP_DETECTED_FILES",
            "走査で見つかったファイルパスの一覧です。デバッグの入力にも使われます。",
        )
        file_layout = QVBoxLayout(grp_files)
        self._file_list = QListWidget()
        self._file_list.setMinimumHeight(180)
        self._file_list.setAlternatingRowColors(False)
        self._main_set_tip(
            self._file_list,
            "TOOLTIP_FILE_LIST",
            "検出されたファイルのフルパス一覧です。",
        )
        file_layout.addWidget(self._file_list)
        self._lbl_detected_file_count = QLabel()
        self._lbl_detected_file_count.setStyleSheet("color: #555; font-size: 11px;")
        self._lbl_detected_file_count.setAlignment(
            Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter
        )
        file_layout.addWidget(self._lbl_detected_file_count)
        self._main_set_tip(
            self._lbl_detected_file_count,
            "TOOLTIP_LBL_FILE_COUNT",
            "検出ファイル件数の表示です。",
        )
        tab_scan_layout.addWidget(grp_files)
        tab_widget.addTab(tab_scan, _u("TAB_SCAN", "基準フォルダ"))
        tab_widget.addTab(tab_items, _u("TAB_ITEMS", "シナリオ"))
        tab_widget.addTab(
            self._create_excel_options_tab(_u, ref_tab=tab_items),
            _u("TAB_EXCEL", "Excel"),
        )
        self._main_set_tip(
            tab_widget,
            "TOOLTIP_TAB_WIDGET_MAIN",
            "基準フォルダ・シナリオ一覧・Excel 設定のタブです。",
        )
        try:
            tab_widget.setTabToolTip(
                0,
                _u(
                    "TOOLTIP_TAB_SCAN",
                    "基準フォルダ・拡張子・検索で対象ファイルを列挙します。",
                ),
            )
            tab_widget.setTabToolTip(
                1,
                _u(
                    "TOOLTIP_TAB_ITEMS",
                    "マスタ項目とシナリオ要約の一覧です。行の移動やシナリオ編集の起点になります。",
                ),
            )
            tab_widget.setTabToolTip(
                2,
                _u(
                    "TOOLTIP_TAB_EXCEL",
                    "マスタへの書き込み先・並べ替えなどのオプションです。",
                ),
            )
        except Exception:
            pass
        self._main_set_tip(
            tab_scan,
            "TOOLTIP_TAB_SCAN",
            "基準フォルダ・拡張子・検索で対象ファイルを列挙します。",
        )
        self._main_set_tip(
            tab_items,
            "TOOLTIP_TAB_ITEMS",
            "マスタ項目とシナリオ要約の一覧です。",
        )
        items_layout.addWidget(tab_widget)
        self._fit_item_table_columns()
        layout.addWidget(grp_items)
        self._file_list_items: list[str] = []
        self._scan_generation: int = 0
        self._scan_busy: bool = False
        self._scan_req_auto_mode: bool = True
        self._scan_req_on_complete: Callable[[int], None] | None = None
        self._scan_thread: QThread | None = None
        self._scan_worker: _FolderScanWorker | None = None
        self._scan_pending_auto: bool = False
        self._excel_unlock_pulse_chain_scheduled: bool = False
        self._excel_create_probe_t0: float = 0.0
        # 制御用ボタン（シナリオ読込/保存、一括実行・キャンセル）
        row_btn = QHBoxLayout()
        btn_load = QPushButton(_u("BTN_SCENARIO_LOAD", "シナリオ読込"))
        btn_load.setAutoDefault(False)
        btn_load.setDefault(False)
        btn_load.clicked.connect(self._on_scenario_load)
        self._main_set_tip(
            btn_load,
            "TOOLTIP_BTN_SCENARIO_LOAD",
            "シナリオ JSON を読み込み、項目ごとの取得設定を復元します。",
        )
        self._btn_scenario_save = QPushButton(_u("BTN_SCENARIO_SAVE", "シナリオ保存"))
        self._btn_scenario_save.setAutoDefault(False)
        self._btn_scenario_save.setDefault(False)
        self._btn_scenario_save.clicked.connect(self._on_scenario_save)
        self._btn_scenario_save.setEnabled(False)
        self._main_set_tip(
            self._btn_scenario_save,
            "TOOLTIP_BTN_SCENARIO_SAVE",
            "現在の取得設定をシナリオファイルに保存します（変更があるとき有効）。",
        )
        row_btn.addWidget(btn_load)
        row_btn.addWidget(self._btn_scenario_save)
        self._btn_debug = QPushButton(_u("BTN_DEBUG", "デバッグ"))
        self._btn_debug.setAutoDefault(False)
        self._btn_debug.setDefault(False)
        tip_dbg = str(self._ui.get("TOOLTIP_DEBUG") or "").strip()
        if tip_dbg:
            self._btn_debug.setToolTip(tip_dbg)
        self._btn_debug.clicked.connect(self._on_debug)
        row_btn.addWidget(self._btn_debug)
        self._btn_scenario_export = QPushButton(_u("BTN_SCENARIO_EXPORT", "シナリオ出力"))
        self._btn_scenario_export.setAutoDefault(False)
        self._btn_scenario_export.setDefault(False)
        tip_exp = str(self._ui.get("TOOLTIP_SCENARIO_EXPORT") or "").strip()
        if tip_exp:
            self._btn_scenario_export.setToolTip(tip_exp)
        self._btn_scenario_export.clicked.connect(self._on_scenario_export)
        row_btn.addWidget(self._btn_scenario_export)
        row_btn.addStretch(1)
        self._btn_batch = QPushButton(_u("BTN_BATCH", "一括実行"))
        self._btn_batch.setAutoDefault(False)
        self._btn_batch.setDefault(False)
        self._btn_batch.clicked.connect(self._on_batch_run)
        btn_cancel = QPushButton(_u("BTN_CANCEL", "キャンセル"))
        btn_cancel.setAutoDefault(False)
        btn_cancel.setDefault(False)
        btn_cancel.clicked.connect(self.reject)
        self._main_set_tip(
            btn_cancel,
            "TOOLTIP_BTN_CANCEL_MAIN",
            "この画面を閉じて Excel に戻ります。",
        )
        row_btn.addWidget(self._btn_batch)
        row_btn.addWidget(btn_cancel)
        layout.addLayout(row_btn)
        layout.addStretch(1)
        try:
            from ui_qt.ui_common import apply_window_config

            apply_window_config(
                self,
                {"WINDOW": self._window_cfg},
                self._parent_hwnd,
                "DATA_AGG_MAIN",
            )
        except Exception:
            pass
        w = int(self._window_cfg.get("DEFAULT_WIDTH") or 720)
        h = int(self._window_cfg.get("DEFAULT_HEIGHT") or 480)
        if w > 0 and h > 0:
            self.resize(w, h)
        if bool(self._ui.get("PREFILL_START_PATH_FROM_LAST_FOLDER")):
            if not (self._edit_start_path.text() or "").strip():
                lf = get_last_folder()
                if lf:
                    self._edit_start_path.setText(lf)
        self._wire_auto_scan_signals()
        self._wire_scenario_dirty_signals()
        self._update_item_count_label()
        self._update_detected_file_count_label()
        # 基準フォルダが空のときは一覧も空。パスありは showEvent 後に非同期走査（UI 表示を先に返す）。
        self._scan_pending_auto = bool((self._edit_start_path.text() or "").strip())
        if not self._scan_pending_auto:
            self._file_list_items = []
            self._file_list.clear()
            self._update_detected_file_count_label()
        self._refresh_scenario_display_label()
        try:
            self.destroyed.connect(self._on_data_agg_main_destroyed)
        except Exception:
            pass

    def _on_data_agg_main_destroyed(self, *_args: Any) -> None:
        try:
            _log_data_agg_main_lifecycle(self, "destroyed")
        except Exception:
            pass

    def _main_ui_disp(self, key: str, default: str = "") -> str:
        """MAIN.UI の表示文字列（改行正規化つき）。"""
        return _ui_disp_str(self._ui or {}, key, default)

    def _main_set_tip(self, w: QWidget | None, key: str, default: str = "") -> None:
        """MAIN.UI の TOOLTIP_* または default をツールチップに設定。"""
        if w is None:
            return
        t = _ui_disp_str(self._ui or {}, key, default).strip()
        if t:
            w.setToolTip(t)

    def _scenario_has_any_registered_source(self) -> bool:
        """マスタ項目が1件以上あり、いずれかの行に有効な取得ソース（sources）が1件以上あるとき True。"""
        try:
            data = self._build_scenario_from_ui()
        except Exception:
            return False
        items = data.get("items") or []
        if not items:
            return False
        for it in items:
            if not isinstance(it, dict):
                continue
            srcs = it.get("sources") or []
            if any(isinstance(s, dict) and s for s in srcs):
                return True
        return False

    def _update_batch_button_enabled(self) -> None:
        """一括実行・デバッグ・シナリオ出力の有効化（出力はシナリオファイル読込も必須）。"""
        btn = getattr(self, "_btn_batch", None)
        dbg = getattr(self, "_btn_debug", None)
        exp = getattr(self, "_btn_scenario_export", None)
        if btn is None and dbg is None and exp is None:
            return
        ok_src = self._scenario_has_any_registered_source()
        ok_path = bool((self._scenario_path or "").strip())
        ok = ok_src
        ok_export = ok_src and ok_path
        tip_ok = str((self._ui or {}).get("TOOLTIP_BATCH") or "").strip()
        tip_dbg_ok = str((self._ui or {}).get("TOOLTIP_DEBUG") or "").strip()
        tip_exp_ok = str((self._ui or {}).get("TOOLTIP_SCENARIO_EXPORT") or "").strip()
        tip_need = str(
            (self._ui or {}).get("TOOLTIP_BATCH_REQUIRES_SCENARIO_SOURCES") or ""
        ).strip()
        if not tip_need:
            tip_need = str(
                (self._ui or {}).get("TOOLTIP_BATCH_REQUIRES_SCENARIO_LOAD") or ""
            ).strip()
        tip_exp_need = str(
            (self._ui or {}).get("TOOLTIP_SCENARIO_EXPORT_REQUIRES_LOAD") or ""
        ).strip()
        if btn is not None:
            btn.setEnabled(ok)
            btn.setToolTip(tip_ok if ok else (tip_need or tip_ok))
        if dbg is not None:
            dbg.setEnabled(ok)
            dbg.setToolTip(tip_dbg_ok if ok else (tip_need or tip_dbg_ok))
        if exp is not None:
            exp.setEnabled(ok_export)
            if ok_export:
                exp.setToolTip(tip_exp_ok or tip_ok)
            elif not ok_src:
                exp.setToolTip(tip_need or tip_exp_ok)
            else:
                exp.setToolTip(tip_exp_need or tip_exp_ok or tip_need)

    def _apply_excel_menu_bar_lock(self, lock: bool) -> bool:
        """メイン表示中は Excel のリボン／メニュー操作を抑止し、閉じたときに解除する。

        lock=True でロックに成功したとき True。コンテキスト未取得などで何もできなければ False。
        lock=False（解除）のあと True。
        """
        from core.core_xlc import (
            excel_try_set_main_commandbars_enabled,
            get_excel_context_from_hwnd,
        )

        try:
            if lock:
                ctx = get_excel_context_from_hwnd(
                    self._parent_hwnd, self._sheet_id
                )
                if not ctx:
                    logger.info(
                        "[DATA_AGG_MENU_LOCK] lock skipped: no ctx parent_hwnd=%s sheet_id=%s",
                        int(self._parent_hwnd or 0),
                        str(self._sheet_id or ""),
                    )
                    try:
                        _data_agg_ui_diag.info(
                            "[DATA_AGG_MENU_LOCK] lock skipped: no ctx parent_hwnd=%s sheet_id=%s",
                            int(self._parent_hwnd or 0),
                            str(self._sheet_id or ""),
                        )
                    except Exception:
                        pass
                    return False
                app, *_ = ctx
                self._excel_menu_lock_app = app
                excel_try_set_main_commandbars_enabled(app, False)
                use_interactive = bool(
                    (self._window_cfg or {}).get("EXCEL_LOCK_INTERACTIVE", True)
                )
                self._excel_lock_interactive_prev = None
                if use_interactive:
                    try:
                        api = getattr(app, "api", None)
                        if api is not None:
                            self._excel_lock_interactive_prev = bool(
                                getattr(api, "Interactive", True)
                            )
                            api.Interactive = False
                    except Exception:
                        self._excel_lock_interactive_prev = None
                logger.info(
                    "[DATA_AGG_MENU_LOCK] lock applied parent_hwnd=%s sheet_id=%s",
                    int(self._parent_hwnd or 0),
                    str(self._sheet_id or ""),
                )
                try:
                    _data_agg_ui_diag.info(
                        "[DATA_AGG_MENU_LOCK] lock applied parent_hwnd=%s sheet_id=%s",
                        int(self._parent_hwnd or 0),
                        str(self._sheet_id or ""),
                    )
                except Exception:
                    pass
                return True
            app = self._excel_menu_lock_app
            if app is not None:
                excel_try_set_main_commandbars_enabled(app, True)
                prev = getattr(self, "_excel_lock_interactive_prev", None)
                if prev is not None:
                    try:
                        ax = getattr(app, "api", None)
                        if ax is not None:
                            ax.Interactive = bool(prev)
                    except Exception:
                        pass
                self._excel_lock_interactive_prev = None
            else:
                logger.info("[DATA_AGG_MENU_LOCK] unlock: no stored app (nothing to re-enable)")
                try:
                    _data_agg_ui_diag.info(
                        "[DATA_AGG_MENU_LOCK] unlock: no stored app (nothing to re-enable)"
                    )
                except Exception:
                    pass
                self._excel_lock_interactive_prev = None
            self._excel_menu_lock_app = None
            logger.info("[DATA_AGG_MENU_LOCK] unlock done")
            try:
                _data_agg_ui_diag.info("[DATA_AGG_MENU_LOCK] unlock done")
            except Exception:
                pass
            return True
        except Exception as ex:
            logger.warning(
                "[DATA_AGG_MENU_LOCK] exception lock=%s parent_hwnd=%s ex=%r",
                lock,
                int(self._parent_hwnd or 0),
                ex,
            )
            try:
                _data_agg_ui_diag.info(
                    "[DATA_AGG_MENU_LOCK] exception lock=%s parent_hwnd=%s ex=%r",
                    lock,
                    int(self._parent_hwnd or 0),
                    ex,
                )
            except Exception:
                pass
            if lock:
                try:
                    ap = self._excel_menu_lock_app
                    prev = getattr(self, "_excel_lock_interactive_prev", None)
                    if ap is not None and prev is not None:
                        ax = getattr(ap, "api", None)
                        if ax is not None:
                            ax.Interactive = bool(prev)
                except Exception:
                    pass
            self._excel_menu_lock_app = None
            self._excel_lock_interactive_prev = None
            return not lock

    def _pulse_excel_unlock_if_excel_lock_off(
        self,
        *,
        _create_dialog_probe: bool = False,
        _probe_t0: float = 0.0,
        _probe_t_prev: float = 0.0,
    ) -> tuple[float, float] | None:
        """
        WINDOW.EXCEL_LOCK が false（子 HWND ロック不要）のとき、Excel 側の操作感を有効寄せにする。
        Win32 の子 HWND 再有効化に加え、COM の CommandBars 有効化と Application.Interactive を試す。
        ensure_front／前面追従の直後に子が再無効化される環境向けに showEvent から短い遅延で複数回呼ぶ。
        """
        ph = int(self._parent_hwnd or 0)
        probe_t0 = float(_probe_t0 or 0.0)
        probe_t_prev = float(_probe_t_prev or probe_t0 or 0.0)

        def _probe_sub(sub_phase: str) -> None:
            nonlocal probe_t_prev
            if not _create_dialog_probe:
                return
            probe_t_prev = _log_data_agg_create_dialog_phase(
                sub_phase,
                t0=probe_t0,
                t_prev=probe_t_prev,
                parent_hwnd=ph,
            )

        if not ph:
            return (probe_t0, probe_t_prev) if _create_dialog_probe else None
        try:
            from shiboken6 import Shiboken

            if not Shiboken.isValid(self):
                return (probe_t0, probe_t_prev) if _create_dialog_probe else None
        except Exception:
            pass
        try:
            from ui_qt.ui_common import want_excel_child_hwnd_lock_while_modal

            if want_excel_child_hwnd_lock_while_modal(self._window_cfg or {}):
                return (probe_t0, probe_t_prev) if _create_dialog_probe else None
        except Exception:
            return (probe_t0, probe_t_prev) if _create_dialog_probe else None
        try:
            from ui_qt.ui_common import enable_excel_window

            enable_excel_window(ph, True)
        except Exception:
            pass
        _probe_sub("pulse_after_enable_win32")
        try:
            from core.core_xlc import (
                excel_try_set_main_commandbars_enabled,
                get_excel_context_from_hwnd,
            )

            ctx = get_excel_context_from_hwnd(ph, self._sheet_id)
            _probe_sub("pulse_after_get_ctx")
            if ctx:
                app, *_rest = ctx
                excel_try_set_main_commandbars_enabled(app, True)
                _probe_sub("pulse_after_cmdbars")
                try:
                    ax = getattr(app, "api", None)
                    if ax is not None:
                        ax.Interactive = True
                except Exception:
                    pass
                _probe_sub("pulse_after_interactive")
        except Exception:
            pass
        try:
            self._apply_excel_menu_bar_lock(False)
            self._excel_menu_bar_lock_applied = False
        except Exception:
            pass
        _probe_sub("pulse_after_menu_unlock")
        return (probe_t0, probe_t_prev) if _create_dialog_probe else None

    def _schedule_excel_unlock_pulse_chain(self) -> None:
        """EXCEL_LOCK=false 時: Win32/COM 解禁を非同期 1 本化（show をブロックしない）。

        QTimer(0) で初回 pulse、90/200/450 ms で再試行（ensure_front 後の無効化取りこぼし緩和）。
        showEvent から複数回呼ばれても 1 チェーンだけ予約する。
        """
        if getattr(self, "_excel_unlock_pulse_chain_scheduled", False):
            return
        ph = int(self._parent_hwnd or 0)
        if not ph:
            return
        try:
            from shiboken6 import Shiboken

            if not Shiboken.isValid(self):
                return
        except Exception:
            pass
        try:
            from ui_qt.ui_common import want_excel_child_hwnd_lock_while_modal

            if want_excel_child_hwnd_lock_while_modal(self._window_cfg or {}):
                return
        except Exception:
            return

        self._excel_unlock_pulse_chain_scheduled = True

        def _run_pulse(*, use_create_probe: bool) -> None:
            try:
                from shiboken6 import Shiboken

                if not Shiboken.isValid(self) or not self.isVisible():
                    return
            except Exception:
                return
            probe_t0 = 0.0
            if use_create_probe:
                probe_t0 = float(getattr(self, "_excel_create_probe_t0", 0) or 0)
            t_prev = time.perf_counter()
            try:
                if use_create_probe and probe_t0 > 0:
                    t_prev = _log_data_agg_create_dialog_phase(
                        "pulse_deferred_enter",
                        t0=probe_t0,
                        t_prev=t_prev,
                        parent_hwnd=ph,
                    )
                probe_out = self._pulse_excel_unlock_if_excel_lock_off(
                    _create_dialog_probe=use_create_probe and probe_t0 > 0,
                    _probe_t0=probe_t0 if use_create_probe else 0.0,
                    _probe_t_prev=t_prev if use_create_probe else 0.0,
                )
                if use_create_probe and probe_t0 > 0:
                    if probe_out is not None:
                        _, t_prev = probe_out
                    _log_data_agg_create_dialog_phase(
                        "pulse_deferred_done",
                        t0=probe_t0,
                        t_prev=t_prev,
                        parent_hwnd=ph,
                    )
                    self._excel_create_probe_t0 = 0.0
            except Exception:
                pass

        def _first_pulse() -> None:
            _run_pulse(use_create_probe=True)

        try:
            QTimer.singleShot(0, _first_pulse)
            for _ms in (90, 200, 450):
                QTimer.singleShot(int(_ms), lambda _u=False: _run_pulse(use_create_probe=_u))
        except Exception:
            self._excel_unlock_pulse_chain_scheduled = False

    def _schedule_deferred_excel_owner_front(self) -> None:
        """DATA_AGG_MAIN は apply_window_config で遅延オーナーが付かないため、表示後に再適用する。

        位置は create_dialog の prepare_dialog（center_on_excel）に任せ、ここではオーナー確定と
        1 回の前面化のみに留めてちらつきを抑える。
        """
        if self._excel_deferred_owner_front_scheduled:
            return
        ph = int(self._parent_hwnd or 0)
        if not ph:
            return
        self._excel_deferred_owner_front_scheduled = True
        try:
            from ui_qt.ui_common import ensure_front, _set_owner_hwnd
        except Exception:
            return

        def _owner() -> None:
            try:
                _set_owner_hwnd(self, ph)
            except Exception:
                pass

        def _front() -> None:
            try:
                ensure_front(self, ph)
            except Exception:
                pass

        QTimer.singleShot(0, _owner)
        QTimer.singleShot(50, _owner)
        QTimer.singleShot(120, _front)

    def _create_excel_options_tab(
        self,
        _u: Any,
        *,
        ref_tab: QWidget,
    ) -> QWidget:
        """
        Excel 出力オプション（出力先・ジャンプ・並べ替え）。文言・ツールチップは ui_data_agg.json の MAIN.UI。
        値の永続化・書込み処理は後続タスク。
        ref_tab: シナリオ等の兄弟タブと同じページ背景色（通常は Base＝白系）に揃えるための参照。
        """
        tab = QWidget()
        outer = QVBoxLayout(tab)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        inner = QWidget()
        vl = QVBoxLayout(inner)
        vl.setContentsMargins(0, 0, 0, 0)

        def _tip_str(key: str, default: str = "") -> str:
            t = str((self._ui or {}).get(key) or "").strip()
            if not t and default:
                t = _normalize_message_newlines(str(default).strip())
            return t

        def _apply_tip(w: QWidget, key: str, default: str = "") -> None:
            t = _tip_str(key, default)
            if t:
                w.setToolTip(t)

        # --- 出力先 ---
        grp_out = QGroupBox(_u("GROUP_EXCEL_OUTPUT", "出力先"))
        _apply_tip(
            grp_out,
            "TOOLTIP_GROUP_EXCEL_OUTPUT",
            "マスターへの Excel 出力の単位と書き込み開始位置の指定です。",
        )
        lay_out = QVBoxLayout(grp_out)
        lay_out.setSpacing(4)
        lay_out.setContentsMargins(8, 8, 8, 8)
        self._excel_output_group = QButtonGroup(self)
        row_rad = QHBoxLayout()
        self._rad_excel_active = QRadioButton(_u("RADIO_EXCEL_ACTIVE_SHEET", "アクティブシート"))
        self._rad_excel_new = QRadioButton(_u("RADIO_EXCEL_NEW_SHEET", "新規シート"))
        self._rad_excel_active.setChecked(True)
        self._excel_output_group.addButton(self._rad_excel_active)
        self._excel_output_group.addButton(self._rad_excel_new)
        _apply_tip(
            self._rad_excel_active,
            "TOOLTIP_EXCEL_ACTIVE_SHEET",
            "現在アクティブなシートへ出力します。",
        )
        _apply_tip(
            self._rad_excel_new,
            "TOOLTIP_EXCEL_NEW_SHEET",
            "ブック内に新しいシートを作成して出力します。",
        )
        row_rad.addWidget(self._rad_excel_active)
        row_rad.addWidget(self._rad_excel_new)
        row_rad.addStretch(1)
        lay_out.addLayout(row_rad)

        row_wm = QHBoxLayout()
        self._lbl_excel_write_mode = QLabel(_u("LABEL_EXCEL_WRITE_MODE", "書込み方式") + ":")
        _apply_tip(
            self._lbl_excel_write_mode,
            "TOOLTIP_EXCEL_LABEL_WRITE_MODE",
            "アクティブシート出力時の書き込み方式のラベルです。",
        )
        self._combo_excel_write_mode = QComboBox()
        self._combo_excel_write_mode.addItem(
            _u("EXCEL_WRITE_MODE_APPEND", "追加"), "append"
        )
        self._combo_excel_write_mode.addItem(
            _u("EXCEL_WRITE_MODE_OVERWRITE", "上書き"), "overwrite"
        )
        self._combo_excel_write_mode.addItem(
            _u("EXCEL_WRITE_MODE_CLEAR_WRITE", "クリア書込み"), "clear_write"
        )
        self._combo_excel_write_mode.addItem(
            _u("EXCEL_WRITE_MODE_ANCHOR", "指定セル"), "anchor_cell"
        )
        _excel_compact_control(self._combo_excel_write_mode, _EXCEL_COMBO_WRITE_MODE_W)
        _apply_tip(
            self._combo_excel_write_mode,
            "TOOLTIP_EXCEL_WRITE_MODE",
            "アクティブシート選択時の書き込み方式です（追加・上書き・クリア書込み・指定セル）。",
        )
        row_wm.addWidget(self._lbl_excel_write_mode)
        row_wm.addWidget(self._combo_excel_write_mode, 0)
        row_wm.addStretch(1)
        lay_out.addLayout(row_wm)

        row_ac = QHBoxLayout()
        self._lbl_excel_anchor = QLabel(_u("LABEL_EXCEL_ANCHOR_CELL", "指定セル") + ":")
        _apply_tip(
            self._lbl_excel_anchor,
            "TOOLTIP_EXCEL_LABEL_ANCHOR",
            "書き込み開始セルを指定するモード用のラベルです。",
        )
        self._edit_excel_anchor_cell = QLineEdit()
        self._edit_excel_anchor_cell.setPlaceholderText("A1")
        _excel_compact_control(self._edit_excel_anchor_cell, _EXCEL_EDIT_ANCHOR_W)
        _apply_tip(
            self._edit_excel_anchor_cell,
            "TOOLTIP_EXCEL_ANCHOR_CELL",
            "結果を書き込む範囲の左上を A1 形式で指定します（例: C5）。",
        )
        row_ac.addWidget(self._lbl_excel_anchor)
        row_ac.addWidget(self._edit_excel_anchor_cell, 0)
        row_ac.addStretch(1)
        lay_out.addLayout(row_ac)

        row_sn = QHBoxLayout()
        self._lbl_excel_sheet_rule = QLabel(_u("LABEL_EXCEL_SHEET_NAME_RULE", "シート名") + ":")
        _apply_tip(
            self._lbl_excel_sheet_rule,
            "TOOLTIP_EXCEL_LABEL_SHEET_RULE",
            "新規シート作成時のシート名ルールのラベルです。",
        )
        self._combo_excel_sheet_name_rule = QComboBox()
        self._combo_excel_sheet_name_rule.addItem(
            _u("EXCEL_SHEET_NAME_SCENARIO_NAME_SEQ", "シナリオ名_連番"),
            "scenario_name_seq",
        )
        self._combo_excel_sheet_name_rule.addItem(
            _u("EXCEL_SHEET_NAME_CUSTOM_INPUT", "シート名入力"),
            "custom_sheet_name",
        )
        _excel_compact_control(self._combo_excel_sheet_name_rule, _EXCEL_COMBO_SHEET_RULE_W)
        _apply_tip(
            self._combo_excel_sheet_name_rule,
            "TOOLTIP_EXCEL_SHEET_NAME_RULE",
            "新規シートの名前の付け方です。",
        )
        self._edit_excel_custom_sheet_name = QLineEdit()
        self._edit_excel_custom_sheet_name.setPlaceholderText(
            _u("PLACEHOLDER_EXCEL_CUSTOM_SHEET_NAME", "シート名を入力")
        )
        _excel_compact_control(self._edit_excel_custom_sheet_name, 220)
        _apply_tip(
            self._edit_excel_custom_sheet_name,
            "TOOLTIP_EXCEL_CUSTOM_SHEET_NAME",
            "「シート名入力」を選んだときのシート名です。",
        )
        self._edit_excel_custom_sheet_name.setVisible(False)
        row_sn.addWidget(self._lbl_excel_sheet_rule)
        row_sn.addWidget(self._combo_excel_sheet_name_rule, 0)
        row_sn.addWidget(self._edit_excel_custom_sheet_name, 0)
        row_sn.addStretch(1)
        lay_out.addLayout(row_sn)

        vl.addWidget(grp_out)

        # --- ジャンプ ---
        grp_jump = QGroupBox(_u("GROUP_EXCEL_JUMP", "ジャンプ"))
        _apply_tip(
            grp_jump,
            "TOOLTIP_GROUP_EXCEL_JUMP",
            "実行後に名前ボックスからジャンプできるよう定義名を登録するかどうかです。",
        )
        lay_jump = QVBoxLayout(grp_jump)
        lay_jump.setSpacing(4)
        lay_jump.setContentsMargins(8, 8, 8, 8)
        self._chk_excel_jump = QCheckBox(_u("CHK_EXCEL_JUMP", "ジャンプ用の名前を登録する"))
        _apply_tip(
            self._chk_excel_jump,
            "TOOLTIP_EXCEL_JUMP",
            "オンにすると出力範囲へ名前でジャンプできるよう登録します。",
        )
        self._chk_excel_jump.setChecked(True)
        lay_jump.addWidget(self._chk_excel_jump)
        vl.addWidget(grp_jump)

        # --- 並べ替え ---
        grp_sort = QGroupBox(_u("GROUP_EXCEL_SORT", "並べ替え（書込み前）"))
        _apply_tip(
            grp_sort,
            "TOOLTIP_GROUP_EXCEL_SORT",
            "マスターへ書き込む直前の並べ替え条件です。",
        )
        lay_sort = QVBoxLayout(grp_sort)
        lay_sort.setSpacing(4)
        lay_sort.setContentsMargins(8, 8, 8, 8)
        self._excel_sort_rows_host = QWidget()
        self._excel_sort_layout = QVBoxLayout(self._excel_sort_rows_host)
        self._excel_sort_layout.setContentsMargins(0, 0, 0, 0)
        lay_sort.addWidget(self._excel_sort_rows_host)
        row_sort_btn = QHBoxLayout()
        btn_add = QPushButton(_u("BTN_EXCEL_SORT_ADD", "条件を追加"))
        btn_rem = QPushButton(_u("BTN_EXCEL_SORT_REMOVE", "最後の条件を削除"))
        btn_add.setMaximumHeight(_EXCEL_SORT_BTN_MAX_H)
        btn_rem.setMaximumHeight(_EXCEL_SORT_BTN_MAX_H)
        _apply_tip(
            btn_add,
            "TOOLTIP_EXCEL_SORT_ADD",
            "並べ替えキーを 1 行追加します。",
        )
        _apply_tip(
            btn_rem,
            "TOOLTIP_EXCEL_SORT_REMOVE",
            "一覧の最後の並べ替え行を削除します。",
        )
        btn_add.clicked.connect(self._on_excel_sort_add_row)
        btn_rem.clicked.connect(self._on_excel_sort_remove_last)
        row_sort_btn.addWidget(btn_add)
        row_sort_btn.addWidget(btn_rem)
        row_sort_btn.addStretch(1)
        lay_sort.addLayout(row_sort_btn)
        vl.addWidget(grp_sort)

        vl.addStretch(1)
        scroll.setWidget(inner)
        self._excel_tab_scroll = scroll
        _apply_tip(
            scroll,
            "TOOLTIP_EXCEL_TAB_SCROLL",
            "Excel 出力オプション全体のスクロール領域です。",
        )
        _apply_tip(
            inner,
            "TOOLTIP_EXCEL_TAB_INNER",
            "出力先・ジャンプ・並べ替えフォームのコンテナです。",
        )
        outer.addWidget(scroll)

        self._excel_sort_row_widgets: list[QWidget] = []
        self._add_excel_sort_row()
        self._rad_excel_active.toggled.connect(self._update_excel_output_visibility)
        self._rad_excel_new.toggled.connect(self._update_excel_output_visibility)
        self._rad_excel_active.toggled.connect(self._mark_scenario_dirty)
        self._rad_excel_new.toggled.connect(self._mark_scenario_dirty)
        self._combo_excel_write_mode.currentIndexChanged.connect(
            self._update_excel_output_visibility
        )
        self._combo_excel_write_mode.currentIndexChanged.connect(self._mark_scenario_dirty)
        self._edit_excel_anchor_cell.textChanged.connect(self._mark_scenario_dirty)
        self._combo_excel_sheet_name_rule.currentIndexChanged.connect(
            self._update_excel_output_visibility
        )
        self._combo_excel_sheet_name_rule.currentIndexChanged.connect(self._mark_scenario_dirty)
        self._edit_excel_custom_sheet_name.textChanged.connect(self._mark_scenario_dirty)
        self._chk_excel_jump.stateChanged.connect(self._mark_scenario_dirty)
        self._update_excel_output_visibility()
        self._refresh_excel_sort_item_combos()
        self._apply_excel_tab_surface_colors(
            scroll, inner, (grp_out, grp_jump, grp_sort), ref_tab=ref_tab
        )
        self._main_set_tip(
            tab,
            "TOOLTIP_TAB_EXCEL",
            "マスタへの書き込み先・並べ替えなどのオプションです。",
        )
        return tab

    def _apply_excel_tab_surface_colors(
        self,
        scroll: QScrollArea,
        inner: QWidget,
        group_boxes: tuple[QGroupBox, ...],
        *,
        ref_tab: QWidget,
    ) -> None:
        """Excel タブ内を兄弟タブ（ref_tab）と同じページ背景（Base）に揃える。外枠は塗らず QTabWidget に任せる。"""
        surface = ref_tab.palette().color(QPalette.ColorRole.Base)
        mid_col = ref_tab.palette().color(QPalette.ColorRole.Mid)
        bg = surface.name()
        midn = mid_col.name()

        inner.setAutoFillBackground(True)
        pal_inner = QPalette(inner.palette())
        pal_inner.setColor(QPalette.ColorRole.Window, surface)
        pal_inner.setColor(QPalette.ColorRole.Base, surface)
        inner.setPalette(pal_inner)

        scroll.setStyleSheet("QScrollArea { background-color: %s; border: none; }" % bg)
        vp = scroll.viewport()
        vp.setAutoFillBackground(True)
        pal_vp = QPalette(vp.palette())
        pal_vp.setColor(QPalette.ColorRole.Window, surface)
        pal_vp.setColor(QPalette.ColorRole.Base, surface)
        vp.setPalette(pal_vp)

        gb_css = (
            "QGroupBox { background-color: %s; border: 1px solid %s; border-radius: 4px; "
            "margin-top: 12px; padding-top: 8px; font-weight: normal; }"
            "QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 4px; }"
        ) % (bg, midn)
        for gb in group_boxes:
            gb.setStyleSheet(gb_css)

        host = getattr(self, "_excel_sort_rows_host", None)
        if host is not None:
            host.setAutoFillBackground(True)
            host.setPalette(pal_inner)
            for row in self._excel_sort_row_widgets:
                row.setAutoFillBackground(True)
                row.setPalette(pal_inner)

    def _update_excel_output_visibility(self) -> None:
        active = self._rad_excel_active.isChecked()
        self._lbl_excel_write_mode.setVisible(active)
        self._combo_excel_write_mode.setVisible(active)
        anchor_mode = (
            active
            and self._combo_excel_write_mode.currentData() == "anchor_cell"
        )
        self._lbl_excel_anchor.setVisible(anchor_mode)
        self._edit_excel_anchor_cell.setVisible(anchor_mode)
        self._lbl_excel_sheet_rule.setVisible(not active)
        self._combo_excel_sheet_name_rule.setVisible(not active)
        custom_rule = (
            self._combo_excel_sheet_name_rule.currentData() == "custom_sheet_name"
        )
        self._edit_excel_custom_sheet_name.setVisible(not active and custom_rule)

    def _excel_sort_row_item_names(self) -> list[str]:
        names: list[str] = []
        for r in range(self._item_table.rowCount()):
            c0 = self._item_table.item(r, 0)
            if c0:
                t = c0.text().strip()
                if t:
                    names.append(t)
        return names

    def _clear_excel_sort_rows(self) -> None:
        """並べ替え行をすべて削除する（シナリオ読込時の差し替え用）。"""
        while self._excel_sort_row_widgets:
            row = self._excel_sort_row_widgets.pop()
            self._excel_sort_layout.removeWidget(row)
            row.deleteLater()

    def _add_excel_sort_row(self, *, skip_refresh: bool = False) -> None:
        _u = lambda k, d: _ui_disp_str(self._ui, k, d)

        def _tip_key(w: QWidget, key: str, default: str = "") -> None:
            t = str((self._ui or {}).get(key) or "").strip()
            if not t and default:
                t = _normalize_message_newlines(str(default).strip())
            if t:
                w.setToolTip(t)

        row = QWidget()
        hl = QHBoxLayout(row)
        hl.setContentsMargins(0, 2, 0, 2)
        lbl = QLabel(_u("LABEL_EXCEL_SORT_ITEM", "項目") + ":")
        cb_item = QComboBox()
        cb_order = QComboBox()
        cb_order.addItem(_u("EXCEL_SORT_ORDER_ASC", "昇順"), "asc")
        cb_order.addItem(_u("EXCEL_SORT_ORDER_DESC", "降順"), "desc")
        chk_nat = QCheckBox(_u("CHK_EXCEL_SORT_NATURAL", "自然順"))
        chk_nat.setChecked(True)
        _excel_compact_control(cb_item, _EXCEL_COMBO_SORT_ITEM_W)
        _excel_compact_control(cb_order, _EXCEL_COMBO_SORT_ORDER_W)
        _tip_key(
            lbl,
            "TOOLTIP_EXCEL_SORT_ROW_LABEL",
            "この並べ替え行でキーとする項目のラベルです。",
        )
        _tip_key(
            cb_item,
            "TOOLTIP_EXCEL_SORT_ITEM",
            "並べ替えのキーにする項目名（マスタ列）です。",
        )
        _tip_key(
            cb_order,
            "TOOLTIP_EXCEL_SORT_ORDER",
            "当該キーの昇順／降順です。",
        )
        _tip_key(
            chk_nat,
            "TOOLTIP_EXCEL_SORT_NATURAL",
            "文字列内の数字を数値として比較する自然順を使う場合にオンにします。",
        )
        hl.addWidget(lbl)
        hl.addWidget(cb_item, 0)
        hl.addWidget(cb_order, 0)
        hl.addWidget(chk_nat)
        hl.addStretch(1)
        self._excel_sort_layout.addWidget(row)
        self._excel_sort_row_widgets.append(row)
        host = getattr(self, "_excel_sort_rows_host", None)
        if host is not None:
            row.setAutoFillBackground(True)
            row.setPalette(host.palette())
        setattr(row, "_excel_sort_cb_item", cb_item)
        setattr(row, "_excel_sort_cb_order", cb_order)
        setattr(row, "_excel_sort_chk_natural", chk_nat)
        cb_item.currentIndexChanged.connect(self._mark_scenario_dirty)
        cb_order.currentIndexChanged.connect(self._mark_scenario_dirty)
        chk_nat.stateChanged.connect(self._mark_scenario_dirty)
        if not skip_refresh:
            self._refresh_excel_sort_item_combos()

    def _on_excel_sort_add_row(self) -> None:
        self._add_excel_sort_row()
        self._mark_scenario_dirty()

    def _on_excel_sort_remove_last(self) -> None:
        if len(self._excel_sort_row_widgets) <= 1:
            return
        row = self._excel_sort_row_widgets.pop()
        self._excel_sort_layout.removeWidget(row)
        row.deleteLater()
        self._mark_scenario_dirty()

    def _excel_options_from_ui(self) -> dict[str, Any]:
        """Excel タブの状態をシナリオ用 dict にする。"""
        sort_keys: list[dict[str, Any]] = []
        for row in getattr(self, "_excel_sort_row_widgets", []) or []:
            cb_item = getattr(row, "_excel_sort_cb_item", None)
            cb_order = getattr(row, "_excel_sort_cb_order", None)
            chk_nat = getattr(row, "_excel_sort_chk_natural", None)
            if not isinstance(cb_item, QComboBox):
                continue
            data = cb_item.currentData()
            item = str(data).strip() if data is not None else ""
            if not item:
                item = (cb_item.currentText() or "").strip()
            order = "asc"
            if isinstance(cb_order, QComboBox):
                od = cb_order.currentData()
                if od in ("asc", "desc"):
                    order = str(od)
            nat = bool(chk_nat.isChecked()) if isinstance(chk_nat, QCheckBox) else False
            sort_keys.append({"item": item, "order": order, "natural": nat})
        if not sort_keys:
            sort_keys = [{"item": "", "order": "asc", "natural": True}]
        out_target = "new_sheet" if self._rad_excel_new.isChecked() else "active_sheet"
        wi = self._combo_excel_write_mode.currentIndex()
        wm = self._combo_excel_write_mode.itemData(wi)
        write_mode = (
            str(wm)
            if wm in ("append", "overwrite", "clear_write", "anchor_cell")
            else "append"
        )
        ns_i = self._combo_excel_sheet_name_rule.currentIndex()
        nsd = self._combo_excel_sheet_name_rule.itemData(ns_i)
        ns_rule = (
            str(nsd)
            if nsd
            in (
                "scenario_name_seq",
                "scenario_datetime",
                "scenario_seq",
                "custom_sheet_name",
            )
            else "scenario_name_seq"
        )
        return {
            "output_target": out_target,
            "write_mode": write_mode,
            "anchor_cell": (self._edit_excel_anchor_cell.text() or "").strip(),
            "new_sheet_name_rule": ns_rule,
            "new_sheet_custom_name": (
                (self._edit_excel_custom_sheet_name.text() or "").strip()
            ),
            "jump_register_name": self._chk_excel_jump.isChecked(),
            "sort_keys": sort_keys,
        }

    def _apply_excel_options_to_ui(self, raw: Any) -> None:
        """シナリオの excel_options を Excel タブに反映する。"""
        from svc import svc_data_agg_scenario as scenario_mod

        opt = scenario_mod.normalize_excel_options(raw)
        self._rad_excel_active.blockSignals(True)
        self._rad_excel_new.blockSignals(True)
        self._combo_excel_write_mode.blockSignals(True)
        self._combo_excel_sheet_name_rule.blockSignals(True)
        self._edit_excel_anchor_cell.blockSignals(True)
        self._edit_excel_custom_sheet_name.blockSignals(True)
        self._chk_excel_jump.blockSignals(True)
        try:
            if opt["output_target"] == "new_sheet":
                self._rad_excel_new.setChecked(True)
            else:
                self._rad_excel_active.setChecked(True)
            wi = self._combo_excel_write_mode.findData(opt["write_mode"])
            if wi >= 0:
                self._combo_excel_write_mode.setCurrentIndex(wi)
            self._edit_excel_anchor_cell.setText(str(opt.get("anchor_cell") or ""))
            ni = self._combo_excel_sheet_name_rule.findData(opt["new_sheet_name_rule"])
            if ni < 0:
                ni = self._combo_excel_sheet_name_rule.findData("scenario_name_seq")
            if ni >= 0:
                self._combo_excel_sheet_name_rule.setCurrentIndex(ni)
            self._edit_excel_custom_sheet_name.setText(
                str(opt.get("new_sheet_custom_name") or "")
            )
            self._chk_excel_jump.setChecked(bool(opt.get("jump_register_name")))
            self._clear_excel_sort_rows()
            sk = opt.get("sort_keys") or []
            if not sk:
                sk = [{"item": "", "order": "asc", "natural": True}]
            for _ in sk:
                self._add_excel_sort_row(skip_refresh=True)
            self._refresh_excel_sort_item_combos()
            for row, spec in zip(self._excel_sort_row_widgets, sk):
                cb_item = getattr(row, "_excel_sort_cb_item", None)
                cb_order = getattr(row, "_excel_sort_cb_order", None)
                chk_nat = getattr(row, "_excel_sort_chk_natural", None)
                if isinstance(cb_item, QComboBox):
                    cb_item.blockSignals(True)
                    iname = str(spec.get("item") or "").strip()
                    ix = cb_item.findData(iname)
                    if ix < 0 and iname:
                        ix = cb_item.findText(iname)
                    if ix >= 0:
                        cb_item.setCurrentIndex(ix)
                    cb_item.blockSignals(False)
                if isinstance(cb_order, QComboBox):
                    cb_order.blockSignals(True)
                    o = str(spec.get("order") or "asc").strip().lower()
                    if o not in ("asc", "desc"):
                        o = "asc"
                    oi = cb_order.findData(o)
                    if oi >= 0:
                        cb_order.setCurrentIndex(oi)
                    cb_order.blockSignals(False)
                if isinstance(chk_nat, QCheckBox):
                    chk_nat.blockSignals(True)
                    chk_nat.setChecked(bool(spec.get("natural")))
                    chk_nat.blockSignals(False)
            self._update_excel_output_visibility()
        finally:
            self._rad_excel_active.blockSignals(False)
            self._rad_excel_new.blockSignals(False)
            self._combo_excel_write_mode.blockSignals(False)
            self._combo_excel_sheet_name_rule.blockSignals(False)
            self._edit_excel_anchor_cell.blockSignals(False)
            self._edit_excel_custom_sheet_name.blockSignals(False)
            self._chk_excel_jump.blockSignals(False)

    def _refresh_excel_sort_item_combos(self) -> None:
        rows = getattr(self, "_excel_sort_row_widgets", None)
        if not rows:
            return
        names = self._excel_sort_row_item_names()
        for row in rows:
            cb = getattr(row, "_excel_sort_cb_item", None)
            if not isinstance(cb, QComboBox):
                continue
            prev = (cb.currentData() if cb.currentData() is not None else cb.currentText())
            prev_s = str(prev).strip() if prev is not None else ""
            cb.blockSignals(True)
            cb.clear()
            cb.addItem("", "")
            for n in names:
                cb.addItem(n, n)
            cb.blockSignals(False)
            if prev_s:
                idx = cb.findData(prev_s)
                if idx < 0:
                    idx = cb.findText(prev_s)
                if idx >= 0:
                    cb.setCurrentIndex(idx)

    def _wire_auto_scan_signals(self) -> None:
        self._chk_recursive.stateChanged.connect(
            lambda _v: self._on_scan_and_mark_dirty(True)
        )
        self._chk_ext_xls.stateChanged.connect(
            lambda _v: self._on_scan_and_mark_dirty(True)
        )
        self._chk_ext_xlsx.stateChanged.connect(
            lambda _v: self._on_scan_and_mark_dirty(True)
        )
        self._chk_ext_xlsm.stateChanged.connect(
            lambda _v: self._on_scan_and_mark_dirty(True)
        )
        self._chk_ext_csv.stateChanged.connect(
            lambda _v: self._on_scan_and_mark_dirty(True)
        )
        self._edit_start_path.returnPressed.connect(
            lambda: self._on_scan_and_mark_dirty(True)
        )

    def _on_scan_and_mark_dirty(self, auto_mode: bool) -> None:
        if self._suppress_scenario_dirty:
            # シナリオ読込など UI 値の一括反映中は、途中状態での再走査を抑止する。
            return
        self._mark_scenario_dirty()
        self._on_scan(auto_mode=auto_mode)

    def _wire_scenario_dirty_signals(self) -> None:
        self._item_table.itemChanged.connect(self._on_item_table_item_changed)
        self._edit_start_path.textChanged.connect(self._mark_scenario_dirty)
        self._edit_keyword.textChanged.connect(self._mark_scenario_dirty)

    def _mark_scenario_dirty(self) -> None:
        if self._suppress_scenario_dirty:
            return
        self._scenario_dirty = True
        self._btn_scenario_save.setEnabled(True)
        self._update_batch_button_enabled()

    def _clear_scenario_dirty(self) -> None:
        self._scenario_dirty = False
        self._btn_scenario_save.setEnabled(False)

    def _refresh_scenario_display_label(self) -> None:
        """読込んだシナリオファイル名（拡張子除く）をタブ上に表示する。"""
        prefix = str((self._ui or {}).get("LABEL_SCENARIO_DISPLAY_NAME") or "シナリオ名：")
        p = (self._scenario_path or "").strip()
        if p:
            self._lbl_scenario_display_name.setText(prefix + Path(p).stem)
        else:
            empty = str((self._ui or {}).get("SCENARIO_NAME_NOT_LOADED") or "（未読込）")
            self._lbl_scenario_display_name.setText(prefix + empty)
        self._update_batch_button_enabled()

    def _on_scenario_clear_all(self) -> None:
        """マスタ項目一覧とシナリオ内容を空にし、未読込相当へ。基準フォルダ・検出一覧もクリア。シナリオ保存は無効のまま。"""
        from svc import svc_data_agg_scenario as scenario_mod

        self._item_table.blockSignals(True)
        try:
            self._item_table.setRowCount(0)
        finally:
            self._item_table.blockSignals(False)
        self._scenario = scenario_mod.create_empty_scenario()
        self._scenario_path = ""
        self._edit_start_path.clear()
        self._file_list_items = []
        self._file_list.clear()
        self._update_detected_file_count_label()
        self._refresh_scenario_display_label()
        self._update_item_count_label()
        self._refresh_excel_sort_item_combos()
        self._fit_item_table_columns()
        self._clear_scenario_dirty()
        self._scenario_save_empty_filename = False

    def _on_scenario_clear_sources_only(self) -> None:
        """項目行は残し、各マスタの取得シナリオ（sources）のみ空にする。"""
        from svc import svc_data_agg_scenario as scenario_mod

        if self._item_table.rowCount() <= 0:
            return
        if not self._scenario:
            self._scenario = scenario_mod.create_empty_scenario()
        items = self._scenario.setdefault("items", [])
        for r in range(self._item_table.rowCount()):
            c0 = self._item_table.item(r, 0)
            disp = (c0.text() if c0 else "").strip()
            nm = disp or ("項目_%s" % (r + 1))
            while len(items) <= r:
                items.append(
                    {
                        "id": "item_%s" % len(items),
                        "name": "項目_%s" % (len(items) + 1),
                        "sources": [],
                        "write_mode": "fill_in",
                    }
                )
            items[r]["name"] = nm
            items[r]["sources"] = []
        if len(items) > self._item_table.rowCount():
            del items[self._item_table.rowCount() :]
        self._refresh_item_summaries_and_link_state()
        self._mark_scenario_dirty()
        self._scenario_save_empty_filename = True

    def _update_item_count_label(self) -> None:
        n = self._item_table.rowCount()
        fmt = str(self._ui.get("LABEL_ITEM_TOTAL_FMT") or "項目総数：%d件").strip()
        self._lbl_item_total.setText(fmt % n)

    def _update_detected_file_count_label(self) -> None:
        if getattr(self, "_scan_busy", False):
            return
        n = len(self._file_list_items)
        fmt = str(self._ui.get("LABEL_DETECTED_FILE_COUNT_FMT") or "ファイル数：%d").strip()
        try:
            self._lbl_detected_file_count.setText(fmt % n)
        except Exception:
            self._lbl_detected_file_count.setText("ファイル数：%d" % n)

    def _set_scan_ui_busy(self, busy: bool) -> None:
        self._scan_busy = bool(busy)
        btn = getattr(self, "_btn_scan_run", None)
        if btn is not None:
            btn.setEnabled(not busy)
        lbl = getattr(self, "_lbl_detected_file_count", None)
        if lbl is None:
            return
        if busy:
            lbl.setText(
                _ui_disp_str(
                    self._ui or {},
                    "LABEL_DETECTED_FILE_COUNT_SCANNING",
                    "走査中…",
                )
            )
        else:
            self._update_detected_file_count_label()

    def _stop_scan_thread(self) -> None:
        th = getattr(self, "_scan_thread", None)
        if th is not None and th.isRunning():
            try:
                th.quit()
                th.wait(5000)
            except Exception:
                pass
        self._scan_thread = None
        self._scan_worker = None

    def _apply_folder_scan_result(
        self,
        generation: int,
        paths: list[str],
        *,
        auto_mode: bool,
        on_complete: Callable[[int], None] | None = None,
    ) -> None:
        if not should_apply_folder_scan_result(generation, self._scan_generation):
            return
        self._set_scan_ui_busy(False)
        self._file_list_items = list(paths)
        self._file_list.clear()
        for fp in self._file_list_items:
            self._file_list.addItem(fp)
        self._update_detected_file_count_label()
        try:
            sp = str(self._get_scan_state().get("start_path") or "").strip() or "."
            rp = Path(sp).resolve()
            if rp.is_dir():
                set_last_folder(str(rp))
        except Exception:
            pass
        if not auto_mode:
            t_scan = _ui_disp_str(self._ui or {}, "BTN_SEARCH_RUN", "検索実行")
            msg = _ui_disp_str(
                self._ui or {},
                "MSG_SCAN_DONE",
                "%s 件のファイルを検出しました。",
            )
            show_info_notice(self, t_scan, msg % len(paths))
        if on_complete is not None:
            on_complete(len(paths))
        try:
            logger.info(
                "[DATA_AGG_UI] folder_scan_async done generation=%s n=%s auto=%s",
                generation,
                len(paths),
                auto_mode,
            )
        except Exception:
            pass

    @Slot(int, object)
    def _on_folder_scan_worker_finished(self, generation: int, paths_obj: object) -> None:
        paths = list(paths_obj) if isinstance(paths_obj, list) else []
        auto_mode = bool(self._scan_req_auto_mode)
        self._apply_folder_scan_result(
            generation,
            paths,
            auto_mode=auto_mode,
            on_complete=self._scan_req_on_complete,
        )
        self._scan_thread = None
        self._scan_worker = None

    @Slot(int, str)
    def _on_folder_scan_worker_failed(self, generation: int, message: str) -> None:
        if not should_apply_folder_scan_result(generation, self._scan_generation):
            return
        self._set_scan_ui_busy(False)
        auto_mode = bool(self._scan_req_auto_mode)
        cb = self._scan_req_on_complete
        if not auto_mode:
            t_scan = _ui_disp_str(self._ui or {}, "BTN_SEARCH_RUN", "検索実行")
            show_warning_notice(
                self,
                t_scan,
                _ui_disp_str(
                    self._ui or {},
                    "MSG_SCAN_FAILED_FMT",
                    "検索に失敗しました: %s",
                )
                % message,
            )
        if callable(cb):
            cb(0)
        self._scan_thread = None
        self._scan_worker = None
        try:
            logger.warning(
                "[DATA_AGG_UI] folder_scan_async failed generation=%s err=%s",
                generation,
                message,
            )
        except Exception:
            pass

    def _request_folder_scan(
        self,
        *,
        auto_mode: bool = False,
        on_complete: Callable[[int], None] | None = None,
    ) -> None:
        """フォルダ走査をバックグラウンドで実行し、完了後に一覧を更新する。"""
        state = self._get_scan_state()
        sp = str(state.get("start_path") or "").strip()
        if not sp:
            self._file_list_items = []
            self._file_list.clear()
            self._update_detected_file_count_label()
            if on_complete is not None:
                on_complete(0)
            return
        exts = list(state.get("extensions") or [])
        if not exts:
            self._file_list_items = []
            self._file_list.clear()
            self._update_detected_file_count_label()
            if not auto_mode:
                t_scan = _ui_disp_str(self._ui or {}, "BTN_SEARCH_RUN", "検索実行")
                show_warning_notice(
                    self,
                    t_scan,
                    _ui_disp_str(
                        self._ui or {},
                        "MSG_SCAN_SELECT_EXT",
                        "拡張子を1つ以上選択してください。（.xls / .xlsx / .xlsm / .csv）",
                    ),
                )
            if on_complete is not None:
                on_complete(0)
            return

        self._scan_generation += 1
        gen = self._scan_generation
        self._scan_req_auto_mode = auto_mode
        self._scan_req_on_complete = on_complete
        self._set_scan_ui_busy(True)
        self._stop_scan_thread()

        thread = QThread(self)
        worker = _FolderScanWorker(gen, state)
        worker.moveToThread(thread)
        thread.started.connect(worker.run)
        worker.finished.connect(self._on_folder_scan_worker_finished)
        worker.failed.connect(self._on_folder_scan_worker_failed)
        worker.finished.connect(thread.quit)
        worker.failed.connect(thread.quit)
        thread.finished.connect(worker.deleteLater)
        thread.finished.connect(thread.deleteLater)
        self._scan_thread = thread
        self._scan_worker = worker
        thread.start()
        try:
            logger.info(
                "[DATA_AGG_UI] folder_scan_async start generation=%s auto=%s path=%s",
                gen,
                auto_mode,
                sp,
            )
        except Exception:
            pass

    def _sync_item_table_master_name_roles(self) -> None:
        """列0の UserRole を表示テキストと一致させる（移動・読込後に呼ぶ）。"""
        for r in range(self._item_table.rowCount()):
            c0 = self._item_table.item(r, 0)
            if c0 is None:
                continue
            nm = (c0.text() or "").strip()
            c0.setData(_ITEM_MASTER_NAME_ROLE, nm or None)

    def _on_item_table_item_changed(self, table_item: QTableWidgetItem) -> None:
        """項目名（列0）変更時は連携/結合/path_item 参照を一括置換しシナリオを整合。"""
        if self._item_table.signalsBlocked():
            return
        col = table_item.column()
        if col != 0:
            self._mark_scenario_dirty()
            return
        row = table_item.row()
        new_name = (table_item.text() or "").strip()
        raw = table_item.data(_ITEM_MASTER_NAME_ROLE)
        old_name = (raw.strip() if isinstance(raw, str) else "") or ""
        if not old_name and self._scenario.get("items") and row < len(self._scenario["items"]):
            old_name = str(self._scenario["items"][row].get("name") or "").strip()
        if old_name and new_name and old_name != new_name:
            self._propagate_master_item_rename_in_scenario(old_name, new_name)
            self._ensure_scenario_item_row_name(row, new_name)
            self._refresh_item_summaries_and_link_state()
        else:
            self._ensure_scenario_item_row_name(row, new_name or old_name)
            if old_name != new_name:
                self._refresh_item_summaries_and_link_state()
        self._item_table.blockSignals(True)
        try:
            table_item.setData(_ITEM_MASTER_NAME_ROLE, new_name or None)
        finally:
            self._item_table.blockSignals(False)
        self._mark_scenario_dirty()
        self._refresh_excel_sort_item_combos()

    def _ensure_scenario_item_row_name(self, row: int, name: str) -> None:
        """self._scenario.items[row].name を保持（行数不足時は拡張）。"""
        from svc import svc_data_agg_scenario as scenario_mod

        if not self._scenario:
            self._scenario = scenario_mod.create_empty_scenario()
        items = self._scenario.setdefault("items", [])
        while len(items) <= row:
            items.append(
                {
                    "id": "item_%s" % len(items),
                    "name": "項目_%s" % (len(items) + 1),
                    "sources": [],
                    "write_mode": "fill_in",
                }
            )
        nm = (name or "").strip() or str(items[row].get("name") or ("項目_%s" % (row + 1))).strip()
        items[row]["name"] = nm

    def _propagate_master_item_rename_in_scenario(self, old: str, new: str) -> None:
        """全項目の全ソースで link/join の item 参照と path_item を旧名→新名で置換する。"""
        old = (old or "").strip()
        new = (new or "").strip()
        if not old or not new or old == new:
            return
        for it in (self._scenario or {}).get("items") or []:
            for src in (it.get("sources") or []):
                if not isinstance(src, dict):
                    continue
                p = ensure_source_ui_block(src)
                for ld in p.get("link_defs") or []:
                    if isinstance(ld, dict) and str(ld.get("item") or "").strip() == old:
                        ld["item"] = new
                for jd in p.get("join_defs") or []:
                    if isinstance(jd, dict) and str(jd.get("item") or "").strip() == old:
                        jd["item"] = new
                pi = str(p.get("path_item") or "").strip()
                if pi == old and not pi.startswith("（主キー"):
                    p["path_item"] = new

    def _refresh_item_summaries_and_link_state(self) -> None:
        """テーブルと self._scenario.items を突き合わせて要約列を全行再生成し参照行の外観を更新する。"""
        items = (self._scenario or {}).get("items") or []
        for r in range(self._item_table.rowCount()):
            c0 = self._item_table.item(r, 0)
            disp = (c0.text() if c0 else "").strip()
            if r < len(items):
                it = items[r]
                if disp:
                    it["name"] = disp
            else:
                it = {
                    "name": disp or ("項目_%s" % (r + 1)),
                    "sources": [],
                    "write_mode": "fill_in",
                }
            summary = self._format_item_summary(it)
            self._item_table.setItem(r, 2, self._create_summary_item(summary))
        self._apply_linked_item_state()
        self._fit_item_table_columns()

    def _on_item_load(self) -> None:
        """選択中のラジオに応じて CSV または シート から項目を取得する。"""
        if self._rad_csv.isChecked():
            self._on_item_read_csv()
        else:
            self._on_item_read_sheet()

    def _on_item_read_csv(self) -> None:
        """CSVファイルから項目名を読込む。"""
        try:
            path, _ = QFileDialog.getOpenFileName(
                self,
                self._main_ui_disp("DIALOG_TITLE_ITEM_CSV", "項目定義CSVを選択"),
                self._file_dialog_initial_dir(),
                self._main_ui_disp("FILE_FILTER_ITEM_CSV", "CSV (*.csv);;すべて (*.*)"),
            )
            if not path:
                return
            text = Path(path).read_text(encoding="utf-8-sig")
            lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
            if not lines:
                show_warning_notice(
                    self,
                    self._main_ui_disp("TITLE_ITEM_CSV_LOAD", "CSV読込"),
                    self._main_ui_disp("MSG_ITEM_CSV_EMPTY", "項目が含まれていません。"),
                )
                return
            set_last_folder(str(Path(path).parent))
            headers = lines[0].split(",") if "," in lines[0] else lines
            self._on_scenario_clear_all()
            self._apply_items_to_table([h.strip() for h in headers if h.strip()])
            show_info_notice(
                self,
                self._main_ui_disp("TITLE_ITEM_CSV_LOAD", "CSV読込"),
                self._main_ui_disp("MSG_ITEM_CSV_LOADED_FMT", "項目を %s 件読込みました。")
                % len(headers),
            )
        except Exception as exc:
            show_warning_notice(
                self,
                self._main_ui_disp("TITLE_ITEM_CSV_LOAD", "CSV読込"),
                self._main_ui_disp("MSG_ITEM_CSV_FAILED_FMT", "読込に失敗しました: %s") % exc,
            )

    def _on_item_read_sheet(self) -> None:
        """押下時点のアクティブシートの1行目から項目名を読込む。"""
        try:
            from core.core_xlc import get_excel_context_from_hwnd
            ctx = get_excel_context_from_hwnd(self._parent_hwnd, "")
            if not ctx:
                show_warning_notice(
                    self,
                    self._main_ui_disp("TITLE_ITEM_SHEET_LOAD", "シート読込"),
                    self._main_ui_disp("MSG_ITEM_SHEET_NO_EXCEL", "Excel に接続できません。"),
                )
                return
            _app, _book, sheet, _hwnd = ctx
            rng = sheet.range((1, 1), (1, 256))
            row1 = rng.value
            if isinstance(row1, list):
                headers = [str(x).strip() if x is not None else "" for x in row1]
            elif row1 is not None:
                headers = [str(row1).strip()]
            else:
                headers = []
            headers = [h for h in headers if h]
            if not headers:
                show_warning_notice(
                    self,
                    self._main_ui_disp("TITLE_ITEM_SHEET_LOAD", "シート読込"),
                    self._main_ui_disp(
                        "MSG_ITEM_SHEET_NO_HEADER", "1行目に項目がありません。"
                    ),
                )
                return
            self._on_scenario_clear_all()
            self._apply_items_to_table(headers)
            show_info_notice(
                self,
                self._main_ui_disp("TITLE_ITEM_SHEET_LOAD", "シート読込"),
                self._main_ui_disp("MSG_ITEM_SHEET_LOADED_FMT", "項目を %s 件読込みました。")
                % len(headers),
            )
        except Exception as exc:
            show_warning_notice(
                self,
                self._main_ui_disp("TITLE_ITEM_SHEET_LOAD", "シート読込"),
                self._main_ui_disp("MSG_ITEM_SHEET_FAILED_FMT", "読込に失敗しました: %s") % exc,
            )

    @staticmethod
    def _summary_table_tooltip(display: str) -> str:
        """要約列ツールチップ: 区切り「 | 」を改行して複数行表示。"""
        return _data_agg_summary_table_tooltip(display)

    def _main_summary_column_tooltip(self, row: int) -> str:
        """
        シナリオ要約列のツールチップ。
        取得ソースが登録されている行はシナリオ編集ソース列と同形式（scenario_source_tooltip_plain）。
        連携先マスタ行・ソース未登録は従来どおり要約セルの「 | 」改行表示。
        """
        c2 = self._item_table.item(row, 2)
        display = (c2.text() if c2 else "") or ""
        fallback = self._summary_table_tooltip(display)
        link_targets = self._collect_link_target_names()
        c0 = self._item_table.item(row, 0)
        nm = (c0.text() if c0 else "").strip()
        if nm in link_targets:
            return fallback
        items = (self._scenario or {}).get("items") or []
        if row < 0 or row >= len(items):
            return fallback
        it = items[row]
        sources = [s for s in (it.get("sources") or []) if isinstance(s, dict)]
        if not sources:
            return fallback
        dn = _global_detail_name_cfg()
        dc = _global_detail_cell_cfg()
        blocks = [
            scenario_source_tooltip_plain(s, dn, detail_cell_cfg=dc) for s in sources
        ]
        blocks = [b for b in blocks if (b or "").strip()]
        if not blocks:
            return fallback
        if len(blocks) == 1:
            return blocks[0]
        return "\n\n".join(blocks)

    def _create_summary_item(self, text: str) -> QTableWidgetItem:
        """シナリオ要約列用の編集不可アイテムを生成する。"""
        item = QTableWidgetItem(text)
        item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
        item.setToolTip(self._summary_table_tooltip(text))
        return item

    def _apply_items_to_table(self, names: list[str]) -> None:
        """項目名リストをテーブルに反映する。既存のシナリオ要約は維持。"""
        prev_items = {}
        for i in range(self._item_table.rowCount()):
            c0 = self._item_table.item(i, 0)
            c2 = self._item_table.item(i, 2)
            if c0:
                prev_items[c0.text().strip()] = c2.text() if c2 else ""
        self._item_table.blockSignals(True)
        try:
            self._item_table.setRowCount(len(names))
            for i, name in enumerate(names):
                self._item_table.setItem(i, 0, QTableWidgetItem(name))
                btn_edit = self._create_scenario_edit_button(i)
                self._item_table.setCellWidget(i, 1, btn_edit)
                summary = prev_items.get(name, "")
                self._item_table.setItem(i, 2, self._create_summary_item(summary))
            self._apply_linked_item_state()
        finally:
            self._item_table.blockSignals(False)
        self._sync_item_table_master_name_roles()
        self._update_item_count_label()
        self._mark_scenario_dirty()
        self._fit_item_table_columns()
        self._refresh_excel_sort_item_combos()

    def _collect_link_target_names(self) -> set[str]:
        """シナリオ内のセル座標連携（link_defs）で参照される項目名を集約する。"""
        names: set[str] = set()
        items = (self._scenario or {}).get("items") or []
        for it in items:
            for src in (it.get("sources") or []):
                if not isinstance(src, dict):
                    continue
                p = source_ui_block(src)
                if isinstance(p, dict):
                    for ld in p.get("link_defs") or []:
                        if isinstance(ld, dict):
                            nm = str(ld.get("item") or "").strip()
                            if nm:
                                names.add(nm)
        return names

    def _is_linked_master_name(self, item_name: str) -> bool:
        nm = (item_name or "").strip()
        if not nm:
            return False
        return nm in self._collect_link_target_names()

    def _collect_join_target_names(self) -> set[str]:
        """シナリオ内の結合キー先 item 名を集約する。"""
        names: set[str] = set()
        for it in (self._scenario or {}).get("items") or []:
            for src in (it.get("sources") or []):
                if not isinstance(src, dict):
                    continue
                p = source_ui_block(src)
                if isinstance(p, dict):
                    for jd in p.get("join_defs") or []:
                        if isinstance(jd, dict):
                            nm = str(jd.get("item") or "").strip()
                            if nm:
                                names.add(nm)
        return names

    def _incoming_link_join_refs_line(self, master_name: str) -> str:
        """当該マスタ項目を参照する連携#N／結合#N／連携(名前) の要約1行。"""
        parts: list[str] = []
        nm = (master_name or "").strip()
        if not nm:
            return ""
        items = (self._scenario or {}).get("items") or []
        for it in items:
            src_item_name = str(it.get("name") or it.get("id") or "").strip() or "項目"
            for idx, src in enumerate(it.get("sources") or []):
                if not isinstance(src, dict):
                    continue
                sn = str(src.get("scenario_name") or "").strip()
                if not sn:
                    sn = "%s_シナリオ%d" % (src_item_name, idx + 1)
                ref_label = "%s_%s" % (src_item_name, sn)
                p = source_ui_block(src)
                if not isinstance(p, dict):
                    continue
                for i, ld in enumerate(p.get("link_defs") or []):
                    if isinstance(ld, dict) and str(ld.get("item") or "").strip() == nm:
                        parts.append("連携#%d：%s" % (i + 1, ref_label))
                for i, jd in enumerate(p.get("join_defs") or []):
                    if isinstance(jd, dict) and str(jd.get("item") or "").strip() == nm:
                        parts.append("結合#%d：%s" % (i + 1, ref_label))
                pi = str(p.get("path_item") or "").strip()
                if pi == nm and not pi.startswith("（主キー"):
                    parts.append("連携(名前)：%s" % ref_label)
        return " | ".join(parts)

    def _apply_linked_item_state(self) -> None:
        """連携／結合参照行の背景・シナリオ編集ボタンを更新する（連携と結合の両方時は連携を優先）。"""
        link_targets = self._collect_link_target_names()
        join_targets = self._collect_join_target_names()
        white = QColor("#FFFFFF")
        black = QColor("#000000")
        for r in range(self._item_table.rowCount()):
            c0 = self._item_table.item(r, 0)
            c2 = self._item_table.item(r, 2)
            nm = (c0.text() if c0 else "").strip()
            linked = nm in link_targets
            joined = nm in join_targets
            if linked:
                bg = _ROW_BG_LINK
                fg = _ROW_FG_LINKED
                tip = "連携項目で参照中のため編集対象外"
            elif joined:
                bg = _ROW_BG_JOIN
                fg = black
                tip = "結合キーで参照中のマスタ項目"
            else:
                bg = white
                fg = black
                tip = ""
            for cell in (c0, c2):
                if cell is None:
                    continue
                cell.setForeground(fg)
                cell.setBackground(bg)
            if c0 is not None:
                c0.setToolTip(tip)
            if c2 is not None:
                c2.setToolTip(self._main_summary_column_tooltip(r))
            btn = self._item_table.cellWidget(r, 1)
            if isinstance(btn, QPushButton):
                btn.setEnabled(not linked)
                if linked:
                    btn.setStyleSheet(_BTN_EDIT_LINKED_DISABLED)
                    btn.setToolTip("連携項目で参照中のため編集不可")
                else:
                    btn.setStyleSheet(_BTN_EDIT_ENABLED)
                    btn.setToolTip("")

    def _get_selected_row_indices(self) -> list[int]:
        """選択中の行インデックスを昇順で返す。連続・歯抜け選択に対応。"""
        indices: set[int] = set()
        sm = self._item_table.selectionModel()
        if sm is not None:
            for idx in sm.selectedRows():
                indices.add(int(idx.row()))
        return sorted(indices)

    def _next_scenario_item_numeric_id(self, items: list[dict[str, Any]]) -> str:
        mx = 0
        for it in items:
            if not isinstance(it, dict):
                continue
            s = str(it.get("id") or "")
            if s.startswith("item_"):
                tail = s[5:]
                if tail.isdigit():
                    mx = max(mx, int(tail))
        return "item_%d" % (mx + 1)

    def _collect_item_names_from_table(self) -> set[str]:
        names: set[str] = set()
        for i in range(self._item_table.rowCount()):
            c0 = self._item_table.item(i, 0)
            if c0:
                names.add(c0.text().strip())
        return names

    def _make_unique_default_item_name(self, existing: set[str]) -> str:
        k = self._item_table.rowCount() + 1
        while True:
            cand = "項目_%d" % k
            if cand not in existing:
                return cand
            k += 1

    def _insert_master_item_row_at(self, insert_at: int) -> None:
        """マスタ項目行を insert_at に挿入（シナリオは空ソース）。"""
        from svc import svc_data_agg_scenario as scenario_mod

        if not self._scenario:
            self._scenario = scenario_mod.create_empty_scenario()
        self._scenario = self._build_scenario_from_ui()
        items = list(self._scenario.get("items") or [])
        n = len(items)
        insert_at = max(0, min(int(insert_at), n))
        names = self._collect_item_names_from_table()
        nm = self._make_unique_default_item_name(names)
        nid = self._next_scenario_item_numeric_id(items)
        new_item: dict[str, Any] = {
            "id": nid,
            "name": nm,
            "sources": [],
            "write_mode": "fill_in",
        }
        items.insert(insert_at, new_item)
        self._scenario["items"] = items
        self._item_table.insertRow(insert_at)
        self._item_table.setItem(insert_at, 0, QTableWidgetItem(nm))
        self._item_table.setCellWidget(
            insert_at, 1, self._create_scenario_edit_button(insert_at)
        )
        self._item_table.setItem(insert_at, 2, self._create_summary_item(""))
        self._apply_linked_item_state()
        self._sync_item_table_master_name_roles()
        self._mark_scenario_dirty()
        self._fit_item_table_columns()
        self._refresh_excel_sort_item_combos()
        self._update_item_count_label()
        self._item_table.selectRow(insert_at)

    def _remove_selected_master_item_rows(self) -> None:
        """選択中のマスタ項目行を削除。"""
        indices = self._get_selected_row_indices()
        if not indices:
            return
        self._scenario = self._build_scenario_from_ui()
        items = list(self._scenario.get("items") or [])
        for r in reversed(indices):
            if 0 <= r < len(items):
                del items[r]
            self._item_table.removeRow(r)
        self._scenario["items"] = items
        self._apply_linked_item_state()
        self._sync_item_table_master_name_roles()
        self._mark_scenario_dirty()
        self._fit_item_table_columns()
        self._refresh_excel_sort_item_combos()
        self._update_item_count_label()
        if self._item_table.rowCount() > 0:
            first = int(indices[0])
            sel = min(first, self._item_table.rowCount() - 1)
            self._item_table.selectRow(max(0, sel))

    def _on_item_table_context_menu(self, pos: QPoint) -> None:
        """メインのマスタ項目表: 行の挿入・削除。"""
        idx = self._item_table.indexAt(pos)
        r = int(idx.row())
        if r >= 0:
            self._item_table.selectRow(r)
        menu = QMenu(self)
        ui = self._ui or {}
        a_up = menu.addAction(
            _ui_disp_str(ui, "CTX_INSERT_ITEM_ABOVE", "上の行を追加")
        )
        a_dn = menu.addAction(
            _ui_disp_str(ui, "CTX_INSERT_ITEM_BELOW", "下の行を追加")
        )
        menu.addSeparator()
        a_del = menu.addAction(
            _ui_disp_str(ui, "CTX_REMOVE_ITEM_ROW", "削除")
        )
        chosen = menu.exec(self._item_table.viewport().mapToGlobal(pos))
        if chosen == a_up:
            ins = r if r >= 0 else 0
            self._insert_master_item_row_at(ins)
        elif chosen == a_dn:
            ins = (r + 1) if r >= 0 else self._item_table.rowCount()
            self._insert_master_item_row_at(ins)
        elif chosen == a_del:
            self._remove_selected_master_item_rows()

    def _on_move_items_up(self) -> None:
        """選択行を上に移動する。連続・歯抜け選択に対応。"""
        indices = self._get_selected_row_indices()
        if not indices or indices[0] == 0:
            return
        self._move_rows(indices, -1)

    def _on_move_items_down(self) -> None:
        """選択行を下に移動する。連続・歯抜け選択に対応。"""
        indices = self._get_selected_row_indices()
        n = self._item_table.rowCount()
        if not indices or indices[-1] >= n - 1:
            return
        self._move_rows(indices, 1)

    def _move_rows(self, indices: list[int], delta: int) -> None:
        """選択行を delta 行分移動する。連続選択はブロック移動、歯抜け選択は歯抜けを維持して移動。"""
        if not indices:
            return
        is_consecutive = (
            len(indices) <= 1
            or all(indices[i + 1] - indices[i] == 1 for i in range(len(indices) - 1))
        )
        if is_consecutive:
            self._move_rows_block(indices, delta)
        else:
            self._move_rows_gap(indices, delta)

    def _move_scenario_items_block(self, indices: list[int], delta: int) -> None:
        """連続選択移動に合わせて self._scenario.items の並びも同期する。"""
        if not self._scenario:
            return
        items = list((self._scenario or {}).get("items") or [])
        if not items:
            return
        picked = [items[i] for i in indices if 0 <= i < len(items)]
        for i in reversed(indices):
            if 0 <= i < len(items):
                del items[i]
        if delta < 0:
            insert_at = indices[0] - 1
        else:
            insert_at = min(indices) + 1
        insert_at = max(0, min(insert_at, len(items)))
        for k, ent in enumerate(picked):
            items.insert(insert_at + k, ent)
        self._scenario["items"] = items

    def _move_scenario_items_gap(self, indices: list[int], delta: int) -> None:
        """歯抜け選択移動に合わせて self._scenario.items の並びも同期する。"""
        if not self._scenario:
            return
        items = list((self._scenario or {}).get("items") or [])
        if not items:
            return
        picked = [items[i] for i in indices if 0 <= i < len(items)]
        new_positions = [r + delta for r in indices]
        for i in reversed(indices):
            if 0 <= i < len(items):
                del items[i]
        for k, ent in enumerate(picked):
            pos = max(0, min(new_positions[k], len(items)))
            items.insert(pos, ent)
        self._scenario["items"] = items

    def _move_rows_block(self, indices: list[int], delta: int) -> None:
        """連続選択行をブロックとして移動する。"""
        self._move_scenario_items_block(indices, delta)
        rows_data: list[tuple[str, str]] = []
        for r in indices:
            c0 = self._item_table.item(r, 0)
            c2 = self._item_table.item(r, 2)
            name = (c0.text() if c0 else "").strip()
            summary = (c2.text() if c2 else "").strip()
            rows_data.append((name, summary))
        for r in reversed(indices):
            self._item_table.removeRow(r)
        n_after = self._item_table.rowCount()
        if delta < 0:
            insert_at = indices[0] - 1
        else:
            insert_at = min(indices) + 1
        insert_at = max(0, min(insert_at, n_after))
        for i, (name, summary) in enumerate(rows_data):
            r = insert_at + i
            self._item_table.insertRow(r)
            self._item_table.setItem(r, 0, QTableWidgetItem(name))
            self._item_table.setCellWidget(r, 1, self._create_scenario_edit_button(r))
            self._item_table.setItem(r, 2, self._create_summary_item(summary))
        self._apply_linked_item_state()
        self._sync_item_table_master_name_roles()
        self._mark_scenario_dirty()
        self._fit_item_table_columns()
        self._refresh_excel_sort_item_combos()
        self._finish_move_selection(insert_at, len(rows_data), delta)

    def _move_rows_gap(self, indices: list[int], delta: int) -> None:
        """歯抜け選択行を歯抜けを維持したまま移動する。"""
        self._move_scenario_items_gap(indices, delta)
        rows_data: list[tuple[str, str]] = []
        for r in indices:
            c0 = self._item_table.item(r, 0)
            c2 = self._item_table.item(r, 2)
            name = (c0.text() if c0 else "").strip()
            summary = (c2.text() if c2 else "").strip()
            rows_data.append((name, summary))
        new_positions = [r + delta for r in indices]
        for r in reversed(indices):
            self._item_table.removeRow(r)
        for k, (name, summary) in enumerate(rows_data):
            insert_at = new_positions[k]
            insert_at = max(0, min(insert_at, self._item_table.rowCount()))
            self._item_table.insertRow(insert_at)
            self._item_table.setItem(insert_at, 0, QTableWidgetItem(name))
            self._item_table.setCellWidget(
                insert_at, 1, self._create_scenario_edit_button(insert_at)
            )
            self._item_table.setItem(insert_at, 2, self._create_summary_item(summary))
        self._apply_linked_item_state()
        self._sync_item_table_master_name_roles()
        self._mark_scenario_dirty()
        self._fit_item_table_columns()
        self._refresh_excel_sort_item_combos()
        new_indices = sorted(new_positions)
        self._finish_move_selection_gap(new_indices, delta)

    def _finish_move_selection(
        self, insert_at: int, count: int, delta: int
    ) -> None:
        """移動後の選択とスクロールを行う（ブロック移動用）。"""
        sm = self._item_table.selectionModel()
        if sm:
            sm.clearSelection()
            for r in range(insert_at, insert_at + count):
                idx = self._item_table.model().index(r, 0)
                sm.select(
                    idx,
                    QItemSelectionModel.SelectionFlag.Select
                    | QItemSelectionModel.SelectionFlag.Rows,
                )
        scroll_row = insert_at if delta < 0 else insert_at + count - 1
        hint = (
            QAbstractItemView.ScrollHint.EnsureVisible
            if delta < 0
            else QAbstractItemView.ScrollHint.PositionAtBottom
        )
        self._item_table.scrollTo(
            self._item_table.model().index(scroll_row, 0), hint
        )

    def _finish_move_selection_gap(
        self, new_indices: list[int], delta: int
    ) -> None:
        """移動後の選択とスクロールを行う（歯抜け移動用）。"""
        sm = self._item_table.selectionModel()
        if sm:
            sm.clearSelection()
            for r in new_indices:
                idx = self._item_table.model().index(r, 0)
                sm.select(
                    idx,
                    QItemSelectionModel.SelectionFlag.Select
                    | QItemSelectionModel.SelectionFlag.Rows,
                )
        if new_indices:
            scroll_row = new_indices[0] if delta < 0 else new_indices[-1]
            hint = (
                QAbstractItemView.ScrollHint.EnsureVisible
                if delta < 0
                else QAbstractItemView.ScrollHint.PositionAtBottom
            )
            self._item_table.scrollTo(
                self._item_table.model().index(scroll_row, 0), hint
            )

    def _on_scenario_edit_button_clicked(self) -> None:
        """編集ボタンから行番号を解決してシナリオ編集を開く（行移動後もずれない）。"""
        btn = self.sender()
        if not isinstance(btn, QPushButton):
            return
        for r in range(self._item_table.rowCount()):
            if self._item_table.cellWidget(r, 1) is btn:
                self._on_scenario_edit(r)
                return

    def _on_item_table_double_clicked(self, row: int, column: int) -> None:
        """項目名列のダブルクリックで編集を開始する。"""
        if row < 0 or column < 0:
            return
        if column != 0:
            return
        c0 = self._item_table.item(row, 0)
        if c0 is None:
            return
        if self._is_linked_master_name((c0.text() if c0 else "").strip()):
            return
        self._item_table.setCurrentItem(c0)
        self._item_table.editItem(c0)

    def _create_scenario_edit_button(self, row: int) -> QPushButton:
        """シナリオ編集ボタンを生成する（スタイル適用）。"""
        _ = row  # 行はクリック時に cellWidget 照合で解決する
        btn = QPushButton(str(self._ui.get("BTN_EDIT_SCENARIO") or "編集").strip())
        btn.setStyleSheet(_BTN_EDIT_ENABLED)
        btn.setCursor(Qt.CursorShape.PointingHandCursor)
        btn.clicked.connect(self._on_scenario_edit_button_clicked)
        return btn

    @staticmethod
    def _item_table_max_line_advance(fm: Any, text: str, extra: int) -> int:
        """複数行セル用に最長行の水平ピクセル幅＋余白。"""
        mx = 0
        for line in str(text or "").split("\n"):
            mx = max(mx, fm.horizontalAdvance(line))
        return mx + extra

    def _resize_item_table_scenario_section_only(self) -> None:
        """列1「シナリオ」幅のみ。ITEM_TABLE_SCENARIO_COLUMN_WIDTH_PX または見出し＋ボタンから算出。"""
        hdr = self._item_table.horizontalHeader()
        fixed = self._ui.get("ITEM_TABLE_SCENARIO_COLUMN_WIDTH_PX")
        try:
            if fixed is not None and int(fixed) > 0:
                hdr.resizeSection(1, int(fixed))
                return
        except (TypeError, ValueError):
            pass
        label = str(self._ui.get("TABLE_HEADER_EDIT") or "シナリオ")
        fm = self._item_table.fontMetrics()
        try:
            pad = int(self._ui.get("ITEM_TABLE_SCENARIO_COLUMN_EXTRA_PADDING_PX") or 28)
        except (TypeError, ValueError):
            pad = 28
        w_hdr = fm.horizontalAdvance(label) + pad
        btn_txt = str(self._ui.get("BTN_EDIT_SCENARIO") or "編集").strip()
        w_btn = fm.horizontalAdvance(btn_txt) + pad
        try:
            floor = int(self._ui.get("ITEM_TABLE_SCENARIO_COLUMN_MIN_WIDTH_PX") or 72)
        except (TypeError, ValueError):
            floor = 72
        hdr.resizeSection(1, max(floor, w_hdr, w_btn))

    def _fit_item_table_columns(self) -> None:
        """項目表3列: 0/2 は内容幅優先、足りない分は横スクロールで見る。"""
        tbl = self._item_table
        hdr = tbl.horizontalHeader()
        ui = self._ui
        fm = tbl.fontMetrics()
        hdr.setStretchLastSection(False)

        hdr.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        tbl.resizeColumnToContents(0)
        w0 = hdr.sectionSize(0)
        try:
            w0_min = int(ui.get("ITEM_TABLE_NAME_COLUMN_MIN_WIDTH_PX") or 64)
        except (TypeError, ValueError):
            w0_min = 64
        w0_max = ui.get("ITEM_TABLE_NAME_COLUMN_MAX_WIDTH_PX")
        try:
            if w0_max is not None:
                w0_max_i = int(w0_max)
                if w0_max_i > 0:
                    w0 = min(w0, w0_max_i)
        except (TypeError, ValueError):
            pass
        w0 = max(w0, w0_min)
        hdr.resizeSection(0, w0)
        hdr.setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)

        self._resize_item_table_scenario_section_only()
        w1 = hdr.sectionSize(1)

        try:
            pad2 = int(ui.get("ITEM_TABLE_SUMMARY_COLUMN_EXTRA_PADDING_PX") or 24)
        except (TypeError, ValueError):
            pad2 = 24
        try:
            w2_floor = int(ui.get("ITEM_TABLE_SUMMARY_COLUMN_MIN_WIDTH_PX") or 120)
        except (TypeError, ValueError):
            w2_floor = 120

        h2 = str(ui.get("TABLE_HEADER_SUMMARY") or "シナリオ要約")
        w2_need = fm.horizontalAdvance(h2) + pad2
        for r in range(tbl.rowCount()):
            it = tbl.item(r, 2)
            if it is not None:
                w2_need = max(w2_need, self._item_table_max_line_advance(fm, it.text(), pad2))
        w2 = max(w2_floor, w2_need)
        hdr.setSectionResizeMode(2, QHeaderView.ResizeMode.Fixed)
        hdr.resizeSection(2, w2)

    def _rescan_paths_from_current_form(self) -> list[str]:
        """基準フォルダ・サブフォルダ・拡張子・キーワードで再走査し、検出一覧と同じ条件のパスを返す。"""
        try:
            from svc import svc_data_agg_scan as scan_mod

            st = self._get_scan_state()
            sp = str(st.get("start_path") or "").strip() or "."
            exts = st.get("extensions") or [".xlsx", ".xlsm", ".csv"]
            if not exts:
                return list(getattr(self, "_file_list_items", None) or [])
            paths = [
                str(p)
                for p in scan_mod.scan_folder(
                    sp,
                    recursive=bool(st.get("recursive")),
                    extensions=tuple(exts),
                    keyword=str(st.get("keyword") or ""),
                )
            ]
            self._file_list_items = paths
            self._file_list.clear()
            for fp in paths:
                self._file_list.addItem(fp)
            self._update_detected_file_count_label()
            return paths
        except Exception:
            return list(getattr(self, "_file_list_items", None) or [])

    def _on_debug(self) -> None:
        """デバッグウィンドウ（§3.1.3）を開く。項目一覧・検出パスを反映（未登録時はデモ）。"""
        _u = lambda k, d: _ui_disp_str(self._ui or {}, k, d)
        t_dbg = _u("TITLE_DEBUG", "デバッグ").strip() or "デバッグ"
        try:
            from ui_qt.ui_data_agg_debug import create_data_agg_debug_dialog

            data = self._build_scenario_from_ui()
            items = data.get("items") or []
            paths = list(getattr(self, "_file_list_items", None) or [])
            dbg = (_get_cfg().get("SCREENS") or {}).get("DEBUG") or {}
            scan_root = (self._edit_start_path.text() or "").strip() or None
            dlg = create_data_agg_debug_dialog(
                self,
                dbg,
                live_items=items if items else None,
                scan_paths=paths if paths else None,
                fixed_mode=1,
                scenario_for_dry_run=data,
                scan_root=scan_root,
            )
            dlg.exec()
        except Exception as exc:
            _data_agg_warn_debug_open_failed(self, t_dbg, exc)

    def _on_scenario_export(self) -> None:
        """読込済みシナリオを、ソース1行単位で Excel シートへ書き出す（既定シート名＝ファイル stem）。"""
        from core.core_xlc import clear_sheet_used_range, get_excel_context_from_hwnd
        from svc import svc_data_agg_scenario as scenario_mod
        from svc import svc_data_agg_write as write_mod
        from svc.data_agg_scenario_export import (
            build_scenario_definition_sheet_matrix_with_headers,
        )

        _u = lambda k, d: str((self._ui or {}).get(k) or d)
        if not (self._scenario_path or "").strip():
            return
        try:
            data = self._build_scenario_from_ui()
        except Exception as exc:
            show_warning_notice(
                self, _u("TITLE_SCENARIO_EXPORT", "シナリオ出力"), str(exc)
            )
            return
        errs = scenario_mod.validate_scenario(data)
        if errs:
            show_warning_notice(
                self,
                _u("TITLE_SCENARIO_EXPORT", "シナリオ出力"),
                _u("MSG_SCENARIO_EXPORT_VALIDATE", "検証エラーのため出力できません。")
                + "\n"
                + "\n".join(str(x) for x in errs[:8]),
            )
            return
        stem = Path(self._scenario_path).stem
        base_name = write_mod.sanitize_excel_tab_name(stem)
        if not base_name:
            base_name = "シナリオ"
        screen_cfg = (_get_cfg().get("SCREENS") or {}).get("SCENARIO_EDIT") or {}
        headers, rows = build_scenario_definition_sheet_matrix_with_headers(
            data.get("items") or [],
            screen_cfg,
            self._ui or {},
        )
        if not rows:
            show_info_notice(
                self,
                _u("TITLE_SCENARIO_EXPORT", "シナリオ出力"),
                _u("MSG_SCENARIO_EXPORT_EMPTY", "出力する取得ソースがありません。"),
            )
            return
        try:
            ctx = get_excel_context_from_hwnd(self._parent_hwnd, self._sheet_id)
        except Exception:
            ctx = None
        if not ctx:
            show_warning_notice(
                self,
                _u("TITLE_SCENARIO_EXPORT", "シナリオ出力"),
                _u("MSG_SCENARIO_EXPORT_NO_EXCEL", "Excel に接続できませんでした。"),
            )
            return
        _app, book, _sheet_master, _hwnd = ctx
        title = _u("TITLE_SCENARIO_EXPORT", "シナリオ出力")
        prompt = _u(
            "MSG_SCENARIO_EXPORT_NAME_PROMPT",
            "同名のシートがあります。別名を入力するか、同名のまま確定して上書き確認へ進んでください。",
        )
        ow_q = _u(
            "MSG_SCENARIO_EXPORT_OVERWRITE",
            "シート「%s」は既に存在します。内容をクリアして上書きしますか？",
        )
        bad_name = _u("MSG_SCENARIO_EXPORT_BAD_NAME", "シート名が無効です。")

        def _sheet_names() -> set[str]:
            return {str(s.name) for s in book.sheets}

        def resolve_target_sheet() -> tuple[Any, str] | None:
            names0 = _sheet_names()
            if base_name not in names0:
                try:
                    ws_new = book.sheets.add(name=base_name, after=book.sheets.active)
                except Exception as e:
                    show_warning_notice(self, title, str(e))
                    return None
                return (ws_new, base_name)
            while True:
                names = _sheet_names()
                sheet_lbl = _ui_disp_str(
                    self._ui or {}, "LABEL_SCENARIO_EXPORT_SHEET", "シート名"
                )
                prompt_body = _ui_disp_str(
                    self._ui or {},
                    "MSG_SCENARIO_EXPORT_NAME_PROMPT",
                    "同名のシートがあります。別名を入力するか、同名のまま確定して上書き確認へ進んでください。",
                )
                input_lbl = (
                    sheet_lbl
                    if not (prompt_body or "").strip()
                    else ("%s\n\n%s" % (sheet_lbl, prompt_body))
                )
                text, ok = QInputDialog.getText(self, title, input_lbl, text=base_name)
                if not ok:
                    return None
                sn = write_mod.sanitize_excel_tab_name(text)
                if not sn:
                    show_warning_notice(self, title, bad_name)
                    continue
                if sn in names:
                    yn = QMessageBox.question(
                        self,
                        title,
                        ow_q % sn,
                        QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                        QMessageBox.StandardButton.No,
                    )
                    if yn != QMessageBox.StandardButton.Yes:
                        continue
                    try:
                        ws_ex = book.sheets[sn]
                    except Exception:
                        show_warning_notice(
                            self,
                            title,
                            _u("MSG_SCENARIO_EXPORT_OPEN_SHEET_FAIL", "シートを開けませんでした。"),
                        )
                        return None
                    try:
                        clear_sheet_used_range(ws_ex)
                    except Exception:
                        pass
                    return (ws_ex, sn)
                try:
                    ws_new = book.sheets.add(name=sn, after=book.sheets.active)
                except Exception as e:
                    show_warning_notice(self, title, str(e))
                    return None
                return (ws_new, sn)

        pair = resolve_target_sheet()
        if pair is None:
            return
        ws_out, used_name = pair
        sheet_title = ("%s %s" % (stem, _u("EXPORT_SHEET_TITLE_SUFFIX", "設定内容一覧"))).strip()
        try:
            write_mod.write_scenario_export_table(
                ws_out, headers, rows, sheet_title=sheet_title
            )
            try:
                ws_out.activate()
            except Exception:
                pass
        except Exception as exc:
            show_warning_notice(
                self,
                title,
                _u("MSG_SCENARIO_EXPORT_WRITE_FAIL", "書き込みに失敗しました。")
                + "\n%s"
                % exc,
            )
            return
        show_done_notice(
            self,
            title,
            _u("MSG_SCENARIO_EXPORT_DONE", "書き出しました。\nシート名: %s") % used_name,
        )

    def _apply_item_edit_result(self, row: int, item: dict[str, Any], result: dict[str, Any]) -> None:
        srcs = result.get("sources") or []
        items = (self._scenario or {}).setdefault("items", [])
        tgt = items[row] if 0 <= row < len(items) else item
        tgt["sources"] = srcs
        tgt.pop("join_path_item_id", None)
        wm = result.get("write_mode")
        if wm is not None and str(wm).strip():
            tgt["write_mode"] = str(wm).strip().lower()
        self._refresh_item_summaries_and_link_state()
        self._sync_item_table_master_name_roles()
        self._mark_scenario_dirty()

    def _on_scenario_edit(self, row: int) -> None:
        """シナリオ編集画面を開く（項目単位）。OK 時にシナリオ更新・要約再描画。"""
        try:
            from svc import svc_data_agg_scenario as scenario_mod
            c0 = self._item_table.item(row, 0)
            if self._is_linked_master_name((c0.text() if c0 else "").strip()):
                return

            if not self._scenario:
                self._scenario = scenario_mod.create_empty_scenario()
            self._scenario = self._build_scenario_from_ui()
            items = self._scenario.get("items") or []
            while len(items) <= row:
                items.append(
                    {
                        "id": "item_%s" % len(items),
                        "name": "項目_%s" % (len(items) + 1),
                        "sources": [],
                        "write_mode": "fill_in",
                    }
                )
                self._scenario["items"] = items
            item = items[row]
            cell = self._item_table.item(row, 0)
            name = (cell.text() if cell else "").strip() or (
                item.get("name") or ("項目_%s" % (row + 1))
            ).strip()
            item_id = str(item.get("id") or ("item_%s" % row)).strip() or ("item_%s" % row)
            screen_cfg = _get_cfg().get("SCREENS") or {}
            scenario_edit_cfg = screen_cfg.get("SCENARIO_EDIT") or {}
            scan_paths_hint: list[str] = list(getattr(self, "_file_list_items", None) or [])
            try:
                from svc import svc_data_agg_scan as scan_mod

                st = self._get_scan_state()
                sp = str(st.get("start_path") or "").strip()
                if sp:
                    exts = st.get("extensions") or [".xlsx", ".xlsm", ".csv"]
                    scan_paths_hint = [
                        str(p)
                        for p in scan_mod.scan_folder(
                            sp,
                            recursive=bool(st.get("recursive")),
                            extensions=tuple(exts),
                            keyword=str(st.get("keyword") or ""),
                        )
                    ]
            except Exception:
                pass
            dlg = _ScenarioEditDialog(
                name,
                item_id,
                item,
                self,
                scenario_edit_cfg,
                items=items,
                scan_paths_hint=scan_paths_hint,
                on_registered=lambda r, it=item, rr=row: self._apply_item_edit_result(rr, it, r),
            )
            dlg.exec()
        except Exception as exc:
            show_warning_notice(
                self,
                "シナリオ編集",
                "シナリオ編集画面を開けませんでした。\n%s" % exc,
            )

    def _on_folder_select(self) -> None:
        """基準フォルダを選択する。"""
        try:
            _lock_xl = False
            try:
                from ui_qt.ui_common import enable_excel_window, want_excel_child_hwnd_lock_while_modal

                _lock_xl = bool(
                    self._parent_hwnd
                    and want_excel_child_hwnd_lock_while_modal(self._window_cfg or {})
                )
                if _lock_xl:
                    enable_excel_window(self._parent_hwnd, False)
            except Exception:
                _lock_xl = False
            try:
                from ui_qt import ui_fld
                cur = (self._edit_start_path.text() or "").strip()
                lf = get_last_folder()
                initial = cur or lf or os.path.expanduser("~")
                cfg = _get_cfg()
                label = ((cfg.get("MAIN") or {}).get("UI") or {}).get("LABEL_BASE_FOLDER") or "基準フォルダ"
                path = ui_fld.show_folder_dialog(
                    self, str(label) + "を選択", initial, (cfg or {}).get("FOLDER") or {}
                )
                if path:
                    self._edit_start_path.setText(path)
                    set_last_folder(path)
                    self._on_scan(auto_mode=True)
            finally:
                if _lock_xl and self._parent_hwnd:
                    try:
                        from ui_qt.ui_common import enable_excel_window

                        enable_excel_window(self._parent_hwnd, True)
                    except Exception:
                        pass
        except Exception as exc:
            show_warning_notice(
                self,
                _ui_disp_str(self._ui or {}, "BTN_FOLDER", "フォルダ選択"),
                _ui_disp_str(
                    self._ui or {},
                    "MSG_FOLDER_SELECT_FAILED_FMT",
                    "フォルダ選択に失敗しました: %s",
                )
                % exc,
            )

    def _on_scan(self, auto_mode: bool = False) -> None:
        """検索実行でフォルダを走査し、検出ファイル一覧を更新する（非同期）。"""
        self._request_folder_scan(auto_mode=auto_mode)

    def _get_scan_state(self) -> dict[str, Any]:
        """現在の走査条件を返す。"""
        exts: list[str] = []
        if self._chk_ext_xls.isChecked():
            exts.append(".xls")
        if self._chk_ext_xlsx.isChecked():
            exts.append(".xlsx")
        if self._chk_ext_xlsm.isChecked():
            exts.append(".xlsm")
        if self._chk_ext_csv.isChecked():
            exts.append(".csv")
        return {
            "start_path": self._edit_start_path.text().strip(),
            "recursive": self._chk_recursive.isChecked(),
            "extensions": exts if exts else [".xlsx", ".xlsm", ".csv"],
            "keyword": self._edit_keyword.text().strip(),
        }

    def _file_dialog_initial_dir(self) -> str:
        """%TEMP%\\csv_tool\\last_folder.txt 等に有効なフォルダがあればそのパス。無ければ空（Qt の従来どおりの初期位置）。"""
        lf = get_last_folder()
        return lf if lf else ""

    def _on_scenario_load(self) -> None:
        """シナリオ読込ダイアログで .json を選択し、項目一覧・対象ファイル一覧を更新する。"""
        try:
            path, _ = QFileDialog.getOpenFileName(
                self,
                self._main_ui_disp("DIALOG_TITLE_SCENARIO_OPEN", "シナリオを読込"),
                self._file_dialog_initial_dir(),
                self._main_ui_disp(
                    "FILE_FILTER_SCENARIO_OPEN",
                    "シナリオ (*.json *.scenario);;すべてのファイル (*.*)",
                ),
            )
            if not path:
                return
            from svc import svc_data_agg_scenario as scenario_mod
            data = scenario_mod.load_scenario(path)
            errs = scenario_mod.validate_scenario(data)
            if errs:
                title_ld = _ui_disp_str(self._ui or {}, "BTN_SCENARIO_LOAD", "シナリオ読込")
                pre = _ui_disp_str(
                    self._ui or {},
                    "MSG_SCENARIO_LOAD_VALIDATE_PREFIX",
                    "シナリオの検証エラー:",
                )
                show_warning_notice(
                    self,
                    title_ld,
                    pre + "\n" + "\n".join(errs[:5]),
                )
                return
            self._scenario = data
            self._scenario_path = path
            self._scenario_save_empty_filename = False
            set_last_folder(str(Path(path).parent))
            self._suppress_scenario_dirty = True
            self._item_table.blockSignals(True)
            try:
                # 項目一覧・シナリオ要約を更新（編集ボタン付き）
                items = data.get("items") or []
                names = [
                    str(it.get("name") or it.get("id") or ("項目_%s" % i))
                    for i, it in enumerate(items)
                ]
                summaries = [self._format_item_summary(it) for it in items]
                self._item_table.setRowCount(len(items))
                for i, (name, summary) in enumerate(zip(names, summaries)):
                    self._item_table.setItem(i, 0, QTableWidgetItem(name))
                    self._item_table.setCellWidget(i, 1, self._create_scenario_edit_button(i))
                    self._item_table.setItem(i, 2, self._create_summary_item(summary))
                self._apply_linked_item_state()
                # 走査条件を反映
                scan = data.get("scan") or {}
                self._edit_start_path.setText(str(scan.get("start_path") or ""))
                if not (self._edit_start_path.text() or "").strip():
                    lf = get_last_folder()
                    if lf:
                        self._edit_start_path.setText(lf)
                self._chk_recursive.setChecked(bool(scan.get("recursive")))
                exts = scan.get("extensions") or [".xlsx", ".xlsm", ".xls", ".csv"]
                ext_set = set(
                    (e.strip().lower() if e.strip().startswith(".") else "." + e.strip().lower())
                    for e in (exts if isinstance(exts, list) else [exts])
                    if e
                )
                self._chk_ext_xls.setChecked(".xls" in ext_set)
                self._chk_ext_xlsx.setChecked(".xlsx" in ext_set)
                self._chk_ext_xlsm.setChecked(".xlsm" in ext_set)
                self._chk_ext_csv.setChecked(".csv" in ext_set)
                self._edit_keyword.setText(str(scan.get("keyword") or ""))
                self._file_list_items = []
                self._file_list.clear()
                self._update_detected_file_count_label()
                self._apply_excel_options_to_ui(data.get("excel_options"))
            finally:
                self._item_table.blockSignals(False)
                self._suppress_scenario_dirty = False
            self._sync_item_table_master_name_roles()
            self._clear_scenario_dirty()
            self._update_item_count_label()
            self._fit_item_table_columns()
            self._refresh_scenario_display_label()
            title_ld = _ui_disp_str(self._ui or {}, "BTN_SCENARIO_LOAD", "シナリオ読込")
            done_msg = _ui_disp_str(
                self._ui or {},
                "MSG_SCENARIO_LOAD_DONE_FMT",
                "シナリオを読込みました。\n項目数: %d\n対象ファイル数: %d",
            )
            items_count = len(items)

            def _after_scenario_load_scan(n_files: int) -> None:
                show_done_notice(
                    self,
                    title_ld,
                    done_msg % (items_count, n_files),
                )

            self._request_folder_scan(auto_mode=True, on_complete=_after_scenario_load_scan)
        except FileNotFoundError as e:
            show_warning_notice(
                self,
                _ui_disp_str(self._ui or {}, "BTN_SCENARIO_LOAD", "シナリオ読込"),
                str(e),
            )
        except Exception as exc:
            show_warning_notice(
                self,
                _ui_disp_str(self._ui or {}, "BTN_SCENARIO_LOAD", "シナリオ読込"),
                _ui_disp_str(
                    self._ui or {},
                    "MSG_SCENARIO_LOAD_FAILED_FMT",
                    "読込に失敗しました: %s",
                )
                % exc,
            )

    def _format_one_source_summary(self, s: dict[str, Any]) -> str:
        """1 ソース分の要約（メイン表のシナリオ要約列用）。内部値は出さず日本語表記に揃える。"""
        stype = (s.get("type") or "cell").strip().lower()
        if stype in ("metadata", "meta", "filename"):
            stype = "name_extract"
        sn_user = str(s.get("scenario_name") or "").strip()
        if stype == "name_extract":
            pb = source_ui_block(s) or {}
            dn = _global_detail_name_cfg()
            segs = name_extract_setting_lines(s, pb, dn, bullet="")
            head = ("%s " % sn_user) if sn_user else ""
            return "%s名前取得 %s" % (head, " | ".join(segs))
        pb = source_ui_block(s) or {}
        fp = str(pb.get("file_pattern") or "").strip()
        fr = str(pb.get("file_name_rule") or "") or "—"
        sheet = str(s.get("sheet_name") or "").strip()
        sr = str(pb.get("sheet_rule") or "") or "—"
        cref = str(s.get("cell_ref") or "").strip()
        cref_d = cref if cref else "（空＝既定）"
        labels = ["N件", "空白まで"]
        if s.get("repeat_until_empty"):
            end_s = labels[1] if len(labels) > 1 else "空白まで"
        else:
            end_s = "%s=%s" % (labels[0], s.get("repeat_max") if s.get("repeat_max") is not None else "—")
        nl = len(pb.get("link_defs") or []) if isinstance(pb.get("link_defs"), list) else 0
        nj = len(pb.get("join_defs") or []) if isinstance(pb.get("join_defs"), list) else 0
        lead = ("%s セル座標から取得" % sn_user) if sn_user else "セル座標から取得"
        bits = [
            lead,
            "ファイル(%s)%s" % (fr, fp or "全件"),
            "シート(%s)%s" % (sr, sheet or "—"),
            "セル座標%s" % cref_d,
            "終了%s" % end_s,
            "連携%d/結合%d" % (nl, nj),
        ]
        return " ".join(bits)

    def _format_item_summary(self, item: dict[str, Any]) -> str:
        """項目のシナリオ要約文字列を生成する。"""
        nm = str(item.get("name") or item.get("id") or "").strip()
        inc = self._incoming_link_join_refs_line(nm)
        parts: list[str] = []
        if inc:
            parts.append(inc)
        sources = item.get("sources") or []
        for s in sources[:6]:
            if isinstance(s, dict):
                parts.append(self._format_one_source_summary(s))
        wm = item.get("write_mode") or "fill_in"
        if wm != "fill_in" and not inc and bool(sources):
            parts.append("書込:%s" % wm)
        out = " | ".join(parts) if parts else ""
        cap = 900
        if len(out) > cap:
            out = out[: cap - 1] + "…"
        return out

    def _on_scenario_save(self) -> None:
        """シナリオ保存ダイアログで保存先を指定し、現在のシナリオを保存する。"""
        try:
            data = self._build_scenario_from_ui()
            if self._scenario_save_empty_filename:
                save_dialog_start = self._file_dialog_initial_dir()
            else:
                save_dialog_start = self._scenario_path or self._file_dialog_initial_dir()
            path, _ = QFileDialog.getSaveFileName(
                self,
                self._main_ui_disp("DIALOG_TITLE_SCENARIO_SAVE", "シナリオを保存"),
                save_dialog_start,
                self._main_ui_disp(
                    "FILE_FILTER_SCENARIO_SAVE",
                    "シナリオ (*.json *.scenario);;すべてのファイル (*.*)",
                ),
            )
            if not path:
                return
            from svc import svc_data_agg_scenario as scenario_mod

            save_errs = scenario_mod.validate_scenario(data)
            if save_errs:
                title_sv = _ui_disp_str(
                    self._ui or {}, "BTN_SCENARIO_SAVE", "シナリオ保存"
                )
                pre_sv = _ui_disp_str(
                    self._ui or {},
                    "MSG_SCENARIO_SAVE_VALIDATE_PREFIX",
                    "保存できません（検証エラー）:",
                )
                show_warning_notice(
                    self,
                    title_sv,
                    pre_sv + "\n" + "\n".join(save_errs[:8]),
                )
                return

            scenario_mod.save_scenario(path, data)
            self._scenario = data
            self._scenario_path = path
            self._scenario_save_empty_filename = False
            set_last_folder(str(Path(path).parent))
            self._clear_scenario_dirty()
            self._refresh_scenario_display_label()
            title_sv = _ui_disp_str(self._ui or {}, "BTN_SCENARIO_SAVE", "シナリオ保存")
            show_done_notice(
                self,
                title_sv,
                _ui_disp_str(
                    self._ui or {},
                    "MSG_SCENARIO_SAVE_DONE_FMT",
                    "保存しました: %s",
                )
                % path,
            )
        except Exception as exc:
            show_warning_notice(
                self,
                _ui_disp_str(self._ui or {}, "BTN_SCENARIO_SAVE", "シナリオ保存"),
                _ui_disp_str(
                    self._ui or {},
                    "MSG_SCENARIO_SAVE_FAILED_FMT",
                    "保存に失敗しました: %s",
                )
                % exc,
            )

    def _build_scenario_from_ui(self) -> dict[str, Any]:
        """UI の状態からシナリオ辞書を組み立てる。"""
        from svc import svc_data_agg_scenario as scenario_mod
        data = self._scenario.copy() if self._scenario else scenario_mod.create_empty_scenario()
        scan = dict(data.get("scan") or {})
        scan["start_path"] = self._edit_start_path.text().strip()
        scan["recursive"] = self._chk_recursive.isChecked()
        exts = []
        if self._chk_ext_xls.isChecked():
            exts.append(".xls")
        if self._chk_ext_xlsx.isChecked():
            exts.append(".xlsx")
        if self._chk_ext_xlsm.isChecked():
            exts.append(".xlsm")
        if self._chk_ext_csv.isChecked():
            exts.append(".csv")
        scan["extensions"] = exts if exts else [".xlsx", ".xlsm", ".csv"]
        scan["keyword"] = self._edit_keyword.text().strip()
        data["scan"] = scan
        data["master_path"] = ""
        data.pop("match_no_key_action", None)
        data.pop("overwrite_row_index", None)
        items: list[dict[str, Any]] = []
        existing = (data.get("items") or [])
        for i in range(self._item_table.rowCount()):
            cell0 = self._item_table.item(i, 0)
            raw_name = (cell0.text() or "").strip() if cell0 else ""
            prev = existing[i] if i < len(existing) else {}
            disp_name = raw_name or str(prev.get("name") or "").strip() or ("項目_%s" % (i + 1))
            it_ent: dict[str, Any] = {
                "id": prev.get("id") or "item_%s" % i,
                "name": disp_name,
                "sources": list(prev.get("sources") or []),
                "write_mode": prev.get("write_mode") or "fill_in",
            }
            items.append(it_ent)
        if items:
            data["items"] = items
        data[scenario_mod.KEY_EXCEL_OPTIONS] = scenario_mod.normalize_excel_options(
            self._excel_options_from_ui()
        )
        return data

    def _start_batch_done_poll_for_sheet(self, sheet_id: str, *, run_id: str = "") -> None:
        """別プロセス一括の完了 pickle をポーリングし、親ダイアログでメッセージを出す。"""
        sid = str(sheet_id or "").strip() or str(self._sheet_id or "")
        self._batch_poll_sheet_id = sid
        self._batch_poll_run_id = str(run_id or "").strip()
        delete_batch_done_notify(sid)
        self._batch_poll_deadline = time.time() + 7200.0
        if self._batch_poll_timer is None:
            self._batch_poll_timer = QTimer(self)
            self._batch_poll_timer.timeout.connect(self._on_batch_done_poll_tick)
        self._batch_poll_timer.start(400)

    def _on_batch_done_poll_tick(self) -> None:
        if time.time() > self._batch_poll_deadline:
            if self._batch_poll_timer is not None:
                self._batch_poll_timer.stop()
            return
        sid = str(getattr(self, "_batch_poll_sheet_id", "") or "").strip() or str(self._sheet_id or "")
        d = try_read_batch_done_notify(sid)
        if not d:
            return
        expect_run_id = str(getattr(self, "_batch_poll_run_id", "") or "").strip()
        got_run_id = str(d.get("run_id") or "").strip()
        # 古い run の通知は無視（先行 run の完了通知が遅れて届く競合対策）
        if expect_run_id and got_run_id and got_run_id != expect_run_id:
            delete_batch_done_notify(sid)
            return
        delete_batch_done_notify(sid)
        if self._batch_poll_timer is not None:
            self._batch_poll_timer.stop()
        title = str(d.get("title") or "データ集約")
        msg = _normalize_message_newlines(str(d.get("message") or ""))
        if d.get("ok", True):
            show_done_notice(self, title, msg)
        else:
            show_warning_notice(self, title, msg)

    def _on_batch_run(self) -> None:
        """一括実行を開始する。"""
        self._run_execution("batch_run")

    def _resolve_live_excel_target(self) -> tuple[int, str]:
        """実行時点の Excel アクティブシート情報（hwnd, sheet_id）を返す。"""
        hwnd = int(self._parent_hwnd or 0)
        sheet_id = str(self._sheet_id or "")
        if hwnd <= 0:
            return hwnd, sheet_id
        try:
            from core.core_xlc import get_excel_context_from_hwnd, get_sheet_prop

            ctx = get_excel_context_from_hwnd(hwnd, "")
            if ctx:
                _app, _book, active_sheet, live_hwnd = ctx
                hwnd = int(live_hwnd or hwnd)
                sid = str(get_sheet_prop(active_sheet, "HC_GUID_B64") or "").strip()
                if sid:
                    sheet_id = sid
        except Exception:
            pass
        return hwnd, sheet_id

    def _run_execution(self, action: str) -> None:
        """一括実行を IPC で svc に依頼する（メイン本番は一括のみ）。"""
        data = self._build_scenario_from_ui()
        items = data.get("items") or []
        if not items:
            msg = _normalize_message_newlines(
                str(self._messages.get("NO_ITEMS") or "項目が定義されていません。").strip()
            )
            show_warning_notice(self, "データ集約", msg)
            return
        if action != "batch_run":
            show_warning_notice(self, "データ集約", "未対応の実行種別です。")
            return
        # 永続パスは「シナリオ保存」のみ。一括は UI スナップショットを一時 JSON に書き子へ渡す。
        scenario_path_persistent = (self._scenario_path or "").strip()
        try:
            from svc import svc_data_agg_scenario as scenario_mod

            pre_errs = scenario_mod.validate_scenario(data)
            if pre_errs:
                show_warning_notice(
                    self,
                    "データ集約",
                    "一括実行できません（検証エラー）:\n" + "\n".join(pre_errs[:8]),
                )
                return
        except Exception as exc:
            show_warning_notice(self, "データ集約", "検証に失敗しました: %s" % exc)
            return
        show_batch_start = bool(self._ui.get("SHOW_BATCH_START_MESSAGE", False))
        notify_parent_dialog = bool(self._ui.get("BATCH_NOTIFY_PARENT_DIALOG", True))
        try:
            excel_opts_runtime = self._excel_options_from_ui()
            if str(excel_opts_runtime.get("output_target") or "active_sheet") == "active_sheet":
                run_parent_hwnd, run_sheet_id = self._resolve_live_excel_target()
            else:
                run_parent_hwnd, run_sheet_id = int(self._parent_hwnd or 0), str(self._sheet_id or "")
            proj_root = Path(__file__).resolve().parents[1]
            from core import runtime_layout

            install_root = runtime_layout.install_root()
            short_runner = runtime_layout.packaged_app_exe("hc_xlwings_short_runner.exe")
            if short_runner is None:
                try:
                    exe = Path(sys.executable).resolve()
                    if exe.name.lower() in ("ui_server.exe", "hc_ui_server.exe"):
                        cand = exe.parent / "hc_xlwings_short_runner.exe"
                        if cand.is_file():
                            short_runner = cand
                            if install_root is None:
                                ir = exe.parent.parent
                                if (ir / "app" / "bin" / "hc_main.exe").is_file():
                                    install_root = ir
                except Exception:
                    pass
            if short_runner is not None:
                py_exe = str(short_runner)
                use_short_runner = True
            else:
                py_exe = sys.executable
                use_short_runner = False
            fd_snap, snapshot_path = tempfile.mkstemp(
                suffix=".json", prefix="data_agg_scenario_", text=False
            )
            try:
                os.write(
                    fd_snap,
                    json.dumps(data, ensure_ascii=False).encode("utf-8"),
                )
            finally:
                os.close(fd_snap)
            payload = {
                "action": "batch_compute",
                "scenario_path": scenario_path_persistent,
                "scenario_snapshot_path": snapshot_path,
                "notify_parent_dialog": notify_parent_dialog,
            }
            batch_run_id = "%s_%s_%s" % (
                int(time.time() * 1000),
                os.getpid(),
                uuid.uuid4().hex[:8],
            )
            payload["batch_run_id"] = batch_run_id
            fd, payload_path = tempfile.mkstemp(suffix=".json", prefix="data_agg_payload_")
            try:
                os.write(fd, json.dumps(payload, ensure_ascii=False).encode("utf-8"))
            finally:
                os.close(fd)
            py_src = (
                "import json\n"
                "import sys\n"
                "from pathlib import Path\n"
                "pp = Path(%r)\n"
                "p = json.loads(pp.read_text(encoding='utf-8'))\n"
                "pp.unlink(missing_ok=True)\n"
                "sys.path.insert(0, %r)\n"
                "from svc.data_agg_batch_compute import run_batch_compute\n"
                "run_batch_compute(parent_hwnd=%s, sheet_id=%r, payload=p)\n"
            ) % (
                str(Path(payload_path)),
                str(proj_root),
                int(run_parent_hwnd),
                str(run_sheet_id),
            )
            fd_py, script_path = tempfile.mkstemp(
                suffix=".py", prefix="data_agg_batch_", text=False
            )
            try:
                os.write(fd_py, ("import sys\n" + py_src).encode("utf-8"))
            finally:
                os.close(fd_py)
            if use_short_runner:
                cmd = [
                    py_exe,
                    "--script-file",
                    str(Path(script_path)),
                ]
            else:
                cmd = [py_exe, "-c", py_src.replace("\n", ";")]
                try:
                    # -c 実行時は script_path を参照しないため即時掃除する。
                    Path(script_path).unlink(missing_ok=True)
                except Exception:
                    pass
            env = os.environ.copy()
            ipc_root = core_env.ipc_dir_raw()
            if ipc_root:
                env["HC_IPC_ROOT"] = ipc_root
                try:
                    from svc.data_agg_cancel import clear_batch_cancel_tombstone  # noqa: WPS433

                    clear_batch_cancel_tombstone(str(run_sheet_id), Path(ipc_root))
                    write_pickle(
                        _batch_active_path(str(run_sheet_id), Path(ipc_root)),
                        {
                            "run_id": str(batch_run_id),
                            "sheet_id": str(run_sheet_id),
                            "scenario_snapshot_path": str(snapshot_path),
                            "ts_ms": int(time.time() * 1000),
                        },
                    )
                except Exception:
                    pass
            if install_root is not None:
                env["HC_INSTALL_ROOT"] = str(install_root)
            env["PYTHONPATH"] = str(proj_root) + (os.pathsep + env.get("PYTHONPATH", ""))
            if use_short_runner and install_root is not None:
                env = runtime_layout.env_with_packaged_dll_search_path(env, install_root)
            if notify_parent_dialog:
                self._start_batch_done_poll_for_sheet(
                    run_sheet_id,
                    run_id=batch_run_id,
                )
            spawn_cwd = str(install_root) if install_root is not None else str(proj_root)
            subprocess.Popen(cmd, cwd=spawn_cwd, env=env)
            if show_batch_start:
                show_info_notice(
                    self,
                    "データ集約",
                    "%s を開始しました。" % "一括実行",
                )
        except Exception as exc:
            show_warning_notice(self, "データ集約", "実行の開始に失敗しました: %s" % exc)

    def showEvent(self, event: QShowEvent) -> None:
        super().showEvent(event)
        want_hwnd_lock = True
        try:
            from ui_qt.ui_common import want_excel_child_hwnd_lock_while_modal

            want_hwnd_lock = want_excel_child_hwnd_lock_while_modal(self._window_cfg or {})
        except Exception:
            want_hwnd_lock = True
        if not want_hwnd_lock:
            try:
                self._schedule_excel_unlock_pulse_chain()
            except Exception:
                pass
        self._schedule_deferred_excel_owner_front()
        lock_on_show = bool(
            (self._window_cfg or {}).get("EXCEL_MENU_BAR_LOCK_ON_SHOW", False)
        ) and bool(want_hwnd_lock)
        if lock_on_show and not self._excel_menu_bar_lock_applied:
            if self._apply_excel_menu_bar_lock(True):
                self._excel_menu_bar_lock_applied = True
            else:

                def _retry_menu_lock() -> None:
                    if self._excel_menu_bar_lock_applied or not self.isVisible():
                        return
                    if not lock_on_show:
                        return
                    if self._apply_excel_menu_bar_lock(True):
                        self._excel_menu_bar_lock_applied = True

                for _delay_ms in (150, 450, 1000):
                    QTimer.singleShot(_delay_ms, _retry_menu_lock)
        if getattr(self, "_scan_pending_auto", False):
            self._scan_pending_auto = False
            self._request_folder_scan(auto_mode=True)

    def resizeEvent(self, event: QResizeEvent) -> None:
        super().resizeEvent(event)
        QTimer.singleShot(0, self._fit_item_table_columns)

    def _teardown_before_hide_main(self) -> None:
        """× の closeEvent と 閉じるの reject の両方から呼ぶ。

        モードレス QDialog は reject が close ではなく hide になり closeEvent を通らないため、
        閉じるボタンだけロック解除が抜けていた。
        """
        try:
            _log_data_agg_main_lifecycle(self, "teardown_enter")
        except Exception:
            pass
        self._apply_excel_menu_bar_lock(False)
        self._excel_menu_bar_lock_applied = False
        self._excel_unlock_pulse_chain_scheduled = False
        self._excel_create_probe_t0 = 0.0
        self._stop_scan_thread()
        try:
            if self._batch_poll_timer is not None:
                self._batch_poll_timer.stop()
        except Exception:
            pass
        # モードレスは reject が hide のみで destroyed が来ないため、共有状態の teardown を明示する。
        try:
            from ui_qt.ui_common import teardown_feature_ui_shared_state

            teardown_feature_ui_shared_state(
                parent_hwnd=int(self._parent_hwnd or 0),
                modeless_widget=self,
                excel_unlock=False,
            )
            try:
                _log_data_agg_main_lifecycle(self, "teardown_feature_ui_shared_state")
            except Exception:
                pass
        except Exception:
            pass
        try:
            _log_data_agg_main_lifecycle(self, "teardown_exit")
        except Exception:
            pass

    def reject(self) -> None:
        try:
            _log_data_agg_main_lifecycle(self, "reject_enter")
        except Exception:
            pass
        self._teardown_before_hide_main()
        super().reject()
        try:
            _log_data_agg_main_lifecycle(self, "reject_after_super")
        except Exception:
            pass

    def closeEvent(self, event: QCloseEvent) -> None:
        """リボン抑止の解除・タイマ停止・モデルレス一覧からの除去。"""
        acc0 = True
        try:
            acc0 = bool(event.isAccepted())
        except Exception:
            pass
        try:
            _log_data_agg_main_lifecycle(self, "close_event_enter", "accepted=%s" % acc0)
        except Exception:
            pass
        self._teardown_before_hide_main()
        super().closeEvent(event)
        acc1 = True
        try:
            acc1 = bool(event.isAccepted())
        except Exception:
            pass
        try:
            _log_data_agg_main_lifecycle(self, "close_event_after_super", "accepted=%s" % acc1)
        except Exception:
            pass

    def hideEvent(self, event: QHideEvent) -> None:
        try:
            _log_data_agg_main_lifecycle(self, "hide_event_enter")
        except Exception:
            pass
        super().hideEvent(event)
        try:
            _log_data_agg_main_lifecycle(self, "hide_event_after_super")
        except Exception:
            pass


class _ScenarioEditDialog(QDialog):
    """
    シナリオ編集ダイアログ。項目単位で取得ソースを設定する（結合キーは詳細ペイン §4）。
    登録時に項目の write_mode（書込みモード）を現在のソース詳細のコンボから同期する。
    シナリオの match_keys はこのダイアログでは変更しない。
    §2.3 データ抽出（セル座標・名前から抽出）、§2.4 データ書き込み・統合に対応。
    """

    @staticmethod
    def _path_item_text_is_legacy_placeholder(s: str) -> bool:
        """名前取得 path_item 未指定・旧プレースホルダ文言。"""
        t = (s or "").strip()
        if not t:
            return True
        if "主キー" in t or "先頭" in t or "項目一覧" in t:
            return True
        return False

    def _scenario_set_tip(self, w: QWidget | None, key: str, default: str = "") -> None:
        if w is None:
            return
        t = _ui_disp_str(self._screen_cfg, key, default).strip()
        if t:
            w.setToolTip(t)

    @staticmethod
    def _combo_select_saved_master_item(cb: QComboBox, itxt: str) -> None:
        """保存済みマスタ項目名をコンボに反映。候補に無ければ追加して選択（登録時に落ちないように）。"""
        t = (itxt or "").strip()
        if not t:
            return
        fl = Qt.MatchFlag.MatchExactly | Qt.MatchFlag.MatchCaseSensitive
        ix = cb.findText(t, fl)
        if ix < 0:
            cb.addItem(t)
            ix = cb.findText(t, fl)
        if ix >= 0:
            cb.setCurrentIndex(ix)

    def __init__(
        self,
        item_name: str,
        item_id: str,
        item_data: dict[str, Any],
        parent: QWidget | None = None,
        screen_cfg: dict[str, Any] | None = None,
        items: list[dict[str, Any]] | None = None,
        scan_paths_hint: list[str] | None = None,
        on_registered: Any = None,
    ) -> None:
        super().__init__(parent)
        try:
            self.setAttribute(Qt.WidgetAttribute.WA_AlwaysShowToolTips, True)
        except Exception:
            pass
        self._screen_cfg = screen_cfg or {}
        try:
            from ui_qt.ui_common import _deep_merge, apply_window_config

            root_win = (_get_cfg().get("WINDOW") or {})
            scen_win = (self._screen_cfg.get("WINDOW") or {})
            win_merged = _deep_merge(dict(root_win), dict(scen_win))
            # Child dialogs should stay on the parent window, not recenter on Excel.
            win_merged["CENTER_ON_EXCEL"] = False
            ph = 0
            pw: QWidget | None = parent
            while pw is not None:
                if hasattr(pw, "_parent_hwnd"):
                    ph = int(getattr(pw, "_parent_hwnd", 0) or 0)
                    break
                pw = pw.parentWidget()
            apply_window_config(self, {"WINDOW": win_merged}, ph, "SCENARIO_EDIT")
        except Exception:
            pass
        self._scan_paths_hint: list[str] = list(scan_paths_hint or [])
        self._item_id = str(item_id or "")
        self._item_name = str(item_name or "項目").strip() or "項目"
        self._item_write_mode_hint = (
            str((item_data or {}).get("write_mode") or "fill_in").strip() or "fill_in"
        )
        self._sources_data: list[dict[str, Any]] = []
        self._registered_display_snapshots: list[dict[str, Any] | None] = []
        self._current_source_index = -1
        self._loading_source_form: bool = False
        self._on_registered = on_registered if callable(on_registered) else None
        self._dirty: bool = False
        self._undo_snapshot: list[dict[str, Any]] | None = None
        self._undo_restore_row: int = -1
        self._master_items_list: list[dict[str, Any]] = list(items or [])
        self._master_item_row = -1
        for _mi, _it in enumerate(self._master_items_list):
            if _it is item_data:
                self._master_item_row = _mi
                break
        if self._master_item_row < 0 and self._item_id:
            for _mi, _it in enumerate(self._master_items_list):
                if str((_it or {}).get("id") or "").strip() == self._item_id:
                    self._master_item_row = _mi
                    break
        _u = lambda k, d: _ui_disp_str(self._screen_cfg, k, d)
        self.setWindowTitle(_u("TITLE", "シナリオ編集"))
        mw = int(self._screen_cfg.get("DIALOG_MIN_WIDTH") or 700)
        mh = int(self._screen_cfg.get("DIALOG_MIN_HEIGHT") or 520)
        # ダイアログの「下限」。ただし子レイアウトの最小サイズの方が大きいと実効最小幅はそちらになる
        # （詳細ペインは ui_data_agg_scenario_layout の QScrollArea で横スクロール可とし、JSON の MIN が効きやすくする）。
        if mw > 0:
            self.setMinimumWidth(mw)
        if mh > 0:
            self.setMinimumHeight(mh)
        root_layout = QVBoxLayout(self)
        splitter = QSplitter(Qt.Orientation.Horizontal)

        left_wrap = QWidget()
        left_wrap.setMinimumWidth(0)
        left_lay = QVBoxLayout(left_wrap)
        left_lay.setContentsMargins(0, 0, 0, 0)
        left_lay.setSpacing(4)
        self._left_splitter = QSplitter(Qt.Orientation.Vertical)
        self._left_splitter.setChildrenCollapsible(False)
        left_top = QWidget()
        left_top.setMinimumHeight(0)
        left_top_lay = QVBoxLayout(left_top)
        left_top_lay.setContentsMargins(0, 0, 0, 0)
        left_top_lay.setSpacing(4)
        lbl_sc_list = QLabel(
            "<b>%s</b>" % _u("LABEL_SCENARIO_LIST", "シナリオ一覧").replace("\n", "<br/>")
        )
        lbl_sc_list.setToolTip(
            _u(
                "TIP_LABEL_SCENARIO_LIST",
                "このマスタ項目に紐づく取得シナリオの一覧です。上から順に評価されます。",
            )
        )
        left_top_lay.addWidget(lbl_sc_list)
        self._sources_table = QTableWidget()
        self._sources_table.setColumnCount(2)
        self._sources_table.verticalHeader().setVisible(False)
        hi_idx = QTableWidgetItem(_u("TABLE_HEADER_INDEX", "#"))
        hi_idx.setToolTip(
            _u("TIP_TABLE_HEADER_INDEX", "シナリオの評価順（行番号）です。")
        )
        hi_name = QTableWidgetItem(_u("TABLE_HEADER_SCENARIO_NAME", "シナリオ名"))
        hi_name.setToolTip(
            _u(
                "TIP_TABLE_HEADER_SCENARIO_NAME",
                "シナリオの表示名です。ツールチップで要約の一部を確認できます。",
            )
        )
        self._sources_table.setHorizontalHeaderItem(0, hi_idx)
        self._sources_table.setHorizontalHeaderItem(1, hi_name)
        self._sources_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        self._sources_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self._sources_table.setColumnWidth(0, 28)
        self._sources_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self._sources_table.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self._sources_table.setAlternatingRowColors(True)
        self._sources_table.setStyleSheet(
            "QTableWidget { alternate-background-color: #FAFAF5; } "
            "QTableWidget::item:selected { background-color: #B0C4DE; } "
            "QTableWidget::item:selected:!active { background-color: #C8D4E0; }"
        )
        self._sources_table.setMinimumHeight(120)
        self._sources_table.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self._scenario_set_tip(
            self._sources_table,
            "TIP_SCENARIO_SOURCES_TABLE",
            "この項目に紐づく取得シナリオの一覧です。行を選ぶと右の詳細が切り替わります。",
        )
        self._sources_table.itemSelectionChanged.connect(self._on_source_selection_changed)
        self._sources_table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self._sources_table.customContextMenuRequested.connect(
            self._on_sources_table_context_menu
        )
        left_top_lay.addWidget(self._sources_table, 1)
        self._scenario_set_tip(
            left_top,
            "TIP_SCENARIO_LEFT_TOP",
            "シナリオ一覧テーブルのエリアです。",
        )
        left_bottom = QWidget()
        left_bottom.setMinimumHeight(0)
        left_bottom_lay = QVBoxLayout(left_bottom)
        left_bottom_lay.setContentsMargins(0, 0, 0, 0)
        left_bottom_lay.setSpacing(2)
        self._summary_preview = QLabel()
        self._summary_preview.setWordWrap(True)
        self._summary_preview.setTextFormat(Qt.TextFormat.PlainText)
        self._summary_preview.setAlignment(
            Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop
        )
        self._summary_preview.setStyleSheet(
            "font-size: 11px; color: #333; padding: 4px; background: #fafafa; "
            "border: 1px solid #ddd; border-radius: 3px;"
        )
        self._summary_preview.setMinimumHeight(48)
        self._summary_preview.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding
        )
        left_bottom_lay.addWidget(self._summary_preview, 1)
        self._scenario_set_tip(
            self._summary_preview,
            "TIP_SCENARIO_SUMMARY_PREVIEW",
            "選択中シナリオの設定要約（プレーンテキスト）です。",
        )

        def _src_btn_min_w(btn: QPushButton) -> None:
            btn.setMinimumSize(0, 0)
            btn.setStyleSheet(
                "QPushButton { padding: 1px 4px; min-width: 0px; }"
            )
            btn.adjustSize()
            tw = btn.fontMetrics().horizontalAdvance(btn.text()) + 8
            mh = btn.minimumSizeHint().width()
            cap = tw + 6
            if mh > cap:
                w = max(20, cap)
            else:
                w = max(20, tw, mh)
            btn.setFixedWidth(w)

        btn_src_up = QPushButton(_u("BTN_SOURCE_UP", "▲"))
        btn_src_up.setToolTip(_u("TIP_SOURCE_UP", "選択中のソースを上へ（取得の評価順）"))
        btn_src_up.clicked.connect(self._on_source_move_up)
        btn_src_up.setAutoDefault(False)
        btn_src_up.setDefault(False)
        _src_btn_min_w(btn_src_up)
        btn_src_dn = QPushButton(_u("BTN_SOURCE_DOWN", "▼"))
        btn_src_dn.setToolTip(_u("TIP_SOURCE_DOWN", "選択中のソースを下へ（取得の評価順）"))
        btn_src_dn.clicked.connect(self._on_source_move_down)
        btn_src_dn.setAutoDefault(False)
        btn_src_dn.setDefault(False)
        _src_btn_min_w(btn_src_dn)
        btn_add = QPushButton(_u("BTN_ADD_SOURCE", "追加"))
        btn_add.setToolTip(
            _u("TIP_ADD_SOURCE", "一覧の末尾に新しいシナリオを追加します。")
        )
        btn_add.clicked.connect(self._on_add_source)
        btn_add.setAutoDefault(False)
        btn_add.setDefault(False)
        _src_btn_min_w(btn_add)
        btn_dup = QPushButton(_u("BTN_DUPLICATE_SOURCE", "複製"))
        btn_dup.setToolTip(_u("TIP_DUPLICATE_SOURCE", "選択中のシナリオをコピーして一覧に追加します。"))
        btn_dup.clicked.connect(self._on_duplicate_source)
        btn_dup.setAutoDefault(False)
        btn_dup.setDefault(False)
        _src_btn_min_w(btn_dup)
        btn_remove = QPushButton(_u("BTN_REMOVE_SOURCE", "削除"))
        btn_remove.setToolTip(
            _u("TIP_REMOVE_SOURCE", "選択中のシナリオを一覧から削除します。")
        )
        btn_remove.clicked.connect(self._on_remove_source)
        btn_remove.setAutoDefault(False)
        btn_remove.setDefault(False)
        _src_btn_min_w(btn_remove)
        self._btn_undo_remove = QPushButton(_u("BTN_UNDO_REMOVE_SOURCE", "Undo"))
        self._btn_undo_remove.setToolTip(
            _u("TIP_UNDO_REMOVE_SOURCE", "直近の削除を元に戻します（利用可能なときのみ）。")
        )
        self._btn_undo_remove.setEnabled(False)
        self._btn_undo_remove.setAutoDefault(False)
        self._btn_undo_remove.setDefault(False)
        self._btn_undo_remove.clicked.connect(self._on_undo_scenario)
        _src_btn_min_w(self._btn_undo_remove)
        fr_ops = QFrame()
        fr_ops.setFrameShape(QFrame.Shape.StyledPanel)
        fr_ops.setSizePolicy(
            QSizePolicy.Policy.Maximum,
            QSizePolicy.Policy.Fixed,
        )
        fr_ops_l = QVBoxLayout(fr_ops)
        fr_ops_l.setContentsMargins(4, 2, 4, 2)
        fr_ops_l.setSpacing(0)
        fr_ops_l.addStretch(1)
        row_ops = QHBoxLayout()
        row_ops.setSpacing(3)
        row_ops.setContentsMargins(0, 0, 0, 0)
        row_ops.addWidget(btn_src_up)
        row_ops.addWidget(btn_src_dn)
        row_ops.addWidget(btn_add)
        row_ops.addWidget(btn_dup)
        row_ops.addWidget(btn_remove)
        row_ops.addWidget(self._btn_undo_remove)
        fr_ops_l.addLayout(row_ops)
        self._scenario_set_tip(
            fr_ops,
            "TIP_SCENARIO_OPS_FRAME",
            "シナリオ行の追加・複製・削除・順序変更・Undo をまとめた操作欄です。",
        )
        left_bottom_lay.addWidget(
            fr_ops,
            0,
            Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignBottom,
        )
        row_dbg = QHBoxLayout()
        row_dbg.addStretch(1)
        self._btn_step = QPushButton(_u("BTN_STEP_SCENARIO", "デバッグ"))
        self._btn_step.setToolTip(
            _u(
                "TIP_STEP_SCENARIO",
                "デバッグを開きます。シナリオフェーズで実ファイル抽出のプレビュー（マスタへは書込みません）。親画面の基準フォルダで再スキャンしたパスを優先します。",
            )
        )
        self._btn_step.clicked.connect(self._on_step_scenario_placeholder)
        self._btn_step.setAutoDefault(False)
        self._btn_step.setDefault(False)
        self._btn_step.setFixedWidth(max(22, self._btn_step.sizeHint().width()))
        row_dbg.addWidget(self._btn_step)
        left_bottom_lay.addLayout(row_dbg)
        self._left_splitter.addWidget(left_top)
        self._left_splitter.addWidget(left_bottom)
        self._scenario_set_tip(
            left_bottom,
            "TIP_SCENARIO_LEFT_BOTTOM",
            "要約プレビューとシナリオ操作ボタンのエリアです。",
        )
        self._left_splitter.setStretchFactor(0, 1)
        self._left_splitter.setStretchFactor(1, 0)
        left_lay.addWidget(self._left_splitter, 1)
        self._scenario_set_tip(
            left_wrap,
            "TIP_SCENARIO_LEFT_WRAP",
            "シナリオ一覧・要約・操作ボタンをまとめた左ペインです。",
        )
        self._scenario_set_tip(
            self._left_splitter,
            "TIP_SCENARIO_LEFT_SPLITTER",
            "上段のシナリオ表と下段の要約・ボタンエリアの高さを調整します。",
        )
        splitter.addWidget(left_wrap)

        right_wrap = QWidget()
        right_wrap.setMinimumWidth(0)
        right_wrap.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding
        )
        right = QVBoxLayout(right_wrap)
        right.setContentsMargins(0, 0, 0, 0)
        right.setSpacing(4)
        lbl_item_hdr = QLabel(
            "<b>%s：%s</b>"
            % (
                _u("LABEL_ITEM_NAME", "項目名").replace("\n", "<br/>"),
                (item_name or "-").replace("\n", "<br/>"),
            )
        )
        lbl_item_hdr.setToolTip(
            _u("TIP_LABEL_ITEM_NAME", "マスタ上のこの列（項目）の名称です。")
        )
        right.addWidget(lbl_item_hdr)
        row_ident = QHBoxLayout()
        lbl_ident = QLabel(_u("LABEL_SCENARIO_NAME", "シナリオ名") + "：")
        lbl_ident.setToolTip(
            _u("TIP_LABEL_SCENARIO_NAME", "一覧で識別するシナリオ名です。")
        )
        row_ident.addWidget(lbl_ident)
        self._edit_scenario_ident = QLineEdit()
        self._scenario_name_default_style = "color: #888;"
        self._edit_scenario_ident.setStyleSheet(self._scenario_name_default_style)
        self._edit_scenario_ident.textChanged.connect(self._on_scenario_name_text_changed)
        self._edit_scenario_ident.returnPressed.connect(lambda: self._edit_scenario_ident.clearFocus())
        row_ident.addWidget(self._edit_scenario_ident, 1)
        self._edit_scenario_ident.setToolTip(
            _u("TIP_SCENARIO_NAME_EDIT", "空のときは自動生成名が一覧に表示されます。")
        )
        right.addLayout(row_ident)
        kind_row_top = QHBoxLayout()
        lbl_kind = QLabel(_u("LABEL_KIND", "種別") + "：")
        lbl_kind.setToolTip(
            _u(
                "TIP_LABEL_KIND",
                "セル座標から取得か名前から取得か。同一項目内では混在できません。",
            )
        )
        kind_row_top.addWidget(lbl_kind)
        self._form_combo_type = QComboBox()
        self._form_combo_type.addItem(_u("SOURCE_TYPE_CELL", "セル座標から取得"), "cell")
        self._form_combo_type.addItem(_u("SOURCE_TYPE_NAME_EXTRACT", "名前から取得"), "name_extract")
        self._form_combo_type.currentIndexChanged.connect(self._on_form_type_changed)
        self._form_combo_type.setMaximumWidth(280)
        self._form_combo_type.setToolTip(
            _u("TIP_FORM_COMBO_TYPE", "取得元の種別です。変更すると詳細フォームが切り替わります。")
        )
        kind_row_top.addWidget(self._form_combo_type)
        kind_row_top.addStretch(1)
        right.addLayout(kind_row_top)
        self._form_stack = QStackedWidget()
        self._form_stack.setMinimumWidth(0)
        self._form_stack.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding
        )
        detail_cell = dict(self._screen_cfg.get("DETAIL_CELL") or {})
        detail_name = dict(self._screen_cfg.get("DETAIL_NAME") or {})
        _dlg_title = (
            _ui_disp_str(self._screen_cfg, "MSGBOX_TITLE", "").strip()
            or _u("TITLE", "シナリオ編集")
        )
        detail_cell.setdefault("MSGBOX_TITLE", _dlg_title)
        detail_name.setdefault("MSGBOX_TITLE", _dlg_title)
        _sw = self._screen_cfg.get("DETAIL_SCROLL_CONTENT_MIN_WIDTH")
        if _sw is not None:
            detail_cell.setdefault("DETAIL_SCROLL_CONTENT_MIN_WIDTH", _sw)
            detail_name.setdefault("DETAIL_SCROLL_CONTENT_MIN_WIDTH", _sw)
        scroll_cell, self._cell_refs = build_scenario_detail_cell_scroll(
            self._item_name, items or [], detail_cell
        )
        self._cell_refs["on_join_group_added"] = self._wire_new_join_def
        self._cell_refs["on_link_group_added"] = self._wire_new_link_def
        self._cell_refs["on_link_group_removed"] = self._on_link_or_join_removed
        self._cell_refs["on_join_group_removed"] = self._on_link_or_join_removed
        scroll_name, self._name_refs = build_scenario_detail_name_scroll(
            self._item_name, items or [], detail_name
        )
        self._detail_scroll_cell = scroll_cell
        self._detail_scroll_name = scroll_name
        self._form_stack.addWidget(scroll_cell)
        self._form_stack.addWidget(scroll_name)
        self._wire_detail_form_signals()
        right.addWidget(self._form_stack, 1)
        self._scenario_set_tip(
            self._form_stack,
            "TIP_SCENARIO_FORM_STACK",
            "種別に応じて「セル座標から取得」または「名前から取得」のフォームを切り替えます。",
        )
        splitter.addWidget(right_wrap)
        self._scenario_set_tip(
            right_wrap,
            "TIP_SCENARIO_RIGHT_WRAP",
            "項目名・種別・取得詳細フォームをまとめた右ペインです。",
        )
        sz = self._screen_cfg.get("SPLITTER_SIZES")
        if isinstance(sz, list) and len(sz) >= 2:
            try:
                splitter.setSizes([int(sz[0]), int(sz[1])])
            except (TypeError, ValueError):
                splitter.setSizes([320, 340])
        else:
            splitter.setSizes([320, 340])
        st0 = int(self._screen_cfg.get("SPLITTER_STRETCH_LEFT") or 1)
        st1 = int(self._screen_cfg.get("SPLITTER_STRETCH_RIGHT") or 1)
        splitter.setStretchFactor(0, max(st0, 0))
        splitter.setStretchFactor(1, max(st1, 0))
        self._scenario_splitter = splitter
        self._scenario_set_tip(
            splitter,
            "TIP_SCENARIO_SPLITTER",
            "左のシナリオ一覧・要約と右の詳細フォームの幅を調整します。",
        )
        self._scenario_right_wrap = right_wrap
        rpref = self._screen_cfg.get("RIGHT_PANE_PREF_WIDTH")
        try:
            if rpref is not None and int(rpref) > 0:
                right_wrap.setMinimumWidth(int(rpref))
        except (TypeError, ValueError):
            pass
        rmax = self._screen_cfg.get("RIGHT_PANE_MAX_WIDTH")
        try:
            if rmax is not None and int(rmax) > 0:
                right_wrap.setMaximumWidth(int(rmax))
        except (TypeError, ValueError):
            pass
        splitter.splitterMoved.connect(self._on_scenario_splitter_moved)
        root_layout.addWidget(splitter, 1)
        row_btn = QHBoxLayout()
        row_btn.addStretch(1)
        self._btn_register = QPushButton(_u("BTN_OK", "登録"))
        self._btn_register.setToolTip(
            _u("TIP_BTN_OK", "変更を項目に反映してダイアログを閉じます。")
        )
        self._btn_register.clicked.connect(self._on_register_clicked)
        self._btn_register.setAutoDefault(False)
        self._btn_register.setDefault(False)
        btn_cancel = QPushButton(_u("BTN_CANCEL", "キャンセル"))
        btn_cancel.setToolTip(
            _u("TIP_BTN_CANCEL", "変更を破棄して閉じます。")
        )
        btn_cancel.clicked.connect(self.reject)
        btn_cancel.setAutoDefault(False)
        btn_cancel.setDefault(False)
        row_btn.addWidget(self._btn_register)
        row_btn.addWidget(btn_cancel)
        root_layout.addLayout(row_btn)
        # 初期表示サイズ: MIN のみでは QLayout.sizeHint が大きいと開いたときの枠が変わらないことがある。
        # DIALOG_DEFAULT_* があれば優先、なければ DIALOG_MIN_* で resize する。
        dw = self._screen_cfg.get("DIALOG_DEFAULT_WIDTH")
        dh = self._screen_cfg.get("DIALOG_DEFAULT_HEIGHT")
        iw = int(dw) if dw is not None else mw
        ih = int(dh) if dh is not None else mh
        if iw > 0 and ih > 0:
            self.resize(iw, ih)
        # 初期データ投入（浅い dict() だと ui_scenario_source_v1 等がソース間で共有され、
        # apply / 入れ替えで A・B の設定が混線する。編集セッションは deepcopy で独立させる。）
        for src in (item_data or {}).get("sources") or []:
            if isinstance(src, dict):
                one = copy.deepcopy(src)
                one.setdefault("registered", True)
                self._sources_data.append(one)
                self._registered_display_snapshots.append(
                    copy.deepcopy(one) if bool(one.get("registered")) else None
                )
        self._refresh_sources_table()
        if self._sources_data:
            self._sources_table.selectRow(0)
            self._current_source_index = 0
            try:
                self._load_source_to_form(0)
            except Exception as exc:
                try:
                    _data_agg_ui_diag.info(
                        "[DATA_AGG_SCENARIO_EDIT] dialog_open load_source_to_form exc row=0 "
                        "item_id=%s err=%s",
                        self._item_id,
                        exc,
                    )
                    _data_agg_ui_diag.info(
                        "[DATA_AGG_SCENARIO_EDIT] dialog_open load traceback\n%s",
                        traceback.format_exc(),
                    )
                except Exception:
                    pass
            _log_scenario_edit_diag(
                "dialog_open n_sources=%s item_id=%s",
                len(self._sources_data),
                self._item_id,
            )
        else:
            self._form_stack.setEnabled(False)
            self._form_combo_type.setEnabled(False)
            self._edit_scenario_ident.clear()
            self._edit_scenario_ident.setEnabled(False)
            _log_scenario_edit_diag("dialog_open empty_sources item_id=%s", self._item_id)
        self._update_step_button_enabled()
        self._update_register_button_state()

    def showEvent(self, event) -> None:
        super().showEvent(event)
        self._center_on_parent_widget()
        ph, _rect = _data_agg_excel_parent_hwnd_rect(self)
        if ph:
            try:
                from ui_qt.ui_common import ensure_front, _set_owner_hwnd

                def _owner() -> None:
                    try:
                        _set_owner_hwnd(self, ph)
                    except Exception:
                        pass

                def _front() -> None:
                    try:
                        ensure_front(self, ph)
                    except Exception:
                        pass

                QTimer.singleShot(0, _owner)
                QTimer.singleShot(50, _owner)
                QTimer.singleShot(120, _front)
            except Exception:
                pass
        QTimer.singleShot(0, self._center_on_parent_widget)
        QTimer.singleShot(160, self._center_on_parent_widget)
        QTimer.singleShot(0, self._sync_left_splitter_sizes)
        QTimer.singleShot(80, self._sync_left_splitter_sizes)
        QTimer.singleShot(0, self._clear_initial_button_focus)

    def _center_on_parent_widget(self) -> None:
        pw = self.parentWidget()
        if pw is None:
            return
        pr = pw.frameGeometry()
        gr = self.frameGeometry()
        x = pr.x() + (pr.width() - gr.width()) // 2
        y = pr.y() + (pr.height() - gr.height()) // 2
        self.move(x, y)

    def _sync_left_splitter_sizes(self) -> None:
        sp = getattr(self, "_left_splitter", None)
        if sp is None:
            return
        h = sp.height()
        if h < 100:
            return
        ratio = float(self._screen_cfg.get("LEFT_SPLITTER_TOP_RATIO", 0.5) or 0.5)
        ratio = max(0.35, min(0.72, ratio))
        top = int(h * ratio)
        bottom = max(h - top, 96)
        sp.setSizes([top, bottom])

    def _clear_initial_button_focus(self) -> None:
        w = self.focusWidget()
        if isinstance(w, QAbstractButton):
            w.clearFocus()

    def _update_register_button_state(self) -> None:
        has = bool(self._sources_data) and self._current_source_index >= 0
        self._btn_register.setEnabled(self._dirty and has)

    def _on_scenario_name_text_changed(self, text: str) -> None:
        if (text or "").strip():
            self._edit_scenario_ident.setStyleSheet("")
        else:
            self._edit_scenario_ident.setStyleSheet(self._scenario_name_default_style)
        if not self._edit_scenario_ident.signalsBlocked():
            self._dirty = True
            self._update_register_button_state()
            self._notify_main_scenario_dirty()

    @staticmethod
    def _resolve_auto_scenario_display_names_for_sources(
        item_name: str, rows: list[dict[str, Any]]
    ) -> list[str]:
        """
        各行の一覧表示用シナリオ名。scenario_name が空の行には、既に使われている名前と重ならない
        「項目名_シナリオN」を上から順に割り当てる（行番号ベースではない）。
        """
        iname = (item_name or "").strip() or "項目"
        n = len(rows)
        out: list[str] = [""] * n
        occupied: set[str] = set()
        for i in range(n):
            sn = str(rows[i].get("scenario_name") or "").strip()
            if sn:
                out[i] = sn
                occupied.add(sn)
        for i in range(n):
            if out[i]:
                continue
            k = 1
            while True:
                cand = "%s_シナリオ%d" % (iname, k)
                k += 1
                if cand not in occupied:
                    out[i] = cand
                    occupied.add(cand)
                    break
        return out

    def _resolve_auto_scenario_display_names(self) -> list[str]:
        return _ScenarioEditDialog._resolve_auto_scenario_display_names_for_sources(
            self._item_name, self._sources_data
        )

    def _default_scenario_name(self, idx: int) -> str:
        names = self._resolve_auto_scenario_display_names()
        if 0 <= idx < len(names):
            return names[idx]
        return "%s_シナリオ%d" % (self._item_name, max(0, idx) + 1)

    def _effective_scenario_name_at_in_list(self, idx: int, data: list[dict[str, Any]]) -> str:
        names = _ScenarioEditDialog._resolve_auto_scenario_display_names_for_sources(
            self._item_name, data
        )
        if 0 <= idx < len(names):
            return names[idx]
        return ""

    def _unique_scenario_name_for_duplicate(self, source_row: int, insert_at: int) -> str:
        src = self._sources_data[source_row]
        base = str(src.get("scenario_name") or "").strip()
        if not base:
            base = self._default_scenario_name(source_row)
        trial = 0
        while trial < 10000:
            cand = (base + "_コピー") if trial == 0 else ("%s_%d" % (base, trial + 1))
            trial += 1
            new_src = copy.deepcopy(src)
            new_src["scenario_name"] = cand
            merged = self._sources_data[:insert_at] + [new_src] + self._sources_data[insert_at:]
            effs = [self._effective_scenario_name_at_in_list(i, merged) for i in range(len(merged))]
            if len(effs) == len(set(effs)):
                return cand
        return base + "_コピー"

    def _is_selected_registered(self) -> bool:
        idx = self._current_source_index
        if idx < 0 or idx >= len(self._sources_data):
            return False
        return bool(self._sources_data[idx].get("registered", False))

    def _update_step_button_enabled(self) -> None:
        if isinstance(getattr(self, "_btn_step", None), QPushButton):
            self._btn_step.setEnabled(self._is_selected_registered())

    def _on_register_clicked(self) -> None:
        if self._current_source_index >= 0:
            self._apply_form_to_source(self._current_source_index, include_scenario_name=True)
            from svc import svc_data_agg_scenario as _scenario_mod

            mr = self._master_item_row
            if mr >= 0 and self._master_items_list:
                items_snap = copy.deepcopy(self._master_items_list)
                if mr < len(items_snap):
                    reg_payload = self.get_item()
                    items_snap[mr]["sources"] = list(reg_payload.get("sources") or [])
                    wm = reg_payload.get("write_mode")
                    if wm is not None and str(wm).strip():
                        items_snap[mr]["write_mode"] = str(wm).strip().lower()
                    _val_errs = _scenario_mod.validate_scenario({"items": items_snap})
                    if _val_errs:
                        t_reg = (
                            _ui_disp_str(self._screen_cfg, "MSGBOX_TITLE", "").strip()
                            or _ui_disp_str(self._screen_cfg, "TITLE", "シナリオ編集")
                        )
                        pre_reg = _ui_disp_str(
                            self._screen_cfg,
                            "MSG_REGISTER_VALIDATE_PREFIX",
                            "登録できません:",
                        )
                        show_warning_notice(
                            self,
                            t_reg,
                            pre_reg + "\n" + "\n".join(_val_errs[:8]),
                        )
                        return
            snap = copy.deepcopy(self._sources_data)
            restore_row = self._current_source_index
            src = self._sources_data[self._current_source_index]
            src["registered"] = True
            if not str(src.get("scenario_name") or "").strip():
                src["scenario_name"] = self._default_scenario_name(self._current_source_index)
            ri = self._current_source_index
            while len(self._registered_display_snapshots) <= ri:
                self._registered_display_snapshots.append(None)
            self._registered_display_snapshots[ri] = copy.deepcopy(src)
            self._refresh_sources_table()
            self._sync_sources_selection_and_form(ri)
            if self._on_registered is not None:
                self._on_registered(self.get_item())
            self._undo_snapshot = snap
            self._undo_restore_row = restore_row
            self._btn_undo_remove.setEnabled(True)
        self._dirty = False
        self._update_register_button_state()
        self._update_step_button_enabled()

    def _on_step_scenario_placeholder(self) -> None:
        import copy

        if not self._is_selected_registered():
            return

        t_reg = (
            _ui_disp_str(self._screen_cfg, "MSGBOX_TITLE", "").strip()
            or _ui_disp_str(self._screen_cfg, "TITLE", "シナリオ編集")
        )
        try:
            from ui_qt.ui_data_agg_debug import create_data_agg_debug_dialog

            dbg = (_get_cfg().get("SCREENS") or {}).get("DEBUG") or {}
            live_items: list[dict[str, Any]] = []
            for i, src in enumerate(self._sources_data):
                if not isinstance(src, dict):
                    continue
                if not src.get("registered"):
                    continue
                snap = (
                    self._registered_display_snapshots[i]
                    if i < len(self._registered_display_snapshots)
                    else None
                )
                base = copy.deepcopy(snap) if snap is not None else copy.deepcopy(src)
                if not str(base.get("scenario_name") or "").strip():
                    base["scenario_name"] = self._default_scenario_name(i)
                live_items.append(
                    {
                        "name": self._item_name,
                        "id": "%s_%d" % (self._item_id, i + 1),
                        "sources": [base],
                    }
                )
            scan_paths = list(self._scan_paths_hint)
            parent = self.parentWidget()
            rescan = (
                getattr(parent, "_rescan_paths_from_current_form", None)
                if parent is not None
                else None
            )
            if callable(rescan):
                scan_paths = list(rescan())
            elif not scan_paths:
                mw = parent
                while mw is not None:
                    if hasattr(mw, "_file_list_items"):
                        scan_paths = list(getattr(mw, "_file_list_items") or [])
                        break
                    mw = mw.parentWidget()
            scan_root_dbg: str | None = None
            mw2 = parent
            while mw2 is not None:
                ep = getattr(mw2, "_edit_start_path", None)
                if ep is not None:
                    scan_root_dbg = (ep.text() or "").strip() or None
                    break
                mw2 = mw2.parentWidget()
            dlg = create_data_agg_debug_dialog(
                self,
                dbg,
                live_items=live_items,
                scan_paths=scan_paths or None,
                fixed_mode=0,
                scan_root=scan_root_dbg,
            )
            dlg.exec()
        except Exception as exc:
            _data_agg_warn_debug_open_failed(self, t_reg, exc)

    def _source_ui_bucket(self, src: dict[str, Any]) -> dict[str, Any]:
        return ensure_source_ui_block(src)

    def _wire_new_join_def(self, jd: dict[str, Any]) -> None:
        """結合キー行追加後に、その行だけフォーム変更シグナルを接続する。"""
        jd["cell"].textChanged.connect(self._on_form_changed)
        jd["row"].valueChanged.connect(self._on_form_changed)
        jd["col"].valueChanged.connect(self._on_form_changed)
        jd["item_combo"].currentIndexChanged.connect(self._on_form_changed)
        jvs = jd.get("value_shape_script")
        if jvs is not None:
            jvs.textChanged.connect(self._on_form_changed)
        for cbx in jd.get("checks") or []:
            cbx.stateChanged.connect(self._on_form_changed)

    def _notify_main_scenario_dirty(self) -> None:
        """シナリオ編集の変更をメインのシナリオ保存ダーティに反映する。"""
        w = self.parentWidget()
        mark_dirty = getattr(w, "_mark_scenario_dirty", None) if w is not None else None
        if callable(mark_dirty):
            mark_dirty()

    def _notify_parent_registered(self) -> None:
        """親のデータ集約画面へ、登録時と同様に sources を即時反映する。"""
        if self._on_registered is not None:
            self._on_registered(self.get_item())

    def _clear_scenario_undo(self) -> None:
        self._undo_snapshot = None
        self._undo_restore_row = -1
        if isinstance(getattr(self, "_btn_undo_remove", None), QPushButton):
            self._btn_undo_remove.setEnabled(False)

    def reject(self) -> None:
        self._clear_scenario_undo()
        super().reject()

    def _on_scenario_splitter_moved(self, _pos: int, _index: int) -> None:
        """右ペイン幅変更に合わせてスクロール内レイアウトを再計算する。"""
        self._resync_right_pane_layout()

    def _resync_right_pane_layout(self) -> None:
        """右ペイン配下の詳細フォームを現在幅へ追従させる。"""
        for attr in ("_detail_scroll_cell", "_detail_scroll_name"):
            sc = getattr(self, attr, None)
            if sc is not None:
                vw = sc.viewport().width()
                cont = sc.widget()
                if cont is not None and vw > 0:
                    cont.setMinimumWidth(0)
                    # ビューポートより不必要に広がらないよう上限を合わせる。
                    cont.setMaximumWidth(vw)
                sc.updateGeometry()
        self._form_stack.updateGeometry()
        rw = getattr(self, "_scenario_right_wrap", None)
        if rw is not None:
            rw.updateGeometry()

    def _on_link_or_join_removed(self, _removed: Any = None) -> None:
        """連携／結合グループ削除時（削除ボタンは _on_form_changed を発火しない）。"""
        if self._loading_source_form:
            return
        self._dirty = True
        self._update_register_button_state()
        if self._current_source_index >= 0:
            self._apply_form_to_source(self._current_source_index, include_scenario_name=False)
        self._notify_main_scenario_dirty()

    def _wire_new_link_def(self, ld: dict[str, Any]) -> None:
        """連携キー行追加後に、その行だけフォーム変更シグナルを接続する。"""
        ld["cell"].textChanged.connect(self._on_form_changed)
        ld["mode_cell"].toggled.connect(self._on_form_changed)
        ld["mode_fixed"].toggled.connect(self._on_form_changed)
        ld["row"].valueChanged.connect(self._on_form_changed)
        ld["col"].valueChanged.connect(self._on_form_changed)
        ld["item_combo"].currentIndexChanged.connect(self._on_form_changed)
        lvs = ld.get("value_shape_script")
        if lvs is not None:
            lvs.textChanged.connect(self._on_form_changed)
        for cbx in ld.get("checks") or []:
            cbx.stateChanged.connect(self._on_form_changed)

    def _wire_detail_form_signals(self) -> None:
        cr = self._cell_refs
        nr = self._name_refs

        cr["file_pattern"].textChanged.connect(self._on_form_changed)
        cr["file_name_rule"].currentIndexChanged.connect(self._on_form_changed)
        for cb in cr["ext_checkboxes"]:
            cb.stateChanged.connect(self._on_form_changed)
        cr["sheet_name"].textChanged.connect(self._on_form_changed)
        cr["sheet_rule"].currentIndexChanged.connect(self._on_form_changed)
        cr["cell_ref"].textChanged.connect(self._on_form_changed)
        cr["row_offset"].valueChanged.connect(self._on_form_changed)
        cr["col_offset"].valueChanged.connect(self._on_form_changed)
        cr["end_mode"].currentIndexChanged.connect(self._on_form_changed)
        cr["n_count"].valueChanged.connect(self._on_form_changed)
        for cbx in cr["cell_checks"]:
            cbx.stateChanged.connect(self._on_form_changed)
        vsc = cr.get("value_shape_script")
        if vsc is not None:
            vsc.textChanged.connect(self._on_form_changed)
        cr["write_mode_cell"].currentIndexChanged.connect(self._on_form_changed)
        for ld in cr["link_defs"]:
            ld["cell"].textChanged.connect(self._on_form_changed)
            ld["mode_cell"].toggled.connect(self._on_form_changed)
            ld["mode_fixed"].toggled.connect(self._on_form_changed)
            ld["row"].valueChanged.connect(self._on_form_changed)
            ld["col"].valueChanged.connect(self._on_form_changed)
            ld["item_combo"].currentIndexChanged.connect(self._on_form_changed)
            lvs = ld.get("value_shape_script")
            if lvs is not None:
                lvs.textChanged.connect(self._on_form_changed)
            for cbx in ld.get("checks") or []:
                cbx.stateChanged.connect(self._on_form_changed)
        for jd in cr["join_defs"]:
            jd["cell"].textChanged.connect(self._on_form_changed)
            jd["row"].valueChanged.connect(self._on_form_changed)
            jd["col"].valueChanged.connect(self._on_form_changed)
            jd["item_combo"].currentIndexChanged.connect(self._on_form_changed)
            jvs = jd.get("value_shape_script")
            if jvs is not None:
                jvs.textChanged.connect(self._on_form_changed)
            for cbx in jd.get("checks") or []:
                cbx.stateChanged.connect(self._on_form_changed)

        nr["search_target"].currentIndexChanged.connect(self._on_form_changed)
        nr["search_cond"].currentIndexChanged.connect(self._on_form_changed)
        nr["search_text"].textChanged.connect(self._on_form_changed)
        nr["pick_search_text"].clicked.connect(self._on_pick_name_extract_search_text)

        def _sm_nm() -> None:
            self._on_name_extract_mode_changed()
            self._on_form_changed()

        nr["start_mode_ui"].currentIndexChanged.connect(_sm_nm)
        nr["extract_mode_extract"].toggled.connect(_sm_nm)
        nr["extract_mode_fixed"].toggled.connect(_sm_nm)
        nr["delimiter"].textChanged.connect(self._on_form_changed)
        nr["block"].valueChanged.connect(self._on_form_changed)
        nr["start_pos"].valueChanged.connect(self._on_form_changed)
        nr["length_mode_ui"].currentIndexChanged.connect(_sm_nm)
        nr["length_value_edit"].textChanged.connect(self._on_form_changed)
        for cbx in nr["name_checks"]:
            cbx.stateChanged.connect(self._on_form_changed)
        vsn = nr.get("value_shape_script")
        if vsn is not None:
            vsn.textChanged.connect(self._on_form_changed)
        nr["write_mode_name"].currentIndexChanged.connect(self._on_form_changed)
        nr["path_item"].currentIndexChanged.connect(self._on_form_changed)

    def _on_name_extract_mode_changed(self) -> None:
        """名前抽出の入力可否制御（抽出/固定値・取得開始位置・取得長さ）。"""
        nr = self._name_refs
        is_fixed = bool(nr["extract_mode_fixed"].isChecked())
        m = nr["start_mode_ui"].currentIndex()
        nr["start_mode_ui"].setEnabled(not is_fixed)
        sob = nr["block"]  # start_or_block 兼用スピン
        sob.setEnabled((not is_fixed) and (m == 1 or m == 2))
        li = nr["length_mode_ui"].currentIndex()
        length_enabled = (not is_fixed) and (m != 2)
        nr["length_mode_ui"].setEnabled(length_enabled)
        nr["length_value_edit"].setEnabled(is_fixed or (length_enabled and li != 2))
        nr["delimiter"].setEnabled((not is_fixed) and m == 2)

    def _on_pick_name_extract_search_text(self) -> None:
        """検索対象に応じて選択ダイアログを開き、検索文字へ反映する。"""
        nr = self._name_refs
        try:
            init_dir = ""
            try:
                init_dir = get_last_folder() or ""
            except Exception:
                init_dir = ""
            target = nr["search_target"].currentIndex()
            if target == 0:
                picked = QFileDialog.getExistingDirectory(
                    self,
                    "フォルダを選択",
                    init_dir,
                )
            else:
                picked, _ = QFileDialog.getOpenFileName(
                    self,
                    "ファイルを選択",
                    init_dir,
                    "すべてのファイル (*.*)",
                )
            if not picked:
                return
            name = Path(picked).name.strip() or str(picked).strip()
            nr["search_text"].setText(name)
        except Exception:
            pass

    def _source_to_row(self, src: dict[str, Any]) -> tuple[str, str]:
        """ソース辞書を一覧用の種別/要約へ変換する。"""
        dc = self._screen_cfg.get("DETAIL_CELL")
        dcell = dc if isinstance(dc, dict) else {}
        return scenario_source_kind_label_and_summary(
            src, self._scenario_edit_detail_name_cfg(), detail_cell_cfg=dcell
        )

    def _refresh_sources_table(self) -> None:
        """ソース一覧テーブルを再描画。"""
        n = len(self._sources_data)
        _log_scenario_edit_diag("refresh_enter n_sources=%s item_id=%s", n, self._item_id)
        try:
            while len(self._registered_display_snapshots) < len(self._sources_data):
                self._registered_display_snapshots.append(None)
            if len(self._registered_display_snapshots) > len(self._sources_data):
                del self._registered_display_snapshots[len(self._sources_data) :]
            self._sources_table.blockSignals(True)
            try:
                self._sources_table.setRowCount(len(self._sources_data))
                dc = self._screen_cfg.get("DETAIL_CELL")
                dcell = dc if isinstance(dc, dict) else {}
                dname = self._scenario_edit_detail_name_cfg()
                disp_names = self._resolve_auto_scenario_display_names()
                for i, src in enumerate(self._sources_data):
                    snap_disp = (
                        self._registered_display_snapshots[i]
                        if i < len(self._registered_display_snapshots)
                        else None
                    )
                    num_it = QTableWidgetItem(str(i + 1))
                    num_it.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                    num_it.setFlags(num_it.flags() & ~Qt.ItemFlag.ItemIsEditable)
                    self._sources_table.setItem(i, 0, num_it)

                    sn = str(src.get("scenario_name") or "").strip()
                    name_disp = disp_names[i] if i < len(disp_names) else self._default_scenario_name(i)
                    cell_it = QTableWidgetItem(name_disp)
                    cell_it.setFlags(cell_it.flags() & ~Qt.ItemFlag.ItemIsEditable)
                    cell_it.setForeground(Qt.GlobalColor.black if sn else Qt.GlobalColor.darkGray)
                    if snap_disp is None:
                        cell_it.setToolTip(
                            _data_agg_summary_table_tooltip("— | （未登録）")
                        )
                    else:
                        cell_it.setToolTip(
                            scenario_source_tooltip_plain(
                                snap_disp, dname, detail_cell_cfg=dcell
                            )
                        )
                    self._sources_table.setItem(i, 1, cell_it)
                self._sources_table.resizeRowsToContents()
            finally:
                self._sources_table.blockSignals(False)
            cr = self._sources_table.currentRow()
            self._update_summary_preview(cr if cr >= 0 else self._current_source_index)
            _log_scenario_edit_diag(
                "refresh_leave n_sources=%s rowCount=%s currentRow=%s item_id=%s",
                len(self._sources_data),
                self._sources_table.rowCount(),
                self._sources_table.currentRow(),
                self._item_id,
            )
        except Exception as exc:
            try:
                _data_agg_ui_diag.info(
                    "[DATA_AGG_SCENARIO_EDIT] refresh_abort n_sources=%s item_id=%s err=%s",
                    len(self._sources_data),
                    self._item_id,
                    exc,
                )
                _data_agg_ui_diag.info(
                    "[DATA_AGG_SCENARIO_EDIT] refresh_abort traceback\n%s",
                    traceback.format_exc(),
                )
            except Exception:
                pass
            raise

    def _normalize_registered_snapshots_len(self) -> None:
        """_sources_data と行数を揃え、移動後もスナップショットと辞書が 1:1 になるようにする。"""
        while len(self._registered_display_snapshots) < len(self._sources_data):
            self._registered_display_snapshots.append(None)
        if len(self._registered_display_snapshots) > len(self._sources_data):
            del self._registered_display_snapshots[len(self._sources_data) :]

    def _prepare_source_row_before_reorder(self, row: int) -> bool:
        """
        ▲▼ 移動直前: 別行を編集中なら先に apply。移動対象行は辞書を正とし load のみ
        （load 直後の apply はウィジェット未反映の取り違えを招くことがあるため行わない）。
        """
        if row < 0 or row >= len(self._sources_data):
            return False
        if 0 <= self._current_source_index < len(self._sources_data) and self._current_source_index != row:
            self._apply_form_to_source(self._current_source_index, include_scenario_name=True)
        self._current_source_index = row
        self._load_source_to_form(row)
        self._normalize_registered_snapshots_len()
        return True

    def _sync_sources_selection_and_form(self, new_row: int) -> None:
        """
        左一覧の選択と右ペイン詳細を new_row に一致させる。
        sources の入れ替え・再描画の直後に selectRow だけすると selectionChanged が走り、
        まだ古い _current_source_index で apply されて「移動した内容が元位置の辞書に残る」混線になる。
        """
        n = len(self._sources_data)
        if new_row < 0 or new_row >= n:
            _log_scenario_edit_diag(
                "sync_skipped invalid new_row=%s n_sources=%s item_id=%s",
                new_row,
                n,
                self._item_id,
            )
            return
        _log_scenario_edit_diag(
            "sync_enter new_row=%s n_sources=%s item_id=%s",
            new_row,
            n,
            self._item_id,
        )
        load_ok = True
        try:
            self._sources_table.blockSignals(True)
            try:
                self._current_source_index = new_row
                try:
                    self._load_source_to_form(new_row)
                except Exception as exc:
                    load_ok = False
                    try:
                        _data_agg_ui_diag.info(
                            "[DATA_AGG_SCENARIO_EDIT] load_source_to_form exc row=%s n_sources=%s "
                            "item_id=%s err=%s",
                            new_row,
                            n,
                            self._item_id,
                            exc,
                        )
                        _data_agg_ui_diag.info(
                            "[DATA_AGG_SCENARIO_EDIT] load_source_to_form traceback\n%s",
                            traceback.format_exc(),
                        )
                    except Exception:
                        pass
                self._sources_table.selectRow(new_row)
            finally:
                self._sources_table.blockSignals(False)
            # blockSignals 中の selectRow では itemSelectionChanged が飛ばないため、
            # 初期表示で右ペインを無効化したまま追加した場合にここで有効化する。
            # load 失敗時も右ペインは有効のままにし、診断ログで追えるようにする。
            self._form_stack.setEnabled(True)
            self._form_combo_type.setEnabled(True)
            self._edit_scenario_ident.setEnabled(True)
            self._update_summary_preview(new_row)
            self._update_step_button_enabled()
            cr = self._sources_table.currentRow()
            _log_scenario_edit_diag(
                "sync_done new_row=%s n_sources=%s load_ok=%s form_stack=%s combo=%s ident=%s "
                "table_currentRow=%s item_id=%s",
                new_row,
                n,
                load_ok,
                self._form_stack.isEnabled(),
                self._form_combo_type.isEnabled(),
                self._edit_scenario_ident.isEnabled(),
                cr,
                self._item_id,
            )
        except Exception as exc:
            try:
                _data_agg_ui_diag.info(
                    "[DATA_AGG_SCENARIO_EDIT] sync_abort new_row=%s n_sources=%s item_id=%s err=%s",
                    new_row,
                    n,
                    self._item_id,
                    exc,
                )
                _data_agg_ui_diag.info(
                    "[DATA_AGG_SCENARIO_EDIT] sync_abort traceback\n%s",
                    traceback.format_exc(),
                )
            except Exception:
                pass
            raise

    def _update_summary_preview(self, row: int) -> None:
        """左ペイン下部の要約（一覧選択行の全文プレビュー）。"""
        _u = lambda k, d: _ui_disp_str(self._screen_cfg, k, d)
        if row < 0 or row >= len(self._sources_data):
            self._summary_preview.clear()
            self._scenario_set_tip(
                self._summary_preview,
                "TIP_SCENARIO_SUMMARY_PREVIEW",
                "選択中シナリオの設定要約（プレーンテキスト）です。",
            )
            return
        disp = (
            self._registered_display_snapshots[row]
            if row < len(self._registered_display_snapshots)
            else None
        )
        if disp is None:
            full_text = "%s\n\n%s" % (_u("LABEL_SUMMARY_FULL", "要約（全文）"), "（未登録）")
            self._summary_preview.setText(full_text)
            self._scenario_set_tip(
                self._summary_preview,
                "TIP_SCENARIO_SUMMARY_PREVIEW",
                "選択中シナリオの設定要約（プレーンテキスト）です。",
            )
            return
        lines = self._detail_lines_for_source(disp, row)
        full_text = "%s\n\n%s" % (_u("LABEL_SUMMARY_FULL", "要約（全文）"), "\n".join(lines))
        self._summary_preview.setText(full_text)
        long_tip = _normalize_message_newlines(full_text).strip()
        if len(long_tip) > 200:
            self._summary_preview.setToolTip(long_tip[:4096])
        else:
            self._scenario_set_tip(
                self._summary_preview,
                "TIP_SCENARIO_SUMMARY_PREVIEW",
                "選択中シナリオの設定要約（プレーンテキスト）です。",
            )

    def _make_new_source_template_row(self) -> dict[str, Any]:
        """新規ソース1行分の既定 dict（種別は先頭行の系統に追随）。"""
        if self._sources_data:
            t0 = str(self._sources_data[0].get("type") or "cell").strip().lower()
            if _scenario_lineage_bucket(t0) == "path":
                return {
                    "type": "name_extract",
                    "registered": False,
                    "source_type": "file_name",
                    "start_mode": "head",
                    "length_mode": "end",
                }
            return {"type": "cell", "sheet_name": "", "cell_ref": "", "registered": False}
        return {"type": "cell", "sheet_name": "", "cell_ref": "", "registered": False}

    def _insert_empty_source_at(self, insert_at: int) -> None:
        """insert_at 位置に空のソース行を挿入（0 .. len まで）。"""
        insert_at = max(0, min(int(insert_at), len(self._sources_data)))
        if 0 <= self._current_source_index < len(self._sources_data):
            self._apply_form_to_source(
                self._current_source_index, include_scenario_name=True
            )
        self._sources_data.insert(insert_at, self._make_new_source_template_row())
        self._registered_display_snapshots.insert(insert_at, None)
        self._refresh_sources_table()
        self._sync_sources_selection_and_form(insert_at)
        self._dirty = True
        self._update_register_button_state()
        self._notify_main_scenario_dirty()

    def _on_sources_table_context_menu(self, pos: QPoint) -> None:
        """ソース一覧の右クリック: 上／下追加・複写・削除（ボタンと同機能）。"""
        idx = self._sources_table.indexAt(pos)
        r = int(idx.row())
        if r >= 0:
            self._sources_table.selectRow(r)
        menu = QMenu(self)
        a_up = menu.addAction(
            _ui_disp_str(self._screen_cfg, "CTX_INSERT_SOURCE_ABOVE", "上の行を追加")
        )
        a_dn = menu.addAction(
            _ui_disp_str(self._screen_cfg, "CTX_INSERT_SOURCE_BELOW", "下の行を追加")
        )
        menu.addSeparator()
        dup_lbl = _ui_disp_str(self._screen_cfg, "CTX_DUPLICATE_SOURCE", "複写")
        a_dup = menu.addAction(dup_lbl)
        menu.addSeparator()
        a_del = menu.addAction(
            _ui_disp_str(self._screen_cfg, "CTX_REMOVE_SOURCE", "削除")
        )
        chosen = menu.exec(self._sources_table.viewport().mapToGlobal(pos))
        if chosen == a_up:
            ins = r if r >= 0 else 0
            self._insert_empty_source_at(ins)
        elif chosen == a_dn:
            ins = (r + 1) if r >= 0 else len(self._sources_data)
            self._insert_empty_source_at(ins)
        elif chosen == a_dup:
            self._on_duplicate_source()
        elif chosen == a_del:
            self._on_remove_source()

    def _on_add_source(self) -> None:
        """取得ソースを追加。"""
        row = self._make_new_source_template_row()
        self._sources_data.append(row)
        self._registered_display_snapshots.append(None)
        _log_scenario_edit_diag(
            "add_source item_id=%s n_sources_after_append=%s",
            self._item_id,
            len(self._sources_data),
        )
        self._refresh_sources_table()
        _log_scenario_edit_diag(
            "after_refresh add_source n_sources=%s table_rowCount=%s table_currentRow=%s item_id=%s",
            len(self._sources_data),
            self._sources_table.rowCount(),
            self._sources_table.currentRow(),
            self._item_id,
        )
        self._sync_sources_selection_and_form(len(self._sources_data) - 1)
        self._dirty = True
        self._update_register_button_state()
        self._notify_main_scenario_dirty()

    def _on_duplicate_source(self) -> None:
        """選択行のシナリオを deepcopy して直下に追加（識別名は重複しないよう付与）。"""
        row = self._sources_table.currentRow()
        if row < 0 or row >= len(self._sources_data):
            return
        if 0 <= self._current_source_index < len(self._sources_data) and self._current_source_index != row:
            self._apply_form_to_source(self._current_source_index, include_scenario_name=True)
        self._current_source_index = row
        self._load_source_to_form(row)
        self._apply_form_to_source(row, include_scenario_name=True)
        insert_at = row + 1
        dup = copy.deepcopy(self._sources_data[row])
        dup["registered"] = False
        dup["scenario_name"] = self._unique_scenario_name_for_duplicate(row, insert_at)
        self._sources_data.insert(insert_at, dup)
        self._registered_display_snapshots.insert(insert_at, None)
        self._refresh_sources_table()
        self._sync_sources_selection_and_form(insert_at)
        self._dirty = True
        self._update_register_button_state()
        self._notify_main_scenario_dirty()

    def _on_remove_source(self) -> None:
        """選択中の取得ソースを削除。"""
        row = self._sources_table.currentRow()
        if 0 <= row < len(self._sources_data):
            snapshot = copy.deepcopy(self._sources_data)
            deleted_index = row
            if self._current_source_index == row:
                self._current_source_index = -1
            elif self._current_source_index > row:
                self._current_source_index -= 1
            del self._sources_data[row]
            if row < len(self._registered_display_snapshots):
                del self._registered_display_snapshots[row]
            self._refresh_sources_table()
            if self._sources_data:
                sel = min(max(0, row), len(self._sources_data) - 1)
                self._sync_sources_selection_and_form(sel)
                self._notify_main_scenario_dirty()
            else:
                self._form_stack.setEnabled(False)
                self._form_combo_type.setEnabled(False)
                self._edit_scenario_ident.setEnabled(False)
                _log_scenario_edit_diag(
                    "remove_source last_deleted pane_disabled item_id=%s", self._item_id
                )
                self._dirty = True
                self._update_register_button_state()
                self._notify_main_scenario_dirty()
            self._undo_snapshot = snapshot
            self._undo_restore_row = deleted_index
            self._btn_undo_remove.setEnabled(True)
            self._notify_parent_registered()
        self._update_step_button_enabled()

    def _on_undo_scenario(self) -> None:
        """直近の削除または登録の直前状態へ1回だけ元に戻す。"""
        snap = self._undo_snapshot
        if snap is None:
            return
        di = self._undo_restore_row
        self._sources_data = copy.deepcopy(snap)
        self._registered_display_snapshots = [
            copy.deepcopy(s) if isinstance(s, dict) and s.get("registered") else None
            for s in self._sources_data
        ]
        self._clear_scenario_undo()
        self._refresh_sources_table()
        self._sources_table.blockSignals(True)
        try:
            if self._sources_data:
                sel = di if 0 <= di < len(self._sources_data) else 0
                self._sources_table.selectRow(sel)
                self._current_source_index = sel
                self._load_source_to_form(sel)
                self._form_stack.setEnabled(True)
                self._form_combo_type.setEnabled(True)
                self._edit_scenario_ident.setEnabled(True)
                self._update_summary_preview(sel)
            else:
                self._current_source_index = -1
                self._form_stack.setEnabled(False)
                self._form_combo_type.setEnabled(False)
                self._edit_scenario_ident.setEnabled(False)
                self._update_summary_preview(-1)
        finally:
            self._sources_table.blockSignals(False)
        self._dirty = True
        self._update_register_button_state()
        self._notify_main_scenario_dirty()
        self._notify_parent_registered()
        self._update_step_button_enabled()

    def _on_source_selection_changed(self) -> None:
        """ソース選択変更時。フォームを保存してから選択行をロード。"""
        row = self._sources_table.currentRow()
        if self._current_source_index >= 0 and self._current_source_index < len(self._sources_data):
            self._apply_form_to_source(self._current_source_index, include_scenario_name=False)
        self._current_source_index = row if 0 <= row < len(self._sources_data) else -1
        if self._current_source_index >= 0:
            try:
                self._load_source_to_form(self._current_source_index)
            except Exception as exc:
                try:
                    _data_agg_ui_diag.info(
                        "[DATA_AGG_SCENARIO_EDIT] selection_changed load exc row=%s item_id=%s err=%s",
                        self._current_source_index,
                        self._item_id,
                        exc,
                    )
                    _data_agg_ui_diag.info(
                        "[DATA_AGG_SCENARIO_EDIT] selection_changed load traceback\n%s",
                        traceback.format_exc(),
                    )
                except Exception:
                    pass
            self._form_stack.setEnabled(True)
            self._form_combo_type.setEnabled(True)
            self._edit_scenario_ident.setEnabled(True)
            self._update_summary_preview(self._current_source_index)
        else:
            if self._sources_data:
                _log_scenario_edit_diag(
                    "selection_changed no_index_but_has_sources table_row=%s n_sources=%s "
                    "item_id=%s",
                    row,
                    len(self._sources_data),
                    self._item_id,
                )
            self._form_stack.setEnabled(False)
            self._form_combo_type.setEnabled(False)
            self._edit_scenario_ident.setEnabled(False)
            self._update_summary_preview(-1)
        self._update_step_button_enabled()

    def _on_source_move_up(self) -> None:
        row = self._sources_table.currentRow()
        if row <= 0 or row >= len(self._sources_data):
            return
        if not self._prepare_source_row_before_reorder(row):
            return
        self._sources_data[row - 1], self._sources_data[row] = (
            self._sources_data[row],
            self._sources_data[row - 1],
        )
        self._registered_display_snapshots[row - 1], self._registered_display_snapshots[row] = (
            self._registered_display_snapshots[row],
            self._registered_display_snapshots[row - 1],
        )
        self._refresh_sources_table()
        self._sync_sources_selection_and_form(row - 1)
        self._dirty = True
        self._update_register_button_state()
        self._notify_main_scenario_dirty()
        self._notify_parent_registered()

    def _on_source_move_down(self) -> None:
        row = self._sources_table.currentRow()
        if row < 0 or row >= len(self._sources_data) - 1:
            return
        if not self._prepare_source_row_before_reorder(row):
            return
        self._sources_data[row + 1], self._sources_data[row] = (
            self._sources_data[row],
            self._sources_data[row + 1],
        )
        self._registered_display_snapshots[row + 1], self._registered_display_snapshots[row] = (
            self._registered_display_snapshots[row],
            self._registered_display_snapshots[row + 1],
        )
        self._refresh_sources_table()
        self._sync_sources_selection_and_form(row + 1)
        self._dirty = True
        self._update_register_button_state()
        self._notify_main_scenario_dirty()
        self._notify_parent_registered()

    def _on_form_type_changed(self) -> None:
        """フォームの種別変更時。スタックを切り替え、ソースの type を更新。"""
        if self._current_source_index < 0:
            return
        new_t = self._form_combo_type.currentData() or "cell"
        new_b = _scenario_lineage_bucket(new_t)
        for j, other in enumerate(self._sources_data):
            if j == self._current_source_index:
                continue
            ob = _scenario_lineage_bucket(str(other.get("type") or "cell"))
            if ob != new_b:
                show_warning_notice(
                    self,
                    _ui_disp_str(self._screen_cfg, "TITLE", "シナリオ編集"),
                    _ui_disp_str(
                        self._screen_cfg,
                        "MSG_MIXED_LINEAGE",
                        "同一項目内では「セル座標から取得」と「名前から取得」を混在できません。別のマスタ項目に分けてください。",
                    ),
                )
                old_t = str(self._sources_data[self._current_source_index].get("type") or "cell").strip().lower()
                if old_t in ("metadata", "meta", "filename"):
                    old_t = "name_extract"
                rev_idx = 0 if old_t == "cell" else 1
                self._form_combo_type.blockSignals(True)
                self._form_combo_type.setCurrentIndex(rev_idx)
                self._form_stack.setCurrentIndex(rev_idx)
                self._form_combo_type.blockSignals(False)
                return
        t = new_t
        idx = {"cell": 0, "name_extract": 1}.get(t, 0)
        self._form_stack.setCurrentIndex(idx)
        self._sources_data[self._current_source_index]["type"] = t
        if t == "cell":
            self._sources_data[self._current_source_index].setdefault("sheet_name", "")
            self._sources_data[self._current_source_index].setdefault("cell_ref", "")
        elif t == "name_extract":
            self._sources_data[self._current_source_index].setdefault("source_type", "file_name")
            self._sources_data[self._current_source_index].setdefault("start_mode", "head")
            self._sources_data[self._current_source_index].setdefault("length_mode", "end")
            self._on_name_extract_mode_changed()
        self._dirty = True
        self._update_register_button_state()
        self._notify_main_scenario_dirty()

    def _on_form_changed(self) -> None:
        """フォームの値変更時。現在のソースに反映してテーブルを更新。"""
        if self._loading_source_form:
            return
        self._dirty = True
        self._update_register_button_state()
        if self._current_source_index >= 0:
            self._apply_form_to_source(self._current_source_index, include_scenario_name=False)
        self._notify_main_scenario_dirty()

    def _block_detail_form_signals(self, block: bool) -> None:
        cr, nr = self._cell_refs, self._name_refs
        for w in (
            cr["file_pattern"],
            cr["file_name_rule"],
            cr["sheet_name"],
            cr["sheet_rule"],
            cr["cell_ref"],
            cr["row_offset"],
            cr["col_offset"],
            cr["end_mode"],
            cr["n_count"],
            cr["write_mode_cell"],
            cr.get("value_shape_script"),
        ):
            if w is not None:
                w.blockSignals(block)
        for cb in cr["ext_checkboxes"]:
            cb.blockSignals(block)
        for cbx in cr["cell_checks"]:
            cbx.blockSignals(block)
        for ld in cr["link_defs"]:
            for k in ("cell", "mode_cell", "mode_fixed", "row", "col", "item_combo"):
                ld[k].blockSignals(block)
            lvs = ld.get("value_shape_script")
            if lvs is not None:
                lvs.blockSignals(block)
            for cbx in ld.get("checks") or []:
                cbx.blockSignals(block)
        for jd in cr["join_defs"]:
            for k in ("cell", "row", "col", "item_combo"):
                jd[k].blockSignals(block)
            jvs = jd.get("value_shape_script")
            if jvs is not None:
                jvs.blockSignals(block)
            for cbx in jd.get("checks") or []:
                cbx.blockSignals(block)
        for w in (
            nr["search_target"],
            nr["search_cond"],
            nr["search_text"],
            nr["pick_search_text"],
            nr["extract_mode_extract"],
            nr["extract_mode_fixed"],
            nr["start_mode_ui"],
            nr["delimiter"],
            nr["block"],
            nr["start_pos"],
            nr["length_mode_ui"],
            nr["length_value_edit"],
            nr["write_mode_name"],
            nr["path_item"],
            nr.get("value_shape_script"),
        ):
            if w is not None:
                w.blockSignals(block)
        for cbx in nr["name_checks"]:
            cbx.blockSignals(block)
        self._form_combo_type.blockSignals(block)
        self._edit_scenario_ident.blockSignals(block)

    def _load_source_to_form(self, idx: int) -> None:
        """ソース idx をフォームにロード（詳細フォーム ↔ UI 保存ブロック `ui_scenario_source_v1`）。"""
        if idx < 0 or idx >= len(self._sources_data):
            return
        src = self._sources_data[idx]
        stype = (src.get("type") or "cell").strip().lower()
        if stype in ("metadata", "meta", "filename"):
            src["type"] = "name_extract"
            if stype == "filename":
                src["source_type"] = "file_name"
            elif src.get("source_type") == "full_path":
                src["source_type"] = "file_name"
            if src.get("delimiter"):
                src["start_mode"] = "delimiter"
            elif src.get("start") is not None:
                src["start_mode"] = "position"
                src["start_value"] = src.get("start")
            else:
                src["start_mode"] = "head"
            if src.get("length") is not None:
                src["length_mode"] = "count"
                src["length_value"] = src.get("length")
            else:
                src["length_mode"] = "end"
            for k in ("start", "length"):
                src.pop(k, None)
            stype = "name_extract"
        elif stype != "name_extract":
            stype = "cell"

        self._loading_source_form = True
        self._block_detail_form_signals(True)
        try:
            i = self._form_combo_type.findData(stype)
            if i >= 0:
                self._form_combo_type.setCurrentIndex(i)
            self._form_stack.setCurrentIndex({"cell": 0, "name_extract": 1}.get(stype, 0))
            p = self._source_ui_bucket(src)

            if stype == "cell":
                r = self._cell_refs
                r["file_pattern"].setText("" if p.get("file_pattern") is None else str(p.get("file_pattern")))
                fn_rule = str(p.get("file_name_rule") or "含む")
                fri = r["file_name_rule"].findText(fn_rule)
                r["file_name_rule"].setCurrentIndex(fri if fri >= 0 else 1)
                want = set(p.get("ext_checked") or [".xlsx", ".xlsm", ".xls"])
                valid_tags = {str(cb.property("ext_tag") or "") for cb in r["ext_checkboxes"]}
                want = {t for t in want if t in valid_tags}
                if not want:
                    want = set(valid_tags)
                for cb in r["ext_checkboxes"]:
                    tag = str(cb.property("ext_tag") or "")
                    cb.setChecked(tag in want)
                r["sheet_name"].setText(str(src.get("sheet_name") or ""))
                rule = str(p.get("sheet_rule") or "左端シート")
                ri = r["sheet_rule"].findText(rule)
                r["sheet_rule"].setCurrentIndex(ri if ri >= 0 else 0)
                r["cell_ref"].setText(str(src.get("cell_ref") or ""))
                r["row_offset"].setValue(int(src.get("row_offset") or 0))
                r["col_offset"].setValue(int(src.get("col_offset") or 0))
                ru = bool(src.get("repeat_until_empty", True))
                rm = src.get("repeat_max")
                labels = r.get("end_mode_labels") or ["N件", "空白まで"]
                blank_lbl = labels[1] if len(labels) > 1 else "空白まで"
                n_lbl = labels[0] if labels else "N件"
                em = r["end_mode"]
                if ru and (rm is None or int(rm or 0) <= 0):
                    ix = em.findText(blank_lbl)
                    em.setCurrentIndex(ix if ix >= 0 else min(1, em.count() - 1))
                else:
                    ix = em.findText(n_lbl)
                    em.setCurrentIndex(ix if ix >= 0 else 0)
                    dc_n = self._screen_cfg.get("DETAIL_CELL")
                    def_n = 1
                    if isinstance(dc_n, dict) and dc_n.get("DEFAULT_N_COUNT") is not None:
                        try:
                            def_n = int(dc_n.get("DEFAULT_N_COUNT") or 1)
                        except (TypeError, ValueError):
                            def_n = 1
                    r["n_count"].setValue(int(rm or def_n))
                sync_guard = r.get("sync_offset_blank_guard")
                if callable(sync_guard):
                    sync_guard()
                else:
                    sync_end = r.get("sync_n_count_for_end")
                    if callable(sync_end):
                        sync_end()
                if "legacy_anchor" not in p and src.get("anchor"):
                    p["legacy_anchor"] = src.get("anchor")
                vsc = r.get("value_shape_script")
                if vsc is not None:
                    vsc.setText(str(p.get("value_shape_script") or ""))
                saved_chk = p.get("cell_checks") or []
                for i2, cbx in enumerate(r["cell_checks"]):
                    cbx.setChecked(self._saved_process_check_at_index(i2, saved_chk))
                llist = p.get("link_defs") or []
                if not isinstance(llist, list):
                    llist = []
                al = r.get("append_link_group")
                rl = r.get("remove_link_group")
                if callable(al) and callable(rl):
                    while len(r["link_defs"]) < len(llist):
                        al()
                    while len(r["link_defs"]) > len(llist) and r["link_defs"]:
                        rl(r["link_defs"][-1])
                dcell = dict(self._screen_cfg.get("DETAIL_CELL") or {})
                link_mode_items = dcell.get("LINK_MODE_ITEMS") or ["セル座標", "固定値"]
                if not isinstance(link_mode_items, list) or len(link_mode_items) < 2:
                    link_mode_items = ["セル座標", "固定値"]
                link_mode_items = [str(x) for x in link_mode_items]
                fixed_mode_lbl = link_mode_items[1]
                for i2, ld in enumerate(r["link_defs"]):
                    if i2 < len(llist) and isinstance(llist[i2], dict):
                        ld["cell"].setText(str(llist[i2].get("cell") or ""))
                        mode_txt = str(llist[i2].get("mode") or "セル座標").strip()
                        if mode_txt == fixed_mode_lbl:
                            ld["mode_fixed"].setChecked(True)
                        else:
                            ld["mode_cell"].setChecked(True)
                        ld["row"].setValue(int(llist[i2].get("row") or 0))
                        ld["col"].setValue(int(llist[i2].get("col") or 0))
                        self._combo_select_saved_master_item(
                            ld["item_combo"], str(llist[i2].get("item") or "")
                        )
                        lnk_vs = ld.get("value_shape_script")
                        if lnk_vs is not None:
                            lnk_vs.setText(str(llist[i2].get("value_shape_script") or ""))
                        chk_vals = llist[i2].get("checks") or []
                        if not isinstance(chk_vals, list):
                            chk_vals = []
                        for ci, cbx in enumerate(ld.get("checks") or []):
                            cbx.setChecked(self._saved_process_check_at_index(ci, chk_vals))
                jlist = p.get("join_defs") or []
                if not isinstance(jlist, list):
                    jlist = []
                aj = r.get("append_join_group")
                rj = r.get("remove_join_group")
                if callable(aj) and callable(rj):
                    while len(r["join_defs"]) < len(jlist):
                        aj()
                    while len(r["join_defs"]) > len(jlist) and r["join_defs"]:
                        rj(r["join_defs"][-1])
                for i2, jd in enumerate(r["join_defs"]):
                    if i2 < len(jlist) and isinstance(jlist[i2], dict):
                        jd["cell"].setText(str(jlist[i2].get("cell") or ""))
                        jd["row"].setValue(int(jlist[i2].get("row") or 0))
                        jd["col"].setValue(int(jlist[i2].get("col") or 0))
                        self._combo_select_saved_master_item(
                            jd["item_combo"], str(jlist[i2].get("item") or "")
                        )
                        jn_vs = jd.get("value_shape_script")
                        if jn_vs is not None:
                            jn_vs.setText(str(jlist[i2].get("value_shape_script") or ""))
                        chk_vals = jlist[i2].get("checks") or []
                        if not isinstance(chk_vals, list):
                            chk_vals = []
                        for ci, cbx in enumerate(jd.get("checks") or []):
                            cbx.setChecked(self._saved_process_check_at_index(ci, chk_vals))
                wm_idx = p.get("write_mode_cell_idx")
                cb_wm = r["write_mode_cell"]
                if isinstance(wm_idx, int) and 0 <= wm_idx < cb_wm.count():
                    cb_wm.setCurrentIndex(wm_idx)
                else:
                    kix = cb_wm.findData(self._item_write_mode_hint)
                    cb_wm.setCurrentIndex(kix if kix >= 0 else 0)
            elif stype == "name_extract":
                nr = self._name_refs
                src_type = str(src.get("source_type") or "file_name")
                nr["search_target"].setCurrentIndex(0 if src_type == "dir_name" else 1)
                sc_raw = str(src.get("search_condition") or "include").strip().lower()
                if sc_raw in ("exact", "equals", "完全一致"):
                    ix = nr["search_cond"].findText("完全一致")
                    nr["search_cond"].setCurrentIndex(ix if ix >= 0 else 0)
                elif sc_raw in ("exclude", "含まない"):
                    ix = nr["search_cond"].findText("含まない")
                    nr["search_cond"].setCurrentIndex(ix if ix >= 0 else min(2, nr["search_cond"].count() - 1))
                else:
                    ix = nr["search_cond"].findText("含む")
                    nr["search_cond"].setCurrentIndex(ix if ix >= 0 else min(1, nr["search_cond"].count() - 1))
                nr["search_text"].setText(str(src.get("search_text") or ""))
                ex_mode = str(p.get("extract_mode") or "extract").strip().lower()
                nr["extract_mode_extract"].setChecked(ex_mode != "fixed")
                nr["extract_mode_fixed"].setChecked(ex_mode == "fixed")
                start_mode = src.get("start_mode")
                if start_mode is None:
                    if src.get("delimiter"):
                        start_mode = "delimiter"
                    elif src.get("start") is not None:
                        start_mode = "position"
                    else:
                        start_mode = "head"
                smap = {"head": 0, "position": 1, "delimiter": 2}
                nr["start_mode_ui"].setCurrentIndex(int(smap.get(str(start_mode), 0)))
                nr["delimiter"].setText(str(src.get("delimiter") or "_"))
                nr["block"].setValue(max(1, int(src.get("part_index") or 1)))
                nr["start_pos"].setValue(max(1, int(src.get("start_value") or src.get("start") or 1)))
                length_mode = src.get("length_mode")
                if length_mode is None and src.get("length") is not None:
                    length_mode = "count"
                elif length_mode is None:
                    length_mode = "end"
                lmap = {"char": 0, "count": 1, "end": 2}
                nr["length_mode_ui"].setCurrentIndex(int(lmap.get(str(length_mode), 2)))
                ln_val = src.get("length_value")
                if ln_val is None:
                    ln_val = src.get("length")
                nr["length_value_edit"].setText(str(ln_val) if ln_val is not None else "")
                vsn = nr.get("value_shape_script")
                if vsn is not None:
                    vsn.setText(str(p.get("value_shape_script") or ""))
                saved_n = p.get("name_checks") or []
                for i2, cbx in enumerate(nr["name_checks"]):
                    cbx.setChecked(self._saved_process_check_at_index(i2, saved_n))
                pit_raw = p.get("path_item")
                path_txt = str(pit_raw).strip() if pit_raw is not None else ""
                if self._path_item_text_is_legacy_placeholder(path_txt):
                    nr["path_item"].setCurrentIndex(0)
                else:
                    primary = str(nr.get("path_item_primary") or "").strip()
                    if primary and path_txt == primary:
                        pix = nr["path_item"].findText(path_txt)
                        nr["path_item"].setCurrentIndex(pix if pix >= 0 else 0)
                    else:
                        self._combo_select_saved_master_item(nr["path_item"], path_txt)
                wm_n = p.get("write_mode_name_idx")
                cb_wn = nr["write_mode_name"]
                if isinstance(wm_n, int) and 0 <= wm_n < cb_wn.count():
                    cb_wn.setCurrentIndex(wm_n)
                else:
                    cb_wn.setCurrentIndex(0)
                self._on_name_extract_mode_changed()

            sn = str(src.get("scenario_name") or "").strip()
            if not sn:
                self._edit_scenario_ident.setPlaceholderText(self._default_scenario_name(idx))
                self._edit_scenario_ident.setText("")
            else:
                self._edit_scenario_ident.setPlaceholderText(self._default_scenario_name(idx))
                self._edit_scenario_ident.setText(sn)
            self._on_scenario_name_text_changed(self._edit_scenario_ident.text())

            self._dirty = False
            self._update_register_button_state()
        finally:
            self._block_detail_form_signals(False)
            self._loading_source_form = False
            try:
                sync_sheet = self._cell_refs.get("sync_sheet_name_enabled")
                if callable(sync_sheet):
                    sync_sheet()
            except Exception:
                pass
            self._resync_right_pane_layout()

    @staticmethod
    def _saved_process_check_at_index(slot: int, saved_vals: list[Any]) -> bool:
        """加工チェック保存値を UI スロットに復元。ラベル改定前の JSON とも位置ベースで互換。"""
        sset = {str(x) for x in (saved_vals or []) if x is not None}
        aliases: dict[int, tuple[str, ...]] = {
            0: ("トリム",),
            1: (
                "全角→半角（英数字・記号）",
                "全角→半角",
            ),
            2: (
                "年月日変換",
                "日付変換",
                "日付変換（yyyy/mm/dd形式に変換）",
                "日付変換 (yyyy/mm/dd) 時刻なし",
                "日付 (yyyy/mm/dd)",
                "日付変換 (yyyy/mm/dd)",
            ),
        }
        for cand in aliases.get(slot, ()):
            if cand in sset:
                return True
        return False

    def _scenario_edit_detail_name_cfg(self) -> dict[str, Any]:
        d = self._screen_cfg.get("DETAIL_NAME")
        return d if isinstance(d, dict) else {}

    def _ja_display_start_mode(self, raw: Any) -> str:
        return _fmt_ne_start_mode(self._scenario_edit_detail_name_cfg(), raw)

    def _ja_display_length_mode(self, raw: Any) -> str:
        return _fmt_ne_length_mode(self._scenario_edit_detail_name_cfg(), raw)

    @staticmethod
    def _ja_display_search_target(raw: Any) -> str:
        st = str(raw or "file_name").strip().lower()
        return "フォルダ名" if st == "dir_name" else "ファイル名"

    @staticmethod
    def _ja_display_search_cond(raw: Any) -> str:
        sc = str(raw or "include").strip().lower()
        if sc in ("exact", "equals", "完全一致"):
            return "完全一致"
        return "含む" if sc == "include" else "含まない"

    def _ja_display_write_mode_name_idx(self, raw_idx: Any) -> str:
        return _fmt_ne_write_mode(self._scenario_edit_detail_name_cfg(), raw_idx)

    def _scenario_name_label_for_summary(self) -> str:
        sl = str(self._screen_cfg.get("LABEL_SCENARIO_NAME") or "シナリオ名").strip()
        for suf in ("：", ":"):
            if sl.endswith(suf):
                sl = sl[:-1].strip()
        return sl or "シナリオ名"

    def _detail_lines_for_source(self, src: dict[str, Any], row_index: int) -> list[str]:
        """要約（全文）用の箇条書き行（上限は呼び出し側でカット）。内部値は日本語表示に変換する。"""
        lbl = self._scenario_name_label_for_summary()
        sn = str(src.get("scenario_name") or "").strip()
        ident = sn if sn else self._default_scenario_name(row_index)
        stype = (src.get("type") or "cell").strip().lower()
        if stype in ("metadata", "meta", "filename"):
            stype = "name_extract"
        if stype == "name_extract":
            pb = source_ui_block(src) or {}
            return name_extract_full_detail_lines(
                self._item_name, lbl, ident, src, pb, self._scenario_edit_detail_name_cfg()
            )
        pb = source_ui_block(src) or {}
        dc = self._screen_cfg.get("DETAIL_CELL")
        detail_cell = dc if isinstance(dc, dict) else {}
        return cell_coordinate_full_detail_lines(
            self._item_name,
            lbl,
            ident,
            src,
            pb,
            detail_cell,
            full_value_shape=True,
        )

    def _apply_form_to_source(self, idx: int, include_scenario_name: bool = True) -> None:
        """フォームの値をソース idx に反映。"""
        if idx < 0 or idx >= len(self._sources_data):
            return
        src = self._sources_data[idx]
        if include_scenario_name:
            src["scenario_name"] = self._edit_scenario_ident.text().strip()
        stype = self._form_combo_type.currentData() or "cell"
        src["type"] = stype
        if stype == "cell":
            r = self._cell_refs
            src["sheet_name"] = r["sheet_name"].text().strip()
            src["cell_ref"] = r["cell_ref"].text().strip()
            src["row_offset"] = int(r["row_offset"].value())
            src["col_offset"] = int(r["col_offset"].value())
            elabels = r.get("end_mode_labels") or ["N件", "空白まで"]
            blank_lbl = elabels[1] if len(elabels) > 1 else "空白まで"
            if r["end_mode"].currentText() == blank_lbl:
                src["repeat_until_empty"] = True
                src["repeat_max"] = None
            else:
                src["repeat_until_empty"] = False
                src["repeat_max"] = int(r["n_count"].value())
            src["repeat_direction"] = "vertical"
            p = self._source_ui_bucket(src)
            src["anchor"] = p.get("legacy_anchor")
            p["file_pattern"] = r["file_pattern"].text()
            p["file_name_rule"] = r["file_name_rule"].currentText()
            p["ext_checked"] = [
                str(cb.property("ext_tag") or "")
                for cb in r["ext_checkboxes"]
                if cb.isChecked() and cb.property("ext_tag")
            ]
            p["sheet_rule"] = r["sheet_rule"].currentText()
            p["cell_checks"] = [cb.text() for cb in r["cell_checks"] if cb.isChecked()]
            vsc = r.get("value_shape_script")
            if vsc is not None:
                p["value_shape_script"] = vsc.text().strip()
            ph_item = str(r.get("join_item_placeholder") or "").strip() or "（マスタ項目を選択）"

            def _item_ok(txt: str) -> bool:
                t = (txt or "").strip()
                return bool(t) and t != ph_item

            def _link_mode_persist(ld0: dict[str, Any]) -> str:
                if ld0.get("mode_fixed") and ld0["mode_fixed"].isChecked():
                    t = ld0["mode_fixed"].text().strip()
                    return t or "固定値"
                t = ld0["mode_cell"].text().strip() if ld0.get("mode_cell") else ""
                return t or "セル座標"

            p["link_defs"] = [
                {
                    "cell": ld["cell"].text(),
                    "mode": _link_mode_persist(ld),
                    "row": ld["row"].value(),
                    "col": ld["col"].value(),
                    "item": ld["item_combo"].currentText().strip(),
                    "checks": [cb.text() for cb in (ld.get("checks") or []) if cb.isChecked()],
                    "value_shape_script": (
                        ld["value_shape_script"].text().strip()
                        if ld.get("value_shape_script") is not None
                        else ""
                    ),
                }
                for ld in r["link_defs"]
                if _item_ok(ld["item_combo"].currentText())
            ]
            p["join_defs"] = [
                {
                    "cell": jd["cell"].text(),
                    "row": jd["row"].value(),
                    "col": jd["col"].value(),
                    "item": jd["item_combo"].currentText().strip(),
                    "checks": [cb.text() for cb in (jd.get("checks") or []) if cb.isChecked()],
                    "value_shape_script": (
                        jd["value_shape_script"].text().strip()
                        if jd.get("value_shape_script") is not None
                        else ""
                    ),
                }
                for jd in r["join_defs"]
                if _item_ok(jd["item_combo"].currentText())
            ]
            p["write_mode_cell_idx"] = int(r["write_mode_cell"].currentIndex())
        elif stype == "name_extract":
            nr = self._name_refs
            src["type"] = "name_extract"
            src["source_type"] = "dir_name" if nr["search_target"].currentIndex() == 0 else "file_name"
            sc_txt = nr["search_cond"].currentText().strip()
            if sc_txt == "完全一致":
                src["search_condition"] = "exact"
            elif sc_txt == "含まない":
                src["search_condition"] = "exclude"
            else:
                src["search_condition"] = "include"
            stxt = nr["search_text"].text().strip()
            src["search_text"] = stxt if stxt else None
            is_fixed = bool(nr["extract_mode_fixed"].isChecked())
            smi = nr["start_mode_ui"].currentIndex()
            src["start_mode"] = ["head", "position", "delimiter"][min(max(smi, 0), 2)]
            src["delimiter"] = nr["delimiter"].text().strip() or None
            src["part_index"] = max(1, int(nr["block"].value()))
            src["start_value"] = max(1, int(nr["start_pos"].value()))
            lmi = nr["length_mode_ui"].currentIndex()
            src["length_mode"] = ["char", "count", "end"][min(max(lmi, 0), 2)]
            ln_txt = nr["length_value_edit"].text().strip()
            lm = str(src["length_mode"])
            if is_fixed:
                src["length_value"] = ln_txt if ln_txt else None
            elif lm == "char":
                src["length_value"] = ln_txt if ln_txt else None
            elif lm == "count":
                try:
                    src["length_value"] = int(ln_txt) if ln_txt else None
                except ValueError:
                    src["length_value"] = None
            else:
                src["length_value"] = None
            src.pop("pattern", None)
            src.pop("replacement", None)
            src.pop("keyword_logic", None)
            src.pop("search_keywords", None)
            for k in ("start", "length"):
                src.pop(k, None)
            p = self._source_ui_bucket(src)
            p["name_checks"] = [cb.text() for cb in nr["name_checks"] if cb.isChecked()]
            vsn = nr.get("value_shape_script")
            if vsn is not None:
                p["value_shape_script"] = vsn.text().strip()
            p["extract_mode"] = "fixed" if is_fixed else "extract"
            p["path_item"] = nr["path_item"].currentText()
            p["write_mode_name_idx"] = int(nr["write_mode_name"].currentIndex())

    def get_item(self) -> dict[str, Any]:
        """編集結果の項目辞書を返す（sources・write_mode）。match_keys は更新しない。"""
        cur = self._current_source_index
        if cur >= 0 and cur < len(self._sources_data):
            self._apply_form_to_source(cur, include_scenario_name=True)
        out_sources: list[dict[str, Any]] = []
        for s in self._sources_data:
            one = copy.deepcopy(s)
            one.pop("registered", None)
            out_sources.append(one)
        return {"sources": out_sources, "write_mode": self._current_write_mode_key()}

    def _current_write_mode_key(self) -> str:
        """現在選択ソースの書込みモード内部キー（セル4種／名前2種）。"""
        cur = self._current_source_index
        if cur < 0 or cur >= len(self._sources_data):
            return self._item_write_mode_hint
        st = str(self._sources_data[cur].get("type") or "cell").strip().lower()
        if st == "name_extract":
            cb = self._name_refs["write_mode_name"]
            d = cb.currentData()
            return str(d) if d else "fill_in"
        cb = self._cell_refs["write_mode_cell"]
        d = cb.currentData()
        return str(d) if d else "fill_in"


class _DataAggDoneDialog(QDialog):
    """一括実行・ステップ実行の完了通知をモーダルで表示するダイアログ。"""

    def __init__(
        self,
        req: dict[str, Any],
        parent_hwnd: int,
        sheet_id: str,
        done_cfg: dict[str, Any],
    ) -> None:
        super().__init__()
        self._req = req or {}
        self._parent_hwnd = int(parent_hwnd or 0)
        self._done_cfg = done_cfg or {}
        title = str(self._req.get("title") or self._done_cfg.get("TITLE") or "データ集約").strip()
        self.setWindowTitle(title)
        message = str(self._req.get("message") or "").strip()
        from ui_qt.ui_common import (
            _icon_size_pixels_from_config,
            _normalize_message_newlines,
            _warning_icon_pixmap,
            apply_window_config,
        )
        lay = QVBoxLayout(self)
        icon_key = str(self._done_cfg.get("ICON") or "").strip()
        if icon_key:
            try:
                sz = _icon_size_pixels_from_config(self._done_cfg.get("ICON_SIZE"), default_pixels=24)
                px = _warning_icon_pixmap(self.style(), icon_key, sz)
                if px is not None:
                    icon_lbl = QLabel(self)
                    icon_lbl.setPixmap(px)
                    icon_lbl.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
                    lay.addWidget(icon_lbl)
            except Exception:
                pass
        msg_lbl = QLabel(_normalize_message_newlines(message) if message else "完了しました。")
        msg_lbl.setWordWrap(True)
        msg_lbl.setMinimumWidth(280)
        lay.addWidget(msg_lbl)
        row_btn = QHBoxLayout()
        row_btn.addStretch(1)
        btn_ok = QPushButton(str(self._done_cfg.get("BTN_OK") or "OK"))
        btn_ok.clicked.connect(self._on_ok)
        row_btn.addWidget(btn_ok)
        lay.addLayout(row_btn)
        try:
            apply_window_config(self, self._done_cfg, self._parent_hwnd, "DONE")
        except Exception:
            pass
        win_cfg = self._done_cfg.get("WINDOW") or {}
        w = int(win_cfg.get("DEFAULT_WIDTH") or 360)
        h = int(win_cfg.get("DEFAULT_HEIGHT") or 120)
        if w > 0 and h > 0:
            self.resize(w, h)

    def _on_ok(self) -> None:
        """OK: Excel 子 HWND ロック解除と前景 nudge を明示してから閉じる（closeEvent 非経路の取りこぼし対策）。"""
        try:
            self.hide()
        except Exception:
            pass
        if self._parent_hwnd:
            try:
                from ui_qt.ui_common import enable_excel_window, focus_excel_after_modal_close

                enable_excel_window(self._parent_hwnd, True)
                focus_excel_after_modal_close(self._parent_hwnd)
            except Exception:
                pass
        self.accept()

    def showEvent(self, event: Any) -> None:
        """表示時に Excel 中央に配置し、WINDOW.EXCEL_LOCK に従い Excel 子 HWND を無効化する。"""
        super().showEvent(event)
        try:
            from ui_qt.ui_notification_sound import play_notification_on_widget

            play_notification_on_widget(self)
        except Exception:
            pass
        if self._parent_hwnd:
            try:
                from ui_qt.ui_common import (
                    center_on_excel,
                    enable_excel_window,
                    want_excel_child_hwnd_lock_while_modal,
                )

                center_on_excel(
                    self,
                    self._parent_hwnd,
                    _excel_rect_tuple_from_req(self._req),
                )
                if want_excel_child_hwnd_lock_while_modal(self._done_cfg.get("WINDOW") or {}):
                    enable_excel_window(self._parent_hwnd, False)
            except Exception:
                pass

    def closeEvent(self, event: Any) -> None:
        if self._parent_hwnd:
            try:
                from ui_qt.ui_common import enable_excel_window
                enable_excel_window(self._parent_hwnd, True)
            except Exception:
                pass
        super().closeEvent(event)


class _DataAggStepPopupDialog(QDialog):
    """ステップ実行用ポップ。現在のステップ番号・項目名・参照ファイル・取得値プレビュー・次へ/中止。"""

    def __init__(
        self,
        req: dict[str, Any],
        parent_hwnd: int,
        sheet_id: str,
        step_cfg: dict[str, Any],
    ) -> None:
        super().__init__()
        self._req = req or {}
        self._parent_hwnd = int(parent_hwnd or 0)
        self._step_cfg = step_cfg or {}
        self._result: dict[str, Any] = {"action": "abort"}
        title = str(req.get("title") or self._step_cfg.get("TITLE") or "ステップ実行").strip()
        self.setWindowTitle(title)
        lay = QVBoxLayout(self)
        step_index = req.get("step_index", 0)
        item_name = str(req.get("item_name") or "").strip()
        lay.addWidget(QLabel("ステップ: %s" % (step_index + 1)))
        lay.addWidget(QLabel("項目: %s" % (item_name or "-")))
        ref_files = req.get("ref_files") or []
        if ref_files:
            lay.addWidget(QLabel("参照ファイル:"))
            for f in ref_files[:5]:
                lay.addWidget(QLabel("  %s" % f))
        preview = req.get("preview_values") or []
        if preview:
            lay.addWidget(QLabel("プレビュー: %s" % (preview[:5],)))
        row_btn = QHBoxLayout()
        row_btn.addStretch(1)
        btn_next = QPushButton(str(self._step_cfg.get("BTN_NEXT") or "次へ"))
        btn_next.clicked.connect(self._on_next)
        btn_abort = QPushButton(str(self._step_cfg.get("BTN_ABORT") or "中止"))
        btn_abort.clicked.connect(self._on_abort)
        row_btn.addWidget(btn_next)
        row_btn.addWidget(btn_abort)
        lay.addLayout(row_btn)
        win_cfg = self._step_cfg.get("WINDOW") or {}
        w = int(win_cfg.get("DEFAULT_WIDTH") or 480)
        h = int(win_cfg.get("DEFAULT_HEIGHT") or 320)
        if w > 0 and h > 0:
            self.resize(w, h)

    def showEvent(self, event: Any) -> None:
        super().showEvent(event)
        ph = int(self._parent_hwnd or 0)
        if not ph:
            return
        try:
            from ui_qt.ui_common import center_on_excel

            center_on_excel(self, ph, _excel_rect_tuple_from_req(self._req))
        except Exception:
            pass

    def _on_next(self) -> None:
        self._result = {"action": "next"}
        self.accept()

    def _on_abort(self) -> None:
        self._result = {"action": "abort"}
        self.reject()

    def get_result(self) -> dict[str, Any]:
        return self._result


class _DataAggProgressWrapper:
    """進捗ダイアログのラップ。show / get_result を提供。"""

    def __init__(self, progress_dlg: Any) -> None:
        self._dlg = progress_dlg

    def show(self) -> None:
        self._dlg.show()
        try:
            from PySide6.QtWidgets import QApplication
            QApplication.processEvents()
        except Exception:
            pass

    def get_result(self) -> dict[str, Any]:
        return getattr(self._dlg, "get_result", lambda: {})()


def create_dialog(
    req_dict: dict[str, Any] | None,
    parent_hwnd: int,
    sheet_id: str,
) -> QDialog | _DataAggProgressWrapper:
    """
    【概要】
      ui_server から呼ばれ、action に応じてメイン画面・進捗・完了・ステップ実行ポップを生成する。
    【引数】
      req_dict: リクエスト（action, message, step_index, item_name, ref_files, preview_values 等）。
      parent_hwnd: 親ウィンドウ（Excel）の HWND。
      sheet_id: 対象シート識別子。
    【戻り値】
      生成したダイアログまたはラッパー。show() / exec() / get_result() のいずれかで表示・結果取得。
    """
    req = req_dict or {}
    action = str(req.get("action", "") or "").strip().lower()
    cfg = _get_cfg()
    main_cfg = (cfg or {}).get("MAIN") or {}
    screens = (cfg or {}).get("SCREENS") or {}

    logger.info(
        "[DATA_AGG_UI] create_dialog enter action=%s parent_hwnd=%s sheet_id=%s ui_pid=%s",
        action,
        int(parent_hwnd or 0),
        str(sheet_id or ""),
        os.getpid(),
    )
    try:
        _data_agg_ui_diag.info(
            "[DATA_AGG_TRACE] create_dialog enter action=%s parent_hwnd=%s sheet_id=%s "
            "wall_perf_s=%.6f ui_pid=%s",
            action,
            int(parent_hwnd or 0),
            str(sheet_id or ""),
            time.perf_counter(),
            os.getpid(),
        )
    except Exception:
        pass

    if action == "main":
        from ui_qt.ui_common import _deep_merge

        root_win = (cfg or {}).get("WINDOW") or {}
        main_win = main_cfg.get("WINDOW") or {}
        win_cfg = _deep_merge(dict(root_win), dict(main_win))
        # シート操作併用: EXCEL_LOCK は JSON に書かずコードで false（want_excel_child_hwnd_lock_while_modal）
        win_cfg = dict(win_cfg)
        win_cfg["EXCEL_LOCK"] = False
        ph = int(parent_hwnd or 0)
        t_create0 = time.perf_counter()
        t_create_prev = t_create0
        dlg = _DataAggMainWindow(req, ph, str(sheet_id or ""), main_cfg, win_cfg)
        t_create_prev = _log_data_agg_create_dialog_phase(
            "main_window_ready",
            t0=t_create0,
            t_prev=t_create_prev,
            parent_hwnd=ph,
        )
        try:
            from ui_qt.ui_common import prepare_dialog_excel_center_before_show

            prepare_dialog_excel_center_before_show(
                dlg, ph, _excel_rect_tuple_from_req(req), win_cfg
            )
        except Exception:
            pass
        t_create_prev = _log_data_agg_create_dialog_phase(
            "prepare_done",
            t0=t_create0,
            t_prev=t_create_prev,
            parent_hwnd=ph,
        )
        try:
            from ui_qt.ipc_file import write_waitform_ready_signal

            write_waitform_ready_signal(int(ph or 0))
        except Exception:
            pass
        t_create_prev = _log_data_agg_create_dialog_phase(
            "waitform_written",
            t0=t_create0,
            t_prev=t_create_prev,
            parent_hwnd=ph,
        )
        dlg._excel_create_probe_t0 = t_create0
        dlg.show()
        t_create_prev = _log_data_agg_create_dialog_phase(
            "show_done",
            t0=t_create0,
            t_prev=t_create_prev,
            parent_hwnd=ph,
        )
        try:
            from ui_qt.ui_common import _keep_modeless

            _keep_modeless(dlg, exclude_from_bulk_close=True)
        except Exception:
            pass
        _log_data_agg_create_dialog_phase(
            "create_dialog_done",
            t0=t_create0,
            t_prev=t_create_prev,
            parent_hwnd=ph,
        )
        return dlg

    if action == "progress":
        from ui_qt.ui_common import (
            create_progress_dialog,
            merge_screen_cfg_window_from_root,
            want_excel_child_hwnd_lock_while_modal,
        )

        progress_cfg = merge_screen_cfg_window_from_root(
            cfg, "PROGRESS", sheet_interaction_excel_unlock=True
        )
        req_progress = dict(req)
        try:
            req_progress["excel_lock"] = want_excel_child_hwnd_lock_while_modal(
                (progress_cfg or {}).get("WINDOW") or {}
            )
        except Exception:
            pass
        main_parent = _find_data_agg_main_window(str(sheet_id or ""), int(parent_hwnd or 0))
        if main_parent is not None:
            req_progress["center_on_parent_widget"] = True
            req_progress["no_native_window"] = True
            req_progress.setdefault("close_parent_when_done", False)
        dlg = create_progress_dialog(
            req_progress,
            int(parent_hwnd or 0),
            parent_widget=main_parent,
            progress_cfg=progress_cfg,
        )
        return _DataAggProgressWrapper(dlg)

    if action == "done":
        from ui_qt.ui_common import merge_screen_cfg_window_from_root

        done_cfg = merge_screen_cfg_window_from_root(
            cfg, "DONE", sheet_interaction_excel_unlock=True
        )
        return _DataAggDoneDialog(req, int(parent_hwnd or 0), str(sheet_id or ""), done_cfg)

    if action == "step_popup":
        step_cfg = screens.get("STEP_POPUP") or {}
        return _DataAggStepPopupDialog(req, int(parent_hwnd or 0), str(sheet_id or ""), step_cfg)

    raise ValueError("ui_data_agg: unknown action %r" % action)
