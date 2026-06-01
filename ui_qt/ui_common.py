# -*- coding: utf-8 -*-
"""
Python: 3.12
Module: ui_qt/ui_common.py
Created: 2026-02-12
Updated: 2026-05-05
Version: 0.2.85
Purpose:
  Qt UI サーバ側の「表示共通」を集約する。
  - owner 設定（Excel HWND）
  - 画面中央表示（Excel ウィンドウ基準）
  - DPI/マルチモニタ対策のための“入口”
  - 最前面/モーダル等の挙動パラメータ適用
  - create_dialog はハブとして action に応じ ui_dialog_* へ委譲。Done/Progress 実装は ui_dialog_done / ui_dialog_progress に移管。

History (latest 3):
  - 0.2.85 (2026-05-05) prepare_dialog_excel_center_before_show: プロパティ _hc_prepare_skip_ensure_front=true のとき prepare 内 ensure_front をスキップ（画面個別のちらつき切り分け用）。
  - 0.2.84 (2026-05-05) ensure_front: 非表示ウィンドウでは SetWindowPos に SWP_SHOWWINDOW を付けない。prepare 中の先行可視化（ちらつき）を抑止。
  - 0.2.83 (2026-05-03) Excel 前景追従（EXCEL_FRONT_FOLLOW／Win32 前景フック）を廃止。apply_window_config は遅延オーナー＋ensure_front＋軽量 raise。merge_screen_cfg_window_from_root に sheet_interaction_excel_unlock。teardown_feature_ui_shared_state からフック停止を削除。
  - 0.2.82 (2026-05-03) WINDOW.EXCEL_LOCK にキー短縮（旧 EXCEL_CHILD_HWND_LOCK_WHILE_MODAL は EXCEL_LOCK 未指定時のみ読取互換）。
  - 0.2.81 (2026-05-03) done_dialog_show_event_on_excel: WINDOW.EXCEL_CHILD_HWND_LOCK_WHILE_MODAL（既定 true）で enable_excel_window(False) をスキップ可。子 HWND ロックは同期の ensure_front 系の後に実行（EXCEL_FRONT_FOLLOW 時の前面取りこぼし緩和）。
  - 0.2.80 (2026-05-03) enable_excel_window: HC_UI_EXCEL_LOCK_DIAG 時に hc_csv_tool.diag.ui_excel_lock へ [UI_EXCEL_LOCK] sym_type=A（子 HWND ロック切替）を出力。ensure_front: ヘルプは attach_input 後に nudge_top_level_to_foreground（症状 B）。
  - 0.2.79 (2026-05-03) ensure_front: プロパティ _hc_help_dialog のとき SetForegroundWindow 後に AttachThreadInput 経由の再試行（core_w32.set_foreground_window_attach_input）。bump_front_follow_deferred_ensure_generation を公開（ヘルプ exec 直前の遅延 ensure 棄却用）。
  - 0.2.78 (2026-05-03) _ff_try_restore_dialog_visibility_for_follow: _front_follow_dialog と不一致なら即 return（閉鎖後の蘇生防止）。ヘルプは ui_help で close 時 stop_front_follow_if_matches。
  - 0.2.77 (2026-05-03) _schedule_ensure_front_from_follow: not_visible 時に setVisible/show・最小化解除・raise/activate を試してから再試行（ヘルプ等が is_hidden のまま exhausted になる事象の緩和）。
  - 0.2.76 (2026-05-03) ウィンドウ前景・QTimer 遅延を config/ui_window_timing.json（core.ui_window_timing）から読込。欠落時は従来既定値。
  - 0.2.75 (2026-05-03) _handle_foreground_event: 前景 PID が Excel（_front_follow_excel_pid）と一致しないときは _front_follow_cooldown_acquire を呼ばず return。第三者プロセスの前景連打でクールダウンが消費され追従が飢える事象の緩和。
  - 0.2.74 (2026-05-03) 0.2.72〜0.73 の EXCEL_FRONT_FOLLOW 実験を撤回。_handle_foreground_event は 300ms クールダウン＋Excel pid 一致時のみ schedule の単一路線に戻す。
  - 0.2.73 (2026-05-03) _handle_foreground_event: Excel 前景一致時の最小間隔抑制等（0.2.74 で撤回）。
  - 0.2.72 (2026-05-03) _handle_foreground_event: Excel 前景クールダウン緩和・done_dialog 遅延 Excel 再有効化等（0.2.74 で撤回）。
  - 0.2.71 (2026-05-03) done_dialog_show_event_on_excel: EXCEL_FRONT_FOLLOW 時も ensure_front＋遅延再試行。_schedule_ensure_front_from_follow: not_visible 時の短い遅延再試行。
  - 0.2.70 (2026-05-03) EXCEL_FRONT_FOLLOW 診断: scheduled_ensure_skip not_visible に winId・is_hidden・window_state・schedule からの経過 ms・dlg_id。schedule_ensure_deferred / start_front_follow ok / ensure_front enter に dlg_id。
  - 0.2.69 (2026-05-03) EXCEL_FRONT_FOLLOW 診断: handle_fg / ensure_front_snap / cooldown_skip に前景 PID の exe パス短縮（core_w32.get_process_image_path_for_diag）。
  - 0.2.68 (2026-05-03) ensure_front: GW_OWNER が 0 のとき Win32 set_owner を再適用（Excel 前面化後に SFW が失敗する事象の緩和）。SetForegroundWindow 失敗時は短い遅延で ensure_front を最大 2 回再試行。
  - 0.2.67 (2026-05-03) ensure_front: hc_csv_tool.diag.front_follow に ensure_front_snap（前景 HWND/PID・親ルート・dlg_owner・sfw_ok）を段階出力。handle_fg の cooldown_skip に fg_hwnd/fg_pid を付与。
  - 0.2.66 (2026-05-03) apply_window_config: TOPMOST（または ALWAYS_IN_FRONT_OF_EXCEL）が true のとき EXCEL_FRONT_FOLLOW を開始しない（表示は TOPMOST 優先）。SHOW_IN_TASKBAR／_hc_show_taskbar／_set_owner_hwnd は変更なし。
  - 0.2.65 (2026-05-02) EXCEL_FRONT_FOLLOW: stop/start で ensure_gen をバンプし遅延 ensure_front を無効化。実行時は dlg 一致・可視・Shiboken。excel_pid=0 で globals クリア。diag: schedule/skip 理由。
  - 0.2.64 (2026-05-02) apply_window_config: EXCEL_FRONT_FOLLOW の destroyed を stop_front_follow_if_matches に変更（進捗破棄でプレビュー追従が止まる退行の修正）。
  - 0.2.63 (2026-05-02) stop_front_follow_if_matches: 進捗終了時にプレビューがフックを引き継いだ場合は stop しない。teardown の stop_front_follow_match_widget 引数。
  - 0.2.62 (2026-05-02) teardown_feature_ui_shared_state: 機能終了時に stop_front_follow・Excel 解除・モードレス削除を共通化（二重呼び可）。ProgressDialog 終了経路から先行呼び出し。
  - 0.2.61 (2026-05-02) WINDOW 設定キー FRONT_FOLLOW を EXCEL_FRONT_FOLLOW に改名（旧キーは読み取り互換）。前面化は TOPMOST と EXCEL_FRONT_FOLLOW の 2 キーで統一。diag タグ [EXCEL_FRONT_FOLLOW]。
  - 0.2.60 (2026-05-02) ensure_front / _handle_foreground_event: Shiboken.isValid で削除済みウィジェットを検出し stop_front_follow。ensure_front 例外時に deleted 系ならフック解除。
  - 0.2.59 (2026-05-02) _remove_from_modeless / stop_front_follow: ゴースト調査用に hc_csv.log へ [MODELESS_REMOVE]・[FRONT_FOLLOW] stop 行（winId・残件数）。
  - 0.2.58 (2026-05-02) _keep_modeless: exclude_from_bulk_close（データ集約メインを _close_all_modeless 対象外）。merge_screen_cfg_window_from_root: MAIN+ルート WINDOW+SCREENS の WINDOW 深マージ。
  - 0.2.57 (2026-05-02) done_dialog_show_event_on_excel: ALWAYS_IN_FRONT 時に ensure_front と遅延再試行（Undo 完了・hd_nr 確認などの前面取りこぼし対策）。
  - 0.2.56 (2026-04-13) apply_window_config: WINDOW.EXCEL_KEEP_FOREGROUND + EXCEL_KEEP_FOREGROUND_POLL_MS。表示中は間隔で ensure_front（Excel→ダイアログの順）。CHOICE 等 _skip_owner_front 画面でも利用可。
  - 0.2.55 (2026-04-13) WS_EX_TOOLWINDOW は core_w32 側でオプトイン化（既定オフ）。apply_window_config で WINDOW.USE_WS_EX_TOOLWINDOW_FOR_TASKBAR をプロパティ化。[UI_CAPTION] サンプルに GWL_EXSTYLE / ws_ex_toolwindow（符号なし hex）。
  - 0.2.54 (2026-04-13) HC_UI_WINDOW_CAPTION_DIAG: apply_window_config で意図フラグ・Win32 GWL_STYLE サンプル（遅延）・remove_min_max 発火ログ（hc_csv_diag.log [UI_CAPTION]）。
  - 0.2.53 (2026-04-13) apply_window_config: SHOW_MINIMIZE/MAXIMIZE 指定時に Qt.Window を併用（親付き QDialog で Windows タイトルバーに最小化/最大化を出しやすくする）。
  - 0.2.52 (2026-04-13) install_ribbon_startup_wait_dismiss_on_first_show: 初回 Show で notify_wait_form_ready（別 UI 分岐対応）。
  - 0.2.51 (2026-04-12) focus_excel_after_modal_close: core_w32.nudge_top_level_to_foreground + 遅延 0/80/200/350ms（AttachThreadInput 経由の SFW 再試行）。
  - 0.2.50 (2026-04-12) focus_excel_after_modal_close: 行／列削除完了などモーダル終了直後に CMD が前面に出る対策（遅延で Excel を nudge）。
  - 0.2.49 (2026-04-10) HC_UI_FG_DIAG: ensure_front / prepare_dialog / set_owner 前後と hc_csv_diag.log へ [UI_FG] 行（HWND・FG・sfw_ok）。
  - 0.2.48 (2026-04-09) prepare_dialog_excel_center_before_show から WA_DontShowOnScreen を除去（表示されない環境への対策）。中央・オーナー・前面は従来どおり 1 回適用。
  - 0.2.47 (2026-04-09) csv_sp 分割 MAIN: プロパティ _hc_csv_sp_split_main で遅延オーナー／再センタタイマー連打を抑止。prepare_dialog_excel_center_before_show 追加（show 前に Excel 中央を 1 回）。
  - 0.2.46 (2026-04-05) get_ui_config2(CSV_MG): ファイル直下 WINDOW を COMMON/MAIN.WINDOW より低優先でマージ（_get_cfg と整合）。
  - 0.2.45 (2026-04-05) apply_dialog_size_for_window_config: DEFAULT_WIDTH/HEIGHT が 0 の軸は sizeHint で確定（子の実寸を反映したあとに呼ぶ）。
  - 0.2.44 (2026-04-05) EXCEL_FRONT_FOLLOW: PROGRESS/DONE/DUPLICATE/WARNING で追従を有効化（遅延センタタイマーはスキップのまま）。CSV_MG_MAIN は EXCEL_FRONT_FOLLOW=true のときのみ追従（レイアウトは create_dialog 側）。
  - 0.2.43 (2026-04-07) CSV_MG_MAIN では EXCEL_FRONT_FOLLOW（start_front_follow）を付けない（チラつき抑制）。
  - 0.2.42 (2026-04-07) apply_window_config: screen_key CSV_MG_MAIN で遅延センタ／前面化タイマーを抑止（create_dialog 側の一括レイアウトと二重化しない）。
"""

from __future__ import annotations

import ctypes
import os
import re
from datetime import datetime
from pathlib import Path
from typing import Any, Optional, Tuple

from PySide6.QtCore import QObject, QTimer, Qt, QEvent
from PySide6.QtWidgets import QMessageBox, QWidget

from ui_qt import ipc_file

from core.ui_window_timing import get_ui_window_timings

# 変数: バージョン情報
__version__ = "0.2.83"


class _RibbonWaitFormDismissOnShow(QObject):
    """リボン経由で表示するウィンドウの初回 Show で VBA WaitForm を閉じる（別 UI 分岐でも同一フック）。"""

    def eventFilter(self, watched: QObject, event: QEvent) -> bool:  # noqa: N802
        if event.type() != QEvent.Type.Show:
            return False
        try:
            watched.removeEventFilter(self)
        except Exception:
            pass
        try:
            self.deleteLater()
        except Exception:
            pass
        try:
            from core.core_cursor import notify_wait_form_ready

            notify_wait_form_ready()
        except Exception:
            pass
        return False


def _resolve_widget_for_ribbon_wait_dismiss(target: Any) -> Optional[QWidget]:
    if isinstance(target, QWidget):
        return target
    inner = getattr(target, "_dlg", None)
    if isinstance(inner, QWidget):
        return inner
    return None


def install_ribbon_startup_wait_dismiss_on_first_show(target: Any) -> None:
    """create_dialog 直後に付与。実際に表示される QWidget の最初の Show で一度だけ notify_wait_form_ready。"""
    w = _resolve_widget_for_ribbon_wait_dismiss(target)
    if w is None:
        return
    if getattr(w, "_hc_ribbon_wait_dismiss_installed", False):
        return
    w._hc_ribbon_wait_dismiss_installed = True
    filt = _RibbonWaitFormDismissOnShow(w)
    filt.setParent(w)
    w.installEventFilter(filt)


class _DataAggBatchProgressCursorOnShow(QObject):
    """本番一括進捗の初回 Show: WaitForm を閉じつつ Excel 砂時計を維持・再武装する。"""

    def __init__(self, sheet_id: str, parent: QObject | None = None) -> None:
        super().__init__(parent)
        self._sheet_id = str(sheet_id or "")

    def eventFilter(self, watched: QObject, event: QEvent) -> bool:  # noqa: N802
        if event.type() != QEvent.Type.Show:
            return False
        try:
            watched.removeEventFilter(self)
        except Exception:
            pass
        try:
            self.deleteLater()
        except Exception:
            pass
        try:
            from core.core_cursor import (
                data_agg_batch_cursor_on,
                notify_wait_form_ready,
            )

            notify_wait_form_ready()
            data_agg_batch_cursor_on(self._sheet_id)
        except Exception:
            pass
        return False


def install_data_agg_batch_progress_cursor_on_show(
    target: Any,
    sheet_id: str,
) -> None:
    """本番一括 progress: 初回 Show で WaitForm 解除 + xlWait 再武装（ui_server の ribbon フックと併用可）。"""
    w = _resolve_widget_for_ribbon_wait_dismiss(target)
    if w is None:
        return
    if getattr(w, "_hc_data_agg_batch_cursor_on_show_installed", False):
        return
    w._hc_data_agg_batch_cursor_on_show_installed = True
    filt = _DataAggBatchProgressCursorOnShow(str(sheet_id or ""), w)
    filt.setParent(w)
    w.installEventFilter(filt)


# ===== ui_common trace (file based, logger independent) =====
_UI_COMMON_TRACE_PATH = (
    Path(os.environ.get("TEMP", ".")) / "csv_tool" / "ui_common_trace.log"
)


def _trace(msg: str) -> None:
    """Best-effort trace to ui_common_trace.log (never raises)."""
    try:
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
        line = f"{ts} {msg}\n"
        if core_log is not None and hasattr(core_log, "append_text_with_cap"):
            core_log.append_text_with_cap(_UI_COMMON_TRACE_PATH, line)
            return
        _UI_COMMON_TRACE_PATH.parent.mkdir(parents=True, exist_ok=True)
        with _UI_COMMON_TRACE_PATH.open("a", encoding="utf-8") as f:
            f.write(line)
    except Exception:
        pass


_trace(f"[loaded] __file__={__file__} version={__version__}")


def _trace_widget_rect(w: QWidget, label: str) -> None:
    """仮説確定用: ウィジェットの現在の画面矩形を取得してログに出す。どの処理で位置が変わったかを追うため。"""
    try:
        hwnd = int(w.winId()) if hasattr(w, "winId") else 0
        if not hwnd:
            _trace(f"[rect] {label} hwnd=0 (winId not ready)")
            return
        if os.name != "nt":
            return
        import ctypes
        from ctypes import wintypes
        r = wintypes.RECT()
        if ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(r)):
            _trace(f"[rect] {label} left={r.left} top={r.top} right={r.right} bottom={r.bottom}")
        else:
            _trace(f"[rect] {label} GetWindowRect failed")
    except Exception as e:
        _trace(f"[rect] {label} error={e!r}")


def _trace_monitor_and_excel_rect(parent_hwnd: int, dialog_w: QWidget) -> None:
    """解析用: Excel 矩形の中心とダイアログがあるモニタの中心を出し、どちらに近いか分かるようにする。"""
    try:
        if os.name != "nt":
            return
        import ctypes
        from ctypes import wintypes
        dialog_left = dialog_top = 0
        try:
            hwnd = int(dialog_w.winId()) if hasattr(dialog_w, "winId") else 0
            if hwnd:
                r = wintypes.RECT()
                if ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(r)):
                    dialog_left = r.left
                    dialog_top = r.top
        except Exception:
            pass
        rect = get_excel_rect(int(parent_hwnd or 0))
        if rect:
            el, et, er, eb = rect
            ecx = (el + er) // 2
            ecy = (et + eb) // 2
            _trace(f"[analyze] Excel rect center=({ecx},{ecy}) dialog topleft=({dialog_left},{dialog_top})")
        MONITOR_DEFAULTTONEAREST = 2
        try:
            class RECT(ctypes.Structure):
                _fields_ = [("left", wintypes.LONG), ("top", wintypes.LONG), ("right", wintypes.LONG), ("bottom", wintypes.LONG)]
            pt = wintypes.POINT()
            pt.x = dialog_left
            pt.y = dialog_top
            mon = ctypes.windll.user32.MonitorFromPoint(pt, MONITOR_DEFAULTTONEAREST)
            if mon:
                class MONITORINFO(ctypes.Structure):
                    _fields_ = [
                        ("cbSize", wintypes.DWORD),
                        ("rcMonitor", RECT),
                        ("rcWork", RECT),
                        ("dwFlags", wintypes.DWORD),
                    ]
                info = MONITORINFO()
                info.cbSize = ctypes.sizeof(MONITORINFO)
                if ctypes.windll.user32.GetMonitorInfoW(mon, ctypes.byref(info)):
                    r = info.rcWork
                    mcx = (r.left + r.right) // 2
                    mcy = (r.top + r.bottom) // 2
                    _trace(f"[analyze] monitor containing dialog: work center=({mcx},{mcy})")
        except Exception as e:
            _trace(f"[analyze] monitor info error={e!r}")
    except Exception as e:
        _trace(f"[analyze] error={e!r}")


# ===========================================================

try:
    from core import core_log
    _log = core_log.get_logger(__name__)
except Exception:  # pragma: no cover
    core_log = None  # type: ignore
    _log = None  # type: ignore

try:
    from core.core_log import get_diag_logger

    _diag_front_follow = get_diag_logger("hc_csv_tool.diag.front_follow")
    _diag_ui_fg = get_diag_logger("hc_csv_tool.diag.ui_fg")
except Exception:  # pragma: no cover
    _diag_front_follow = None  # type: ignore[misc, assignment]
    _diag_ui_fg = None  # type: ignore[misc, assignment]

def _ff_diag(msg: str, *args: object) -> None:
    if _diag_front_follow is None:
        return
    try:
        _diag_front_follow.info(msg, *args)
    except Exception:
        pass


def _dlg_hwnd_for_fg_diag(w: Optional[QWidget]) -> int:
    try:
        if w is not None and hasattr(w, "winId"):
            return int(w.winId() or 0)
    except Exception:
        pass
    return 0


def _ff_fg_exe_for_pid(pid: int) -> str:
    """前景 PID の実行ファイルパス短縮（診断用。取得失敗時は空）。"""
    try:
        p = int(pid or 0)
        if not p or _w32 is None or not hasattr(_w32, "get_process_image_path_for_diag"):
            return ""
        return str(_w32.get_process_image_path_for_diag(p) or "")
    except Exception:
        return ""


def _ff_diag_ensure_front_snapshot(
    phase: str,
    parent_hwnd: int,
    dlg_hwnd: int,
    *,
    sfw_ok: Optional[bool] = None,
) -> None:
    """
    ensure_front 内の前景・PID・トップルート・ダイアログ GW_OWNER を
    hc_csv_tool.diag.front_follow に出す（HC_UI_FG_DIAG 不要。Z 順切り分け用）。
    """
    if _diag_front_follow is None or os.name != "nt":
        return
    try:
        ph = int(parent_hwnd or 0)
        dh = int(dlg_hwnd or 0)
        fg = 0
        if _w32 is not None and hasattr(_w32, "get_foreground_window"):
            fg = int(_w32.get_foreground_window() or 0)
        fg_pid = 0
        ph_pid = 0
        dh_pid = 0
        dlg_owner = 0
        pr = 0
        fr = 0
        if _w32 is not None:
            if fg:
                fg_pid = int(_w32.get_window_pid(fg) or 0)
            if ph:
                ph_pid = int(_w32.get_window_pid(ph) or 0)
                pr = int(_w32.get_root_window(ph) or 0)
            if dh:
                dh_pid = int(_w32.get_window_pid(dh) or 0)
                dlg_owner = int(_w32.get_owner_hwnd(dh) or 0)
            if fg:
                fr = int(_w32.get_root_window(fg) or 0)
        fg_eq_dlg = int(bool(fg and dh and fg == dh))
        fg_root_eq_parent_root = int(bool(fr and pr and fr == pr))
        fg_exe = _ff_fg_exe_for_pid(fg_pid)
        if sfw_ok is None:
            _ff_diag(
                "[EXCEL_FRONT_FOLLOW] ensure_front_snap phase=%s fg=%s fg_pid=%s fg_exe=%r "
                "parent=%s parent_pid=%s parent_root=%s dlg=%s dlg_pid=%s dlg_owner=%s fg_root=%s "
                "fg_eq_dlg=%s fg_root_eq_parent_root=%s",
                phase,
                fg,
                fg_pid,
                fg_exe,
                ph,
                ph_pid,
                pr,
                dh,
                dh_pid,
                dlg_owner,
                fr,
                fg_eq_dlg,
                fg_root_eq_parent_root,
            )
        else:
            _ff_diag(
                "[EXCEL_FRONT_FOLLOW] ensure_front_snap phase=%s fg=%s fg_pid=%s fg_exe=%r "
                "parent=%s parent_pid=%s parent_root=%s dlg=%s dlg_pid=%s dlg_owner=%s fg_root=%s "
                "fg_eq_dlg=%s fg_root_eq_parent_root=%s sfw_ok=%s",
                phase,
                fg,
                fg_pid,
                fg_exe,
                ph,
                ph_pid,
                pr,
                dh,
                dh_pid,
                dlg_owner,
                fr,
                fg_eq_dlg,
                fg_root_eq_parent_root,
                int(bool(sfw_ok)),
            )
    except Exception:
        pass


def ui_fg_diag_enabled() -> bool:
    """HC_UI_FG_DIAG=1 のとき True（hc_csv_diag.log の単独有効化にも使う）。"""
    try:
        from core import core_env

        return core_env.ui_fg_diag_enabled()
    except Exception:
        return False


def log_ui_fg_phase(
    phase: str,
    parent_hwnd: int,
    w: Optional[QWidget] = None,
    *,
    sfw_ok: Optional[bool] = None,
    extra: str = "",
) -> None:
    """Excel 前面／Z 順の切り分け用。HC_UI_FG_DIAG=1 かつ Windows のとき hc_csv_diag.log に 1 行。"""
    if os.name != "nt" or _diag_ui_fg is None or not ui_fg_diag_enabled():
        return
    try:
        from core import core_w32 as _cw

        line = _cw.format_ui_fg_diag_line(
            phase,
            int(parent_hwnd or 0),
            _dlg_hwnd_for_fg_diag(w),
            sfw_ok=sfw_ok,
            extra=str(extra or ""),
        )
        _diag_ui_fg.info("%s", line)
    except Exception:
        pass

try:
    from core import core_cst as cst
except Exception:  # pragma: no cover
    try:
        from core import hc_cst as cst  # type: ignore
    except Exception:
        cst = object()  # type: ignore

# Win32 API: core_w32 を優先（親子関係・タスクバー非表示・前面追従に使用）
try:
    from core import core_w32 as _w32
except Exception:  # pragma: no cover
    try:
        from core import hc_w32 as _w32  # type: ignore
    except Exception:
        _w32 = None  # type: ignore


def _normalize_message_newlines(text: str) -> str:
    """共通仕様: \\n および改行→改行、\\t およびタブ→4文字空白。文末\\nは改行として有効（レイアウト側で空欄を追加すること）。"""
    if not text:
        return text
    s = str(text)
    if "\\n" in s:
        s = s.replace("\\n", "\n")
    if "\\t" in s:
        s = s.replace("\\t", "    ")
    # タブ文字は 4 文字分の空白として表示（共通仕様）
    if "\t" in s:
        s = s.replace("\t", "    ")
    return s


def apply_common_window_flags(w: QWidget) -> None:
    """
    Method Name : apply_common_window_flags
    Arguments   : w (QWidget)
    Return      : None
    機能概要    : hc_cst.UI_COMMON を参照して、基本フラグを適用する。
    """
    # 変数: 共通UI設定辞書
    ui = getattr(cst, "UI_COMMON", {}) or {}
    win = ui.get("WINDOW") if isinstance(ui.get("WINDOW"), dict) else {}
    # 変数: 最前面フラグの真偽値（旧 JSON キー ALWAYS_IN_FRONT_OF_EXCEL も読み取り互換）
    topmost = bool(win.get("TOPMOST", False) or win.get("ALWAYS_IN_FRONT_OF_EXCEL", False))
    # 変数: ウィンドウフラグのベース
    flags = Qt.Dialog

    # 判定コメント: TOPMOSTが要求されている場合
    if topmost:
        # 【目的】対象ウィンドウをOSの最前面に固定するため
        flags |= Qt.WindowStaysOnTopHint

    # 命令分離: ウィンドウへのフラグ適用
    w.setWindowFlags(flags)


def apply_common_window_style(w: QWidget, parent_hwnd: int) -> None:
    """
    Method Name : apply_common_window_style
    Arguments   : w (QWidget), parent_hwnd (int)
    Return      : None
    機能概要    : 共通の表示スタイルを適用する（flags + Excel中央配置）。
    """
    # 【目的】共通フラグの適用と中央配置を一括で行うため
    apply_common_window_flags(w)
    ui = getattr(cst, "UI_COMMON", {}) or {}
    win = ui.get("WINDOW") if isinstance(ui.get("WINDOW"), dict) else {}
    if bool(win.get("CENTER_ON_EXCEL", False)):
        center_on_excel(w, int(parent_hwnd or 0))


# ------------------------------------------------------------------------------
# Window config apply (screen-specific)
# ------------------------------------------------------------------------------


def _ui_caption_diag_enabled() -> bool:
    try:
        from core import core_env

        return core_env.ui_window_caption_diag_enabled()
    except Exception:
        return False


def _log_ui_caption_apply_enter(
    screen_key: str,
    parent_hwnd: int,
    win: dict,
    show_min: bool,
    show_max: bool,
    schedule_remove_min_max: bool,
    flags: Qt.WindowType,
) -> None:
    if not _ui_caption_diag_enabled():
        return
    try:
        from core.core_log import get_diag_logger

        log = get_diag_logger("hc_csv_tool.diag.ui_caption")
        keys = sorted(win.keys()) if isinstance(win, dict) else []
        keys_s = ",".join(keys[:48])
        if len(keys) > 48:
            keys_s += ",..."
        raw_sm = win.get("SHOW_MINIMIZE", "<missing>") if isinstance(win, dict) else "<n/a>"
        raw_sx = win.get("SHOW_MAXIMIZE", "<missing>") if isinstance(win, dict) else "<n/a>"
        try:
            qf = int(flags)
        except Exception:
            qf = -1
        log.info(
            "[UI_CAPTION] apply_enter screen_key=%s parent_hwnd=%s "
            "win_SHOW_MINIMIZE=%r win_SHOW_MAXIMIZE=%r resolved_show_min=%s resolved_show_max=%s "
            "schedule_win32_remove_min_max=%s qt_windowFlags=0x%x win_key_count=%d keys=[%s]",
            screen_key,
            int(parent_hwnd or 0),
            raw_sm,
            raw_sx,
            show_min,
            show_max,
            schedule_remove_min_max,
            qf,
            len(keys),
            keys_s,
        )
    except Exception:
        pass


def _log_ui_caption_hwnd_sample(
    w: QWidget, phase: str, screen_key: str, parent_hwnd: int
) -> None:
    if os.name != "nt" or not _ui_caption_diag_enabled():
        return
    try:
        from core.core_log import get_diag_logger

        log = get_diag_logger("hc_csv_tool.diag.ui_caption")
        hwnd = int(w.winId()) if hasattr(w, "winId") else 0
        st, hm, hx = (0, False, False)
        owner = 0
        ex_u = 0
        ex_tb = False
        if hwnd and _w32 is not None:
            if hasattr(_w32, "get_window_caption_style_summary"):
                st, hm, hx = _w32.get_window_caption_style_summary(hwnd)
            if hasattr(_w32, "get_window_exstyle_toolwindow"):
                _ex, ex_tb = _w32.get_window_exstyle_toolwindow(hwnd)
                ex_u = _ex & 0xFFFFFFFF
            if hasattr(_w32, "get_owner_hwnd"):
                try:
                    owner = int(_w32.get_owner_hwnd(hwnd) or 0)
                except Exception:
                    owner = 0
        st_u = st & 0xFFFFFFFF
        log.info(
            "[UI_CAPTION] sample phase=%s screen_key=%s parent_hwnd=%s hwnd=%s "
            "gwl_style=0x%08x ws_min_box=%s ws_max_box=%s gwl_exstyle=0x%08x ws_ex_toolwindow=%s gw_owner=%s",
            phase,
            screen_key,
            int(parent_hwnd or 0),
            hwnd,
            st_u,
            hm,
            hx,
            ex_u,
            ex_tb,
            owner,
        )
    except Exception:
        pass


def _schedule_ui_caption_hwnd_samples(
    w: QWidget, screen_key: str, parent_hwnd: int
) -> None:
    if os.name != "nt" or not _ui_caption_diag_enabled():
        return
    try:
        from PySide6.QtCore import QTimer

        _log_ui_caption_hwnd_sample(
            w, "t0_after_setWindowFlags", screen_key, parent_hwnd
        )
        _tw_cap = get_ui_window_timings()
        for _ms in _tw_cap.ui_caption_diagnostic_hwnd_sample_delays_ms:
            _msi = int(_ms)
            QTimer.singleShot(
                _msi,
                lambda m=_msi: _log_ui_caption_hwnd_sample(
                    w, f"t{m}ms", screen_key, parent_hwnd
                ),
            )
    except Exception:
        pass


def apply_window_config(
    w: QWidget, ui_cfg: dict, parent_hwnd: int, screen_key: str = ""
) -> None:
    """
    Method Name : apply_window_config
    Arguments   : w (QWidget), ui_cfg (dict), parent_hwnd (int), screen_key (str)
    Return      : None
    機能概要    : UI_COMMON + 画面固有の WINDOW 設定を統合して適用する。
    WINDOW.TOPMOST または ALWAYS_IN_FRONT_OF_EXCEL（読み取り互換）が true のとき WindowStaysOnTopHint を付与する。
    表示直後に遅延オーナー・ensure_front・軽量 raise で Excel オーナー兄弟として前面に寄せる（前景フックは使用しない）。
    SHOW_IN_TASKBAR は _hc_show_taskbar のみでオーナー経路と連動し本メソッドでは変更しない。
    """
    # 変数: ウィンドウ設定の辞書
    win = (ui_cfg or {}).get("WINDOW") or {}
    # 変数: 各種ボタン・タスクバーの表示制御フラグ（未指定時は非表示＝子画面向け）
    show_min = bool(win.get("SHOW_MINIMIZE", False))
    show_max = bool(win.get("SHOW_MAXIMIZE", False))
    # 変数: 最前面フラグ（TOPMOST またはヘルプ等用の ALWAYS_IN_FRONT_OF_EXCEL）
    topmost = bool(win.get("TOPMOST", False) or win.get("ALWAYS_IN_FRONT_OF_EXCEL", False))
    # 変数: タスクバー表示（既定は非表示）
    show_taskbar = bool(win.get("SHOW_IN_TASKBAR", False))
    # 変数: ×閉じるボタン表示（既定は表示。false で進捗画面などで非表示）
    show_close = bool(win.get("SHOW_CLOSE_BUTTON", True))

    # 変数: ウィンドウフラグの設定
    flags = Qt.Dialog
    # 親付き QDialog でも Windows で最小化/最大化ヒントをタイトルバーに出しやすくする
    if show_min or show_max:
        flags |= Qt.Window
    flags |= Qt.WindowTitleHint | Qt.WindowSystemMenuHint
    if show_close:
        flags |= Qt.WindowCloseButtonHint

    # 判定コメント: 最小化ボタン表示
    if show_min:
        flags |= Qt.WindowMinimizeButtonHint
    # 判定コメント: 最大化ボタン表示
    if show_max:
        flags |= Qt.WindowMaximizeButtonHint
    # 判定コメント: 最前面表示
    if topmost:
        flags |= Qt.WindowStaysOnTopHint

    # NOTE:
    #  - タスクバー非表示は主に Win32 owner（_set_owner_hwnd）。WS_EX_TOOLWINDOW はタイトルバー min/max を
    #    消すことがあるため既定では付けない（HC_USE_WS_EX_TOOLWINDOW_FOR_TASKBAR または WINDOW の
    #    USE_WS_EX_TOOLWINDOW_FOR_TASKBAR でオプトイン）。
    #  - Qt.Tool を強制するとボタン/システムメニュー等が変化しやすいため、ここでは付与しない

    schedule_remove_min_max = (
        not show_min
        and not show_max
        and _w32 is not None
        and hasattr(_w32, "set_window_style_remove_min_max")
    )
    _log_ui_caption_apply_enter(
        screen_key,
        parent_hwnd,
        win if isinstance(win, dict) else {},
        show_min,
        show_max,
        schedule_remove_min_max,
        flags,
    )

    # 命令分離: フラグの適用
    w.setWindowFlags(flags)
    _schedule_ui_caption_hwnd_samples(w, screen_key, parent_hwnd)

    # 【目的】最小化/最大化ボタンを確実に非表示にするため（Qtフラグだけでは出る環境がある）
    if schedule_remove_min_max:
        try:
            from PySide6.QtCore import QTimer

            sk = str(screen_key or "")

            def _remove_min_max() -> None:
                try:
                    hwnd = int(w.winId()) if hasattr(w, "winId") else 0
                    if hwnd:
                        if _ui_caption_diag_enabled():
                            try:
                                from core.core_log import get_diag_logger

                                get_diag_logger("hc_csv_tool.diag.ui_caption").info(
                                    "[UI_CAPTION] fired set_window_style_remove_min_max hwnd=%s screen_key=%s",
                                    hwnd,
                                    sk,
                                )
                            except Exception:
                                pass
                        _w32.set_window_style_remove_min_max(hwnd)
                except Exception:
                    pass

            QTimer.singleShot(
                int(get_ui_window_timings().apply_window_min_max_buttons_remove_delay_ms),
                _remove_min_max,
            )
        except Exception:
            pass

    # 【目的】Win32側のowner/タスクバー制御のため、表示方針を保持するため
    try:
        w.setProperty("_hc_show_taskbar", bool(show_taskbar))
    except Exception:
        pass
    try:
        if isinstance(win, dict) and "USE_WS_EX_TOOLWINDOW_FOR_TASKBAR" in win:
            w.setProperty(
                "_hc_use_ws_ex_toolwindow_for_taskbar",
                bool(win.get("USE_WS_EX_TOOLWINDOW_FOR_TASKBAR")),
            )
    except Exception:
        pass

    # 変数: リサイズ可否フラグ
    resizable = bool(win.get("RESIZABLE", True))
    # 変数: 初期サイズ（DEFAULT_WIDTH/DEFAULT_HEIGHT が 0 または未指定 = オートサイズ＝adjustSize/sizeHint に任せる）
    default_w = int(win.get("DEFAULT_WIDTH") or 0)
    default_h = int(win.get("DEFAULT_HEIGHT") or 0)

    # 判定コメント: リサイズ不可の場合
    if not resizable:
        if default_w > 0 and default_h > 0:
            # 【目的】設定で指定されたサイズで固定（進捗画面等でオートサイズ化を防ぐ）
            try:
                w.setFixedSize(default_w, default_h)
            except Exception:
                pass
        else:
            # 【目的】コンテンツサイズに合わせて固定化するため
            w.adjustSize()
            sh = w.sizeHint()
            if sh.width() > 0 and sh.height() > 0:
                w.setFixedSize(sh)
    else:
        # 【目的】自由なリサイズを許可するため
        w.setMinimumSize(0, 0)
        w.setMaximumSize(16777215, 16777215)
        if default_w > 0 and default_h > 0:
            try:
                w.resize(default_w, default_h)
            except Exception:
                pass

    # 変数: 起動位置指定文字列
    pos = str(win.get("STARTUP_POSITION") or "").strip().lower()
    # 変数: ジオメトリ保存用のキー
    storage_key = str(win.get("STORAGE_KEY") or screen_key or "").strip() or str(
        screen_key or ""
    )
    # DONE/PROGRESS/WARNING/DUPLICATE/REPORT/HELP/CHOICE は生成側で show 前に中央・オーナーを済ませるため、遅延タイマーによる二重配置を避ける
    _skip_owner_front = screen_key in (
        "DONE",
        "PROGRESS",
        "WARNING",
        "DUPLICATE",
        "REPORT",
        "HELP",
        "CHOICE",
    )
    # CSV_MG_MAIN: ui_csv_mg.create_dialog が show 前にオーナー・中央・前面化を一括実行する。
    # csv_sp 分割 MAIN: 同上（ui_server で prepare_dialog_excel_center_before_show を 1 回）。
    # DATA_AGG_MAIN / HD_NR_CONFIRM: 各 create_dialog が show/exec 前に prepare を 1 回。
    _csv_sp_split_main = False
    try:
        _csv_sp_split_main = bool(w.property("_hc_csv_sp_split_main"))
    except Exception:
        _csv_sp_split_main = False
    _defer_layout_to_creator = (
        screen_key == "CSV_MG_MAIN"
        or screen_key == "DATA_AGG_MAIN"
        or screen_key == "HD_NR_CONFIRM"
        or (
            _csv_sp_split_main and str(screen_key or "").strip().upper() == "MAIN"
        )
    )
    _trace(f"[apply_window_config] screen_key={screen_key!r} _skip_owner_front={_skip_owner_front} _defer_layout_to_creator={_defer_layout_to_creator} CENTER_ON_EXCEL={bool(win.get('CENTER_ON_EXCEL', False))}")

    # 判定コメント: 前回位置記憶の場合
    if pos == "remember_last":
        # 【目的】保存された位置の復元を試みるため
        if not restore_geometry(storage_key, w):
            if (
                not _skip_owner_front
                and not _defer_layout_to_creator
                and bool(win.get("CENTER_ON_EXCEL", False))
            ):
                center_on_excel(w, parent_hwnd)
    else:
        # 判定コメント: 明示的な中央配置指定がある場合
        if not _skip_owner_front and not _defer_layout_to_creator and (
            bool(win.get("CENTER_ON_EXCEL", False)) or pos in (
                "center_on_excel",
                "center",
            )
        ):
            center_on_excel(w, parent_hwnd)

    # 【目的】イベントループ開始後に確実にExcelへ所有権を紐付け、OS標準のZオーダー連動に完全に委ねるため
    # DONE/PROGRESS のときは _set_owner_hwnd / ensure_front をスキップ（オーナー設定で OS が位置を上書きするため）
    from PySide6.QtCore import QTimer

    if not _skip_owner_front and not _defer_layout_to_creator:
        def _apply_owner():
            _set_owner_hwnd(w, parent_hwnd)

        def _apply_front():
            ensure_front(w, parent_hwnd)

        _tw_aw = get_ui_window_timings()
        _o_ms = _tw_aw.apply_window_owner_hwnd_timer_ms
        _f_ms = _tw_aw.apply_window_ensure_front_timer_ms
        QTimer.singleShot(int(_o_ms[0]), _apply_owner)
        QTimer.singleShot(int(_o_ms[1]), _apply_owner)
        QTimer.singleShot(int(_f_ms[0]), _apply_front)
        QTimer.singleShot(int(_o_ms[2]), _apply_owner)
        QTimer.singleShot(int(_o_ms[3]), _apply_owner)
        QTimer.singleShot(int(_f_ms[1]), _apply_front)

    # 【目的】owner 設定後などで OS がウィンドウ位置をずらす場合があるため、中央指定時は遅延で再配置する。DONE/PROGRESS は各ダイアログ側で専用処理。
    if (
        not _skip_owner_front
        and not _defer_layout_to_creator
        and bool(win.get("CENTER_ON_EXCEL", False))
    ):
        QTimer.singleShot(
            int(get_ui_window_timings().apply_window_center_on_excel_recenter_delay_ms),
            lambda: center_on_excel(w, parent_hwnd),
        )

    # ツール同士では直近表示を手前に（前景フックは廃止。軽量の raise のみ）
    if not _skip_owner_front and not _defer_layout_to_creator:

        def _nudge_tool_window_front() -> None:
            try:
                w.raise_()
                w.activateWindow()
            except Exception:
                pass

        QTimer.singleShot(0, _nudge_tool_window_front)

    # 【目的】他アプリが前面に出たあとも、一定間隔で Excel→ダイアログの順に戻す（CHOICE 等 _skip_owner_front でも有効）
    excel_keep_fg = bool(win.get("EXCEL_KEEP_FOREGROUND", False))
    try:
        excel_keep_poll = int(win.get("EXCEL_KEEP_FOREGROUND_POLL_MS", 900) or 900)
    except (TypeError, ValueError):
        excel_keep_poll = 900
    excel_keep_poll = max(250, min(8000, excel_keep_poll))
    ph_keep = int(parent_hwnd or 0)
    if excel_keep_fg and ph_keep and _w32 is not None:

        def _arm_excel_keep_foreground() -> None:
            old_t = getattr(w, "_hc_excel_keep_fg_timer", None)
            if old_t is not None:
                try:
                    old_t.stop()
                    old_t.deleteLater()
                except Exception:
                    pass
            t = QTimer(w)
            t.setInterval(excel_keep_poll)

            def _tick_keep_excel_fg() -> None:
                try:
                    if not w.isVisible():
                        return
                    ensure_front(w, ph_keep)
                except Exception:
                    pass

            t.timeout.connect(_tick_keep_excel_fg)  # type: ignore[attr-defined]
            t.start()
            setattr(w, "_hc_excel_keep_fg_timer", t)
            QTimer.singleShot(
                int(get_ui_window_timings().excel_keep_foreground_initial_tick_immediate_ms),
                _tick_keep_excel_fg,
            )

            def _stop_excel_keep_fg(*_: object) -> None:
                try:
                    t.stop()
                except Exception:
                    pass
                try:
                    t.deleteLater()
                except Exception:
                    pass

            try:
                w.destroyed.connect(_stop_excel_keep_fg)  # type: ignore[attr-defined]
            except Exception:
                pass

        QTimer.singleShot(
            int(get_ui_window_timings().excel_keep_foreground_arm_timer_delay_ms),
            _arm_excel_keep_foreground,
        )


def excel_rect_tuple_from_req(
    req: Optional[dict[str, Any]],
) -> Optional[Tuple[int, int, int, int]]:
    """svc 送信の excel_rect（GetWindowRect 4 整数）を tuple に正規化。無効時は None。"""
    if not isinstance(req, dict):
        return None
    er = req.get("excel_rect")
    if er is None:
        return None
    try:
        if len(er) >= 4:
            return (int(er[0]), int(er[1]), int(er[2]), int(er[3]))
    except (TypeError, ValueError):
        pass
    return None


def prepare_dialog_excel_center_before_show(
    w: QWidget,
    parent_hwnd: int,
    rect_override: Optional[Tuple[int, int, int, int]] = None,
    window_cfg: Optional[dict] = None,
) -> None:
    """
    exec/show 前にサイズ・HWND を確定し、Excel 中央＋オーナー＋前面を 1 回だけ適用する。

    window_cfg に WINDOW 相当の辞書を渡した場合、CENTER_ON_EXCEL が false なら中央配置をスキップする。
    未指定時は従来どおり中央配置する。

    DEFAULT_WIDTH / DEFAULT_HEIGHT のいずれかが正のときは apply_dialog_size_for_window_config で
    サイズを確定する（先頭の無条件 adjustSize だと apply_window_config や __init__ の resize が潰れるため）。

    WA_DontShowOnScreen は Windows 上で exec 後も表示されない事例があったため使わない。
    """
    ph = int(parent_hwnd or 0)
    if not ph:
        return
    win = window_cfg if isinstance(window_cfg, dict) else {}
    center_ok = bool(win.get("CENTER_ON_EXCEL", True))
    log_ui_fg_phase("prepare:enter", ph, w)
    dw = int(win.get("DEFAULT_WIDTH") or 0)
    dh = int(win.get("DEFAULT_HEIGHT") or 0)
    if dw > 0 or dh > 0:
        try:
            apply_dialog_size_for_window_config(w, win)
        except Exception:
            pass
    else:
        try:
            w.adjustSize()
            w.updateGeometry()
        except Exception:
            pass
    try:
        if hasattr(w, "winId"):
            w.winId()
    except Exception:
        pass
    if center_ok:
        try:
            center_on_excel(w, ph, rect_override)
        except Exception:
            pass
    log_ui_fg_phase("prepare:after_center", ph, w)
    try:
        _set_owner_hwnd(w, ph)
    except Exception:
        pass
    log_ui_fg_phase("prepare:after_owner", ph, w)
    _skip_ensure = False
    try:
        _skip_ensure = bool(w.property("_hc_prepare_skip_ensure_front"))
    except Exception:
        _skip_ensure = False
    if not _skip_ensure:
        try:
            ensure_front(w, ph)
        except Exception:
            pass
    log_ui_fg_phase("prepare:after_ensure_front", ph, w)


def want_excel_child_hwnd_lock_while_modal(win_cfg: dict) -> bool:
    """
    WINDOW.EXCEL_LOCK が false のときだけロックしない。
    未指定時は true（従来どおりモーダル表示中は Excel 子 HWND を無効化＝操作不能）。
    互換: 旧キー EXCEL_CHILD_HWND_LOCK_WHILE_MODAL は EXCEL_LOCK が JSON に無いときのみ参照。
    """
    w = win_cfg or {}
    if "EXCEL_LOCK" in w:
        return bool(w.get("EXCEL_LOCK", True))
    return bool(w.get("EXCEL_CHILD_HWND_LOCK_WHILE_MODAL", True))


def done_dialog_show_event_on_excel(
    w: QWidget,
    parent_hwnd: int,
    req: dict,
    screen_cfg: dict,
) -> None:
    """
    完了／警告／ヘルプ等モーダルの showEvent 共通処理。
    WINDOW.CENTER_ON_EXCEL（既定 true）で Excel 基準の中央配置、
    TOPMOST / ALWAYS_IN_FRONT_OF_EXCEL で raise＋activate に加え、
    prepare 相当の ensure_front と短い遅延での再試行を行う（SetForegroundWindow の失敗緩和）。
    TOPMOST でないモーダルも parent があるときは ensure_front＋遅延再試行する（Excel 背後に残るのを緩和）。
    WINDOW.EXCEL_LOCK（既定 true）が false のときは enable_excel_window(False) を呼ばない（各画面の WINDOW に手動指定。旧キー互換あり）。
    ロックは同期の ensure_front／raise の後に行い、前面化のあとで子 HWND を無効化する。
    """
    ph = int(parent_hwnd or 0)
    win_cfg = (screen_cfg or {}).get("WINDOW") or {}
    want_lock = want_excel_child_hwnd_lock_while_modal(win_cfg)
    want_front = bool(win_cfg.get("TOPMOST") or win_cfg.get("ALWAYS_IN_FRONT_OF_EXCEL"))
    if ph:
        try:
            if bool(win_cfg.get("CENTER_ON_EXCEL", True)):
                center_on_excel(w, ph, excel_rect_tuple_from_req(req))
        except Exception:
            pass
    if want_front:
        try:
            w.raise_()
            w.activateWindow()
        except Exception:
            pass
        if ph:

            def _ensure_if_visible() -> None:
                try:
                    if w.isVisible():
                        ensure_front(w, ph)
                except Exception:
                    pass

            try:
                ensure_front(w, ph)
            except Exception:
                pass
            for _ms in get_ui_window_timings().done_dialog_show_on_excel_ensure_front_extra_delays_ms:
                QTimer.singleShot(int(_ms), _ensure_if_visible)
    elif ph and not want_front:
        try:
            w.raise_()
            w.activateWindow()
        except Exception:
            pass

        def _ensure_follow_if_visible() -> None:
            try:
                if w.isVisible():
                    ensure_front(w, ph)
            except Exception:
                pass

        try:
            ensure_front(w, ph)
        except Exception:
            pass
        for _ms in get_ui_window_timings().done_dialog_show_on_excel_ensure_front_extra_delays_ms:
            QTimer.singleShot(int(_ms), _ensure_follow_if_visible)
    if ph and want_lock:
        try:
            enable_excel_window(ph, False)
        except Exception:
            pass


def apply_dialog_size_for_window_config(w: QWidget, win_cfg: dict) -> None:
    """
    DEFAULT_WIDTH / DEFAULT_HEIGHT が 0 の軸はレイアウトの sizeHint を使い、
    正の値の軸はそのピクセルでウィンドウサイズを確定する。
    子ウィジェットの最小サイズ・内容反映後に呼ぶこと。
    """
    win = win_cfg or {}
    dw = int(win.get("DEFAULT_WIDTH") or 0)
    dh = int(win.get("DEFAULT_HEIGHT") or 0)
    try:
        w.updateGeometry()
    except Exception:
        pass
    try:
        lay = w.layout()
        if lay is not None:
            lay.activate()
    except Exception:
        pass
    try:
        w.adjustSize()
    except Exception:
        pass
    try:
        sh = w.sizeHint()
        rw = dw if dw > 0 else max(int(sh.width()), 1)
        rh = dh if dh > 0 else max(int(sh.height()), 1)
        w.resize(rw, rh)
    except Exception:
        pass


def _reapply_win32_owner_if_missing(w: QWidget, owner_hwnd: int) -> None:
    """
    GetWindow(GW_OWNER) が 0 のときだけ Win32 set_owner を再適用する。
    bring_to_front 前後で GW_OWNER が外れる／未設定のまま SFW されると sfw_ok=0 になり得るため。
    """
    try:
        owner = int(owner_hwnd or 0)
        if not owner or _w32 is None or os.name != "nt":
            return
        try:
            if bool(w.property("_hc_show_taskbar")):
                return
        except Exception:
            pass
        hwnd = int(w.winId()) if hasattr(w, "winId") else 0
        if not hwnd or not hasattr(_w32, "get_owner_hwnd") or not hasattr(_w32, "set_owner"):
            return
        cur = int(_w32.get_owner_hwnd(hwnd) or 0)
        if cur != 0:
            return
        owner_root = int(owner)
        if hasattr(_w32, "get_root_window"):
            try:
                owner_root = int(_w32.get_root_window(owner) or owner)
            except Exception:
                owner_root = int(owner)
        _w32.set_owner(hwnd, owner_root)
        _ff_diag(
            "[EXCEL_FRONT_FOLLOW] ensure_front reapplied_set_owner hwnd=%s owner_root=%s",
            hwnd,
            owner_root,
        )
        try:
            from PySide6.QtGui import QWindow

            excel_window = QWindow.fromWinId(owner_root)
            dialog_window = w.windowHandle()
            if excel_window and dialog_window:
                dialog_window.setTransientParent(excel_window)
        except Exception:
            pass
    except Exception:
        pass


def ensure_front(w: QWidget, parent_hwnd: int, *, _ff_retry: int = 0) -> None:
    """
    Method Name : ensure_front
    Arguments   : w (QWidget), parent_hwnd (int)
    Return      : None
    機能概要    : 常にExcel前面→ダイアログ前面の順で最前面化する（ダイアログがExcelの手前に見えるように）。
    GW_OWNER が欠ける環境では set_owner を再適用する。SetForegroundWindow が失敗した場合は短い遅延で数回まで再試行する。
    """
    _trace(f"[ensure_front:enter] parent_hwnd={parent_hwnd}")
    dlg_vis = False
    try:
        dlg_vis = bool(w.isVisible()) if w is not None else False
    except Exception:
        dlg_vis = False
    _ff_diag(
        "[EXCEL_FRONT_FOLLOW] ensure_front enter parent_hwnd=%s dlg=%s dlg_id=%s dialog_visible=%s",
        int(parent_hwnd or 0),
        type(w).__name__ if w is not None else None,
        id(w) if w is not None else 0,
        dlg_vis,
    )
    owner = int(parent_hwnd or 0)
    _ff_diag_ensure_front_snapshot(
        "ensure_front:before_bring_excel", owner, _dlg_hwnd_for_fg_diag(w)
    )
    log_ui_fg_phase("ensure_front:enter", owner, w, extra=f"dlg_visible={int(dlg_vis)}")
    if w is not None:
        try:
            from shiboken6 import Shiboken

            if not Shiboken.isValid(w):
                _ff_diag(
                    "[EXCEL_FRONT_FOLLOW] ensure_front abort deleted_widget parent_hwnd=%s",
                    int(parent_hwnd or 0),
                )
                return
        except Exception:
            pass
    _reapply_win32_owner_if_missing(w, owner)
    try:
        # 【目的】先にExcelを前面化し、その上にダイアログを表示するため
        if owner and _w32 is not None:
            _w32.bring_to_front(owner)
    except Exception:
        pass
    _reapply_win32_owner_if_missing(w, owner)
    log_ui_fg_phase("ensure_front:after_bring_excel", owner, w)
    _ff_diag_ensure_front_snapshot(
        "ensure_front:after_bring_excel", owner, _dlg_hwnd_for_fg_diag(w)
    )
    try:
        # 【目的】Qtネイティブ機能での前面化を試行するため
        w.raise_()
        w.activateWindow()
    except Exception:
        pass

    try:
        # 変数: ダイアログのHWND
        hwnd = int(w.winId()) if hasattr(w, "winId") else 0

        # 判定コメント: HWNDが正常取得できた場合
        if hwnd:
            _ff_diag_ensure_front_snapshot("ensure_front:after_qt_raise", owner, hwnd)
            log_ui_fg_phase("ensure_front:before_swp_sfw", owner, w)
            # 変数: API定数群（user32 は ctypes で利用可能）
            HWND_TOP = 0
            SWP_NOMOVE = 0x0002
            SWP_NOSIZE = 0x0001
            SWP_SHOWWINDOW = 0x0040

            # 【目的】最前面化（TOPMOSTのトグルは行わない：×ボタンのhover/描画が赤固定化する事象の回避）
            # prepare（show/exec 前）中に SWP_SHOWWINDOW を使うと先行表示が起きうるため、
            # 可視状態のときだけ付与する。
            _swp_flags = SWP_NOMOVE | SWP_NOSIZE
            try:
                if bool(w.isVisible()):
                    _swp_flags |= SWP_SHOWWINDOW
            except Exception:
                pass
            ctypes.windll.user32.SetWindowPos(
                hwnd, HWND_TOP, 0, 0, 0, 0, _swp_flags
            )
            sfw_ok: Optional[bool] = None
            if _w32 is not None and hasattr(_w32, "set_foreground_window_result"):
                sfw_ok = _w32.set_foreground_window_result(hwnd)
            else:
                try:
                    sfw_ok = bool(ctypes.windll.user32.SetForegroundWindow(hwnd))
                except Exception:
                    sfw_ok = False
            log_ui_fg_phase("ensure_front:after_sfw", owner, w, sfw_ok=sfw_ok)
            _ff_diag_ensure_front_snapshot(
                "ensure_front:after_sfw", owner, hwnd, sfw_ok=sfw_ok
            )
            _help_use_attach = False
            try:
                _help_use_attach = bool(w.property("_hc_help_dialog"))
            except Exception:
                _help_use_attach = False
            if _help_use_attach and hwnd:
                try:
                    from core import core_w32 as _cw_help_fg

                    _cw_help_fg.set_foreground_window_attach_input(hwnd)
                    if hasattr(_cw_help_fg, "nudge_top_level_to_foreground"):
                        _cw_help_fg.nudge_top_level_to_foreground(int(hwnd))
                except Exception:
                    pass
            _trace(f"[ensure_front:done] widget_hwnd={hwnd} parent_hwnd={parent_hwnd}")
            _trace_widget_rect(w, "ensure_front:直後(前面化で位置が変わったか)")
            _ff_diag(
                "[EXCEL_FRONT_FOLLOW] ensure_front done parent_hwnd=%s widget_hwnd=%s sfw_ok=%s",
                int(parent_hwnd or 0),
                hwnd,
                int(bool(sfw_ok)) if sfw_ok is not None else -1,
            )
            _tw_ef = get_ui_window_timings()
            _disable_retry = False
            try:
                _disable_retry = bool(w.property("_hc_disable_ensure_front_retry"))
            except Exception:
                _disable_retry = False
            if (
                (not _disable_retry)
                and sfw_ok is False
                and int(_ff_retry) < int(_tw_ef.ensure_front_set_foreground_fail_max_ff_retry_exclusive)
            ):
                try:
                    from PySide6.QtCore import QTimer

                    rr = int(_ff_retry) + 1
                    ph_snap = int(owner)
                    _retry_ms = int(_tw_ef.ensure_front_set_foreground_fail_retry_ms)

                    def _retry_ef() -> None:
                        try:
                            ensure_front(w, ph_snap, _ff_retry=rr)
                        except Exception:
                            pass

                    QTimer.singleShot(_retry_ms, _retry_ef)
                except Exception:
                    pass
        else:
            _trace(f"[ensure_front:skip] widget_hwnd=0 (winId not ready?) parent_hwnd={parent_hwnd}")
            log_ui_fg_phase("ensure_front:skip_no_winid", owner, w)
            _ff_diag(
                "[EXCEL_FRONT_FOLLOW] ensure_front skip_no_winId parent_hwnd=%s",
                int(parent_hwnd or 0),
            )
    except Exception as e:
        _trace(f"[ensure_front:error] {e!r}")
        _ff_diag("[EXCEL_FRONT_FOLLOW] ensure_front error parent_hwnd=%s exc=%r", int(parent_hwnd or 0), e)
        try:
            str(e)
        except Exception:
            pass
        return


def ensure_dialog_front_of_excel(
    w: QWidget,
    parent_hwnd: int,
    rect_override: Optional[Tuple[int, int, int, int]] = None,
) -> None:
    """
    ダイアログを Excel のオーナーにし、Excel 中央・前面に表示する。
    完了通知が本番でモニタ中央になる事象対策（テストスクリプトと同様の set_owner + 再配置）。
    """
    ph = int(parent_hwnd or 0)
    try:
        hwnd = int(w.winId()) if hasattr(w, "winId") else 0
    except Exception:
        hwnd = 0
    if ph and hwnd and _w32 is not None:
        try:
            root = _w32.get_root_window(ph)
            _w32.set_owner(hwnd, root)
            _w32.bring_to_front(root)
            center_on_excel(w, ph, rect_override)
        except Exception:
            pass
    try:
        w.raise_()
        w.activateWindow()
    except Exception:
        pass
    if hwnd:
        try:
            import ctypes
            ctypes.windll.user32.SetForegroundWindow(hwnd)
            if _w32 is not None and hasattr(_w32, "nudge_to_front"):
                _w32.nudge_to_front(hwnd)
        except Exception:
            pass


# ------------------------------------------------------------------------------
# Excel foreground follow (Windows)
# ------------------------------------------------------------------------------


def _set_owner_hwnd(w: QWidget, owner_hwnd: int) -> None:
    """
    Method Name : _set_owner_hwnd
    Arguments   : w (QWidget), owner_hwnd (int)
    Return      : None
    機能概要    : Qtネイティブ機能およびWin32APIを用いて、ExcelへのHWND紐付け（親子化）を行う。
    NOTE        : マルチモニタ+DPI で SetOwner 後に位置がずれることがあるため、可能なら本呼び出しの前に
                  adjustSize / center_on_excel で位置を決め、show は 1 回に留める（ちらつき抑制）。
    """
    _trace(f"[_set_owner_hwnd:enter] owner_hwnd={owner_hwnd}")
    try:
        # 変数: 対象ダイアログとExcelのHWND
        hwnd = int(w.winId()) if hasattr(w, "winId") else 0
        owner = int(owner_hwnd or 0)

        # 判定コメント: どちらかのHWNDが欠落している場合
        if not hwnd or not owner:
            return

        # 変数: タスクバー表示要求（Trueの場合はowner設定を行わない）
        try:
            show_taskbar = bool(w.property("_hc_show_taskbar"))
        except Exception:
            show_taskbar = False

        if show_taskbar:
            return

        # タスクバー非表示を安定させるため、オーナーは必ずルート（トップレベル）ウィンドウにする
        owner_root = int(owner)
        if _w32 is not None and hasattr(_w32, "get_root_window"):
            try:
                owner_root = _w32.get_root_window(owner)
            except Exception:
                pass

        try:
            from PySide6.QtGui import QWindow

            # 変数: QtのQWindowオブジェクト（ルートで親子関係を構築）
            excel_window = QWindow.fromWinId(owner_root)
            dialog_window = w.windowHandle()

            # 判定コメント: 両方のQWindowオブジェクトが取得できた場合
            if excel_window and dialog_window:
                # 【目的】Qtフレームワーク内部での親ウィンドウ関係を構築するため
                dialog_window.setTransientParent(excel_window)
        except Exception:
            pass

        try:
            if _w32 is not None:
                # 【目的】Windows OSレベル（Win32API）での確実なowner紐付けとタスクバー抑止を行うため（ルート指定）
                _w32.set_owner(hwnd, owner_root)
                if _log is not None:
                    try:
                        _log.debug("[OWNER] set_owner ok hwnd=%s owner_root=%s", hwnd, owner_root)
                    except Exception:
                        pass
                # Qt が表示後に拡張スタイルを戻す環境向けにタスクバー抑止を遅延再適用
                _tb_hwnd = int(hwnd)
                if hasattr(_w32, "apply_taskbar_hiding_extended_style"):
                    try:
                        from PySide6.QtCore import QTimer

                        def _tb_reapply() -> None:
                            try:
                                h2 = int(w.winId()) if hasattr(w, "winId") else 0
                                if h2:
                                    _w32.apply_taskbar_hiding_extended_style(h2, w)  # type: ignore[attr-defined]
                            except Exception:
                                pass

                        for _tb_ms in get_ui_window_timings().owner_taskbar_extended_style_reapply_delays_ms:
                            QTimer.singleShot(int(_tb_ms), _tb_reapply)
                    except Exception:
                        try:
                            if _tb_hwnd:
                                _w32.apply_taskbar_hiding_extended_style(_tb_hwnd, w)  # type: ignore[attr-defined]
                        except Exception:
                            pass
                try:
                    gw = (
                        int(_w32.get_owner_hwnd(hwnd))
                        if hasattr(_w32, "get_owner_hwnd")
                        else 0
                    )
                    log_ui_fg_phase(
                        "set_owner_after",
                        owner,
                        w,
                        extra=f"owner_root={owner_root} gw_owner={gw} match={int(gw == int(owner_root))}",
                    )
                except Exception:
                    pass
        except Exception as e:
            if _log is not None:
                try:
                    _log.warning("[OWNER] set_owner failed: %s", e)
                except Exception:
                    pass
    except Exception:
        return


def ensure_owner_and_front(w: QWidget, owner_hwnd: int) -> None:
    """
    ダイアログ表示直後に呼ぶ。Excel との親子関係を設定し、Excel→ダイアログの順で前面化する。
    親子関係によりタスクバーにダイアログが出なくなり、常に Excel の手前に表示される。
    """
    _set_owner_hwnd(w, owner_hwnd)
    ensure_front(w, owner_hwnd)


def apply_tooltip_if_set(widget, cfg: dict, key: str = "TOOLTIP") -> None:
    """
    設定にツールチップが定義されていればウィジェットに設定する。空・未定義なら何もしない。
    共通仕様: \\n・\\t を有効にするため _normalize_message_newlines を適用。前後は空白・タブのみ strip。
    """
    if widget is None or not isinstance(cfg, dict):
        return
    raw = str(cfg.get(key) or cfg.get(key.lower()) or "").strip(" \t\r")
    tip = _normalize_message_newlines(raw)
    if tip:
        try:
            widget.setToolTip(tip)
        except Exception:
            pass


def _get_progress_config() -> dict:
    """進捗画面の既定設定。get_ui_config2('CSV_MG', 'MAIN') を基準に SCREENS.PROGRESS をマージ。progress_cfg 未指定時のフォールバック（後方互換）。"""
    try:
        main = get_ui_config2("CSV_MG", "MAIN")
        progress = (main.get("SCREENS") or {}).get("PROGRESS") or {}
        return _deep_merge(main, progress)
    except Exception:
        return {}


def _get_done_config() -> dict:
    """CSV_MG.MAIN を基準に SCREENS.DONE をマージ。MAIN で設定した画面制御が標準で継承される。"""
    try:
        main = get_ui_config2("CSV_MG", "MAIN")
        done = (main.get("SCREENS") or {}).get("DONE") or {}
        return _deep_merge(main, done)
    except Exception:
        return {}


def get_ui_config2(feature_key: str, screen_key: str) -> dict:
    """新UI定義の設定を返す。CSV_MG は config/ui_csv_mg.json のみ参照（外部のみ・救済なし）。"""
    base = getattr(cst, "UI_COMMON", {}) or {}
    fk = str(feature_key or "").strip().upper()
    sk = str(screen_key or "").strip().upper()
    # CSV_MG: 外部ファイルのみ。失敗時は get_ui_config_from_file_required が UiConfigLoadError を発生
    if fk == "CSV_MG":
        load_required = getattr(cst, "get_ui_config_from_file_required", None)
        if not callable(load_required):
            if _log is not None:
                _log.error("[get_ui_config2] CSV_MG requires get_ui_config_from_file_required")
            return _deep_merge(base, {})
        feature = load_required("CSV_MG")
    else:
        screens = getattr(cst, "UI_SCREENS", {}) or {}
        feature = screens.get(fk) if isinstance(screens, dict) else None
        if not isinstance(feature, dict):
            if _log is not None:
                _log.debug("[get_ui_config2] no feature for fk=%s", fk)
            return _deep_merge(base, {})
    common = feature.get("COMMON") if isinstance(feature.get("COMMON"), dict) else {}
    screen = feature.get(sk) if isinstance(feature.get(sk), dict) else {}
    if not isinstance(screen, dict) and _log is not None:
        _log.debug("[get_ui_config2] no screen for sk=%s feature_keys=%s", sk, list(feature.keys()))
    out = _deep_merge(base, common)
    out = _deep_merge(out, screen)
    # CSV_MG: ファイル直下 WINDOW を _get_cfg と同様に最下位レイヤとして合成（COMMON/MAIN より上書きされうる）
    if fk == "CSV_MG" and isinstance(feature, dict):
        win_root = feature.get("WINDOW") or {}
        if isinstance(win_root, dict) and win_root:
            win_existing = out.get("WINDOW") or {}
            if isinstance(win_existing, dict):
                out["WINDOW"] = _deep_merge(dict(win_root), win_existing)
    return out

def _deep_merge(a: dict, b: dict) -> dict:
    """
    Method Name : _deep_merge
    Arguments   : a (dict), b (dict)
    Return      : dict
    機能概要    : 辞書を深い階層まで再帰的にマージする（引数bの内容を優先）。
    """
    # 変数: ベースとなる辞書のコピー
    out = dict(a)
    for k, v in (b or {}).items():
        # 判定コメント: 両方の値が辞書型である場合
        if isinstance(v, dict) and isinstance(out.get(k), dict):
            # 【目的】ネストされた辞書も再帰的に統合するため
            out[k] = _deep_merge(out[k], v)
        else:
            # 【目的】スカラー値等はそのまま上書きするため
            out[k] = v
    return out


def merge_screen_cfg_window_from_root(
    cfg: dict, screen_key: str, *, sheet_interaction_excel_unlock: bool = False
) -> dict:
    """
    MAIN・ルート WINDOW・SCREENS.<screen_key> をマージした画面用辞書を返す。
    ネストした WINDOW は root → MAIN.WINDOW → SCREENS.*.WINDOW の順で深マージ（衝突は後勝ち）。
    日付変換など SCREENS のみ渡していた完了／警告で、ルートの SHOW_IN_TASKBAR 等が効くようにする。
    sheet_interaction_excel_unlock: True のとき WINDOW.EXCEL_LOCK を False に上書き（JSON に書かずコードで統一する画面向け）。
    """
    main = (cfg or {}).get("MAIN") or {}
    if not isinstance(main, dict):
        main = {}
    root_w = dict((cfg or {}).get("WINDOW") or {})
    mw_raw = main.get("WINDOW")
    main_w = dict(mw_raw) if isinstance(mw_raw, dict) else {}
    base_win = _deep_merge(root_w, main_w)
    raw = ((cfg or {}).get("SCREENS") or {}).get(screen_key) or {}
    if not isinstance(raw, dict):
        raw = {}
    screen = dict(raw)
    sw_raw = screen.get("WINDOW")
    sw = dict(sw_raw) if isinstance(sw_raw, dict) else {}
    screen["WINDOW"] = _deep_merge(base_win, sw)
    merged = _deep_merge(main, screen)
    if sheet_interaction_excel_unlock:
        wout = dict(merged.get("WINDOW") or {})
        wout["EXCEL_LOCK"] = False
        merged["WINDOW"] = wout
    return merged


def get_ui_config(screen_key: str) -> dict:
    """
    Method Name : get_ui_config
    Arguments   : screen_key (str)
    Return      : dict
    機能概要    : hc_cst の UI_COMMON と 画面別設定(UI_*) をマージして返却する。
    """
    # 変数: 共通設定の取得
    base = getattr(cst, "UI_COMMON", {}) or {}
    # 変数: 対象画面キーの整形
    sk = str(screen_key or "").strip().upper()
    # 変数: 画面固有設定の取得
    screen = getattr(cst, f"UI_{sk}", {}) or {}
    # 命令分離: マージ実行
    return _deep_merge(base, screen)


def position_widget_at_excel_center(w: QWidget, parent_hwnd: int, size: int = 100) -> None:
    """
    ウィジェットを Excel ウィンドウの中央に配置する（ファイル選択などの親ウィンドウ用）。
    size は幅・高さ（ピクセル）。OS標準ダイアログのオーナーを中央に置くことで、ダイアログが Excel 付近に表示されやすくする。
    """
    if not int(parent_hwnd or 0):
        return

    def _do_position() -> None:
        rect = get_excel_rect(int(parent_hwnd or 0))
        if not rect:
            return
        left, top, right, bottom = rect
        cx = (left + right) // 2
        cy = (top + bottom) // 2
        x = cx - size // 2
        y = cy - size // 2
        try:
            hwnd = int(w.winId()) if hasattr(w, "winId") else 0
            if hwnd and os.name == "nt":
                import ctypes
                from ctypes import wintypes
                SWP_NOZORDER = 0x0004
                ctypes.windll.user32.SetWindowPos(hwnd, None, x, y, size, size, SWP_NOZORDER)
                return
        except Exception:
            pass
        try:
            w.setGeometry(x, y, size, size)
        except Exception:
            pass

    _with_thread_dpi_physical(_do_position)


def center_on_excel(
    w: QWidget, parent_hwnd: int, rect_override: Optional[Tuple[int, int, int, int]] = None
) -> None:
    """
    Method Name : center_on_excel
    Arguments   : w (QWidget), parent_hwnd (int), rect_override=送信時点の Excel 矩形（指定時は get_excel_rect を使わない）
    Return      : None
    機能概要    : Excelのウィンドウ領域を取得し、その中央に対象ウィジェットを配置する。
    NOTE        : rect_override があればそれを使う（リクエスト送信時点の矩形で、UI 表示時より正確な場合がある）。
    """
    ph = int(parent_hwnd or 0)

    def _do_center() -> None:
        rect = rect_override
        if not rect and ph:
            rect = get_excel_rect(ph)
        _trace(f"[center_on_excel] parent_hwnd={ph} rect_override={rect_override is not None} rect={rect!r}")
        if not rect:
            return
        try:
            if int(w.width()) <= 0 or int(w.height()) <= 0:
                w.adjustSize()
        except Exception:
            pass
        center_on_rect(w, tuple(rect))

    _with_thread_dpi_physical(_do_center)


# 変数: ジオメトリ保存用ディレクトリ名
_GEOM_DIRNAME = "geometry"


def _geom_path(screen_key: str) -> Path:
    """保存先パスの生成"""
    # 変数: コントロールディレクトリ配下のパス構築
    root = ipc_file.get_control_dir() / _GEOM_DIRNAME
    root.mkdir(parents=True, exist_ok=True)
    # 変数: ファイル名として安全な文字列にサニタイズ
    safe = re.sub(r"[^A-Za-z0-9_\-]+", "_", str(screen_key or ""))
    return root / f"{safe}.pkl"


def save_geometry(screen_key: str, w: QWidget) -> None:
    """
    Method Name : save_geometry
    Arguments   : screen_key (str), w (QWidget)
    Return      : None
    機能概要    : ウィンドウの現在位置とサイズをPickleとして保存する。
    """
    try:
        # 変数: 現在のジオメトリ情報
        g = w.geometry()
        # 変数: シリアライズ用の辞書
        data = {
            "x": int(g.x()),
            "y": int(g.y()),
            "w": int(g.width()),
            "h": int(g.height()),
        }
        # 【目的】次回起動時の復元用に情報を永続化するため
        ipc_file.write_pickle(_geom_path(screen_key), data)
    except Exception:
        return


def restore_geometry(screen_key: str, w: QWidget) -> bool:
    """
    Method Name : restore_geometry
    Arguments   : screen_key (str), w (QWidget)
    Return      : bool (成功時 True)
    機能概要    : 保存済みのウィンドウ位置・サイズ情報を復元する。
    """
    try:
        # 変数: ターゲットのパス
        p = _geom_path(screen_key)
        # 判定コメント: ファイルが存在しない、または空の場合
        if not p.exists() or p.stat().st_size <= 0:
            return False

        # 変数: 保存情報の読み込み
        d = ipc_file.read_pickle(p)
        if not isinstance(d, dict):
            return False

        # 変数: 値の抽出
        x = int(d.get("x", 0))
        y = int(d.get("y", 0))
        ww = int(d.get("w", 0))
        hh = int(d.get("h", 0))

        # 判定コメント: 有効なサイズ情報がある場合
        if ww > 0 and hh > 0:
            w.resize(ww, hh)

        # 命令分離: 位置の復元
        w.move(x, y)
        return True
    except Exception:
        return False


def _with_thread_dpi_physical(fn, *args, **kwargs):
    """スレッドを Per-Monitor DPI v2 に一時切り替え、GetWindowRect/SetWindowPos を物理ピクセルで統一する。マルチモニタ＋混在DPI でずれ防止。"""
    prev_ctx = None
    try:
        if hasattr(ctypes.windll.user32, "SetThreadDpiAwarenessContext"):
            DPI_PER_MONITOR_V2 = -4
            prev_ctx = ctypes.windll.user32.SetThreadDpiAwarenessContext(ctypes.c_void_p(DPI_PER_MONITOR_V2))
        return fn(*args, **kwargs)
    finally:
        if prev_ctx is not None and hasattr(ctypes.windll.user32, "SetThreadDpiAwarenessContext"):
            try:
                ctypes.windll.user32.SetThreadDpiAwarenessContext(prev_ctx)
            except Exception:
                pass


def get_excel_rect(parent_hwnd: int) -> Optional[Tuple[int, int, int, int]]:
    """
    Method Name : get_excel_rect
    Arguments   : parent_hwnd (int)
    Return      : Optional[Tuple[int, int, int, int]]
    機能概要    : 指定されたHWNDのスクリーン座標上の矩形(左, 上, 右, 下)を取得する。
    """
    if not int(parent_hwnd or 0):
        return None
    try:
        import ctypes
        from ctypes import wintypes

        rect = wintypes.RECT()
        ok = ctypes.windll.user32.GetWindowRect(int(parent_hwnd), ctypes.byref(rect))
        if not ok:
            return None
        return (int(rect.left), int(rect.top), int(rect.right), int(rect.bottom))
    except Exception:
        return None


def center_on_rect(w: QWidget, rect: Tuple[int, int, int, int]) -> None:
    """
    Method Name : center_on_rect
    Arguments   : w (QWidget), rect (Tuple)  … 基準矩形 (左, 上, 右, 下) 物理ピクセル
    Return      : None
    機能概要    : 対象ウィジェットを与えられた矩形領域の幾何学的な中心へ移動する。
    NOTE        : Excel 矩形は GetWindowRect（物理ピクセル）。Qt の move() は論理ピクセルで
                  DPI によりずれるため、HWND 取得可能なときは SetWindowPos で物理ピクセル配置に統一する。
    """
    left, top, right, bottom = rect
    try:
        hwnd = int(w.winId()) if hasattr(w, "winId") else 0
    except Exception:
        hwnd = 0

    if hwnd:
        try:
            import ctypes
            from ctypes import wintypes
            r = wintypes.RECT()
            if ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(r)):
                cw = max(int(r.right - r.left), 1)
                ch = max(int(r.bottom - r.top), 1)
                x = int(left + (right - left - cw) / 2)
                y = int(top + (bottom - top - ch) / 2)
                SWP_NOSIZE = 0x0001
                SWP_NOZORDER = 0x0004
                SWP_NOACTIVATE = 0x0010
                ctypes.windll.user32.SetWindowPos(hwnd, None, x, y, 0, 0, SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE)
                _trace(f"[center_on_rect] SetWindowPos hwnd={hwnd} x={x} y={y} size={cw}x{ch}")
                # w.move(x,y) は論理ピクセル扱いのため DPI でずれ、SetWindowPos の物理ピクセルを上書きしてしまうため行わない
                if hasattr(w, "winId"):
                    _trace_widget_rect(w, "center_on_rect:直後(SetWindowPosが効いたか)")
                return
        except Exception as e:
            _trace(f"[center_on_rect] SetWindowPos failed: {e!r}")

    # フォールバック: Qt の move（論理ピクセル・枠は frameGeometry）
    try:
        fr = w.frameGeometry()
        cw = max(int(fr.width()), int(w.width()), 1)
        ch = max(int(fr.height()), int(w.height()), 1)
    except Exception:
        cw = max(int(w.width()), 1)
        ch = max(int(w.height()), 1)
    x = int(left + (right - left - cw) / 2)
    y = int(top + (bottom - top - ch) / 2)
    w.move(x, y)
    _trace(f"[center_on_rect] Qt move x={x} y={y}")


def center_on_parent_widget(child: QWidget, parent: QWidget) -> None:
    """
    子ダイアログを親 QWidget の矩形（グローバル座標）の中央に配置する。
    データ集約デバッグ進捗のように、Excel ではなくアプリ画面中央に寄せたい場合に使う。
    """
    try:
        g = parent.frameGeometry()
        center_on_rect(child, (g.left(), g.top(), g.right(), g.bottom()))
    except Exception as e:
        _trace(f"[center_on_parent_widget] error: {e!r}")


def enable_excel_window(hwnd: int, enabled: bool) -> None:
    """
    Method Name : enable_excel_window
    Arguments   : hwnd (int), enabled (bool)
    Return      : None
    機能概要    : Excel本体のZオーダー連動を保ちつつ、内部の構成部品（リボンやセル等の子HWND）のみをロック/解除する。
                 Excelは階層が深いため、孫以下も含めて再帰的に列挙する。
    """
    _trace(f"[enable_excel_window:enter] hwnd={hwnd} enabled={enabled}")
    if not int(hwnd or 0):
        return
    if _w32 is None:
        return
    try:
        root = int(hwnd)
        seen: set[int] = set()
        q: list[int] = [root]

        # 【目的】深い階層のUI要素（リボン等）も対象に含めるため
        while q:
            h = q.pop(0)
            try:
                children = _w32.enum_child_windows(int(h))
            except Exception:
                children = []
            for ch in children or []:
                ih = int(ch)
                if ih and ih not in seen:
                    seen.add(ih)
                    q.append(ih)
            # 【目的】異常な増殖を防止するため（フェイルセーフ）
            if len(seen) > 20000:
                break

        if not seen:
            _trace("[enable_excel_window:targets] count=0")
            return

        _trace(
            f"[enable_excel_window:targets] count={len(seen)} sample={list(seen)[:10]}"
        )
        _w32.enable_windows(list(seen), bool(enabled))
        try:
            from core import core_env

            if core_env.ui_excel_lock_diag_enabled():
                from core.core_log import get_diag_logger

                get_diag_logger("hc_csv_tool.diag.ui_excel_lock").info(
                    "[UI_EXCEL_LOCK] sym_type=A enabled=%d root_hwnd=%s win32_targets=%d",
                    int(bool(enabled)),
                    int(root),
                    int(len(seen)),
                )
        except Exception:
            pass
    except Exception:
        return


def focus_excel_after_modal_close(parent_hwnd: int) -> None:
    """
    モーダル（QDialog.exec）終了直後に、フォーカスがコンソール（CMD）等へ逃げるのを抑える。
    Qt のイベントループで複数回・段階的に遅延し、Excel ルート HWND へ nudge_top_level_to_foreground
    （SetWindowPos / BringWindowToTop / SetForegroundWindow、失敗時は AttachThreadInput 経由の再試行）をかける。
    """
    ph = int(parent_hwnd or 0)
    if not ph or os.name != "nt" or _w32 is None:
        return

    def _nudge() -> None:
        try:
            root = int(ph)
            if hasattr(_w32, "get_root_window"):
                try:
                    root = int(_w32.get_root_window(ph))
                except Exception:
                    pass
            if hasattr(_w32, "nudge_top_level_to_foreground"):
                _w32.nudge_top_level_to_foreground(root)  # type: ignore[attr-defined]
            else:
                _w32.bring_to_front(root)
        except Exception:
            pass

    try:
        for _fx_ms in get_ui_window_timings().focus_excel_after_modal_nudge_delays_ms:
            QTimer.singleShot(int(_fx_ms), _nudge)
    except Exception:
        _nudge()


def _shutdown_requested() -> bool:
    """終了フラグファイルの有無を確認する。"""
    try:
        return ipc_file.get_shutdown_flag_path().exists()
    except Exception:
        return False


def _pid_alive(pid: int) -> bool:
    """
    Method Name : _pid_alive
    Arguments   : pid (int)
    Return      : bool
    機能概要    : WindowsAPIを用いて、指定PIDのプロセスが生存しているか判定する。
    """
    try:
        # 変数: プロセス問い合わせ用のアクセス権限定数
        PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
        # 変数: プロセスハンドルの取得
        handle = ctypes.windll.kernel32.OpenProcess(
            PROCESS_QUERY_LIMITED_INFORMATION, 0, int(pid)
        )

        # 判定コメント: ハンドルが0（取得失敗）の場合は終了とみなす
        if handle == 0:
            return False

        # 命令分離: 取得したハンドルの解放
        ctypes.windll.kernel32.CloseHandle(handle)
        return True
    except Exception:
        # 【目的】アクセス権限不足等のエラー時は、フェイルセーフとして生存中とみなすため
        return True


class UiShutdownGuard(QObject):
    """
    クラス名: UiShutdownGuard
    概要: UI表示中における終了要求(shutdown flag)や親プロセスの死活監視を行い、異常時は安全にUIを閉じる。
    """

    def __init__(
        self,
        dialog: QWidget,
        parent_hwnd: int,
        *,
        poll_ms: int = 200,
        excel_lock_enabled: bool = False,
        parent_pid: int = 0,
        excel_pid: int = 0,
        on_shutdown=None,
    ) -> None:
        super().__init__(dialog)
        # 変数: 各種パラメタの初期化
        self._dialog = dialog
        self._parent_hwnd = int(parent_hwnd or 0)
        self._poll_ms = max(int(poll_ms), 50)
        self._excel_lock_enabled = bool(excel_lock_enabled)
        self._parent_pid = int(parent_pid or 0)
        self._excel_pid = int(excel_pid or 0)
        self._on_shutdown = on_shutdown
        self._timer: QTimer | None = None
        self._stopped = False

    def start(self) -> None:
        """監視タイマーを開始する。"""
        if self._stopped:
            return
        if self._timer is None:
            self._timer = QTimer(self)
            self._timer.setInterval(self._poll_ms)
            self._timer.timeout.connect(self._tick)  # type: ignore[attr-defined]
        self._timer.start()

    def stop_and_unlock(self) -> None:
        """監視を停止し、必要であればロック解除を行う。"""
        if self._stopped:
            return
        self._stopped = True
        try:
            if self._timer is not None:
                self._timer.stop()
        except Exception:
            pass

    def _tick(self) -> None:
        """
        Method Name : _tick
        Arguments   : None
        Return      : None
        機能概要    : 定周期で終了フラグおよび関連プロセスの死活を監視する。
        """
        if self._stopped:
            return

        # 判定コメント: 終了フラグが検知された場合
        if _shutdown_requested():
            # 【目的】安全なシャットダウンシーケンスに移行するため
            self._handle_shutdown()
            return

        # 変数: 監視対象PID群
        ppid = self._parent_pid
        epid = self._excel_pid

        # 判定コメント: 親プロセスが指定されており、かつ死んでいる場合
        if ppid and (not _pid_alive(ppid)):
            self._handle_shutdown()
            return

        # 判定コメント: Excelプロセスが指定されており、かつ死んでいる場合
        if epid and (not _pid_alive(epid)):
            self._handle_shutdown()
            return

    def _handle_shutdown(self) -> None:
        """
        Method Name : _handle_shutdown
        Arguments   : None
        Return      : None
        機能概要    : コールバックの実行やダイアログの破棄等、シャットダウン処理を完遂する。
        """
        try:
            # 判定コメント: 独自のシャットダウン処理が定義されている場合
            if callable(self._on_shutdown):
                self._on_shutdown()
            else:
                # 判定コメント: rejectメソッドを持つ(QDialog等)場合
                if hasattr(self._dialog, "reject"):
                    getattr(self._dialog, "reject")()
                else:
                    self._dialog.close()
        finally:
            # 【目的】どのような終了経路でもタイマーの停止を確実に行うため
            self.stop_and_unlock()


# =============================================================================
# Progress / Done 用ヘルパ（ui_dialog_progress / ui_dialog_done から参照）
# =============================================================================

from PySide6.QtWidgets import QStyle  # noqa: E402

# 変数: モデルレスダイアログをガベージコレクションから保護するリスト
_MODELLESS_DIALOGS: list[object] = []


def _keep_modeless(obj: object, *, exclude_from_bulk_close: bool = False) -> None:
    """モデルレスダイアログをリストに保持し、GC から保護する。

    exclude_from_bulk_close=True のウィンドウは _close_all_modeless の一括 close 対象外
    （長寿命モードレスのデータ集約メイン等。進捗ダイアログは従来どおり False）。
    """
    if exclude_from_bulk_close:
        try:
            setattr(obj, "_hc_exclude_from_bulk_close", True)
        except Exception:
            pass
    try:
        _MODELLESS_DIALOGS.append(obj)
    except Exception:
        return


def _close_all_modeless() -> None:
    """進捗などのモデルレス画面を閉じる（ベストエフォート）。

    _hc_exclude_from_bulk_close が付いたウィンドウは閉じずリストに残す（データ集約メイン）。
    """
    try:
        remaining: list[object] = []
        for d in list(_MODELLESS_DIALOGS):
            try:
                if getattr(d, "_hc_exclude_from_bulk_close", False):
                    remaining.append(d)
                    continue
                if hasattr(d, "close"):
                    d.close()
            except Exception:
                pass
        try:
            _MODELLESS_DIALOGS[:] = remaining
        except Exception:
            _MODELLESS_DIALOGS.clear()
    except Exception:
        return


def _remove_from_modeless(obj: object) -> None:
    """モデルレス一覧から1件だけ削除（進捗クローズ時に二重クローズを防ぐ）。"""
    try:
        had = obj in _MODELLESS_DIALOGS
        wid = 0
        try:
            if hasattr(obj, "winId"):
                wid = int(obj.winId())  # type: ignore[attr-defined]
        except Exception:
            wid = 0
        if had:
            _MODELLESS_DIALOGS.remove(obj)
        if _log is not None:
            try:
                _log.info(
                    "[MODELESS_REMOVE] type=%s winId=%s had=%s remaining=%s",
                    type(obj).__name__,
                    wid,
                    had,
                    len(_MODELLESS_DIALOGS),
                )
            except Exception:
                pass
    except Exception:
        pass


def teardown_feature_ui_shared_state(
    *,
    parent_hwnd: int = 0,
    modeless_widget: Optional[QWidget] = None,
    excel_unlock: bool = False,
) -> None:
    """
    機能終了時に Qt UI サーバ内の共有状態をベストエフォートで片付ける。
    （任意で）enable_excel_window(True)・（任意で）モードレス一覧からの削除。
    二重呼び可。_close_all_modeless は含めない（他画面巻き込み防止）。
    """
    ph = int(parent_hwnd or 0)
    if excel_unlock and ph:
        try:
            enable_excel_window(ph, True)
        except Exception:
            pass
    if modeless_widget is not None:
        try:
            _remove_from_modeless(modeless_widget)
        except Exception:
            pass


# アイコンサイズ S/M/L のピクセル既定値（ICON_SIZE 未指定時のフォールバックに利用）
_ICON_SIZE_S = 16
_ICON_SIZE_M = 24
_ICON_SIZE_L = 32


def _icon_size_pixels_from_config(icon_size_value, default_pixels: int = 24) -> int:
    """
    ICON_SIZE 設定値（S/M/L または数値）をピクセル数に変換する。
    S=16, M=24, L=32。数値の場合は 12〜48 にクランプ。未設定・空は default_pixels。
    """
    if icon_size_value is None:
        return default_pixels
    s = str(icon_size_value).strip().upper()
    if not s:
        return default_pixels
    if s == "S":
        return _ICON_SIZE_S
    if s == "M":
        return _ICON_SIZE_M
    if s == "L":
        return _ICON_SIZE_L
    try:
        n = int(float(s))
        return max(12, min(48, n))
    except (ValueError, TypeError):
        return default_pixels


def _warning_icon_pixmap(style, icon_key: str, size: int = 18):
    """ICON 設定値（Warning/Info 等）から QPixmap を返す。未対応は None。"""
    key = str(icon_key or "").strip().lower()
    if not key:
        return None
    sp = None
    if key in ("warning", "warn"):
        sp = QStyle.StandardPixmap.SP_MessageBoxWarning
    elif key in ("information", "info"):
        sp = QStyle.StandardPixmap.SP_MessageBoxInformation
    elif key in ("critical", "error"):
        sp = QStyle.StandardPixmap.SP_MessageBoxCritical
    elif key == "question":
        sp = QStyle.StandardPixmap.SP_MessageBoxQuestion
    if sp is None:
        return None
    try:
        return style.standardIcon(sp).pixmap(size, size)
    except Exception:
        return None


def create_warning_dialog(req: dict, parent_hwnd: int, warning_cfg: dict):
    """共通ワーニングダイアログを生成する（ui_dialog_warning ラッパー）。"""
    try:
        from ui_qt.ui_dialog_warning import create_warning_dialog as _impl
    except Exception:
        # フォールバック: 直接 WarningDialog を探す（開発時用）
        from ui_qt.ui_dialog_warning import WarningDialog as _Warn  # type: ignore

        return _Warn(req, int(parent_hwnd or 0), warning_cfg or {})
    return _impl(req, int(parent_hwnd or 0), warning_cfg or {})


def _progress_center_fallback(w: QWidget) -> None:
    """ProgressDialog 用: 不透明化直後に Excel 中央へ再配置（フォールバック）。送信時点の excel_rect があれば使用。"""
    try:
        _trace_widget_rect(w, "Progress:fallback呼び出し前")
        if getattr(w, "_center_on_parent_widget", False):
            pw = w.parentWidget()
            if pw is not None:
                center_on_parent_widget(w, pw)
                _trace("[_progress_center_fallback] center_on_parent_widget done")
                _trace_widget_rect(w, "Progress:fallback center_on_parent直後")
                return
        ph = getattr(w, "_parent_hwnd", 0) or 0
        rect_override = getattr(w, "_excel_rect", None)
        if ph or rect_override:
            center_on_excel(w, ph, rect_override)
            _trace("[_progress_center_fallback] center_on_excel done")
            _trace_widget_rect(w, "Progress:fallback center_on_excel直後")
    except Exception as e:
        _trace(f"[_progress_center_fallback] error: {e!r}")


def _progress_done_recenter(w: QWidget) -> None:
    """ProgressDialog 用: 「完了」表示に切り替えたあと 1 フレーム程度で前面化＋中央を 1 回。送信時点の excel_rect があれば使用。"""
    try:
        if getattr(w, "_center_on_parent_widget", False):
            pw = w.parentWidget()
            if pw is not None:
                center_on_parent_widget(w, pw)
                _trace("[_progress_done_recenter] center_on_parent_widget done")
                _trace_widget_rect(w, "Progress:DONE 50ms後 parent")
                return
        ph = getattr(w, "_parent_hwnd", 0) or 0
        rect_override = getattr(w, "_excel_rect", None)
        if ph or rect_override:
            if ph:
                ensure_front(w, ph)
            center_on_excel(w, ph, rect_override)
            _trace("[_progress_done_recenter] center_on_excel done")
            _trace_widget_rect(w, "Progress:DONE 50ms後")
    except Exception as e:
        _trace(f"[_progress_done_recenter] error: {e!r}")


def _done_center_fallback(w: QWidget, parent_hwnd: int) -> None:
    """DoneDialog 用: 不透明化直後に Excel 中央へ再配置（フォールバック）。"""
    try:
        _trace_widget_rect(w, "Done:fallback呼び出し前")
        if parent_hwnd:
            center_on_excel(w, parent_hwnd)
            _trace("[_done_center_fallback] center_on_excel done")
            _trace_widget_rect(w, "Done:fallback center_on_excel直後")
    except Exception as e:
        _trace(f"[_done_center_fallback] error: {e!r}")


def create_dialog(
    req_dict: dict, parent_hwnd: int, sheet_id: str, parent_widget: Optional[QWidget] = None
):  # noqa: ARG001
    """
    Method Name : create_dialog
    Arguments   : req_dict (dict), parent_hwnd (int), sheet_id: str, parent_widget: 親ウィジェット（サブ画面で結合画面を背後に残す場合に指定）
    Return      : QDialog
    機能概要    : ui_server からのディスパッチを受け、アクションに応じたダイアログを生成する。
    """
    # 変数: アクション文字列の取得と整形
    action = str(req_dict.get("action", "") or "").strip().lower()

    # 判定: 完了画面の要求
    if action == "done":
        if parent_widget is None:
            _close_all_modeless()
        return create_done_dialog(req_dict, int(parent_hwnd or 0), parent_widget, None)

    # 判定: 進捗画面の要求
    if action == "progress":
        return create_progress_dialog(req_dict, int(parent_hwnd or 0), parent_widget, None)

    # 未知のアクションは空の完了画面を返す（安全側）
    return create_done_dialog({"items": []}, int(parent_hwnd or 0), parent_widget)


def create_done_dialog(
    req: dict,
    parent_hwnd: int,
    parent_widget: Optional[QWidget] = None,
    done_cfg: Optional[dict] = None,
):
    """
    【概要】
        共通完了通知ダイアログを生成する（ui_dialog_done へ委譲）。
    【補足】
        done_cfg を渡すとその設定を使用。未指定時は _get_done_config() を使用。
    """
    from ui_qt.ui_dialog_done import create_done_dialog as _impl
    return _impl(req, int(parent_hwnd or 0), parent_widget, done_cfg)


def create_progress_dialog(
    req_dict: dict,
    parent_hwnd: int,
    parent_widget: Optional[QWidget] = None,
    progress_cfg: Optional[dict] = None,
):
    """
    【概要】
        共通進捗ダイアログ（ProgressDialog）を生成するラッパー。
    【補足】
        progress_cfg 未指定時は _get_progress_config()（従来の CSV_MG 参照）で既定設定を使用する。
        実装本体は ui_qt.ui_dialog_progress 側に委譲しつつ、モデルレス保護リストで GC から保護する。
    """
    from ui_qt.ui_dialog_progress import create_progress_dialog as _impl
    return _impl(req_dict, int(parent_hwnd or 0), parent_widget, progress_cfg)


def show_info_notice(parent: QWidget | None, title: str, text: str) -> int:
    """お知らせ通知（Information）。"""
    from ui_qt.ui_notification_sound import play_notification_sound

    play_notification_sound("info")
    return QMessageBox.information(parent, title, text)


def show_done_notice(parent: QWidget | None, title: str, text: str) -> int:
    """終了通知（完了・Information 表示）。"""
    from ui_qt.ui_notification_sound import play_notification_sound

    play_notification_sound("done")
    return QMessageBox.information(parent, title, text)


def show_warning_notice(parent: QWidget | None, title: str, text: str) -> int:
    """お知らせ通知（Warning アイコン・お知らせ音）。"""
    from ui_qt.ui_notification_sound import play_notification_sound

    play_notification_sound("info")
    return QMessageBox.warning(parent, title, text)


def show_error_notice(parent: QWidget | None, title: str, text: str) -> int:
    """エラー通知。"""
    from ui_qt.ui_notification_sound import play_notification_sound

    play_notification_sound("error")
    return QMessageBox.critical(parent, title, text)


# 恒久的なインポート診断（ベストエフォート）
try:
    ipc_file.log_module_loaded(__name__, str(Path(__file__).resolve()), __version__)
except Exception:
    pass
