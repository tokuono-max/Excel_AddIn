# -*- coding: utf-8 -*-
"""
Python: 3.12
Module: ui_qt/ui_dupli.py
Created: 2026-03-06
Updated: 2026-05-05
Version: 1.3.18
Purpose:
  重複チェックの UI（進捗・完了通知・レポート）。設定は config/ui_dupli.json 必須。
History (latest 3):
  - 1.3.18 (2026-05-05) モードA: 外枠幅は JSON（REPORT.WINDOW.DEFAULT/MIN/MAX）を優先。列の強制縮小を抑止し、長文は横スクロールで閲覧可能に調整。
  - 1.3.17 (2026-05-05) 重複一覧の横スクロールバーを常時表示に変更（非表示に見える環境対策）。
  - 1.3.16 (2026-05-05) 重複一覧セルの省略表示（...）を無効化。値はそのまま表示し、横スクロールで全文確認できるよう調整。
"""
from __future__ import annotations

import logging
import os
import time
from pathlib import Path
from typing import Any, Optional

from ui_qt import ipc_file

from PySide6.QtCore import QEvent, Qt, QTimer
from PySide6.QtGui import QColor, QBrush
from PySide6.QtWidgets import (
    QAbstractItemView,
    QDialog,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QPushButton,
    QSizePolicy,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
)

__version__ = "1.3.18"

# レポート: changeEvent で ensure_front を打つ間隔（連打抑制）
_DUPLI_REPORT_ENSURE_THROTTLE_SEC = 0.72
# 表示直後・前面化後に Win32/COM の操作有効が取り残される環境向け（ui_data_agg と同系）
_DUPLI_REPORT_EXCEL_UNLOCK_PULSE_MS = (0, 130, 400)

_XL_CALC_MANUAL_UI = -4135
_XL_CALC_AUTO_UI = -4105

# highlight_clear 矩形ループの診断（大量 COM 時の心拍）
_HLCLR_LOG_EVERY_N_RECTS = 100
_HLCLR_HEARTBEAT_SEC = 5.0
# ui_server の while が req_hlclr を pop し進捗を show してから COM 解除に入るまでの猶予（ms）
_HLCLR_UI_LEAD_MS = 180
# hlclr 進捗 pickle の更新間隔（矩形件数のステップ。大きいほど COM 負荷は下がる）
_HLCLR_PROGRESS_STRIDE = 8
# 上記に加え、最長これだけ間隔が空いたら UI 向けに 1 回書き込む（長時間 1 矩形で止まった見えを防ぐ）
_HLCLR_PROGRESS_UI_INTERVAL_SEC = 0.22

_log_dupli_ui = logging.getLogger(__name__)

try:
    from core.core_log import get_diag_logger

    _diag_ui_dupli = get_diag_logger("hc_csv_tool.diag.ui_dupli")
except Exception:  # pragma: no cover
    _diag_ui_dupli = None  # type: ignore[misc, assignment]


def _dupli_ui_diag(msg: str, *args: object) -> None:
    """hc_csv_diag.log 向け。"""
    if _diag_ui_dupli is None:
        return
    try:
        _diag_ui_dupli.info(msg, *args)
    except Exception:
        pass

# モードレスで開いたレポートダイアログをリストで保持し、GC で回収されないようにする
_open_report_dialogs: list[QDialog] = []

_DUPLI_CFG_CACHE: dict[str, Any] | None = None


def _get_cfg() -> dict[str, Any]:
    """
    重複チェック用の画面設定を config/ui_dupli.json から読み込む。
    読込失敗時は UiConfigLoadError が発生する（救済なし）。
    """
    global _DUPLI_CFG_CACHE
    if _DUPLI_CFG_CACHE is not None:
        return _DUPLI_CFG_CACHE
    from core import core_cst as cst

    _DUPLI_CFG_CACHE = cst.get_ui_config_from_file_required("dupli")
    return _DUPLI_CFG_CACHE


def _resolve_xlwings_book(app: Any, book_name: str) -> Any:
    """book_name があれば一致するブックを探す。無ければ active。"""
    bn = str(book_name or "").strip()
    if not bn:
        return app.books.active
    try:
        for b in app.books:
            if str(b.name or "").strip() == bn:
                return b
    except Exception:
        pass
    try:
        return app.books[bn]
    except Exception:
        pass
    _log_dupli_ui.warning(
        "[ui_dupli] book %r not found, using active book (highlight_clear / goto)",
        bn,
    )
    return app.books.active


def _dupli_report_column_content_width(tbl: QTableWidget, col: int) -> int:
    """1 列ぶんのセル表示文字のみの最大ピクセル幅（ヘッダは含めない）。"""
    try:
        fm = tbl.fontMetrics()
    except Exception:
        return 48
    m = 48
    for ri in range(tbl.rowCount()):
        it = tbl.item(ri, col)
        if it is None:
            continue
        t = str(it.text() or "")
        if not t:
            continue
        try:
            m = max(m, int(fm.horizontalAdvance(t)))
        except Exception:
            m = max(m, len(t) * 8)
    return m


def _dupli_report_header_text_width(tbl: QTableWidget, col: int) -> int:
    """ヘッダ文字列の表示幅（ピクセル）。"""
    try:
        fm = tbl.fontMetrics()
    except Exception:
        return 0
    try:
        hi = tbl.horizontalHeaderItem(col)
    except Exception:
        hi = None
    txt = str(hi.text() or "").strip() if hi is not None else ""
    if not txt:
        return 0
    try:
        return max(0, int(fm.horizontalAdvance(txt)))
    except Exception:
        return max(0, len(txt) * 8)


def _dupli_rebalance_widths_for_min_column(
    widths: list[int], target_idx: int, min_target: int, floor_each: int = 44
) -> list[int]:
    """
    総和を維持したまま target_idx 列に最低幅を割り当てる。
    不足分は他列から floor_each まで削って再配分する。
    """
    if not widths:
        return list(widths)
    n = len(widths)
    if target_idx < 0 or target_idx >= n:
        return list(widths)
    out = [max(floor_each, int(w)) for w in widths]
    need = max(0, int(min_target) - int(out[target_idx]))
    if need <= 0:
        return out
    got = 0
    donors = sorted(
        [i for i in range(n) if i != target_idx],
        key=lambda i: out[i],
        reverse=True,
    )
    for di in donors:
        if got >= need:
            break
        reducible = max(0, int(out[di]) - int(floor_each))
        if reducible <= 0:
            continue
        take = min(reducible, need - got)
        out[di] -= int(take)
        got += int(take)
    if got > 0:
        out[target_idx] += int(got)
    return out


def _dupli_scale_natural_widths_to_budget(
    widths: list[int], budget: int, floor_each: int = 44
) -> list[int]:
    """自然幅の合計が budget を超えるとき、余裕の大きい列ほど削るよう比例縮小（各列は floor_each 以上）。"""
    if not widths or budget <= 0:
        return list(widths)
    n = len(widths)
    base = [max(floor_each, int(w)) for w in widths]
    if sum(base) <= budget:
        return base
    slack = [b - floor_each for b in base]
    st = sum(slack)
    if st <= 0:
        return [floor_each] * n
    rem = budget - n * floor_each
    if rem < 0:
        return [floor_each] * n
    out = [floor_each + int(rem * s / st) for s in slack]
    d = budget - sum(out)
    order = sorted(range(n), key=lambda i: slack[i], reverse=True)
    k = 0
    while d > 0 and k < 10000:
        out[order[k % n]] += 1
        d -= 1
        k += 1
    return out


def _goto_scroll_margins_from_cfg() -> tuple[int, int]:
    """config/ui_dupli.json の GOTO_MARGIN_ROWS / GOTO_MARGIN_COLS（既定 2,1）。"""
    try:
        c = _get_cfg()
        mr = int(c.get("GOTO_MARGIN_ROWS", 2) or 2)
        mc = int(c.get("GOTO_MARGIN_COLS", 1) or 1)
    except (TypeError, ValueError):
        mr, mc = 2, 1
    return max(0, int(mr)), max(0, int(mc))


def _goto_highlight_bgr_from_cfg(cfg: dict[str, Any]) -> int:
    """ジャンプ先の一時グレー（BGR 整数）。GOTO_HIGHLIGHT.BGR 未指定時は薄いグレー。"""
    gh = cfg.get("GOTO_HIGHLIGHT") if isinstance(cfg.get("GOTO_HIGHLIGHT"), dict) else {}
    try:
        v = int(gh.get("BGR", 0) or 0)
    except (TypeError, ValueError):
        v = 0
    if v:
        return v
    r = g = b = 221
    return r + (g * 256) + (b * 65536)


def _goto_excel_cell(
    parent_hwnd: int,
    a1: str,
    *,
    book_name: str = "",
    sheet_name: str = "",
) -> Optional[tuple[int, int, int, int]]:
    """
    指定した Excel ウィンドウ（HWND）で、セルアドレス（A1 形式）に選択を移動する。
    シート名付き（例: Sheet1!A10）も可。book_name / sheet_name があれば処理対象ブック・シートを固定する。
    縦・横それぞれ、余白付きで表示に収まらないときだけ ActiveWindow（同一ブック時）の ScrollRow / ScrollColumn を更新する。
    最後に Excel を前面化して選択枠が見えやすくする。
    成功時は選択レンジの (r1,c1,r2,c2) を返す。
    """
    raw_full = (a1 or "").strip()
    if not int(parent_hwnd or 0) or not raw_full:
        return None
    sheet_from_addr = ""
    cell_ref = raw_full
    if "!" in raw_full:
        left, right = raw_full.split("!", 1)
        sheet_from_addr = left.strip().strip("'")
        cell_ref = right.strip()
    if not cell_ref:
        return None
    sn = sheet_from_addr or str(sheet_name or "").strip()
    try:
        import ctypes

        from xlwings import App
        from xlwings._xlwindows import App as WinApp

        app = App(impl=WinApp(xl=int(parent_hwnd)))
        book = _resolve_xlwings_book(app, book_name)
        try:
            book.activate()
        except Exception:
            pass
        try:
            app.activate(steal_focus=True)
        except Exception:
            pass
        if sn:
            try:
                sh = book.sheets[sn]
            except Exception:
                sh = book.sheets.active
                _log_dupli_ui.warning(
                    "[ui_dupli] goto: sheet %r not found, using active sheet", sn
                )
        else:
            sh = book.sheets.active
        try:
            sh.activate()
        except Exception:
            pass
        rng = sh.range(cell_ref)
        try:
            r1 = int(rng.api.Row)
            c1 = int(rng.api.Column)
            r2 = r1 + int(rng.api.Rows.Count) - 1
            c2 = c1 + int(rng.api.Columns.Count) - 1
        except Exception:
            try:
                r1, c1 = int(rng.row), int(rng.column)
                r2 = r1 + int(rng.rows.count) - 1
                c2 = c1 + int(rng.columns.count) - 1
            except Exception:
                r1, c1, r2, c2 = 1, 1, 1, 1
        mr, mc = _goto_scroll_margins_from_cfg()
        api = app.api
        win = None
        try:
            awb = api.ActiveWorkbook
            if awb is not None and _dupli_workbook_names_match(
                str(awb.Name), str(book.api.Name)
            ):
                win = api.ActiveWindow
        except Exception:
            win = None
        if win is None:
            try:
                win = book.api.Windows(1)
            except Exception:
                win = None
        scroll_row_applied = False
        scroll_col_applied = False
        if win is not None:
            try:
                vis = win.VisibleRange
                vr1 = int(vis.Row)
                vc1 = int(vis.Column)
                H = int(vis.Rows.Count)
                W = int(vis.Columns.Count)
                vr2 = vr1 + H - 1
                vc2 = vc1 + W - 1
                T1 = max(1, r1 - mr)
                T2 = r2 + mr
                C1 = max(1, c1 - mc)
                C2 = c2 + mc
                vert_ok = vr1 <= T1 and vr2 >= T2
                horiz_ok = vc1 <= C1 and vc2 >= C2
                if not vert_ok:
                    s_low = max(1, T2 - H + 1)
                    s_high = T1
                    new_sr = s_low if s_low <= s_high else max(1, T1)
                    win.ScrollRow = new_sr
                    scroll_row_applied = True
                if not horiz_ok:
                    s_low_c = max(1, C2 - W + 1)
                    s_high_c = C1
                    new_sc = s_low_c if s_low_c <= s_high_c else max(1, C1)
                    win.ScrollColumn = new_sc
                    scroll_col_applied = True
            except Exception:
                pass
        try:
            rng.select()
        except Exception:
            _dupli_ui_diag(
                "[UI_DUPLI_GOTO] fail select parent_hwnd=%s ref=%r book=%r sheet=%r",
                parent_hwnd,
                cell_ref,
                str(getattr(book, "name", "") or ""),
                str(getattr(sh, "name", "") or ""),
            )
            return None
        # 二重 steal_focus はレポート等の Qt 窓の Z 順を不安定にしやすい。セル選択後は Win32 のみで Excel 前面化。
        try:
            ctypes.windll.user32.SetForegroundWindow(int(parent_hwnd))
        except Exception:
            pass
        _dupli_ui_diag(
            "[UI_DUPLI_GOTO] ok parent_hwnd=%s ref=%r book=%r sheet=%r "
            "quad=(%s,%s,%s,%s) margin=(%s,%s) scroll_row=%s scroll_col=%s",
            parent_hwnd,
            cell_ref,
            str(getattr(book, "name", "") or ""),
            str(getattr(sh, "name", "") or ""),
            r1,
            c1,
            r2,
            c2,
            mr,
            mc,
            scroll_row_applied,
            scroll_col_applied,
        )
        return (r1, c1, r2, c2)
    except Exception as exc:
        _dupli_ui_diag(
            "[UI_DUPLI_GOTO] exception parent_hwnd=%s ref=%r exc=%r",
            parent_hwnd,
            raw_full,
            exc,
        )
        try:
            import ctypes

            ctypes.windll.user32.SetForegroundWindow(int(parent_hwnd))
        except Exception:
            pass
        return None


def _dupli_clear_range_fill(rng: Any) -> None:
    """重複ハイライト用の塗りつぶしを落とす（xlwings Range）。"""
    try:
        rng.color = None
    except Exception:
        pass
    try:
        interior = rng.api.Interior
        interior.Pattern = -4142  # xlNone
    except Exception:
        pass
    try:
        rng.api.Interior.TintAndShade = 0
    except Exception:
        pass
    for idx in (-4105, -4146):  # xlAutomatic 相当（環境差の吸収）
        try:
            rng.api.Interior.ColorIndex = idx
        except Exception:
            pass


def _dupli_fill_bgr_from_cfg(cfg: dict[str, Any]) -> int:
    """svc_dupli._highlight_bgr と同じ BGR 整数（fill_bgr 未指定時のフォールバック）。"""
    from core import core_cst as cst

    hi = cfg.get("HIGHLIGHT") or {}
    if hi.get("USE_CORE_ERR_BG"):
        t = getattr(cst, "ERR_BG_COLOR", (255, 200, 200))
        if isinstance(t, (list, tuple)) and len(t) >= 3:
            r, g, b = int(t[0]), int(t[1]), int(t[2])
            return r + (g * 256) + (b * 65536)
    rgb = hi.get("RGB")
    if isinstance(rgb, (list, tuple)) and len(rgb) >= 3:
        r, g, b = int(rgb[0]), int(rgb[1]), int(rgb[2])
        return r + (g * 256) + (b * 65536)
    return 180 + (200 * 256) + (255 * 65536)


def _dupli_intersect_visible_quad(
    vr1: int, vc1: int, vr2: int, vc2: int, quad: list[int]
) -> Optional[list[int]]:
    """軸平行矩形 quad と可視矩形の交差。無ければ None。"""
    if len(quad) < 4:
        return None
    r1, c1, r2, c2 = int(quad[0]), int(quad[1]), int(quad[2]), int(quad[3])
    ar1 = max(r1, vr1)
    ac1 = max(c1, vc1)
    ar2 = min(r2, vr2)
    ac2 = min(c2, vc2)
    if ar1 > ar2 or ac1 > ac2:
        return None
    return [ar1, ac1, ar2, ac2]


def _dupli_expand_visible_bounds(
    vr1: int, vc1: int, vr2: int, vc2: int, margin_rows: int, margin_cols: int
) -> tuple[int, int, int, int]:
    mr = max(0, int(margin_rows))
    mc = max(0, int(margin_cols))
    return (
        max(1, vr1 - mr),
        max(1, vc1 - mc),
        vr2 + mr,
        vc2 + mc,
    )


def _dupli_workbook_names_match(a: str, b: str) -> bool:
    """ブック表示名のゆるい一致（拡張子の有無差を吸収）。"""

    def _norm(x: str) -> str:
        t = str(x or "").strip().lower()
        if "." in t:
            t = t.rsplit(".", 1)[0]
        return t

    return _norm(a) == _norm(b)


def _dupli_hlclr_progress_write(path: Path, obj: dict[str, Any]) -> None:
    try:
        o = dict(obj)
        if "seq" not in o:
            try:
                prev = ipc_file.read_pickle(path)
                sq = int(prev.get("seq", -1)) + 1 if isinstance(prev, dict) else 0
            except Exception:
                sq = 0
            o["seq"] = sq
        ipc_file.write_pickle(path, o)
    except Exception:
        pass


def _hlclr_progress_path(sheet_id: str) -> Path:
    d = Path(ipc_file.get_ipc_root()) / "progress"
    d.mkdir(parents=True, exist_ok=True)
    sid = str(sheet_id or "").strip() or "_"
    return d / f"progress_dupli_hlclr_{sid}.pkl"


def _submit_hlclr_progress_ui(parent_hwnd: int, sheet_id: str, progress_path: Path) -> None:
    try:
        from svc.svc_host import ensure_ui_server

        ensure_ui_server()
        excel_rect = None
        if int(parent_hwnd or 0) and os.name == "nt":
            try:
                import ctypes
                from ctypes import wintypes

                r = wintypes.RECT()
                if ctypes.windll.user32.GetWindowRect(int(parent_hwnd), ctypes.byref(r)):
                    excel_rect = [int(r.left), int(r.top), int(r.right), int(r.bottom)]
            except Exception:
                pass
        res_dir = Path(ipc_file.get_ipc_root()) / "result"
        res_dir.mkdir(parents=True, exist_ok=True)
        ts_ms = int(time.time() * 1000)
        result_path = str(res_dir / f"res_progress_dupli_hlclr_{ts_ms}_{os.getpid()}.pkl")
        req_dict: dict[str, Any] = {
            "action": "progress",
            "progress_path": str(progress_path),
            "phase_total": 1,
            "excel_lock": False,
            "no_native_window": True,
            "done_delay_ms": 450,
        }
        if excel_rect is not None:
            req_dict["excel_rect"] = excel_rect
        payload = {
            "parent_hwnd": int(parent_hwnd),
            "result_path": result_path,
            "ready_path": "",
            "sheet_id": str(sheet_id or "").strip() or "_",
            "log_path": "",
            "action": "progress",
            "module": "ui_qt.ui_dupli",
            "req_dict": req_dict,
        }
        req_path = ipc_file.get_request_dir() / f"req_hlclr_{ts_ms}_{os.getpid()}.pkl"
        ipc_file.write_pickle(req_path, payload)
    except Exception as exc:
        _log_dupli_ui.warning("[ui_dupli] hlclr progress UI request failed: %s", exc)


def _dupli_hlclr_finish_deferred(dlg: Any, parent_hwnd: int) -> None:
    """レポート遅延解除チェーンの終了: Excel 前面化・deleteLater・フラグ更新。"""
    try:
        dlg._highlight_cleared = True
    except Exception:
        pass
    try:
        dlg._hl_clear_running = False
    except Exception:
        pass
    ph = int(parent_hwnd or 0)
    if ph:
        try:
            from core import core_w32

            core_w32.bring_to_front(ph)
        except Exception:
            try:
                import ctypes

                ctypes.windll.user32.SetForegroundWindow(int(ph))
            except Exception:
                pass
    try:
        dlg.deleteLater()
    except Exception:
        pass
    try:
        from PySide6.QtWidgets import QApplication

        app = QApplication.instance()
        if app is not None:
            app.processEvents()
    except Exception:
        pass


def _clear_dupli_highlight(
    parent_hwnd: int, payload: dict[str, Any] | None, *, progress_path: Path | None = None
) -> None:
    """svc_dupli が付与した重複セルの背景色を除去。sheet_name があればそのシートを対象にする。"""
    if not payload or not int(parent_hwnd or 0):
        _dupli_ui_diag(
            "[UI_DUPLI_HLCLR] skip no_payload_or_hwnd parent_hwnd=%r has_payload=%s",
            parent_hwnd,
            bool(payload),
        )
        return
    t_start = time.perf_counter()
    keys_sorted = sorted(str(k) for k in payload.keys())
    rects_in = payload.get("rects")
    n_rect = len(rects_in) if isinstance(rects_in, list) else 0
    runs_in = payload.get("runs")
    n_runs = len(runs_in) if isinstance(runs_in, list) else 0
    _dupli_ui_diag(
        "[UI_DUPLI_HLCLR] start parent_hwnd=%s keys=%s book_name=%r sheet_name=%r rects=%s runs=%s progress=%s",
        parent_hwnd,
        keys_sorted,
        str(payload.get("book_name") or ""),
        str(payload.get("sheet_name") or ""),
        n_rect,
        n_runs,
        bool(progress_path),
    )
    phase_clear = ""
    if progress_path is not None:
        try:
            c = _get_cfg()
            phase_clear = str((c.get("MESSAGES") or {}).get("PHASE_CLEAR") or "ハイライトを解除しています...")
        except Exception:
            phase_clear = "ハイライトを解除しています..."

    def _pe() -> None:
        try:
            from PySide6.QtWidgets import QApplication

            app = QApplication.instance()
            if app is not None:
                app.processEvents()
        except Exception:
            pass

    try:
        try:
            from xlwings import App
            from xlwings._xlwindows import App as WinApp

            app = App(impl=WinApp(xl=int(parent_hwnd)))
            book = _resolve_xlwings_book(app, str(payload.get("book_name") or ""))
            try:
                book.activate()
            except Exception:
                pass
            sn = str(payload.get("sheet_name") or "").strip()
            if sn:
                try:
                    sheet = book.sheets[sn]
                except Exception:
                    try:
                        sheet = book.sheets.active
                        _log_dupli_ui.warning(
                            "[ui_dupli] highlight_clear: sheet %r not found, using active sheet", sn
                        )
                    except Exception as exc:
                        _log_dupli_ui.warning("[ui_dupli] highlight_clear: no sheet: %s", exc)
                        _dupli_ui_diag(
                            "[UI_DUPLI_HLCLR] exit_early reason=no_sheet exc=%r elapsed_ms=%s",
                            exc,
                            int((time.perf_counter() - t_start) * 1000),
                        )
                        return
            else:
                sheet = book.sheets.active
        except Exception as exc:
            _log_dupli_ui.warning("[ui_dupli] highlight_clear: xlwings open failed: %s", exc)
            _dupli_ui_diag(
                "[UI_DUPLI_HLCLR] exit_early reason=xlwings_open_failed exc=%r elapsed_ms=%s",
                exc,
                int((time.perf_counter() - t_start) * 1000),
            )
            return

        resolved_book = str(getattr(sheet.book, "name", "") or "")
        resolved_sheet = str(getattr(sheet, "name", "") or "")
        _dupli_ui_diag(
            "[UI_DUPLI_HLCLR] resolved book=%r sheet=%r", resolved_book, resolved_sheet
        )

        api_app = app.api
        prev_screen = True
        prev_calc: Any = _XL_CALC_AUTO_UI
        api_frozen = False
        try:
            prev_screen = api_app.ScreenUpdating
            prev_calc = api_app.Calculation
            api_app.ScreenUpdating = False
            api_app.Calculation = _XL_CALC_MANUAL_UI
            api_frozen = True
        except Exception:
            pass

        try:
            t_loop = time.perf_counter()
            rects = payload.get("rects")
            cleared_rects = 0
            failed_rects = 0
            if isinstance(rects, list) and rects:
                n_rect_total = len(rects)
                last_hb = t_loop
                last_prog_ui = t_loop
                for ni, quad in enumerate(rects):
                    if not isinstance(quad, (list, tuple)) or len(quad) < 4:
                        continue
                    r1, c1, r2, c2 = int(quad[0]), int(quad[1]), int(quad[2]), int(quad[3])
                    try:
                        _dupli_clear_range_fill(sheet.range((r1, c1), (r2, c2)))
                        cleared_rects += 1
                    except Exception as exc:
                        failed_rects += 1
                        _log_dupli_ui.debug("[ui_dupli] highlight_clear rect failed: %s", exc)
                        _dupli_ui_diag(
                            "[UI_DUPLI_HLCLR] rect_fail quad=%r exc=%r", list(quad), exc
                        )
                    done_i = ni + 1
                    now = time.perf_counter()
                    if progress_path is not None and (
                        done_i % _HLCLR_PROGRESS_STRIDE == 0
                        or done_i == n_rect_total
                        or (now - last_prog_ui) >= _HLCLR_PROGRESS_UI_INTERVAL_SEC
                    ):
                        _dupli_hlclr_progress_write(
                            progress_path,
                            {
                                "status": "RUN",
                                "hide_cancel_button": True,
                                "phase": phase_clear,
                                "msg": phase_clear,
                                "pct": int(100 * done_i / max(n_rect_total, 1)),
                                "done": done_i,
                                "total": n_rect_total,
                            },
                        )
                        last_prog_ui = now
                        _pe()
                    if done_i % _HLCLR_LOG_EVERY_N_RECTS == 0:
                        _dupli_ui_diag(
                            "[UI_DUPLI_HLCLR] loop_progress processed=%s/%s cleared=%s failed=%s elapsed_ms=%s",
                            done_i,
                            n_rect_total,
                            cleared_rects,
                            failed_rects,
                            int((time.perf_counter() - t_loop) * 1000),
                        )
                    if now - last_hb >= _HLCLR_HEARTBEAT_SEC:
                        _dupli_ui_diag(
                            "[UI_DUPLI_HLCLR] heartbeat processed=%s/%s cleared=%s failed=%s loop_elapsed_ms=%s total_elapsed_ms=%s",
                            done_i,
                            n_rect_total,
                            cleared_rects,
                            failed_rects,
                            int((now - t_loop) * 1000),
                            int((now - t_start) * 1000),
                        )
                        last_hb = now
                _dupli_ui_diag(
                    "[UI_DUPLI_HLCLR] rects_done cleared=%s failed=%s sample_first=%r loop_elapsed_ms=%s total_elapsed_ms=%s",
                    cleared_rects,
                    failed_rects,
                    rects[0] if rects else None,
                    int((time.perf_counter() - t_loop) * 1000),
                    int((time.perf_counter() - t_start) * 1000),
                )
                if cleared_rects > 0:
                    _dupli_ui_diag(
                        "[UI_DUPLI_HLCLR] exit_ok path=rects cleared=%s total_elapsed_ms=%s",
                        cleared_rects,
                        int((time.perf_counter() - t_start) * 1000),
                    )
                    return

            runs = payload.get("runs") or []
            if not runs:
                if isinstance(rects, list) and rects and cleared_rects == 0:
                    _log_dupli_ui.warning(
                        "[ui_dupli] highlight_clear: rects 指定だが 1 件もクリアできず runs もなし"
                    )
                    _dupli_ui_diag("[UI_DUPLI_HLCLR] no_runs_fallback rects_all_failed")
                _dupli_ui_diag(
                    "[UI_DUPLI_HLCLR] exit reason=no_runs_payload total_elapsed_ms=%s",
                    int((time.perf_counter() - t_start) * 1000),
                )
                return
            y1 = int(payload.get("y1", 0))
            x1 = int(payload.get("x1", 0))
            xn = max(1, int(payload.get("xn", 1)))
            cleared_runs = 0
            n_run_total = len(runs) if isinstance(runs, list) else 0
            t_runs = time.perf_counter()
            for ri, se in enumerate(runs):
                if not isinstance(se, (list, tuple)) or len(se) < 2:
                    continue
                s_rel, e_rel = int(se[0]), int(se[1])
                ay_lo = y1 + s_rel
                ay_hi = y1 + e_rel
                try:
                    _dupli_clear_range_fill(sheet.range((ay_lo, x1), (ay_hi, x1 + xn - 1)))
                    cleared_runs += 1
                except Exception as exc:
                    _log_dupli_ui.debug("[ui_dupli] highlight_clear run failed: %s", exc)
                    _dupli_ui_diag("[UI_DUPLI_HLCLR] run_fail se=%r exc=%r", list(se), exc)
                if (ri + 1) % max(1, _HLCLR_LOG_EVERY_N_RECTS) == 0:
                    _dupli_ui_diag(
                        "[UI_DUPLI_HLCLR] runs_progress processed=%s/%s cleared_runs=%s runs_elapsed_ms=%s",
                        ri + 1,
                        n_run_total,
                        cleared_runs,
                        int((time.perf_counter() - t_runs) * 1000),
                    )
            _dupli_ui_diag(
                "[UI_DUPLI_HLCLR] runs_done cleared_runs=%s runs_elapsed_ms=%s total_elapsed_ms=%s",
                cleared_runs,
                int((time.perf_counter() - t_runs) * 1000),
                int((time.perf_counter() - t_start) * 1000),
            )
        finally:
            if api_frozen:
                try:
                    api_app.Calculation = prev_calc
                except Exception:
                    pass
                try:
                    api_app.ScreenUpdating = prev_screen
                except Exception:
                    pass
    finally:
        if progress_path is not None:
            try:
                c = _get_cfg()
                msg_done = str((c.get("MESSAGES") or {}).get("HIGHLIGHT_CLEAR_DONE") or "完了")
                _dupli_hlclr_progress_write(
                    progress_path,
                    {
                        "status": "DONE",
                        "show_done_dialog": False,
                        "phase": msg_done,
                        "seq": 999,
                    },
                )
            except Exception:
                pass
            _pe()
            _pe()


class _DupliProgressWrapper:
    """
    進捗ダイアログ（ui_common.create_progress_dialog の戻り値）をラップし、
    show / get_result を svc_dupli 側から扱いやすくする。
    """

    def __init__(self, progress_dlg: Any) -> None:
        self._dlg = progress_dlg

    def show(self) -> None:
        """進捗ダイアログを表示し、1 回イベント処理して描画を反映する。"""
        self._dlg.show()
        try:
            from PySide6.QtWidgets import QApplication

            QApplication.processEvents()
        except Exception:
            pass

    def get_result(self) -> dict[str, Any]:
        """進捗ダイアログの結果辞書を返す。未実装の場合は空辞書。"""
        return getattr(self._dlg, "get_result", lambda: {})()


class DupliReportDialog(QDialog):
    """
    重複検出結果を一覧表示するモードレスダイアログ。
    行・座標・内容のテーブルと「セルへ移動」「閉じる」ボタンを持つ。ダブルクリックで該当セルに移動。
    """

    def __init__(
        self,
        req: dict[str, Any],
        parent_hwnd: int,
        cfg: dict[str, Any],
        sheet_id: str = "",
    ) -> None:
        super().__init__()
        try:
            self.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose, True)
        except Exception:
            pass
        self._req = req or {}
        self._parent_hwnd = int(parent_hwnd or 0)
        self._sheet_id = str(sheet_id or "").strip()
        self._highlight_cleared = False
        self._hl_clear_running = False
        _hc = self._req.get("highlight_clear")
        _hc = _hc if isinstance(_hc, dict) else {}
        self._book_name = str(_hc.get("book_name") or "").strip()
        self._sheet_name = str(_hc.get("sheet_name") or "").strip()
        self._viewport_follow = bool(_hc.get("viewport_follow"))
        self._sidecar_path: Optional[Path] = None
        self._hl_rects: list[list[int]] = []
        self._vp_applied: list[list[int]] = []
        self._viewport_teardown_done = False
        self._vp_timer: Optional[QTimer] = None
        self._vp_skip_bounds: Optional[tuple[int, int, int, int]] = None
        self._vp_skip_goto: Optional[tuple[int, int, int, int]] = None
        self._vp_skip_paint: tuple[tuple[int, int, int, int], ...] = ()
        self._dupli_cfg = cfg
        _vp_cfg = cfg.get("VIEWPORT_HIGHLIGHT") or {}
        try:
            self._vp_margin_rows = int(_vp_cfg.get("MARGIN_ROWS", 2) or 2)
        except (TypeError, ValueError):
            self._vp_margin_rows = 2
        try:
            self._vp_margin_cols = int(_vp_cfg.get("MARGIN_COLS", 1) or 1)
        except (TypeError, ValueError):
            self._vp_margin_cols = 1
        try:
            self._vp_poll_ms = int(_vp_cfg.get("POLL_MS", 450) or 450)
        except (TypeError, ValueError):
            self._vp_poll_ms = 450
        self._vp_poll_ms = max(80, min(2000, self._vp_poll_ms))
        _rp = str(_hc.get("rects_path") or "").strip()
        if _rp:
            self._sidecar_path = Path(_rp)
            try:
                blob = ipc_file.read_pickle(self._sidecar_path)
                if isinstance(blob, dict):
                    rl = blob.get("rects")
                    if isinstance(rl, list):
                        self._hl_rects = [
                            [int(q[0]), int(q[1]), int(q[2]), int(q[3])]
                            for q in rl
                            if isinstance(q, (list, tuple)) and len(q) >= 4
                        ]
            except Exception as exc:
                _log_dupli_ui.warning("[ui_dupli] viewport: sidecar load failed %s: %s", _rp, exc)
        try:
            self._fill_bgr = int(_hc.get("fill_bgr") or 0)
        except (TypeError, ValueError):
            self._fill_bgr = 0
        if self._fill_bgr == 0:
            self._fill_bgr = _dupli_fill_bgr_from_cfg(cfg)
        self._goto_hl_bgr = _goto_highlight_bgr_from_cfg(cfg)
        self._goto_hl_quad: Optional[list[int]] = None
        _rn_embed = len(_hc["rects"]) if isinstance(_hc.get("rects"), list) else -1
        _run_n = len(_hc["runs"]) if isinstance(_hc.get("runs"), list) else 0
        _dupli_ui_diag(
            "[UI_DUPLI_REPORT] open parent_hwnd=%s sheet_id=%s highlight_clear_keys=%s rects_embed_n=%s rects_loaded_n=%s runs_n=%s viewport=%s sidecar=%r",
            self._parent_hwnd,
            self._sheet_id,
            sorted(str(k) for k in _hc.keys()),
            _rn_embed,
            len(self._hl_rects),
            _run_n,
            self._viewport_follow,
            _rp,
        )
        rep = ((cfg.get("SCREENS") or {}).get("REPORT") or {})
        self._link_col = int(self._req.get("link_col", 1))
        if self._link_col < 0:
            self._link_col = 0
        _mbw = rep.get("MODE_B_WINDOW")
        _mbw = _mbw if isinstance(_mbw, dict) else {}
        _win_rep = dict(rep.get("WINDOW") or {})
        if self._link_col == 2 and _mbw:
            for _wk in ("DEFAULT_WIDTH", "DEFAULT_HEIGHT"):
                if _wk in _mbw and _mbw[_wk]:
                    try:
                        _win_rep[_wk] = int(_mbw[_wk])
                    except (TypeError, ValueError):
                        pass
        title = str(self._req.get("title") or "").strip() or str(rep.get("TITLE_TEMPLATE") or "重複レポート")
        self.setWindowTitle(title)

        from ui_qt.ui_common import apply_tooltip_if_set, apply_window_config

        apply_window_config(self, {"WINDOW": _win_rep}, self._parent_hwnd, "REPORT")
        self._hc_prepare_window_cfg = dict(_win_rep)
        if self._link_col == 2 and _mbw:
            try:
                _mw = int(_mbw.get("MIN_WIDTH") or 0)
                _mh = int(_mbw.get("MIN_HEIGHT") or 0)
                if _mw > 0 or _mh > 0:
                    self.setMinimumSize(max(_mw, 0), max(_mh, 0))
            except (TypeError, ValueError):
                pass
        apply_tooltip_if_set(self, rep, "TOOLTIP")

        intro_text = str(self._req.get("report_intro") or "").strip()
        dup_n = int(self._req.get("dup_count") or 0)
        count_tpl = str(self._req.get("count_caption_template") or "").strip() or "検出総数: {count} 件"
        try:
            count_line = count_tpl.format(count=dup_n)
        except Exception:
            count_line = f"検出総数: {dup_n} 件"

        headers: list[dict[str, Any]] = list(self._req.get("headers") or rep.get("COLUMNS") or [])
        rows: list[list[Any]] = list(self._req.get("rows") or [])
        self._addresses = [str(x) for x in (self._req.get("addresses") or [])]

        ncol = len(headers) if headers else (max((len(r) for r in rows), default=0) or 3)
        ncol = max(ncol, 1)

        labels: list[str] = []
        for i in range(ncol):
            if i < len(headers):
                h = headers[i]
                labels.append(str(h.get("label") or h.get("key") or f"C{i+1}"))
            else:
                labels.append(f"C{i+1}")

        tbl = QTableWidget(len(rows), ncol)
        tbl.setHorizontalHeaderLabels(labels)
        tbl.horizontalHeader().setDefaultAlignment(
            Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter
        )
        fixed_widths: list[int] = []
        for ci in range(ncol):
            fw = 0
            if ci < len(headers):
                try:
                    fw = int(headers[ci].get("width", 0) or 0)
                except (TypeError, ValueError):
                    fw = 0
            fixed_widths.append(fw)
            if fw > 0:
                tbl.setColumnWidth(ci, max(48, fw))
        tbl.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        tbl.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        tbl.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        # セル値の省略記号(...)を出さず、横スクロールで全文を確認できるようにする。
        tbl.setWordWrap(False)
        tbl.setTextElideMode(Qt.TextElideMode.ElideNone)
        tbl.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        tbl.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)
        tt_tbl = str(rep.get("TABLE_TOOLTIP") or "").strip()
        if tt_tbl:
            tbl.setToolTip(tt_tbl)

        for ci in range(min(ncol, len(headers))):
            ht = str(headers[ci].get("tooltip") or headers[ci].get("TOOLTIP") or "").strip()
            if ht:
                it = tbl.horizontalHeaderItem(ci)
                if it is not None:
                    it.setToolTip(ht)

        # 重複行の先頭列を薄いオレンジで区別
        dup_brush = QBrush(QColor(255, 245, 230))
        for ri, row in enumerate(rows):
            for ci in range(ncol):
                val = row[ci] if ci < len(row) else ""
                it = QTableWidgetItem(str(val))
                it.setFlags(it.flags() & ~Qt.ItemFlag.ItemIsEditable)
                if ci == 0:
                    it.setBackground(dup_brush)
                tbl.setItem(ri, ci, it)

        for ci, fw in enumerate(fixed_widths):
            if fw > 0:
                tbl.setColumnWidth(ci, max(48, fw))

        _hdr_rep = tbl.horizontalHeader()
        _is_mode_b = self._link_col == 2
        if _is_mode_b:
            try:
                _vmax = int(_mbw.get("VALUE_COL_MAX_WIDTH") or 420)
            except (TypeError, ValueError):
                _vmax = 420
            _vmax = max(120, _vmax)
            tbl.setWordWrap(True)
            _hdr = tbl.horizontalHeader()
            _hdr.setSectionResizeMode(0, QHeaderView.ResizeMode.Interactive)
            _w0 = tbl.columnWidth(0)
            tbl.setColumnWidth(0, min(_vmax, max(120, _w0)))
            if ncol > 1:
                _hdr.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
            tbl.resizeRowsToContents()

        lbl_intro: Optional[QLabel] = None
        btn_goto = QPushButton(str(rep.get("BTN_GOTO") or "セルへ移動"))
        btn_close = QPushButton(str(rep.get("BTN_CLOSE") or "閉じる"))
        tt_goto = str(rep.get("BTN_GOTO_TOOLTIP") or "").strip()
        tt_close = str(rep.get("BTN_CLOSE_TOOLTIP") or "").strip()
        if tt_goto:
            btn_goto.setToolTip(tt_goto)
        if tt_close:
            btn_close.setToolTip(tt_close)
        btn_goto.clicked.connect(lambda: self._do_goto(tbl))
        btn_close.clicked.connect(self._on_report_close_clicked)

        lay_h = QHBoxLayout()
        lay_h.addStretch(1)
        lay_h.addWidget(btn_goto)
        lay_h.addWidget(btn_close)
        lay = QVBoxLayout(self)
        try:
            _mw_cap0 = int(((rep.get("WINDOW") or {}).get("MAX_WIDTH")) or 920)
        except Exception:
            _mw_cap0 = 920
        _mw_cap0 = max(260, min(1200, int(_mw_cap0))) - 40
        if intro_text:
            lbl_intro = QLabel(intro_text)
            try:
                lbl_intro.setMaximumWidth(int(_mw_cap0))
            except Exception:
                pass
            lbl_intro.setWordWrap(True)
            try:
                lbl_intro.setTextFormat(Qt.TextFormat.PlainText)
            except Exception:
                pass
            tt_intro = str(rep.get("REPORT_INTRO_TOOLTIP") or "").strip()
            if tt_intro:
                lbl_intro.setToolTip(tt_intro)
            lay.addWidget(lbl_intro)
        lbl_count = QLabel(count_line)
        try:
            lbl_count.setMaximumWidth(int(_mw_cap0))
        except Exception:
            pass
        try:
            lbl_count.setTextFormat(Qt.TextFormat.PlainText)
        except Exception:
            pass
        tt_count = str(rep.get("COUNT_CAPTION_TOOLTIP") or "").strip()
        if tt_count:
            lbl_count.setToolTip(tt_count)
        lay.addWidget(lbl_count)
        lay.addWidget(tbl)
        lay.addLayout(lay_h)

        self._lbl_report_intro = lbl_intro
        self._lbl_report_count = lbl_count
        self._tbl = tbl
        tbl.cellDoubleClicked.connect(self._on_cell_double)
        _open_report_dialogs.append(self)  # GC 対策でリストに登録
        self.finished.connect(self._on_finished)
        self._vp_timer = QTimer(self)
        self._vp_timer.timeout.connect(self._viewport_highlight_tick)  # type: ignore[attr-defined]
        self._vp_timer.setInterval(self._vp_poll_ms)
        self._nudge_report_last_sec: float = 0.0
        if not _is_mode_b:
            self._rep_report_for_layout = dict(rep)
            self._apply_mode_a_report_table_layout(
                tbl, rep, ncol, fixed_widths, intro_text, count_line
            )
            self._clamp_mode_a_labels_to_dialog_width()
            self._mode_a_layout_snap = (intro_text, count_line, list(fixed_widths), ncol)
            try:
                QTimer.singleShot(0, self._reapply_mode_a_report_table_layout_deferred)
            except Exception:
                pass

    def _clamp_mode_a_labels_to_dialog_width(self) -> None:
        """説明・件数ラベルが1行想定の sizeHint で外枠を押し広げないよう、幅をダイアログに合わせる。"""
        try:
            w = int(self.width())
            if w <= 0:
                return
            cap = max(200, w - 32)
            li = getattr(self, "_lbl_report_intro", None)
            lc = getattr(self, "_lbl_report_count", None)
            if li is not None:
                li.setMaximumWidth(int(cap))
            if lc is not None:
                lc.setMaximumWidth(int(cap))
        except Exception:
            pass

    def _reapply_mode_a_report_table_layout_deferred(self) -> None:
        """prepare_dialog_excel_center_before_show の adjustSize 後に列・外枠幅を再適用する。"""
        if int(getattr(self, "_link_col", 0) or 0) == 2:
            return
        sp = getattr(self, "_mode_a_layout_snap", None)
        if not sp or len(sp) < 4:
            return
        intro, count, fw, nc = sp[0], sp[1], sp[2], int(sp[3])
        try:
            tbl = getattr(self, "_tbl", None)
            rep = getattr(self, "_rep_report_for_layout", None)
            if tbl is None or not isinstance(rep, dict):
                return
            self._apply_mode_a_report_table_layout(tbl, rep, nc, fw, intro, count)
            self._clamp_mode_a_labels_to_dialog_width()
        except Exception:
            pass

    def _apply_mode_a_report_table_layout(
        self,
        tbl: QTableWidget,
        rep: dict[str, Any],
        ncol: int,
        fixed_widths: list[int],
        intro_text: str,
        count_line: str,
    ) -> None:
        """
        モード A: 列幅はセル表示値＋余白。外枠幅は tw_body（列合計＋TABLE_CHROME_WIDTH）＋
        メインレイアウトの左右マージン＋DIALOG_WIDTH_FUDGE。Qt ヘッダ length() や説明文幅は使わない。
        MIN〜MAX に収める。intro_text/count_line は呼び出し互換のため残す（幅計算には使わない）。
        """
        win = rep.get("WINDOW") if isinstance(rep.get("WINDOW"), dict) else {}
        try:
            wmax = int(win.get("MAX_WIDTH") or 0)
        except (TypeError, ValueError):
            wmax = 0
        if wmax <= 0:
            wmax = 920
        try:
            wmin = int(win.get("MIN_WIDTH") or 0)
        except (TypeError, ValueError):
            wmin = 0
        if wmin <= 0:
            wmin = 400
        wmax = max(wmin + 80, int(wmax))
        try:
            summary_cap = int(rep.get("SUMMARY_COL_MAX_WIDTH") or 0)
        except (TypeError, ValueError):
            summary_cap = 0
        # 0 以下は上限制限なし（長文は横スクロールで閲覧する）。
        if summary_cap > 0:
            summary_cap = max(120, min(10000, int(summary_cap)))
        try:
            col_pad = int(rep.get("TABLE_COLUMN_PAD") or 28)
        except (TypeError, ValueError):
            col_pad = 28
        col_pad = max(8, min(80, int(col_pad)))
        try:
            dlg_fudge = int(rep.get("DIALOG_WIDTH_FUDGE", 12))
        except (TypeError, ValueError):
            dlg_fudge = 12
        dlg_fudge = max(0, min(64, int(dlg_fudge)))
        try:
            table_chrome = int(rep.get("TABLE_CHROME_WIDTH") or 72)
        except (TypeError, ValueError):
            table_chrome = 72
        table_chrome = max(40, min(200, int(table_chrome)))

        margins_lr = 20
        try:
            lay_m = self.layout()
            if lay_m is not None:
                try:
                    lay_m.activate()
                except Exception:
                    pass
                _cm = lay_m.contentsMargins()
                margins_lr = int(_cm.left() + _cm.right())
        except Exception:
            pass
        if margins_lr <= 0:
            margins_lr = 20

        try:
            self.setMaximumWidth(int(wmax))
        except Exception:
            pass
        try:
            self.setMinimumWidth(int(min(wmax, wmin)))
        except Exception:
            pass

        hdr = tbl.horizontalHeader()
        if ncol <= 1:
            hdr.setStretchLastSection(True)
            hdr.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
            try:
                dw = int(min(wmax, max(wmin, int(win.get("DEFAULT_WIDTH") or wmin))))
                self.resize(dw, max(200, int(self.height())))
                setattr(self, "_mode_a_target_width", int(dw))
            except Exception:
                pass
            return

        # 4列構成などでは最終列 Stretch にしない（余白が「重複内容」列に溜まるのを防ぐ）。5列以上で最終列のみ Stretch。
        stretch_last = ncol >= 5
        bulk_end = ncol - 1 if stretch_last else ncol

        naturals: list[int] = []
        summary_header_min = 0
        for ci in range(bulk_end):
            if ci < len(fixed_widths) and fixed_widths[ci] > 0:
                naturals.append(max(48, int(fixed_widths[ci])))
                continue
            cw = _dupli_report_column_content_width(tbl, ci) + col_pad
            w = max(52, int(cw))
            if ci == 2:
                summary_header_min = _dupli_report_header_text_width(tbl, ci) + col_pad + 10
                summary_header_min = max(120, int(summary_header_min))
                if summary_cap > 0:
                    w = min(w, summary_cap)
                w = max(w, summary_header_min)
            naturals.append(w)

        budget_inner = int(wmax) - int(table_chrome) - int(margins_lr) - int(dlg_fudge)
        if stretch_last:
            budget_inner -= 72
        budget_inner = max(len(naturals) * 44 + 8, int(budget_inner))

        # モードA（非 Stretch）は列を予算内に圧縮しない。
        # 圧縮すると横スクロール対象が消え、長文末尾を辿れなくなるため。
        if stretch_last and sum(naturals) > budget_inner:
            naturals = _dupli_scale_natural_widths_to_budget(naturals, budget_inner)
        if (
            not stretch_last
            and ncol >= 3
            and len(naturals) >= 3
            and int(summary_header_min) > 0
        ):
            naturals = _dupli_rebalance_widths_for_min_column(
                naturals, 2, int(summary_header_min), floor_each=44
            )

        for ci, w in enumerate(naturals):
            tbl.setColumnWidth(ci, int(w))

        if stretch_last:
            hdr.setStretchLastSection(True)
            for ci in range(bulk_end):
                hdr.setSectionResizeMode(ci, QHeaderView.ResizeMode.Interactive)
            hdr.setSectionResizeMode(ncol - 1, QHeaderView.ResizeMode.Stretch)
        else:
            hdr.setStretchLastSection(False)
            for ci in range(ncol):
                hdr.setSectionResizeMode(ci, QHeaderView.ResizeMode.Interactive)

        try:
            tbl.updateGeometry()
        except Exception:
            pass
        # 一覧の列幅合計＋表まわり固定幅（行ヘッダ・罫線・スクロールバー予備など）。ヘッダ length() は使わない。
        cols_sum = int(sum(naturals))
        tw_body = cols_sum + int(table_chrome)
        tw_body = max(56, int(tw_body))

        try:
            wdef = int(win.get("DEFAULT_WIDTH") or 0)
        except (TypeError, ValueError):
            wdef = 0
        base_w = int(wdef) if wdef > 0 else int(tw_body + int(margins_lr) + int(dlg_fudge))
        dlg_w = int(min(wmax, max(wmin, base_w)))
        try:
            setattr(self, "_mode_a_target_width", int(dlg_w))
        except Exception:
            pass

        try:
            cap_lbl = max(160, int(dlg_w) - 32)
            li = getattr(self, "_lbl_report_intro", None)
            lc = getattr(self, "_lbl_report_count", None)
            if li is not None:
                li.setMaximumWidth(cap_lbl)
            if lc is not None:
                lc.setMaximumWidth(cap_lbl)
        except Exception:
            pass

        try:
            h0 = int(self.height())
        except Exception:
            h0 = 0
        if h0 < 120:
            try:
                h0 = int(win.get("DEFAULT_HEIGHT") or 500)
            except Exception:
                h0 = 500
        try:
            self.resize(dlg_w, int(h0))
        except Exception:
            pass
        self._clamp_mode_a_labels_to_dialog_width()

        try:
            tbl.setMinimumWidth(56)
            tbl.setMaximumWidth(16777215)
            tbl.setSizePolicy(
                QSizePolicy.Policy.Expanding,
                QSizePolicy.Policy.Expanding,
            )
        except Exception:
            pass

    def _pulse_excel_for_sheet_edit(self) -> None:
        """レポート表示中もシートで編集できるよう、子 HWND 再有効化と Interactive を試す。"""
        try:
            if not self.isVisible():
                return
        except Exception:
            return
        ph = int(self._parent_hwnd or 0)
        if not ph or os.name != "nt":
            return
        try:
            from shiboken6 import Shiboken

            if not Shiboken.isValid(self):
                return
        except Exception:
            pass
        try:
            from ui_qt.ui_common import enable_excel_window

            enable_excel_window(ph, True)
        except Exception:
            pass
        try:
            from core.core_xlc import (
                excel_try_set_main_commandbars_enabled,
                get_excel_context_from_hwnd,
            )

            ctx = get_excel_context_from_hwnd(ph, self._sheet_id)
            if ctx:
                app, *_rest = ctx
                excel_try_set_main_commandbars_enabled(app, True)
                try:
                    ax = getattr(app, "api", None)
                    if ax is not None:
                        ax.Interactive = True
                except Exception:
                    pass
        except Exception:
            pass

    def _nudge_report_above_excel(self, _reason: str = "") -> None:
        ph = int(self._parent_hwnd or 0)
        if not ph or not self.isVisible():
            return
        try:
            now = time.monotonic()
            if (
                now - float(getattr(self, "_nudge_report_last_sec", 0.0) or 0.0)
                < _DUPLI_REPORT_ENSURE_THROTTLE_SEC
            ):
                return
        except Exception:
            pass
        try:
            from ui_qt.ui_common import ensure_front

            ensure_front(self, ph)
            self._nudge_report_last_sec = time.monotonic()
        except Exception:
            pass

    def changeEvent(self, event) -> None:  # type: ignore[override]
        try:
            if event.type() == QEvent.Type.ActivationChange and self.isActiveWindow():
                if int(self._parent_hwnd or 0):

                    def _nudge() -> None:
                        self._nudge_report_above_excel("activation_change")

                    QTimer.singleShot(0, _nudge)
        except Exception:
            pass
        super().changeEvent(event)

    def event(self, e) -> bool:  # type: ignore[override]
        try:
            et = e.type()
            if et in (QEvent.Type.WindowActivate, QEvent.Type.WindowDeactivate):
                try:
                    vis = bool(self.isVisible())
                except Exception:
                    vis = False
                try:
                    act = bool(self.isActiveWindow())
                except Exception:
                    act = False
                _dupli_ui_diag(
                    "[UI_DUPLI_REPORT_FG] %s parent_hwnd=%s sheet_id=%s visible=%s active_window=%s",
                    "WindowActivate" if et == QEvent.Type.WindowActivate else "WindowDeactivate",
                    self._parent_hwnd,
                    self._sheet_id,
                    vis,
                    act,
                )
            elif et in (QEvent.Type.FocusIn, QEvent.Type.FocusOut):
                try:
                    vis = bool(self.isVisible())
                except Exception:
                    vis = False
                _fr = ""
                try:
                    from PySide6.QtGui import QFocusEvent

                    if isinstance(e, QFocusEvent):
                        _fr = str(e.reason())
                except Exception:
                    pass
                _dupli_ui_diag(
                    "[UI_DUPLI_REPORT_FG] %s parent_hwnd=%s sheet_id=%s visible=%s focus_reason=%s",
                    "FocusIn" if et == QEvent.Type.FocusIn else "FocusOut",
                    self._parent_hwnd,
                    self._sheet_id,
                    vis,
                    _fr,
                )
        except Exception:
            pass
        return super().event(e)

    def showEvent(self, event) -> None:  # type: ignore[override]
        super().showEvent(event)
        ph = int(self._parent_hwnd or 0)
        if ph and os.name == "nt":
            for _ms in _DUPLI_REPORT_EXCEL_UNLOCK_PULSE_MS:
                try:
                    QTimer.singleShot(int(_ms), self._pulse_excel_for_sheet_edit)
                except Exception:
                    pass
        if self._viewport_follow and self._hl_rects and self._parent_hwnd:
            try:
                QTimer.singleShot(0, self._viewport_highlight_tick)
            except Exception:
                pass
            try:
                self._vp_timer.start()
            except Exception:
                pass

    def _strip_goto_highlight_fill(self) -> None:
        """ジャンプ一時グレーのみ COM で落とす（quad をクリア）。次の viewport ティックで重複色は復帰しうる。"""
        gq = self._goto_hl_quad
        if not gq or len(gq) < 4:
            self._goto_hl_quad = None
            return
        if not int(self._parent_hwnd or 0):
            self._goto_hl_quad = None
            return
        try:
            from xlwings import App
            from xlwings._xlwindows import App as WinApp

            app = App(impl=WinApp(xl=int(self._parent_hwnd)))
            book = _resolve_xlwings_book(app, self._book_name)
            sn = str(self._sheet_name or "").strip()
            if sn:
                try:
                    sh = book.sheets[sn]
                except Exception:
                    sh = book.sheets.active
            else:
                sh = book.sheets.active
            r1, c1, r2, c2 = int(gq[0]), int(gq[1]), int(gq[2]), int(gq[3])
            _dupli_clear_range_fill(sh.range((r1, c1), (r2, c2)))
        except Exception as exc:
            _dupli_ui_diag(
                "[UI_DUPLI_GOTO_HL] strip_failed parent_hwnd=%s exc=%r",
                self._parent_hwnd,
                exc,
            )
        self._goto_hl_quad = None

    def _apply_goto_gray_com(self, quad: list[int]) -> None:
        if len(quad) < 4 or not int(self._parent_hwnd or 0):
            return
        try:
            from xlwings import App
            from xlwings._xlwindows import App as WinApp

            app = App(impl=WinApp(xl=int(self._parent_hwnd)))
            book = _resolve_xlwings_book(app, self._book_name)
            sn = str(self._sheet_name or "").strip()
            if sn:
                try:
                    sh = book.sheets[sn]
                except Exception:
                    sh = book.sheets.active
            else:
                sh = book.sheets.active
            r1, c1, r2, c2 = int(quad[0]), int(quad[1]), int(quad[2]), int(quad[3])
            sh.range((r1, c1), (r2, c2)).color = int(self._goto_hl_bgr)
        except Exception as exc:
            _dupli_ui_diag(
                "[UI_DUPLI_GOTO_HL] apply_gray_failed parent_hwnd=%s exc=%r",
                self._parent_hwnd,
                exc,
            )

    def _teardown_viewport_highlight(self) -> None:
        try:
            self._strip_goto_highlight_fill()
        except Exception:
            pass
        if not self._viewport_follow:
            return
        if getattr(self, "_viewport_teardown_done", False):
            return
        self._viewport_teardown_done = True
        try:
            if self._vp_timer is not None:
                self._vp_timer.stop()
        except Exception:
            pass
        try:
            self._viewport_clear_applied_ranges()
        except Exception as exc:
            _log_dupli_ui.warning("[ui_dupli] viewport teardown clear: %s", exc)
        try:
            sp = self._sidecar_path
            if sp is not None and sp.is_file():
                sp.unlink(missing_ok=True)
        except OSError:
            pass
        self._highlight_cleared = True

    def _vp_invalidate_skip_cache(self) -> None:
        """表示クリア・シート切替後は VisibleRange スキップ判定を無効化する。"""
        self._vp_skip_bounds = None
        self._vp_skip_goto = None
        self._vp_skip_paint = ()

    def _viewport_clear_applied_ranges(self) -> None:
        """直前ティックで付与した範囲だけ Interior クリア。"""
        self._vp_invalidate_skip_cache()
        if not self._vp_applied or not int(self._parent_hwnd or 0):
            self._vp_applied = []
            return
        quads = list(self._vp_applied)
        self._vp_applied = []
        try:
            from xlwings import App
            from xlwings._xlwindows import App as WinApp

            app = App(impl=WinApp(xl=int(self._parent_hwnd)))
            book = _resolve_xlwings_book(app, self._book_name)
            sn = str(self._sheet_name or "").strip()
            if sn:
                try:
                    sh = book.sheets[sn]
                except Exception:
                    sh = book.sheets.active
            else:
                sh = book.sheets.active
            for quad in quads:
                if len(quad) < 4:
                    continue
                r1, c1, r2, c2 = int(quad[0]), int(quad[1]), int(quad[2]), int(quad[3])
                try:
                    _dupli_clear_range_fill(sh.range((r1, c1), (r2, c2)))
                except Exception:
                    pass
        except Exception as exc:
            _dupli_ui_diag(
                "[UI_DUPLI_VP] clear_applied_failed parent_hwnd=%s exc=%r",
                self._parent_hwnd,
                exc,
            )

    def _viewport_highlight_tick(self) -> None:
        if (
            not self._viewport_follow
            or getattr(self, "_viewport_teardown_done", False)
            or self._highlight_cleared
            or not int(self._parent_hwnd or 0)
            or not self._hl_rects
        ):
            return
        try:
            from xlwings import App
            from xlwings._xlwindows import App as WinApp

            app = App(impl=WinApp(xl=int(self._parent_hwnd)))
            book = _resolve_xlwings_book(app, self._book_name)
            api = app.api
            ab = getattr(api, "ActiveWorkbook", None)
            if ab is None:
                self._viewport_clear_applied_ranges()
                return
            try:
                if not _dupli_workbook_names_match(str(ab.Name), str(book.api.Name)):
                    self._viewport_clear_applied_ranges()
                    return
            except Exception:
                self._viewport_clear_applied_ranges()
                return
            sn = str(self._sheet_name or "").strip()
            try:
                ash = book.sheets.active
            except Exception:
                self._vp_invalidate_skip_cache()
                return
            if sn and str(ash.name) != sn:
                self._viewport_clear_applied_ranges()
                return
            sh = ash
            aw = api.ActiveWindow
            if aw is None:
                self._vp_invalidate_skip_cache()
                return
            vis = aw.VisibleRange
            vr1 = int(vis.Row)
            vc1 = int(vis.Column)
            vr2 = vr1 + int(vis.Rows.Count) - 1
            vc2 = vc1 + int(vis.Columns.Count) - 1
            bounds = (vr1, vc1, vr2, vc2)
            er1, ec1, er2, ec2 = _dupli_expand_visible_bounds(
                vr1, vc1, vr2, vc2, self._vp_margin_rows, self._vp_margin_cols
            )
            to_paint: list[list[int]] = []
            for quad in self._hl_rects:
                hit = _dupli_intersect_visible_quad(er1, ec1, er2, ec2, quad)
                if hit is not None:
                    to_paint.append(hit)
            gq_pre = self._goto_hl_quad
            if gq_pre and len(gq_pre) >= 4:
                goto_sig = (
                    int(gq_pre[0]),
                    int(gq_pre[1]),
                    int(gq_pre[2]),
                    int(gq_pre[3]),
                )
            else:
                goto_sig = None
            paint_sig = tuple(
                (int(q[0]), int(q[1]), int(q[2]), int(q[3])) for q in to_paint if len(q) >= 4
            )
            if (
                self._vp_skip_bounds == bounds
                and self._vp_skip_goto == goto_sig
                and self._vp_skip_paint == paint_sig
            ):
                return
            prev_screen = True
            prev_calc: Any = _XL_CALC_AUTO_UI
            try:
                prev_screen = api.ScreenUpdating
                prev_calc = api.Calculation
                api.ScreenUpdating = False
                api.Calculation = _XL_CALC_MANUAL_UI
                for quad in self._vp_applied:
                    if len(quad) < 4:
                        continue
                    r1, c1, r2, c2 = int(quad[0]), int(quad[1]), int(quad[2]), int(quad[3])
                    try:
                        _dupli_clear_range_fill(sh.range((r1, c1), (r2, c2)))
                    except Exception:
                        pass
                self._vp_applied = []
                for quad in to_paint:
                    r1, c1, r2, c2 = int(quad[0]), int(quad[1]), int(quad[2]), int(quad[3])
                    try:
                        sh.range((r1, c1), (r2, c2)).color = int(self._fill_bgr)
                        self._vp_applied.append([r1, c1, r2, c2])
                    except Exception:
                        pass
                gq = self._goto_hl_quad
                if gq and len(gq) >= 4:
                    hit_g = _dupli_intersect_visible_quad(er1, ec1, er2, ec2, gq)
                    if hit_g is not None:
                        gr1, gc1, gr2, gc2 = (
                            int(hit_g[0]),
                            int(hit_g[1]),
                            int(hit_g[2]),
                            int(hit_g[3]),
                        )
                        try:
                            sh.range((gr1, gc1), (gr2, gc2)).color = int(self._goto_hl_bgr)
                        except Exception:
                            pass
            finally:
                try:
                    api.Calculation = prev_calc
                except Exception:
                    pass
                try:
                    api.ScreenUpdating = prev_screen
                except Exception:
                    pass
            self._vp_skip_bounds = bounds
            self._vp_skip_goto = goto_sig
            self._vp_skip_paint = paint_sig
        except Exception as exc:
            self._vp_invalidate_skip_cache()
            _dupli_ui_diag(
                "[UI_DUPLI_VP] tick_failed parent_hwnd=%s sheet_id=%s exc=%r",
                self._parent_hwnd,
                self._sheet_id,
                exc,
            )

    def _on_report_close_clicked(self) -> None:
        """閉じるボタン: accept() は closeEvent を経由しないため、× と同じ close() 経路に寄せる。"""
        try:
            self.setResult(QDialog.DialogCode.Accepted)
        except Exception:
            pass
        self.close()

    def _clear_highlight_once(self) -> None:
        if self._highlight_cleared:
            _dupli_ui_diag(
                "[UI_DUPLI_HLCLR] skip already_cleared parent_hwnd=%s sheet_id=%s",
                self._parent_hwnd,
                self._sheet_id,
            )
            return
        if getattr(self, "_hl_clear_running", False):
            return
        self._hl_clear_running = True
        try:
            pl = self._req.get("highlight_clear")
            if isinstance(pl, dict):
                _rk = sorted(str(k) for k in pl.keys())
                _rlist = pl.get("rects")
                _rn = len(_rlist) if isinstance(_rlist, list) else -1
                _runs = pl.get("runs")
                _rnn = len(_runs) if isinstance(_runs, list) else 0
                _dupli_ui_diag(
                    "[UI_DUPLI_HLCLR] invoke_clear parent_hwnd=%s sheet_id=%s keys=%s rects_n=%s runs_n=%s",
                    self._parent_hwnd,
                    self._sheet_id,
                    _rk,
                    _rn,
                    _rnn,
                )
                if isinstance(_rlist, list) and len(_rlist) > 0:
                    _log_dupli_ui.warning(
                        "[ui_dupli] sync _clear_highlight_once with rects>0 unexpected; use closeEvent defer path"
                    )
                    _clear_dupli_highlight(self._parent_hwnd, pl)
                else:
                    _clear_dupli_highlight(self._parent_hwnd, pl)
            else:
                _dupli_ui_diag(
                    "[UI_DUPLI_HLCLR] no_dict_payload parent_hwnd=%s sheet_id=%s type=%r",
                    self._parent_hwnd,
                    self._sheet_id,
                    type(pl).__name__,
                )
            self._highlight_cleared = True
        except Exception as exc:
            _log_dupli_ui.warning("[ui_dupli] highlight_clear on report close failed: %s", exc)
        finally:
            self._hl_clear_running = False

    def closeEvent(self, event) -> None:  # type: ignore[override]
        """閉じるときに着色解除のうえ Excel を有効化・前面化し、枠残りを抑える。"""
        try:
            from ui_qt.ui_common import teardown_feature_ui_shared_state

            teardown_feature_ui_shared_state(
                parent_hwnd=int(self._parent_hwnd or 0),
                modeless_widget=self,
                excel_unlock=False,
            )
        except Exception:
            pass
        try:
            self._strip_goto_highlight_fill()
        except Exception:
            pass
        try:
            _wa = bool(self.testAttribute(Qt.WidgetAttribute.WA_DeleteOnClose))
        except Exception:
            _wa = False
        _dupli_ui_diag(
            "[UI_DUPLI_REPORT] closeEvent enter parent_hwnd=%s sheet_id=%s wa_delete_on_close=%s highlight_already_cleared=%s",
            self._parent_hwnd,
            self._sheet_id,
            _wa,
            self._highlight_cleared,
        )
        self._teardown_viewport_highlight()
        pl0 = self._req.get("highlight_clear")
        rlist0 = pl0.get("rects") if isinstance(pl0, dict) else None
        defer_hlclr = (
            isinstance(pl0, dict)
            and isinstance(rlist0, list)
            and len(rlist0) > 0
            and not self._highlight_cleared
            and not self._hl_clear_running
        )
        if defer_hlclr:
            self._hl_clear_running = True
            try:
                self.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose, False)
            except Exception:
                pass
            try:
                event.accept()
            except Exception:
                pass
            ph = int(self._parent_hwnd or 0)
            sid = str(self._sheet_id or "").strip() or "_"
            pl_copy: dict[str, Any] = dict(pl0)
            dlg = self
            if ph:
                try:
                    from ui_qt.ui_common import enable_excel_window

                    enable_excel_window(ph, True)
                except Exception:
                    pass
            try:
                self.hide()
            except Exception:
                pass
            try:
                self.setWindowOpacity(0.0)
            except Exception:
                pass

            def _phase1() -> None:
                try:
                    cfg = _get_cfg()
                    msg_clear = str(
                        (cfg.get("MESSAGES") or {}).get("PHASE_CLEAR") or "ハイライトを解除しています..."
                    )
                    prog = _hlclr_progress_path(sid)
                    rects = pl_copy.get("rects")
                    n_rect = len(rects) if isinstance(rects, list) else 0
                    if n_rect <= 0:
                        try:
                            _clear_dupli_highlight(ph, pl_copy)
                        except Exception as exc:
                            _log_dupli_ui.warning("[ui_dupli] hlclr phase1 empty rects clear: %s", exc)
                        _dupli_hlclr_finish_deferred(dlg, ph)
                        return
                    _dupli_ui_diag(
                        "[UI_DUPLI_HLCLR] defer phase1 submit progress parent_hwnd=%s sheet_id=%s rects=%s",
                        ph,
                        sid,
                        n_rect,
                    )
                    _dupli_hlclr_progress_write(
                        prog,
                        {
                            "status": "RUN",
                            "phase": msg_clear,
                            "msg": msg_clear,
                            "pct": 0,
                            "done": 0,
                            "total": n_rect,
                        },
                    )
                    _submit_hlclr_progress_ui(ph, sid, prog)

                    def _phase2() -> None:
                        try:
                            _clear_dupli_highlight(ph, pl_copy, progress_path=prog)
                        except Exception as exc:
                            _log_dupli_ui.warning("[ui_dupli] hlclr phase2 COM clear failed: %s", exc)
                        finally:
                            _dupli_hlclr_finish_deferred(dlg, ph)

                    QTimer.singleShot(_HLCLR_UI_LEAD_MS, _phase2)
                except Exception as exc:
                    _log_dupli_ui.warning("[ui_dupli] hlclr phase1 failed: %s", exc)
                    try:
                        _clear_dupli_highlight(ph, pl_copy)
                    except Exception:
                        pass
                    _dupli_hlclr_finish_deferred(dlg, ph)

            QTimer.singleShot(0, _phase1)
            try:
                from PySide6.QtWidgets import QApplication

                QApplication.processEvents()
            except Exception:
                pass
            super().closeEvent(event)
            return

        self._clear_highlight_once()
        try:
            event.accept()
        except Exception:
            pass
        ph = self._parent_hwnd
        if ph:
            try:
                from ui_qt.ui_common import enable_excel_window

                enable_excel_window(ph, True)
            except Exception:
                pass
            try:
                from core import core_w32

                core_w32.bring_to_front(ph)
            except Exception:
                try:
                    import ctypes

                    ctypes.windll.user32.SetForegroundWindow(int(ph))
                except Exception:
                    pass
        try:
            self.hide()
        except Exception:
            pass
        try:
            self.deleteLater()
        except Exception:
            pass
        try:
            from PySide6.QtWidgets import QApplication

            QApplication.processEvents()
        except Exception:
            pass
        super().closeEvent(event)

    def _on_finished(self, _result: int) -> None:
        """ダイアログ閉鎖時にリストから自身を外す。"""
        try:
            _open_report_dialogs.remove(self)
        except ValueError:
            pass

    def _addr_for_row(self, row_idx: int) -> str:
        """指定行インデックスに対応する Excel セルアドレス（A1 形式）を返す。"""
        if 0 <= row_idx < len(self._addresses):
            return self._addresses[row_idx]
        row = self._req.get("rows") or []
        if row_idx < len(row) and self._link_col < len(row[row_idx]):
            return str(row[row_idx][self._link_col])
        return ""

    def _goto_row(self, row_idx: int) -> None:
        """指定行のセルアドレスに Excel で選択を移動する。"""
        addr = self._addr_for_row(row_idx)
        _dupli_ui_diag(
            "[UI_DUPLI_GOTO_ROW] parent_hwnd=%s sheet_id=%s row_idx=%s addr=%r highlight_cleared=%s",
            self._parent_hwnd,
            self._sheet_id,
            row_idx,
            addr,
            self._highlight_cleared,
        )
        if addr:
            self._strip_goto_highlight_fill()
            quad = _goto_excel_cell(
                self._parent_hwnd,
                addr,
                book_name=self._book_name,
                sheet_name=self._sheet_name,
            )
            if quad:
                self._goto_hl_quad = [
                    int(quad[0]),
                    int(quad[1]),
                    int(quad[2]),
                    int(quad[3]),
                ]
                self._apply_goto_gray_com(self._goto_hl_quad)
            try:
                QTimer.singleShot(0, self._viewport_highlight_tick)
            except Exception:
                pass
            # _goto_excel_cell が Excel を SetForegroundWindow するため、TOPMOST でもレポートが背後に回る環境がある。JSON で TOPMOST 優先のため追従が付かないので ensure_front で戻す。
            ph = int(self._parent_hwnd or 0)
            if ph:

                def _nudge_report_after_goto() -> None:
                    try:
                        from ui_qt.ui_common import ensure_front

                        ensure_front(self, ph)
                    except Exception:
                        pass

                try:
                    from core.ui_window_timing import get_ui_window_timings

                    for _ms in get_ui_window_timings().dupli_report_after_cell_goto_ensure_front_delays_ms:
                        QTimer.singleShot(int(_ms), _nudge_report_after_goto)
                except Exception:
                    _nudge_report_after_goto()
            try:
                for _ms in (80, 350):
                    QTimer.singleShot(int(_ms), self._pulse_excel_for_sheet_edit)
            except Exception:
                pass

    def _on_cell_double(self, row: int, _col: int) -> None:
        """テーブルセルをダブルクリックしたときに、その行のセルへ移動する。"""
        self._goto_row(row)

    def _do_goto(self, tbl: QTableWidget) -> None:
        """「セルへ移動」ボタン押下時: 現在選択中の行のセルへ移動する。"""
        r = tbl.currentRow()
        if r >= 0:
            self._goto_row(r)

    def get_result(self) -> dict[str, Any]:
        """閉じた際の結果。常に OK。"""
        return {"status": "OK", "rc": 0}


def create_dialog(
    req_dict: dict[str, Any] | None,
    parent_hwnd: int,
    sheet_id: str,
) -> _DupliProgressWrapper | DupliReportDialog:
    """
    【概要】
        ui_server から呼ばれ、action に応じて進捗・完了通知・レポートのいずれかのダイアログを生成する。
    【補足】
        設定は config/ui_dupli.json。progress / dupli_done / dupli_report の各 action を処理する。
    """
    req = req_dict or {}
    action = str(req.get("action", "") or "").strip().lower()

    if action == "progress":
        # 進捗ダイアログは ui_common の共通部品を使用し、PROGRESS 設定をマージ
        from ui_qt.ui_common import _deep_merge, create_progress_dialog

        cfg = _get_cfg()
        main = (cfg or {}).get("MAIN") or {}
        progress = ((cfg or {}).get("SCREENS") or {}).get("PROGRESS") or {}
        progress_cfg = _deep_merge(main, progress)
        dlg = create_progress_dialog(
            req, int(parent_hwnd or 0), parent_widget=None, progress_cfg=progress_cfg
        )
        try:
            from ui_qt.ui_common import apply_tooltip_if_set

            apply_tooltip_if_set(dlg, progress_cfg, "TOOLTIP")
        except Exception:
            pass
        return _DupliProgressWrapper(dlg)

    if action == "dupli_report":
        # 重複一覧をモードレスで表示
        cfg = _get_cfg()
        dlg = DupliReportDialog(req, int(parent_hwnd or 0), cfg, sheet_id=str(sheet_id or ""))
        ph = int(parent_hwnd or 0)
        try:
            from ui_qt.ipc_file import write_waitform_ready_signal

            write_waitform_ready_signal(ph)
        except Exception:
            pass
        try:
            from ui_qt.ui_common import excel_rect_tuple_from_req, prepare_dialog_excel_center_before_show

            _pw = dict(getattr(dlg, "_hc_prepare_window_cfg", None) or {})
            # sizeHint（広いラベル）に頼らず、列＋行ヘッダで決めた幅を維持する
            _tw = int(getattr(dlg, "_mode_a_target_width", 0) or 0)
            if _tw > 0:
                _pw["DEFAULT_WIDTH"] = _tw
            else:
                _pw["DEFAULT_WIDTH"] = 0
            prepare_dialog_excel_center_before_show(
                dlg, ph, excel_rect_tuple_from_req(req), _pw
            )
        except Exception:
            pass
        return dlg

    if action == "dupli_done":
        # 完了通知をモーダルで表示（重複あり/なしの結果メッセージ）
        cfg = _get_cfg()
        done_cfg = (cfg.get("SCREENS") or {}).get("DONE") or {}
        return _DupliDoneDialog(req, int(parent_hwnd or 0), str(sheet_id or ""), done_cfg)

    raise ValueError(f"ui_dupli: unknown action {action!r}")


class _DupliDoneDialog(QDialog):
    """
    重複チェック完了時の通知をモーダルで表示するダイアログ。
    SCREENS.DONE の TITLE / ICON / ICON_SIZE / BTN_OK / WINDOW に従い、
    アイコン・メッセージ・OK ボタン・Excel 中央表示を適用する。
    """

    def __init__(
        self,
        req: dict[str, Any],
        parent_hwnd: int,
        sheet_id: str,
        done_cfg: dict[str, Any],
    ) -> None:
        super().__init__()
        try:
            self.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose, True)
        except Exception:
            pass
        self._req = req or {}
        self._parent_hwnd = int(parent_hwnd or 0)
        self._done_cfg = done_cfg or {}
        title = str(self._req.get("title") or self._done_cfg.get("TITLE") or "重複チェック").strip()
        self.setWindowTitle(title)
        message = str(self._req.get("message") or "").strip()

        from ui_qt.ui_common import (
            _icon_size_pixels_from_config,
            _normalize_message_newlines,
            _warning_icon_pixmap,
            apply_tooltip_if_set,
            apply_window_config,
        )

        lay = QVBoxLayout(self)
        # JSON で ICON が指定されていれば標準アイコンを表示
        icon_key = str(self._done_cfg.get("ICON") or "").strip()
        if icon_key:
            try:
                sz = _icon_size_pixels_from_config(self._done_cfg.get("ICON_SIZE"), default_pixels=24)
                px = _warning_icon_pixmap(self.style(), icon_key, sz)
                if px is not None:
                    icon_lbl = QLabel(self)
                    icon_lbl.setPixmap(px)
                    icon_lbl.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
                    icon_tip = str(self._done_cfg.get("ICON_TOOLTIP") or "").strip()
                    if icon_tip:
                        icon_lbl.setToolTip(icon_tip)
                    lay.addWidget(icon_lbl)
            except Exception:
                pass
        msg_lbl = QLabel(_normalize_message_newlines(message) if message else "完了しました。")
        msg_lbl.setWordWrap(True)
        msg_lbl.setMinimumWidth(280)
        msg_lbl.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        try:
            msg_lbl.setTextFormat(Qt.TextFormat.PlainText)
        except Exception:
            pass
        msg_tip = str(self._done_cfg.get("MSG_TOOLTIP") or "").strip()
        if msg_tip:
            msg_lbl.setToolTip(msg_tip)
        lay.addWidget(msg_lbl)
        lay.addStretch(1)
        row_btn = QHBoxLayout()
        row_btn.addStretch(1)
        btn_label = str(self._done_cfg.get("BTN_OK") or "OK").strip()
        btn_ok = QPushButton(btn_label or "OK")
        btn_tip = str(self._done_cfg.get("BTN_OK_TOOLTIP") or "").strip()
        if btn_tip:
            btn_ok.setToolTip(btn_tip)
        btn_ok.clicked.connect(self._on_ok)
        row_btn.addWidget(btn_ok)
        lay.addLayout(row_btn)

        try:
            apply_window_config(self, self._done_cfg, self._parent_hwnd, "DONE")
        except Exception:
            pass
        apply_tooltip_if_set(self, self._done_cfg, "TOOLTIP")
        # サイズは WINDOW の DEFAULT_WIDTH / DEFAULT_HEIGHT を優先
        win_cfg = self._done_cfg.get("WINDOW") or {}
        w = int(win_cfg.get("DEFAULT_WIDTH") or 0)
        h = int(win_cfg.get("DEFAULT_HEIGHT") or 0)
        if w > 0 and h > 0:
            self.resize(w, h)
        else:
            self.adjustSize()

    def _on_ok(self) -> None:
        """OK 押下: 先に非表示・Excel 有効化し、イベントを回してから次ティックで accept（同期 accept だと枠だけ残ることがある）。"""
        try:
            self.hide()
        except Exception:
            pass
        try:
            self.setVisible(False)
        except Exception:
            pass
        if self._parent_hwnd:
            try:
                from ui_qt.ui_common import enable_excel_window

                enable_excel_window(self._parent_hwnd, True)
            except Exception:
                pass
        try:
            from PySide6.QtWidgets import QApplication

            app = QApplication.instance()
            if app is not None:
                for _ in range(3):
                    app.processEvents()
        except Exception:
            pass
        try:
            QTimer.singleShot(0, self.accept)
        except Exception:
            self.accept()

    def showEvent(self, event) -> None:
        """表示時に Excel 中央に配置し、Excel ウィンドウを無効化する。"""
        super().showEvent(event)
        try:
            from ui_qt.ui_notification_sound import play_notification_on_widget

            play_notification_on_widget(self)
        except Exception:
            pass
        if self._parent_hwnd:
            try:
                from ui_qt.ui_common import center_on_excel, enable_excel_window, excel_rect_tuple_from_req

                center_on_excel(self, self._parent_hwnd, excel_rect_tuple_from_req(self._req))
                enable_excel_window(self._parent_hwnd, False)
            except Exception:
                pass

    def closeEvent(self, event) -> None:
        """× や accept 後: Excel 有効化・前面化、hide／deleteLater、DeferredDelete 消化（レポート閉鎖と同系のゴースト枠対策）。"""
        try:
            event.accept()
        except Exception:
            pass
        ph = self._parent_hwnd
        if ph:
            try:
                from ui_qt.ui_common import enable_excel_window

                enable_excel_window(ph, True)
            except Exception:
                pass
            try:
                from core import core_w32

                core_w32.bring_to_front(ph)
            except Exception:
                try:
                    import ctypes

                    ctypes.windll.user32.SetForegroundWindow(int(ph))
                except Exception:
                    pass
        try:
            self.hide()
        except Exception:
            pass
        try:
            self.deleteLater()
        except Exception:
            pass
        try:
            from PySide6.QtWidgets import QApplication

            QApplication.processEvents()
        except Exception:
            pass
        super().closeEvent(event)

    def exec(self) -> int:
        """モーダル実行。戻り値は整数で返す。"""
        return int(super().exec())

    def get_result(self) -> dict[str, Any]:
        """閉じた際の結果。OK で閉じた場合は rc=1。"""
        return {"status": "OK", "rc": 1}
