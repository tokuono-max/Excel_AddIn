# -*- coding: utf-8 -*-
"""
Python: 3.10+
Module: svc/svc_hd_nr
Created: 2026-03-05
Updated: 2026-04-06
Version: 2.4.9
Purpose:
  行整形（ヘッダブロック横結合）。選択行をヘッダブロックとし、行を横列に結合してデータ領域を整形する。
  データ領域は一括読込→メモリ内で一括 reshape（numpy）→一括書込で高速化。チャンクループは行わない。  UI: ui_hd_nr + ui_common。JSON: config/ui_hd_nr.json。

History (latest 3):
  - 2.4.9 (2026-06-06) ハング緩和: clear/autofit/DONE を suspend with 内に統一（indent 修正）。
  - 2.4.8 (2026-06-06) ハング緩和: 書込〜DONE を ScreenUpdating 復帰前に完了。restore_on_exit=False + wait_after_progress_done。
  - 2.4.7 (2026-04-06) HC_LOG_PERF: [HD_NR_PERF] phase / cumulative_ms。診断: [HD_NR_TRACE]。
  - 2.4.6 (2026-03-10) 不足セル背景: 範囲単位の range.color = (r,g,b) のみに簡略化（セル単位フォールバック削除）。
  - 2.4.5 (2026-03-10) 不足セル: ジャンプと背景色を分離。ジャンプ失敗時も背景色を実行。
  - 2.4.4 (2026-03-09) 不足セル背景: xlwings 基本形 range.color = (r,g,b) に変更。
  - 2.4.3 (2026-03-09) 不足セル背景: 範囲単位で range.color = 整数 に統一。
  - 2.4.2 (2026-03-09) 不足セル背景: Excel Interior.Color は BGR のため BGR 整数で設定。セル単位で COM Interior を設定。
  - 2.4.0 (2026-03-09) ヘッダを選択先頭行に設置・データは次行から・以前の行は残す。枠固定は結合ヘッダ行。不足時は先頭不足セルにジャンプし不足セルのみ背景色。有効データ最終行を header_first_row+len(output_2d) に統一。JSON: SHORTAGE_CELL_BG_RGB。
  - 2.3.6 (2026-03-09) 不足行背景色を config/ui_hd_nr.json の DATA_SHORTAGE.SHORTAGE_ROW_BG_RGB に移行。
"""
from __future__ import annotations

import os
import sys
import time
from pathlib import Path
from typing import Any, List, Optional, Tuple

_path_here = os.path.abspath(os.path.dirname(__file__))
_path_root = os.path.dirname(_path_here)
if _path_root not in sys.path:
    sys.path.insert(0, _path_root)

from core.core_log import get_diag_logger, get_logger, get_perf_logger
from core.core_progress_wait import wait_after_progress_done
from ui_qt.ipc_file import get_ipc_root, get_request_dir, read_pickle, write_pickle
from svc.svc_host import ensure_ui_server

logger = get_logger(__name__)
_hd_nr_diag = get_diag_logger("hc_csv_tool.diag.hd_nr")
_perf_hd_nr = get_perf_logger("svc.svc_hd_nr.perf")
__version__ = "2.4.9"


def _elapsed_ms_hd_nr(since: float) -> int:
    return max(0, int((time.perf_counter() - since) * 1000))


def _trace_hd_nr(phase: str, t0: float, **kv: object) -> None:
    try:
        if kv:
            _hd_nr_diag.info(
                "[HD_NR_TRACE] phase=%s cumulative_ms=%d %s",
                phase,
                _elapsed_ms_hd_nr(t0),
                " ".join("%s=%s" % (k, v) for k, v in kv.items()),
            )
        else:
            _hd_nr_diag.info("[HD_NR_TRACE] phase=%s cumulative_ms=%d", phase, _elapsed_ms_hd_nr(t0))
    except Exception:
        pass


def _perf_nr(phase: str, t0: float, **kv: object) -> None:
    try:
        if kv:
            _perf_hd_nr.info(
                "[HD_NR_PERF] phase=%s cumulative_ms=%d %s",
                phase,
                _elapsed_ms_hd_nr(t0),
                " ".join("%s=%s" % (k, v) for k, v in kv.items()),
            )
        else:
            _perf_hd_nr.info("[HD_NR_PERF] phase=%s cumulative_ms=%d", phase, _elapsed_ms_hd_nr(t0))
    except Exception:
        pass

try:
    from core import core_xlc as xlc
    from core import core_stat
    from core import core_w32 as w32
except ImportError:
    xlc = None  # type: ignore[assignment]
    core_stat = None  # type: ignore[assignment]
    w32 = None  # type: ignore[assignment]

# 不足データのセル値（空欄とする）
_SHORTAGE_PLACEHOLDER: Any = None
try:
    from core import core_cst as _cst
    from core.core_cst import get_ui_config_from_file_required as _get_ui_config
except ImportError:
    _cst = None  # type: ignore[assignment]
    _get_ui_config = None  # type: ignore[assignment]


def _clear_range(rng: Any, *, clear_all: bool = False) -> None:
    """範囲をクリア。clear_all=True のときは Clear（内容+書式）で UsedRange を縮め、そうでなければ ClearContents。"""
    if clear_all:
        try:
            clear_fn = getattr(rng, "clear", None) or getattr(rng, "Clear", None)
            if callable(clear_fn):
                clear_fn()
                return
        except Exception:
            pass
    try:
        clear_fn = getattr(rng, "clear_contents", None) or getattr(rng, "ClearContents", None)
        if callable(clear_fn):
            clear_fn()
        else:
            rng.value = None
    except Exception:
        try:
            rng.value = None
        except Exception:
            pass


def _autofit_output_range(ptr_s: Any, last_row: int, max_col: int, sheet_name: str, start_row: int = 1) -> None:
    """整形後の出力範囲で列のオートフィットを行う。行数が core_cst.AUTOFIT_MAX_ROWS を超えるときのみスキップ。start_row 以降のみ対象。"""
    if last_row <= 0 or max_col <= 0 or start_row > last_row:
        return
    max_rows = int(getattr(_cst, "AUTOFIT_MAX_ROWS", 100000) or 100000)
    if last_row - start_row + 1 > max_rows:
        return
    try:
        rng = ptr_s.range((start_row, 1), (last_row, max_col))
        cols = getattr(rng, "columns", None)
        if cols is not None:
            af = getattr(cols, "autofit", None) or getattr(cols, "AutoFit", None)
            if callable(af):
                af()
    except Exception:
        pass


def _freeze_first_row(ptr_s: Any, freeze_after_row: int = 1) -> None:
    """指定行の直後でウィンドウ枠を固定する。freeze_after_row=1 で1行目固定、=5 で5行目まで固定。xlwings の freeze_panes を優先、不可時は COM。"""
    if freeze_after_row < 1:
        return
    try:
        fp = getattr(ptr_s, "freeze_panes", None)
        if fp is not None:
            unfreeze = getattr(fp, "unfreeze", None)
            freeze_at = getattr(fp, "freeze_at", None)
            if callable(unfreeze):
                unfreeze()
            if callable(freeze_at):
                freeze_at(f"{freeze_after_row + 1}:{freeze_after_row + 1}")
                return
    except Exception:
        pass
    try:
        api = getattr(ptr_s, "api", None)
        if api is None:
            return
        book = getattr(ptr_s, "book", None)
        xl_app = getattr(book, "app", None) if book else None
        if xl_app is None:
            return
        app_api = getattr(xl_app, "api", None)
        if app_api is None:
            return
        aw = getattr(app_api, "ActiveWindow", None)
        if aw is None:
            return
        aw.FreezePanes = False
        api.Range(f"A{freeze_after_row + 1}").Select()
        aw.FreezePanes = True
    except Exception:
        pass


def _get_sheet(book: Any, sheet_id: str) -> Any:
    """sheet_id またはアクティブシートを返す。"""
    if sheet_id and xlc:
        sh = xlc.find_sheet_by_guid(book, sheet_id)
        if sh is not None:
            return sh
    try:
        return book.sheets.active
    except Exception:
        return None


def _get_selected_row_numbers(book: Any) -> List[int]:
    """選択範囲から行番号の一覧を取得（重複除く・ソート済み）。"""
    out: List[int] = []
    try:
        sel = book.app.selection
        if sel is None:
            return []
        api = getattr(sel, "api", None)
        if api is not None:
            n_areas = getattr(api.Areas, "Count", 0) or 0
            for i in range(1, int(n_areas) + 1):
                area = api.Areas(i)
                r1 = int(getattr(area, "Row", 0))
                r_count = int(getattr(area.Rows, "Count", 0))
                for r in range(r1, r1 + r_count):
                    out.append(r)
        if not out:
            r1 = int(getattr(sel, "row", 0))
            r_count = int(getattr(sel.rows, "count", 0))
            for r in range(r1, r1 + r_count):
                out.append(r)
    except Exception:
        pass
    return sorted(set(out))


def _get_window_rect(hwnd: int) -> tuple[int, int, int, int] | None:
    """Excel HWND の GetWindowRect（left, top, right, bottom）。UI の excel_rect 用。"""
    if not int(hwnd or 0) or os.name != "nt":
        return None
    try:
        import ctypes
        from ctypes import wintypes

        r = wintypes.RECT()
        if ctypes.windll.user32.GetWindowRect(int(hwnd), ctypes.byref(r)):
            return (int(r.left), int(r.top), int(r.right), int(r.bottom))
    except Exception:
        pass
    return None


def _submit_ui_and_wait(
    parent_hwnd: int,
    sheet_id: str,
    action: str,
    req_dict: Optional[dict] = None,
    timeout_sec: float = 60.0,
) -> Optional[dict]:
    """UI リクエストを投入し、result_path の結果を待つ。"""
    req_dir = get_request_dir()
    res_dir = Path(get_ipc_root()) / "results"
    res_dir.mkdir(parents=True, exist_ok=True)
    ts_ms = int(time.time() * 1000)
    result_path = res_dir / f"res_hd_nr_{action}_{ts_ms}_{os.getpid()}.pkl"
    sheet_id_safe = (str(sheet_id or "").strip() or "_")
    rd: dict[str, Any] = dict(req_dict or {}, action=action)
    er_u = _get_window_rect(int(parent_hwnd or 0))
    if er_u is not None:
        rd["excel_rect"] = list(er_u)
    payload = {
        "parent_hwnd": int(parent_hwnd),
        "result_path": str(result_path),
        "ready_path": "",
        "sheet_id": sheet_id_safe,
        "action": action,
        "module": "ui_qt.ui_hd_nr",
        "req_dict": rd,
    }
    req_path = req_dir / f"req_hd_nr_{action}_{ts_ms}_{os.getpid()}.pkl"
    try:
        write_pickle(req_path, payload)
    except Exception as exc:
        logger.warning("[HD_NR] UI 送信失敗: %s", exc)
        return None
    t0 = time.time()
    while (time.time() - t0) < timeout_sec:
        if result_path.exists() and result_path.stat().st_size > 0:
            try:
                return read_pickle(result_path)
            except Exception:
                pass
        time.sleep(0.05)
    return None


def _submit_progress_ui(
    parent_hwnd: int,
    sheet_id: str,
    progress_path: Path,
    phase_total: int,
) -> None:
    """進捗画面表示を UI サーバへ依頼（モデルレス）。"""
    req_dir = get_request_dir()
    res_dir = Path(get_ipc_root()) / "results"
    res_dir.mkdir(parents=True, exist_ok=True)
    ts_ms = int(time.time() * 1000)
    result_path = str(res_dir / f"res_hd_nr_progress_{ts_ms}_{os.getpid()}.pkl")
    sheet_id_safe = (str(sheet_id or "").strip() or "_")
    pr: dict[str, Any] = {
        "action": "progress",
        "progress_path": str(progress_path),
        "phase_total": phase_total,
        "excel_lock": True,
        "no_native_window": True,
    }
    er_p = _get_window_rect(int(parent_hwnd or 0))
    if er_p is not None:
        pr["excel_rect"] = list(er_p)
    payload = {
        "parent_hwnd": int(parent_hwnd),
        "result_path": result_path,
        "ready_path": "",
        "sheet_id": sheet_id_safe,
        "action": "progress",
        "module": "ui_qt.ui_hd_nr",
        "req_dict": pr,
    }
    req_path = req_dir / f"req_hd_nr_progress_{ts_ms}_{os.getpid()}.pkl"
    try:
        write_pickle(req_path, payload)
    except Exception as exc:
        logger.warning("[HD_NR] 進捗UI 送信失敗: %s", exc)


def insert_header(
    book: Any,
    sheet_id: str = "",
    target_hwnd: Optional[int] = None,
    excel_hwnd: Optional[int] = None,
    **kwargs: Any,
) -> None:
    """
    行整形：選択行をヘッダブロックとして横列に結合し、データ領域を整形する。
    複数行未選択時はワーニング。ヘッダ確認で開始後に整形。完了時は完了通知、不足時はデータ不足通知。
    """
    t_flow = time.perf_counter()
    _perf_nr("enter", t_flow)
    _trace_hd_nr("enter", t_flow)

    hwnd = int(target_hwnd or excel_hwnd or 0)
    if book is None:
        logger.warning("[HD_NR] 対象ブックなし")
        _perf_nr("abort_no_book", t_flow)
        _trace_hd_nr("abort_no_book", t_flow)
        return
    try:
        if not hwnd:
            hwnd = int(getattr(book.app, "hwnd", 0))
    except Exception:
        pass
    ptr_s = _get_sheet(book, sheet_id)
    if ptr_s is None:
        logger.warning("[HD_NR] 対象シートなし sheet_id=%s", sheet_id)
        _perf_nr("abort_no_sheet", t_flow, sheet_id=sheet_id or "")
        _trace_hd_nr("abort_no_sheet", t_flow, sheet_id=sheet_id or "")
        if w32 and hwnd:
            try:
                w32.bring_to_front(hwnd)
            except Exception:
                pass
        return

    ensure_ui_server()
    sheet_name = str(getattr(ptr_s, "name", "") or "").strip() or "Sheet1"
    logger.info("[HD_NR] 開始 sheet_id=%s シート=%s", sheet_id or "", sheet_name)

    selected = _get_selected_row_numbers(book)
    if len(selected) < 2:
        logger.info("[HD_NR] 複数行未選択")
        _perf_nr("early_few_rows_selected", t_flow, n=len(selected))
        _trace_hd_nr("early_few_rows_selected", t_flow, n=len(selected))
        _submit_ui_and_wait(hwnd, sheet_id, "hd_nr_warning")
        if w32 and hwnd:
            try:
                w32.bring_to_front(hwnd)
            except Exception:
                pass
        return

    res = _submit_ui_and_wait(hwnd, sheet_id, "hd_nr_confirm", req_dict={"selected_rows": selected})
    if not res or not res.get("start"):
        logger.info("[HD_NR] ユーザーキャンセル")
        _perf_nr("user_cancel_confirm", t_flow)
        _trace_hd_nr("user_cancel_confirm", t_flow)
        if w32 and hwnd:
            try:
                w32.bring_to_front(hwnd)
            except Exception:
                pass
        return

    # 共通仕様: 破壊的処理の直前で Undo 用スナップショットを保存（元に戻すで復元可能にする）
    try:
        from svc.svc_undo import save_undo_snapshot
        save_undo_snapshot(book, sheet_id=sheet_id, target_hwnd=hwnd, excel_hwnd=hwnd)
    except Exception as e:
        logger.warning("[HD_NR] save_undo_snapshot failed (undo unavailable): %s", e)

    _perf_nr("after_user_confirm", t_flow, n_rows=len(selected))
    _trace_hd_nr("after_user_confirm", t_flow, n_rows=len(selected))

    try:
        ur = getattr(ptr_s, "used_range", None)
        if ur is None:
            logger.warning("[HD_NR] 使用範囲なし")
            _perf_nr("abort_no_used_range", t_flow)
            _trace_hd_nr("abort_no_used_range", t_flow)
            return
        nr = getattr(ur, "rows", None)
        nc = getattr(ur, "columns", None)
        if nr is None or nc is None:
            _perf_nr("abort_used_range_shape", t_flow)
            _trace_hd_nr("abort_used_range_shape", t_flow)
            return
        last_row = int(nr.count)
        ncols = int(nc.count)
    except Exception as e:
        logger.warning("[HD_NR] 使用範囲読込失敗: %s", e)
        _perf_nr("abort_used_range", t_flow)
        _trace_hd_nr("abort_used_range", t_flow)
        return

    header_rows = selected
    n_header = len(header_rows)
    header_max_col = 1
    for r in header_rows:
        try:
            row_val = ptr_s.range((r, 1), (r, ncols)).value
            if isinstance(row_val, (list, tuple)):
                for c in range(len(row_val) - 1, -1, -1):
                    if c + 1 > header_max_col and (row_val[c] is not None and str(row_val[c]).strip()):
                        header_max_col = c + 1
                        break
            elif row_val is not None and str(row_val).strip():
                header_max_col = max(header_max_col, 1)
        except Exception:
            pass

    combined_header: List[Any] = []
    for r in header_rows:
        try:
            row_val = ptr_s.range((r, 1), (r, header_max_col)).value
            if isinstance(row_val, (list, tuple)):
                combined_header.extend(list(row_val))
            else:
                combined_header.append(row_val)
        except Exception:
            combined_header.extend([None] * header_max_col)

    # 結合ヘッダは選択先頭行に設定。先頭行以前はそのまま残す。
    header_first_row = min(header_rows)
    try:
        ptr_s.range((header_first_row, 1), (header_first_row, len(combined_header))).value = [combined_header]
    except Exception as e:
        logger.exception("[HD_NR] 書込失敗 ヘッダ: %s", e)
        _perf_nr("abort_header_write", t_flow)
        _trace_hd_nr("abort_header_write", t_flow)
        return

    data_start_row = max(header_rows) + 1
    if data_start_row > last_row:
        logger.info("[HD_NR] データ領域なし 完了")
        _perf_nr("early_no_data_region", t_flow)
        _trace_hd_nr("early_no_data_region", t_flow)
        _submit_done_and_wait(hwnd, sheet_id, 1, 1)
        if w32 and hwnd:
            try:
                w32.bring_to_front(hwnd)
            except Exception:
                pass
        return

    total_data_rows = last_row - data_start_row + 1
    total_chunks = (total_data_rows + n_header - 1) // n_header
    progress_path = Path(get_ipc_root()) / "progress" / f"progress_hd_nr_{sheet_id or '_'}.pkl"
    progress_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        progress_path.unlink(missing_ok=True)
    except Exception:
        pass
    _PROGRESS_PHASES = 3  # シート読込 / 整形 / シート書込み
    _submit_progress_ui(hwnd, sheet_id, progress_path, _PROGRESS_PHASES)
    time.sleep(0.5)

    def _write_progress(done: int, total: int, phase: str) -> None:
        try:
            write_pickle(progress_path, {
                "status": "RUN",
                "phase_i": done,
                "phase_total": total,
                "phase": phase,
                "done": done,
                "total": total,
                "pct": int(100 * done / total) if total else 0,
            })
        except Exception:
            pass

    # 一括読込: データ領域を1回で取得
    data_2d: List[List[Any]] = []
    try:
        raw = ptr_s.range((data_start_row, 1), (last_row, header_max_col)).value
        if raw is None:
            data_2d = []
        elif not isinstance(raw, (list, tuple)):
            data_2d = [[raw]]
        elif len(raw) == 0:
            data_2d = []
        elif not isinstance(raw[0], (list, tuple)):
            data_2d = [list(raw)]
        else:
            data_2d = [list(row) if isinstance(row, (list, tuple)) else [row] for row in raw]
    except Exception as e:
        logger.warning("[HD_NR] 読込失敗: %s", e)
        _perf_nr("abort_data_read", t_flow)
        _trace_hd_nr("abort_data_read", t_flow)
        if core_stat:
            try:
                core_stat.set_status_info(ptr_s, f"ERROR: データ読込失敗 Detail: {e}")
            except Exception:
                pass
        return
    logger.info("[HD_NR] 読込 行数=%s 列数=%s", len(data_2d), header_max_col)
    _perf_nr("after_data_matrix_read", t_flow, rows=len(data_2d), cols=header_max_col)
    _trace_hd_nr("after_data_matrix_read", t_flow, rows=len(data_2d), cols=header_max_col)
    _write_progress(1, _PROGRESS_PHASES, "シート読込中")

    # メモリ内で全チャンクを整形し output_2d を構築。
    # 一括 reshape: データを一度 numpy に載せ、reshape 一発で (n_chunks, n_header*header_max_col) にし、一括で output_2d を得る。ループでチャンクごとに触らない。
    output_2d: List[List[Any]] = []
    shortage_output_rows: List[int] = []  # 不足を書いた出力行（1-based 行番号）
    _use_vectorized = False
    try:
        import numpy as np
        _use_vectorized = True
    except ImportError:
        pass

    if _use_vectorized and data_2d:
        try:
            import numpy as np
            # 行ごとに header_max_col に正規化した 2 次元配列を用意
            rows_norm: List[List[Any]] = []
            for r in data_2d:
                row = list(r) if isinstance(r, (list, tuple)) else [r]
                if len(row) < header_max_col:
                    row.extend([None] * (header_max_col - len(row)))
                rows_norm.append(row[:header_max_col])
            arr = np.array(rows_norm, dtype=object)
            total_data_rows = len(arr)
            need_pad = max(0, total_chunks * n_header - total_data_rows)
            if need_pad > 0:
                shortage_output_rows.append(header_first_row + total_chunks)
                pad = np.full((need_pad, header_max_col), None, dtype=object)
                arr = np.vstack([arr, pad])
            # 一括 reshape: (N, header_max_col) → (n_chunks, n_header * header_max_col)
            arr = arr.reshape(total_chunks, n_header * header_max_col)
            if need_pad > 0:
                last_row_fill_start = (n_header - need_pad) * header_max_col
                arr[total_chunks - 1, last_row_fill_start:] = None
            output_2d = arr.tolist()
            _write_progress(2, _PROGRESS_PHASES, "整形中")
        except Exception as e:
            logger.warning("[HD_NR] vectorized reshape failed, fallback to list loop: %s", e)
            output_2d = []
            shortage_output_rows = []
            _use_vectorized = False

    if not _use_vectorized or not output_2d:
        output_2d = []
        shortage_output_rows = []
        for chunk_i in range(total_chunks):
            start_i = chunk_i * n_header
            end_i = min(start_i + n_header, len(data_2d))
            chunk_rows = data_2d[start_i:end_i]
            need = len(chunk_rows)
            _write_progress(chunk_i + 1, total_chunks, "整形中")

            row_data: List[Any] = []
            for r in chunk_rows:
                row = list(r) if isinstance(r, (list, tuple)) else [r]
                if len(row) < header_max_col:
                    row.extend([None] * (header_max_col - len(row)))
                row_data.extend(row)
            if need < n_header:
                shortage_output_rows.append(header_first_row + 1 + len(output_2d))
                row_data.extend([None] * (n_header * header_max_col - len(row_data)))
            if len(row_data) < n_header * header_max_col:
                row_data.extend([None] * (n_header * header_max_col - len(row_data)))
            output_2d.append(row_data)
        _write_progress(2, _PROGRESS_PHASES, "整形中")

    _perf_nr("after_reshape", t_flow, out_rows=len(output_2d))
    _trace_hd_nr("after_reshape", t_flow, out_rows=len(output_2d))

    # 一括書込（ScreenUpdating 復帰は DONE 後に遅延）。データはヘッダ次の行から。
    data_start_write_row = header_first_row + 1
    out_rows = header_first_row + len(output_2d)
    try:
        with xlc.suspend_sheet_updates(ptr_s, restore_on_exit=False):
            if output_2d and xlc:
                try:
                    xlc.write_chunk(ptr_s, data_start_write_row, 1, output_2d)
                except Exception as e:
                    logger.exception("[HD_NR] 書込失敗: %s", e)
                    _perf_nr("abort_chunk_write", t_flow)
                    _trace_hd_nr("abort_chunk_write", t_flow)
                    if core_stat:
                        try:
                            core_stat.set_status_info(ptr_s, f"ERROR: 書込失敗 Detail: {e}")
                        except Exception:
                            pass
                    return

            # 整形後の不要行（旧データの残り）をクリア
            clear_from = data_start_write_row + len(output_2d)
            out_cols = n_header * header_max_col
            clear_cols = max(ncols, out_cols)
            if clear_from <= last_row:
                try:
                    rng_old = ptr_s.range((clear_from, 1), (last_row, clear_cols))
                    _clear_range(rng_old, clear_all=True)
                except Exception:
                    pass
            last_data_col = -1
            for row in output_2d:
                for c in range(len(row) - 1, -1, -1):
                    if row[c] is not None and str(row[c]).strip():
                        last_data_col = max(last_data_col, c)
                        break
            if last_data_col >= 0 and last_data_col + 1 < out_cols:
                try:
                    _clear_range(ptr_s.range((header_first_row, last_data_col + 2), (out_rows, out_cols)), clear_all=True)
                except Exception:
                    pass
            fit_cols = (last_data_col + 1) if last_data_col >= 0 else out_cols
            if xlc and out_rows > 0 and fit_cols > 0:
                try:
                    xlc.clear_used_range_overflow(ptr_s, out_rows, fit_cols)
                except Exception:
                    pass
            try:
                _autofit_output_range(ptr_s, out_rows, fit_cols or out_cols, sheet_name, start_row=header_first_row)
                _freeze_first_row(ptr_s, freeze_after_row=header_first_row)
            except Exception:
                pass
            _write_progress(3, _PROGRESS_PHASES, "シート書込み中")

            if shortage_output_rows:
                try:
                    write_pickle(progress_path, {
                        "status": "DONE",
                        "phase_i": _PROGRESS_PHASES,
                        "phase_total": _PROGRESS_PHASES,
                        "done": _PROGRESS_PHASES,
                        "total": _PROGRESS_PHASES,
                        "pct": 100,
                        "show_done_dialog": False,
                    })
                except Exception:
                    pass
                wait_after_progress_done(min_sec=0.5)
            else:
                try:
                    write_pickle(progress_path, {
                        "status": "DONE",
                        "phase_i": _PROGRESS_PHASES,
                        "phase_total": _PROGRESS_PHASES,
                        "done": _PROGRESS_PHASES,
                        "total": _PROGRESS_PHASES,
                        "pct": 100,
                        "show_done_dialog": True,
                        "done_items": [{"no": 1, "name": "行整形", "rows": total_chunks}],
                        "done_detail_text": f"シート名：{sheet_name}\\n整形ブロック数：{total_chunks}",
                    })
                except Exception:
                    pass
                wait_after_progress_done(min_sec=1.0)
    finally:
        xlc.restore_screen_updating(ptr_s)

    _perf_nr("after_sheet_write_and_fit", t_flow, out_rows=out_rows)
    _trace_hd_nr("after_sheet_write_and_fit", t_flow, out_rows=out_rows)

    if shortage_output_rows:
        # 進捗クローズ後に不足行表示
        # 不足先頭セルにジャンプし、不足セル全てに背景色を付与（xlwings 基本形: range.color = (r, g, b)）
        try:
            ptr_s.activate()
        except Exception:
            pass
        r, g, b = 255, 255, 224  # 既定: 薄い黄 (RGB)
        if _get_ui_config:
            try:
                cfg = _get_ui_config("hd_nr")
                data_shortage = (cfg.get("SCREENS") or {}).get("DATA_SHORTAGE") or {}
                if isinstance(data_shortage, dict):
                    rgb = data_shortage.get("SHORTAGE_CELL_BG_RGB") or data_shortage.get("SHORTAGE_ROW_BG_RGB")
                else:
                    rgb = None
                if isinstance(rgb, (list, tuple)) and len(rgb) >= 3:
                    r, g, b = int(rgb[0]) & 0xFF, int(rgb[1]) & 0xFF, int(rgb[2]) & 0xFF
            except Exception:
                pass
        rgb_tuple = (r, g, b)
        # 不足行ごとに不足セル範囲を求め、先頭不足セルを特定
        first_short_row = None
        first_short_col = None
        ranges_to_highlight: List[Tuple[int, int, int]] = []  # (row_1b, col_start_1b, col_end_1b)
        for row_1b in sorted(shortage_output_rows):
            i = row_1b - data_start_write_row
            if i < 0 or i >= len(output_2d):
                continue
            row_data = output_2d[i]
            last_filled = -1
            for c in range(len(row_data) - 1, -1, -1):
                if row_data[c] is not None and str(row_data[c]).strip():
                    last_filled = c
                    break
            col_start_1b = last_filled + 2
            col_end_1b = len(row_data)
            if col_start_1b <= col_end_1b:
                ranges_to_highlight.append((row_1b, col_start_1b, col_end_1b))
                if first_short_row is None or (row_1b < first_short_row) or (row_1b == first_short_row and (first_short_col is None or col_start_1b < first_short_col)):
                    first_short_row = row_1b
                    first_short_col = col_start_1b
        # ジャンプ: Select/ScrollIntoView は COM で「名前が不明です」になることがあるため別 try。失敗しても背景色は実施する。
        if first_short_row is not None and first_short_col is not None:
            try:
                rng_first = ptr_s.range((first_short_row, first_short_col), (first_short_row, first_short_col))
                api_first = getattr(rng_first, "api", None)
                if api_first is not None:
                    api_first.Select()
                    api_first.ScrollIntoView()
            except Exception:
                pass
        # 背景色: 範囲単位で range.color = (r, g, b)（xlwings 基本形）
        for row_1b, col_start_1b, col_end_1b in ranges_to_highlight:
            try:
                ptr_s.range((row_1b, col_start_1b), (row_1b, col_end_1b)).color = rgb_tuple
            except Exception as e_c:
                logger.warning("[HD_NR] 不足セル背景色 設定失敗 row=%s cols %s-%s: %s", row_1b, col_start_1b, col_end_1b, e_c)
        first_short = shortage_output_rows[0] if shortage_output_rows else 0
        # 不足発生行・不足行を「***行」形式で表示
        shortage_lines = "、".join(f"{r}行" for r in shortage_output_rows)
        msg = f"不足発生行: {first_short}行\n不足行: {shortage_lines}"
        _submit_ui_and_wait(hwnd, sheet_id, "hd_nr_data_shortage", {"msg": msg})

    if w32 and hwnd:
            try:
                w32.bring_to_front(hwnd)
            except Exception:
                pass
    logger.info("[HD_NR] 完了 シート=%s ブロック数=%s 不足行数=%s", sheet_name, total_chunks, len(shortage_output_rows))
    _perf_nr("flow_end", t_flow, chunks=total_chunks, shortage=len(shortage_output_rows))
    _trace_hd_nr("flow_end", t_flow, chunks=total_chunks, shortage=len(shortage_output_rows))


def _submit_done_and_wait(
    parent_hwnd: int,
    sheet_id: str,
    done_count: int,
    total_count: int,
) -> None:
    """完了通知のみ表示（進捗を経由しない場合用）。"""
    req_dir = get_request_dir()
    res_dir = Path(get_ipc_root()) / "results"
    res_dir.mkdir(parents=True, exist_ok=True)
    ts_ms = int(time.time() * 1000)
    result_path = str(res_dir / f"res_hd_nr_done_{ts_ms}_{os.getpid()}.pkl")
    dr: dict[str, Any] = {
        "action": "done",
        "items": [{"no": 1, "name": "行整形", "rows": done_count}],
    }
    er_dn = _get_window_rect(int(parent_hwnd or 0))
    if er_dn is not None:
        dr["excel_rect"] = list(er_dn)
    payload = {
        "parent_hwnd": int(parent_hwnd),
        "result_path": result_path,
        "ready_path": "",
        "sheet_id": str(sheet_id or ""),
        "action": "done",
        "module": "ui_qt.ui_hd_nr",
        "req_dict": dr,
    }
    req_path = req_dir / f"req_hd_nr_done_{ts_ms}_{os.getpid()}.pkl"
    try:
        write_pickle(req_path, payload)
        t0 = time.time()
        while (time.time() - t0) < 30.0:
            if Path(result_path).exists() and Path(result_path).stat().st_size > 0:
                return
            time.sleep(0.05)
    except Exception as exc:
        logger.warning("[HD_NR] 完了UI 送信失敗: %s", exc)


normalize_header = insert_header
