# -*- coding: utf-8 -*-
"""
Python: 3.12+
Module: svc/svc_csv_sv.py
Created: 2026-03-05
Updated: 2026-06-04
Version: 1.3.10
Purpose:
  CSV保存（Qt UIサーバ方式 / 2プロセス分離）。
  - UI表示は ui_qt/ui_csv_sv（ファイル「名前を付けて保存」ダイアログ）で行う。
  - svc は Excel 操作と業務処理に専念し、UIとは IPC(Pickle) で通信する。
  - 基準フォルダは last_folder で共有。進捗の分母は総行数で統一。進捗閉じ時に Excel 操作有効化。
  - 完了時は csv_ld と同様の完了通知（シート名・ファイル名・容量・行数）を表示。

History (latest 3):
  - 1.3.10 (2026-06-13) 進捗 UI 共通設定（poll/creep）と即時表示連携。保存先確定直後の砂時計 ON。
  - 1.3.9 (2026-06-04) 保存: 既定で画面上の表示文字列をチャンク Copy→クリップボードで読込。文字列行列を csv.writer で直接出力（日付正規化不要）。HC_CSV_SV_USE_VALUE_READ=1 で従来 .value 経路。
  - 1.3.8 (2026-06-04) 性能: 保存前チェックを先頭行サンプルのみに変更（全 UsedRange 読込廃止）。大容量は日付正規化スキップ。保存計測ログ分割。
  - 1.3.6 (2026-06-04) 保存終了時に EnableEvents=True を保証（restore_excel_host_after_operation）。シート切替イベント復帰。
  - 1.3.5 (2026-06-04) 保存処理中の砂時計 ON（保存先確定後〜完了）。長いシート読込中は tick で再武装。
  - 1.3.4 (2026-04-09) 進捗 IPC に excel_rect（Excel HWND の GetWindowRect）を付与。svc_csv_ld / svc_csv_sp / svc_csv_mg と同様の中央寄せ基準に統一。
  - 1.3.3 (2026-04-07) 保存ダイアログ後・警告 UI 後・処理完了後に bring_to_front で Excel 前面復帰。ui_csv_sv ネイティブダイアログ終了時も同対応。
  - 1.3.2 (2026-04-06) `dialog_wait_ms`（ui_ipc 確定〜result_ok/cancel）。`HC_EXCEL_HWND` を core.core_env 経由に統一。
  - 1.3.1 (2026-04-06) 運用・診断ログ（phase / ms / req 相関）。WaitForm: book・シートなしで notify_wait_form_ready。無データ警告に ready_path＋_watch_ready。ScreenUpdating を csv_ld と同じ抑止意味に修正。
  - 1.3.0 (2026-03-09) 完了通知を表示。表示内容は csv_ld と同様（シート名・ファイル名・容量・データ/総数行）。
  - 1.2.0 (2026-03-05) 基準フォルダ・進捗分母を総数に・excel_lock で閉じ時に操作有効。
  - 1.1.0 (2026-03-05) 同上。
"""
from __future__ import annotations

import csv
import ctypes
import os
import re
import threading
import time
from datetime import date, datetime
from pathlib import Path
from typing import Any

from core import core_env
from core.core_log import get_diag_logger, get_logger
from core.excel_display_read import read_range_display_text_matrix, use_display_text_for_csv_save
from core.excel_host_restore import restore_excel_host_after_operation
from core.excel_perf_mode import set_excel_performance_mode
from core.core_cursor import (
    notify_ui_ready,
    notify_wait_form_ready,
    progress_dialog_wait_cursor_on,
)
from core.csv_tool_progress_ui import enrich_progress_req_dict
from ui_qt.ipc_file import get_ipc_root, get_last_folder, get_request_dir, read_pickle, set_last_folder, write_pickle
from svc.svc_host import ensure_ui_server

__version__ = "1.3.9"

try:
    from core import core_xlc as xlc
except Exception:
    xlc = None  # type: ignore[assignment]

try:
    from core import core_stat
except Exception:
    core_stat = None  # type: ignore[assignment]

try:
    from core import core_w32 as _w32
except Exception:
    _w32 = None  # type: ignore[assignment]

logger = get_logger(__name__)
_sv_diag = get_diag_logger("hc_csv_tool.diag.csv_sv")

READ_CHUNK_SIZE: int = 10000
# 保存前の「データあり」判定で先頭行だけ COM 取得する最大列数（全 UsedRange 読込を避ける）
_SV_VALID_SAMPLE_MAX_COLS: int = 128
# この行数超は日付正規化をスキップ（旧版同等の速度）。HC_CSV_SV_FORCE_DATE_NORMALIZE=1 で強制ON
_SV_DATE_NORMALIZE_MAX_ROWS_DEFAULT: int = 50_000
_CSV_DATE_TEXT_RE = re.compile(
    r"^\d{4}[-/.]\d{1,2}[-/.]\d{1,2}(?:[ T]\d{1,2}:\d{2}(?::\d{2})?)?$"
)


def _elapsed_ms(since: float) -> int:
    return max(0, int((time.perf_counter() - since) * 1000))


def _bring_excel_to_front(hwnd: int) -> None:
    """ネイティブ UI や別プロセス表示のあと Excel が背面に回る事象の緩和（ベストエフォート）。"""
    if not int(hwnd or 0) or _w32 is None or os.name != "nt":
        return
    try:
        _w32.bring_to_front(int(hwnd))
    except Exception:
        pass


def _sv_trace(fmt: str, *args: object) -> None:
    try:
        _sv_diag.info(fmt, *args)
    except Exception:
        pass


def _progress_path(sheet_id: str) -> Path:
    root = Path(get_ipc_root())
    d = root / "progress"
    d.mkdir(parents=True, exist_ok=True)
    return d / f"progress_sv_{sheet_id}.pkl"


def _progress_write(path: Path, obj: dict[str, Any]) -> None:
    try:
        write_pickle(path, obj)
    except Exception:
        pass


_PROGRESS_SEQ: dict[str, int] = {}


def _progress_key(path: Path) -> str:
    try:
        return str(path.resolve())
    except Exception:
        return str(path)


def _progress_write_monotonic(path: Path, obj: dict[str, Any]) -> None:
    """UI の seq 順序保証用に単調増加 seq を付与して進捗を書く。"""
    key = _progress_key(path)
    n = _PROGRESS_SEQ.get(key, -1) + 1
    _PROGRESS_SEQ[key] = n
    merged = dict(obj)
    merged["seq"] = n
    _progress_write(path, merged)


def _calc_sv_read_pct(done_rows: int, total_rows: int) -> int:
    """Excel 読込フェーズ用バー目標（全体の 0–49%）。"""
    if total_rows <= 0:
        return 0
    return min(49, int(done_rows * 49 / total_rows))


def _get_window_rect(hwnd: int) -> tuple[int, int, int, int] | None:
    """Excel 等の HWND 矩形を取得（進捗の excel_rect 用）。svc_csv_ld と同様。"""
    if not int(hwnd or 0) or os.name != "nt":
        return None
    try:
        from ctypes import wintypes

        r = wintypes.RECT()
        if ctypes.windll.user32.GetWindowRect(int(hwnd), ctypes.byref(r)):
            return (int(r.left), int(r.top), int(r.right), int(r.bottom))
    except Exception:
        pass
    return None


def _submit_progress_ui(
    parent_hwnd: int, sheet_id: str, progress_path: Path, phase_total: int
) -> None:
    """ui_server へ進捗ウィンドウ表示を要求する（モデルレス）。送信時点の Excel 矩形を渡し中央配置の基準にする。"""
    try:
        ph = int(parent_hwnd or 0)
        excel_rect = _get_window_rect(ph)
        req_dir = get_request_dir()
        res_dir = Path(get_ipc_root()) / "result"
        res_dir.mkdir(parents=True, exist_ok=True)
        ts_ms = int(time.time() * 1000)
        result_path = str(res_dir / f"res_progress_sv_{ts_ms}_{os.getpid()}.pkl")
        req_inner: dict[str, Any] = enrich_progress_req_dict(
            {
                "action": "progress",
                "progress_path": str(progress_path),
                "phase_total": int(phase_total),
                "excel_lock": True,
            },
            done_delay_ms=400,
            no_native_window=True,
        )
        if excel_rect is not None:
            req_inner["excel_rect"] = list(excel_rect)
        payload = {
            "parent_hwnd": ph,
            "result_path": result_path,
            "ready_path": "",
            "sheet_id": str(sheet_id),
            "log_path": "",
            "action": "progress",
            "module": "ui_qt.ui_csv_sv",
            "req_dict": req_inner,
        }
        req_path = req_dir / f"req_{ts_ms}_{os.getpid()}.pkl"
        write_pickle(req_path, payload)
    except Exception:
        pass


def _submit_request_dict(req_dict: dict[str, Any]) -> Path:
    """ui_server への要求を request_dir に pickle で投げる。"""
    req_dir = get_request_dir()
    req_dir.mkdir(parents=True, exist_ok=True)
    ts_ms = int(time.time() * 1000)
    req_path = req_dir / f"req_{ts_ms}_{os.getpid()}_{threading.get_ident()}.pkl"
    write_pickle(req_path, req_dict)
    return req_path


def _cell_has_value(cell: Any) -> bool:
    """1セルに値が入っているか（None・空文字・空白のみは無効）。"""
    if cell is None:
        return False
    if isinstance(cell, str):
        return cell.strip() != ""
    return True


def _matrix_sample_has_value(val: Any) -> bool:
    """COM から得た 1 行分（またはスカラー）に有効な値があるか。"""
    if val is None:
        return False
    if not isinstance(val, (list, tuple)):
        return _cell_has_value(val)
    for row in val:
        if row is None:
            continue
        if not isinstance(row, (list, tuple)):
            if _cell_has_value(row):
                return True
            continue
        for cell in row:
            if _cell_has_value(cell):
                return True
    return False


def _csv_sv_use_display_text() -> bool:
    """use_display_text_for_csv_save のエイリアス（テスト互換）。"""
    return use_display_text_for_csv_save()


def _should_normalize_dates_for_save(row_count: int) -> bool:
    """大容量シートでは全セル日付正規化をスキップして保存時間を短縮する。"""
    if core_env.truthy(core_env.get("HC_CSV_SV_SKIP_DATE_NORMALIZE")):
        return False
    if core_env.truthy(core_env.get("HC_CSV_SV_FORCE_DATE_NORMALIZE")):
        return True
    try:
        max_rows = int(
            core_env.get("HC_CSV_SV_DATE_NORMALIZE_MAX_ROWS")
            or _SV_DATE_NORMALIZE_MAX_ROWS_DEFAULT
        )
    except (TypeError, ValueError):
        max_rows = _SV_DATE_NORMALIZE_MAX_ROWS_DEFAULT
    return row_count <= max(0, max_rows)


def _sheet_has_valid_data(ptr_s: Any) -> bool:
    """
    シートに有効なデータがあるか（保存ダイアログ表示前の軽量チェック）。
    UsedRange の行数・列数のみ取得し、先頭行の一部列だけ COM 読込（全範囲読込は行わない）。
    """
    if ptr_s is None:
        return False
    try:
        api_range = ptr_s.used_range
        if api_range is None:
            return False
        row_count = int(getattr(api_range.rows, "count", 0) or 0)
        col_count = int(getattr(api_range.columns, "count", 0) or 0)
        if row_count < 1 or col_count < 1:
            return False
        row_start = int(getattr(api_range, "row", 1) or 1)
        col_start = int(getattr(api_range, "column", 1) or 1)
        sample_cols = min(col_count, _SV_VALID_SAMPLE_MAX_COLS)
        sample = ptr_s.range(
            (row_start, col_start),
            (row_start, col_start + sample_cols - 1),
        )
        val = sample.options(ndim=2).value
        if _matrix_sample_has_value(val):
            return True
        if sample_cols < col_count:
            tail_start = col_start + col_count - sample_cols
            tail = ptr_s.range(
                (row_start, tail_start),
                (row_start, col_start + col_count - 1),
            )
            return _matrix_sample_has_value(tail.options(ndim=2).value)
        return False
    except Exception:
        return False


def _get_formatted_size(str_path: str) -> str:
    try:
        sz = os.path.getsize(str_path)
        if sz >= 1048576:
            return f"{sz / 1048576:.2f} MB"
        return f"{sz / 1024:.1f} KB"
    except Exception:
        return "不明"


def _normalize_csv_date_cell(val: Any) -> Any:
    """CSV 出力時の日付表記を YYYY/MM/DD に統一する。"""
    if val is None:
        return val
    if isinstance(val, datetime):
        return val.strftime("%Y/%m/%d")
    if isinstance(val, date):
        return val.strftime("%Y/%m/%d")
    if isinstance(val, str):
        raw = val.strip()
        if not raw:
            return val
        if raw.startswith("'"):
            raw = raw[1:]
        if not _CSV_DATE_TEXT_RE.match(raw):
            return val
        try:
            import pandas as pd  # type: ignore

            ts = pd.to_datetime(raw, errors="coerce")
            if pd.isna(ts):
                return val
            return ts.strftime("%Y/%m/%d")
        except Exception:
            return val
    return val


def _normalize_matrix_dates_for_csv(matrix_2d: list[list[Any]]) -> list[list[Any]]:
    if not matrix_2d:
        return matrix_2d
    out: list[list[Any]] = []
    for row in matrix_2d:
        if isinstance(row, list):
            out.append([_normalize_csv_date_cell(v) for v in row])
        elif isinstance(row, tuple):
            out.append([_normalize_csv_date_cell(v) for v in row])
        else:
            out.append([_normalize_csv_date_cell(row)])
    return out


def _read_matrix_safe(
    sheet_ptr: Any,
    progress_path: Path | None = None,
    total_rows: int = 0,
    sheet_id: str = "",
) -> list[list[Any]]:
    """シートをチャンクで読み、2次元リストで返す。progress_path 指定時は進捗を書き込む。"""
    api_range = sheet_ptr.used_range
    val_row_start = api_range.row
    val_col_start = api_range.column
    val_row_count = api_range.rows.count
    val_col_count = api_range.columns.count
    list_total: list[list[Any]] = []
    for i_offset in range(0, val_row_count, READ_CHUNK_SIZE):
        rows_to_read = min(READ_CHUNK_SIZE, val_row_count - i_offset)
        curr_range = sheet_ptr.range(
            (val_row_start + i_offset, val_col_start),
            (
                val_row_start + i_offset + rows_to_read - 1,
                val_col_start + val_col_count - 1,
            ),
        )
        list_chunk = curr_range.options(ndim=2).value
        list_total.extend(list_chunk)
        if progress_path is not None and total_rows > 0:
            done = len(list_total)
            _progress_write_monotonic(
                progress_path,
                {
                    "status": "RUN",
                    "phase_i": 1,
                    "phase": "データ読込中",
                    "done": done,
                    "total": total_rows,
                    "pct": _calc_sv_read_pct(done, total_rows),
                },
            )
    return list_total


def _read_matrix_display_text(
    sheet_ptr: Any,
    progress_path: Path | None = None,
    total_rows: int = 0,
    sheet_id: str = "",
) -> list[list[str]]:
    """シートをチャンクで読み、画面上の表示文字列の 2 次元リストで返す。"""
    api_range = sheet_ptr.used_range
    val_row_start = int(getattr(api_range, "row", 1) or 1)
    val_col_start = int(getattr(api_range, "column", 1) or 1)
    val_row_count = int(getattr(api_range.rows, "count", 0) or 0)
    val_col_count = int(getattr(api_range.columns, "count", 0) or 0)
    list_total: list[list[str]] = []
    for i_offset in range(0, val_row_count, READ_CHUNK_SIZE):
        rows_to_read = min(READ_CHUNK_SIZE, val_row_count - i_offset)
        list_chunk = read_range_display_text_matrix(
            sheet_ptr,
            row_start=val_row_start + i_offset,
            col_start=val_col_start,
            n_rows=rows_to_read,
            n_cols=val_col_count,
        )
        list_total.extend(list_chunk)
        if xlc is not None:
            try:
                xlc.yield_to_excel()
            except Exception:
                pass
        if progress_path is not None and total_rows > 0:
            done = len(list_total)
            _progress_write_monotonic(
                progress_path,
                {
                    "status": "RUN",
                    "phase_i": 1,
                    "phase": "データ読込中",
                    "done": done,
                    "total": total_rows,
                    "pct": _calc_sv_read_pct(done, total_rows),
                },
            )
    return list_total


def _write_string_matrix_to_csv(str_save_path: str, matrix: list[list[str]]) -> None:
    with open(str_save_path, "w", encoding="utf-8-sig", newline="") as f_out:
        writer = csv.writer(f_out, quoting=csv.QUOTE_MINIMAL)
        for row in matrix:
            writer.writerow(row)


def _watch_ready(ready_path: str, sheet_id: str, t_load0: float) -> None:
    """READY_UI 監視（非同期）。受信したら notify_ui_ready（砂時計・WaitForm 含む）。"""
    p = Path(ready_path)
    while True:
        if p.exists() and p.stat().st_size > 0:
            try:
                d = read_pickle(p)
            except Exception:
                time.sleep(0.05)
                continue
            st = str(d.get("status", "")).strip().upper()
            if st == "READY_UI":
                notify_ui_ready(cancel_reason="READY_UI")
                logger.info(
                    "[CSV_SV] phase=ready_ui elapsed_ms=%s sheet_id=%s",
                    _elapsed_ms(t_load0),
                    sheet_id or "",
                )
                _sv_trace(
                    "[CSV_SV_TRACE] phase=ready_ui sheet_id=%s elapsed_since_enter_ms=%s wall_perf_s=%.6f",
                    sheet_id or "",
                    _elapsed_ms(t_load0),
                    time.perf_counter(),
                )
            return
        time.sleep(0.05)


def _do_save_csv(
    book: Any,
    sheet_id: str,
    str_save_path: str,
    ptr_s: Any,
    progress_path: Path | None = None,
    parent_hwnd: int = 0,
    sheet_id_for_progress: str = "",
    t_load0: float = 0.0,
    progress_ui_already_shown: bool = False,
) -> None:
    """保存パス確定後の実行（Excel読込・CSV書込・通知）。進捗表示ありの場合は progress_path を渡す。分母は総行数で統一。"""
    t_do0 = time.perf_counter()
    if ptr_s is None:
        return
    try:
        _do_save_csv_body(
            book,
            sheet_id,
            str_save_path,
            ptr_s,
            progress_path=progress_path,
            parent_hwnd=parent_hwnd,
            sheet_id_for_progress=sheet_id_for_progress,
            t_load0=t_load0,
            t_do0=t_do0,
            progress_ui_already_shown=progress_ui_already_shown,
        )
    except Exception as ex_save:
        logger.error("[CSV_SV] 保存処理失敗: %s", ex_save, exc_info=True)
        if progress_path is not None:
            _progress_write_monotonic(
                progress_path,
                {
                    "status": "ERROR",
                    "phase": "CSV保存エラー",
                    "msg": str(ex_save),
                    "detail": str(ex_save),
                    "pct": 0,
                },
            )
        raise
    finally:
        restore_excel_host_after_operation(
            parent_hwnd,
            sheet_id or sheet_id_for_progress,
            getattr(book, "app", None),
        )


def _do_save_csv_body(
    book: Any,
    sheet_id: str,
    str_save_path: str,
    ptr_s: Any,
    *,
    progress_path: Path | None = None,
    parent_hwnd: int = 0,
    sheet_id_for_progress: str = "",
    t_load0: float = 0.0,
    t_do0: float = 0.0,
    progress_ui_already_shown: bool = False,
) -> None:
    logger.info(
        "[CSV_SV] phase=do_save_enter file=%s sheet_id=%s elapsed_since_enter_ms=%s",
        os.path.basename(str_save_path),
        sheet_id or "",
        _elapsed_ms(t_load0) if t_load0 else 0,
    )
    _sv_trace(
        "[CSV_SV_TRACE] phase=do_save_enter file=%s sheet_id=%s wall_perf_s=%.6f",
        os.path.basename(str_save_path),
        sheet_id or "",
        time.perf_counter(),
    )
    total_steps = 2
    api_range = ptr_s.used_range
    val_row_count = api_range.rows.count
    val_total = max(1, val_row_count)

    if progress_path is not None and not progress_ui_already_shown and parent_hwnd:
        _submit_progress_ui(parent_hwnd, sheet_id or sheet_id_for_progress or "_", progress_path, total_steps)

    if progress_path is not None:
        _progress_write_monotonic(
            progress_path,
            {"status": "RUN", "phase_i": 1, "phase": "データ読込中", "done": 0, "total": val_total, "pct": 0},
        )
    use_display_text = _csv_sv_use_display_text()
    set_excel_performance_mode(book.app, True, disable_events=False)
    t_read0 = time.perf_counter()
    try:
        if use_display_text:
            list_matrix_2d = _read_matrix_display_text(
                ptr_s,
                progress_path=progress_path,
                total_rows=val_row_count if progress_path else 0,
                sheet_id=sheet_id or sheet_id_for_progress,
            )
        else:
            list_matrix_2d = _read_matrix_safe(
                ptr_s,
                progress_path=progress_path,
                total_rows=val_row_count if progress_path else 0,
                sheet_id=sheet_id or sheet_id_for_progress,
            )
    finally:
        set_excel_performance_mode(book.app, False, disable_events=False)
    read_ms = _elapsed_ms(t_read0)
    logger.info(
        "[CSV_SV] phase=read_matrix_done read_ms=%s rows=%s display_text=%s since_enter_ms=%s",
        read_ms,
        len(list_matrix_2d) if list_matrix_2d else 0,
        use_display_text,
        _elapsed_ms(t_do0),
    )
    _sv_trace(
        "[CSV_SV_TRACE] phase=read_matrix_done read_ms=%s wall_perf_s=%.6f",
        read_ms,
        time.perf_counter(),
    )

    if not list_matrix_2d:
        logger.warning("[CSV_SV] データなし シート=%s", ptr_s.name)
        if progress_path is not None:
            _progress_write_monotonic(
                progress_path,
                {
                    "status": "DONE",
                    "phase_i": total_steps,
                    "phase": "完了",
                    "done": 0,
                    "total": val_total,
                    "pct": 100,
                },
            )
        _bring_excel_to_front(parent_hwnd)
        return

    if progress_path is not None:
        _progress_write_monotonic(
            progress_path,
            {
                "status": "RUN",
                "phase_i": 2,
                "phase": "ファイル保存中",
                "done": val_total,
                "total": val_total,
                "pct": 50,
            },
        )

    row_count_saved = len(list_matrix_2d)
    normalize_ms = 0
    dataframe_ms = 0
    do_normalize = False

    if progress_path is not None:
        _progress_write_monotonic(
            progress_path,
            {
                "status": "RUN",
                "phase_i": 2,
                "phase": "ファイル保存中",
                "done": val_total,
                "total": val_total,
                "pct": 60,
            },
        )

    if use_display_text:
        t_csv0 = time.perf_counter()
        _write_string_matrix_to_csv(str_save_path, list_matrix_2d)  # type: ignore[arg-type]
        to_csv_ms = _elapsed_ms(t_csv0)
        write_ms = to_csv_ms
        logger.info(
            "[CSV_SV] phase=write_csv_done display_text=1 to_csv_ms=%s write_total_ms=%s "
            "do_save_ms=%s since_enter_ms=%s rows=%s",
            to_csv_ms,
            write_ms,
            _elapsed_ms(t_do0),
            _elapsed_ms(t_load0) if t_load0 else 0,
            row_count_saved,
        )
        _sv_trace(
            "[CSV_SV_TRACE] phase=write_csv_done display_text=1 to_csv_ms=%s wall_perf_s=%.6f",
            to_csv_ms,
            time.perf_counter(),
        )
    else:
        import pandas as pd

        do_normalize = _should_normalize_dates_for_save(row_count_saved)
        t_norm0 = time.perf_counter()
        if do_normalize:
            matrix_for_csv = _normalize_matrix_dates_for_csv(list_matrix_2d)
        else:
            matrix_for_csv = list_matrix_2d
            logger.info(
                "[CSV_SV] date_normalize skipped rows=%s (HC_CSV_SV_DATE_NORMALIZE_MAX_ROWS or SKIP)",
                row_count_saved,
            )
        normalize_ms = _elapsed_ms(t_norm0)

        if progress_path is not None:
            _progress_write_monotonic(
                progress_path,
                {
                    "status": "RUN",
                    "phase_i": 2,
                    "phase": "ファイル保存中",
                    "done": val_total,
                    "total": val_total,
                    "pct": 75,
                },
            )

        t_df0 = time.perf_counter()
        df = pd.DataFrame(matrix_for_csv)
        dataframe_ms = _elapsed_ms(t_df0)

        t_csv0 = time.perf_counter()
        df.to_csv(
            str_save_path,
            encoding="utf-8-sig",
            index=False,
            header=False,
            quoting=csv.QUOTE_MINIMAL,
        )
        to_csv_ms = _elapsed_ms(t_csv0)
        write_ms = normalize_ms + dataframe_ms + to_csv_ms
        logger.info(
            "[CSV_SV] phase=write_csv_done normalize_ms=%s dataframe_ms=%s to_csv_ms=%s "
            "write_total_ms=%s do_save_ms=%s since_enter_ms=%s rows=%s normalize=%s",
            normalize_ms,
            dataframe_ms,
            to_csv_ms,
            write_ms,
            _elapsed_ms(t_do0),
            _elapsed_ms(t_load0) if t_load0 else 0,
            row_count_saved,
            do_normalize,
        )
        _sv_trace(
            "[CSV_SV_TRACE] phase=write_csv_done normalize_ms=%s dataframe_ms=%s to_csv_ms=%s wall_perf_s=%.6f",
            normalize_ms,
            dataframe_ms,
            to_csv_ms,
            time.perf_counter(),
        )

    if progress_path is not None and use_display_text:
        _progress_write_monotonic(
            progress_path,
            {
                "status": "RUN",
                "phase_i": 2,
                "phase": "ファイル保存中",
                "done": val_total,
                "total": val_total,
                "pct": 75,
            },
        )

    str_fn_saved = os.path.basename(str_save_path)
    val_rows_saved = len(list_matrix_2d)
    str_size_saved = _get_formatted_size(str_save_path)
    ptr_s_name = getattr(ptr_s, "name", "") or "Sheet1"

    # 完了通知は csv_ld と同様の情報（シート名・ファイル名・容量・行数）を表示する
    data_rows = max(0, val_rows_saved - 1) if val_rows_saved else 0
    done_detail_text = (
        f"シート名：{ptr_s_name}\n"
        f"ファイル名：{str_fn_saved}\n"
        f"容量：{str_size_saved}\n"
        f"データ：{data_rows} 行\n"
        f"総数(ヘッダ含)：{val_rows_saved} 行"
    )

    try:
        book.app.api.StatusBar = f"CSV保存終了｜{str_fn_saved} を保存しました。"
    except Exception:
        pass
    if progress_path is not None:
        _progress_write_monotonic(
            progress_path,
            {
                "status": "DONE",
                "phase_i": total_steps,
                "phase": "完了",
                "done": val_total,
                "total": val_total,
                "pct": 100,
                "show_done_dialog": True,
                "done_items": [
                    {"no": 1, "name": str_fn_saved, "rows": val_rows_saved},
                ],
                "done_detail_text": done_detail_text,
            },
        )
    logger.info("[CSV_SV] 完了 ファイル=%s 行数=%s", str_fn_saved, val_rows_saved)


def _watch_result(
    result_path: str,
    book: Any,
    sheet_id: str,
    ptr_s: Any,
    parent_hwnd: int = 0,
    t_load0: float = 0.0,
    t_dialog_wait_start: float | None = None,
) -> None:
    """UI結果を待ち、OK かつ path があれば _do_save_csv を実行。

    t_dialog_wait_start:
        ui_ipc 確定直後の perf_counter。指定時は result_ok/cancel で dialog_wait_ms をログする
        （保存／警告ダイアログの操作待ち＋ポーリング、機械処理の前後比較用）。
    """
    p = Path(result_path)
    while True:
        if p.exists() and p.stat().st_size > 0:
            try:
                d = read_pickle(p)
            except Exception:
                time.sleep(0.05)
                continue
            status = str(d.get("status", "")).strip().upper()
            path = str(d.get("path", "")).strip()
            dialog_wait_ms = (
                _elapsed_ms(t_dialog_wait_start) if t_dialog_wait_start is not None else None
            )
            if status == "OK" and path:
                if dialog_wait_ms is not None:
                    logger.info(
                        "[CSV_SV] phase=result_ok file=%s elapsed_ms=%s dialog_wait_ms=%s sheet_id=%s",
                        os.path.basename(path),
                        _elapsed_ms(t_load0) if t_load0 else 0,
                        dialog_wait_ms,
                        sheet_id or "",
                    )
                    _sv_trace(
                        "[CSV_SV_TRACE] phase=result_ok path=%s elapsed_since_enter_ms=%s dialog_wait_ms=%s wall_perf_s=%.6f",
                        path,
                        _elapsed_ms(t_load0) if t_load0 else 0,
                        dialog_wait_ms,
                        time.perf_counter(),
                    )
                else:
                    logger.info(
                        "[CSV_SV] phase=result_ok file=%s elapsed_ms=%s sheet_id=%s",
                        os.path.basename(path),
                        _elapsed_ms(t_load0) if t_load0 else 0,
                        sheet_id or "",
                    )
                    _sv_trace(
                        "[CSV_SV_TRACE] phase=result_ok path=%s elapsed_since_enter_ms=%s wall_perf_s=%.6f",
                        path,
                        _elapsed_ms(t_load0) if t_load0 else 0,
                        time.perf_counter(),
                    )
                try:
                    set_last_folder(os.path.dirname(path))
                except Exception:
                    pass
                try:
                    progress_dialog_wait_cursor_on(str(sheet_id or "progress"))
                except Exception:
                    pass
                _bring_excel_to_front(parent_hwnd)
                progress_path = _progress_path(sheet_id or "_")
                _do_save_csv(
                    book,
                    sheet_id,
                    path,
                    ptr_s,
                    progress_path=progress_path,
                    parent_hwnd=parent_hwnd,
                    sheet_id_for_progress=sheet_id or "",
                    t_load0=t_load0,
                    progress_ui_already_shown=True,
                )
            else:
                if dialog_wait_ms is not None:
                    logger.info(
                        "[CSV_SV] phase=result_cancel status=%s elapsed_ms=%s dialog_wait_ms=%s sheet_id=%s",
                        status,
                        _elapsed_ms(t_load0) if t_load0 else 0,
                        dialog_wait_ms,
                        sheet_id or "",
                    )
                    _sv_trace(
                        "[CSV_SV_TRACE] phase=result_cancel status=%s elapsed_since_enter_ms=%s dialog_wait_ms=%s wall_perf_s=%.6f",
                        status,
                        _elapsed_ms(t_load0) if t_load0 else 0,
                        dialog_wait_ms,
                        time.perf_counter(),
                    )
                else:
                    logger.info(
                        "[CSV_SV] phase=result_cancel status=%s elapsed_ms=%s sheet_id=%s",
                        status,
                        _elapsed_ms(t_load0) if t_load0 else 0,
                        sheet_id or "",
                    )
                    _sv_trace(
                        "[CSV_SV_TRACE] phase=result_cancel status=%s elapsed_since_enter_ms=%s wall_perf_s=%.6f",
                        status,
                        _elapsed_ms(t_load0) if t_load0 else 0,
                        time.perf_counter(),
                    )
            _bring_excel_to_front(parent_hwnd)
            return
        time.sleep(0.05)


def save_csv(book: Any, sheet_id: str = "") -> None:
    """
    CSV保存のエントリ。Qt UIサーバで保存先選択 → 結果取得後に保存実行。
    """
    t_load0 = time.perf_counter()
    logger.info(
        "[CSV_SV] 開始 sheet_id=%s svc_pid=%s",
        sheet_id or "",
        os.getpid(),
    )
    _sv_trace(
        "[CSV_SV_TRACE] phase=enter sheet_id=%s svc_pid=%s wall_perf_s=%.6f",
        sheet_id or "",
        os.getpid(),
        time.perf_counter(),
    )

    if book is None:
        logger.error("[CSV_SV] book=None のため中断（WaitForm 解除を試行）")
        try:
            notify_wait_form_ready(book=book)
        except Exception:
            pass
        return

    ptr_s = None
    try:
        if sheet_id and xlc is not None:
            ptr_s = xlc.find_sheet_by_guid(book, sheet_id)
        if ptr_s is None:
            ptr_s = book.sheets.active
    except Exception as e:
        logger.warning("[CSV_SV] シート解決: %s", e)
        ptr_s = book.sheets.active if book else None

    if ptr_s is None:
        logger.error("[CSV_SV] 対象シートなし（WaitForm 解除を試行）")
        try:
            notify_wait_form_ready(book=book)
        except Exception:
            pass
        return

    ensure_ui_server()

    t_valid0 = time.perf_counter()
    has_valid = _sheet_has_valid_data(ptr_s)
    valid_ms = _elapsed_ms(t_valid0)
    logger.info(
        "[CSV_SV] phase=sheet_valid_check valid_check_ms=%s has_valid=%s (sample_row_only)",
        valid_ms,
        has_valid,
    )
    _sv_trace(
        "[CSV_SV_TRACE] phase=sheet_valid_check valid_check_ms=%s has_valid=%s wall_perf_s=%.6f",
        valid_ms,
        has_valid,
        time.perf_counter(),
    )

    if not has_valid:
        # 有効なデータがなければワーニングを表示し、OKでExcel操作有効のまま戻る
        ensure_ui_server()
        logger.info(
            "[CSV_SV] phase=after_ensure_ui_server branch=no_valid_data elapsed_ms=%s",
            _elapsed_ms(t_load0),
        )
        parent_hwnd = 0
        try:
            parent_hwnd = int(getattr(book.app, "hwnd", 0))
        except Exception:
            pass
        ipc_root = Path(get_ipc_root())
        warn_res_path = str(ipc_root / "results" / f"res_sv_warn_{sheet_id or '_'}_{int(time.time()*1000)}.pkl")
        warn_ready_path = str(ipc_root / "ready" / f"ready_sv_warn_{sheet_id or '_'}_{int(time.time()*1000)}.pkl")
        Path(warn_res_path).parent.mkdir(parents=True, exist_ok=True)
        Path(warn_ready_path).parent.mkdir(parents=True, exist_ok=True)
        try:
            Path(warn_res_path).unlink(missing_ok=True)
            Path(warn_ready_path).unlink(missing_ok=True)
        except Exception:
            pass
        req_warn = {
            "parent_hwnd": parent_hwnd,
            "result_path": warn_res_path,
            "ready_path": warn_ready_path,
            "sheet_id": str(sheet_id) if sheet_id else "_",
            "log_path": "",
            "action": "csv_sv_warning",
            "module": "ui_qt.ui_csv_sv",
            "req_dict": {
                "action": "csv_sv_warning",
                "title": "",
                "message": "",
            },
        }
        req_warn_path = _submit_request_dict(req_warn)
        logger.info(
            "[CSV_SV] ui_ipc ok kind=no_valid_warn req=%s sheet_id=%s hwnd=%s",
            req_warn_path.name,
            sheet_id or "",
            parent_hwnd,
        )
        _sv_trace(
            "[CSV_SV_TRACE] ui_ipc ok kind=no_valid_warn req=%s elapsed_ms=%s wall_perf_s=%.6f",
            req_warn_path.name,
            _elapsed_ms(t_load0),
            time.perf_counter(),
        )
        t_dialog_wait_warn = time.perf_counter()
        th_warn_ready = threading.Thread(
            target=_watch_ready,
            args=(warn_ready_path, str(sheet_id or ""), t_load0),
            name=f"ready_watch_sv_warn_{sheet_id}",
            daemon=True,
        )
        th_warn_ready.start()
        _watch_result(
            warn_res_path,
            None,
            str(sheet_id or ""),
            None,
            parent_hwnd,
            t_load0,
            t_dialog_wait_start=t_dialog_wait_warn,
        )
        logger.info(
            "[CSV_SV] phase=no_valid_warn_flow_done elapsed_ms=%s",
            _elapsed_ms(t_load0),
        )
        return

    logger.info(
        "[CSV_SV] phase=after_ensure_ui_server elapsed_ms=%s",
        _elapsed_ms(t_load0),
    )
    _sv_trace(
        "[CSV_SV_TRACE] phase=after_ensure_ui_server elapsed_ms=%s wall_perf_s=%.6f",
        _elapsed_ms(t_load0),
        time.perf_counter(),
    )

    parent_hwnd = 0
    try:
        parent_hwnd = int(getattr(book.app, "hwnd", 0))
        core_env.set_excel_hwnd_for_spawn(parent_hwnd)
    except Exception as e:
        logger.warning("[CSV_SV] Excel HWND: %s", e)
    if not parent_hwnd:
        logger.warning("[CSV_SV] parent_hwnd=0")

    ipc_root = Path(get_ipc_root())
    res_path = str(ipc_root / "results" / f"res_sv_{sheet_id or '_'}_{int(time.time()*1000)}.pkl")
    ready_path = str(ipc_root / "ready" / f"ready_sv_{sheet_id or '_'}_{int(time.time()*1000)}.pkl")
    Path(res_path).parent.mkdir(parents=True, exist_ok=True)
    Path(ready_path).parent.mkdir(parents=True, exist_ok=True)
    try:
        Path(res_path).unlink(missing_ok=True)
        Path(ready_path).unlink(missing_ok=True)
    except Exception:
        pass

    default_name = getattr(ptr_s, "name", None) or getattr(ptr_s, "Name", None) or "Sheet1"
    req_dict = {
        "parent_hwnd": parent_hwnd,
        "result_path": res_path,
        "ready_path": ready_path,
        "sheet_id": str(sheet_id) if sheet_id else "_",
        "log_path": "",
        "action": "csv_sv",
        "module": "ui_qt.ui_csv_sv",
        "req_dict": {
            "action": "csv_sv",
            "title": "名前を付けてCSVを保存",
            "default_name": default_name,
            "initial_dir": get_last_folder(),
        },
    }

    req_path = _submit_request_dict(req_dict)
    logger.info(
        "[CSV_SV] ui_ipc ok req=%s sheet_id=%s hwnd=%s",
        req_path.name,
        sheet_id or "",
        parent_hwnd,
    )
    _sv_trace(
        "[CSV_SV_TRACE] ui_ipc ok req=%s sheet_id=%s hwnd=%s elapsed_ms=%s wall_perf_s=%.6f",
        req_path.name,
        sheet_id or "",
        parent_hwnd,
        _elapsed_ms(t_load0),
        time.perf_counter(),
    )

    t_dialog_wait_save = time.perf_counter()
    th_ready = threading.Thread(
        target=_watch_ready,
        args=(ready_path, str(sheet_id or ""), t_load0),
        name=f"ready_watch_sv_{sheet_id}",
        daemon=True,
    )
    th_ready.start()

    _watch_result(
        res_path,
        book,
        str(sheet_id or ""),
        ptr_s,
        parent_hwnd,
        t_load0,
        t_dialog_wait_start=t_dialog_wait_save,
    )
    logger.info(
        "[CSV_SV] phase=save_csv_flow_done elapsed_ms=%s sheet_id=%s",
        _elapsed_ms(t_load0),
        sheet_id or "",
    )
    _sv_trace(
        "[CSV_SV_TRACE] phase=save_csv_flow_done elapsed_ms=%s wall_perf_s=%.6f",
        _elapsed_ms(t_load0),
        time.perf_counter(),
    )
