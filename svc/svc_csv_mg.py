# -*- coding: utf-8 -*-
"""
Python: 3.12+
Module: svc/svc_csv_mg.py
Created: 2026-02-11
Updated: 2026-06-22 (JST)
Version: 1.4.29
Purpose:
  CSV結合（Qt UIサーバ方式 / 2プロセス分離）。
  - UI表示は ui_qt/ui_server.py（Qt UIサーバ）で行う。
  - svc は Excel 操作と業務処理に専念し、UIとは IPC(Pickle) で通信する。

History (latest 3):
  - 1.4.29 (2026-06-22) 結合直前の Book/Sheet 解決を _attach_book+GUID に変更（apps.active 廃止）。
  - 1.4.28 (2026-06-13) 進捗表示: Excel書き込み→CSVファイル結合中。
  - 1.4.27 (2026-06-13) 進捗 UI 共通設定（poll/creep）とファイル確定直後の砂時計 ON。
  - 1.4.26 (2026-06-06) ハング緩和: 書込〜DONE を ScreenUpdating 復帰前に完了。restore_on_exit=False + wait_after_progress_done。
  - 1.4.24 (2026-06-04) 結合処理中の砂時計 ON（ファイル確定後〜完了）。Excel 書込みループで tick 再武装。
  - 1.4.23 (2026-05-05) progress の parent_hwnd を環境変数依存から引数へ統一。progress_closed_path ACK 待ちを追加し、進捗クローズ後に完了通知/再表示へ遷移。
  - 1.4.22 (2026-04-09) 結合メイン IPC と done_then_merge に excel_rect（Excel HWND の GetWindowRect）を付与。進捗と同じ送信時点矩形で中央寄せを統一。
  - 1.4.21 (2026-04-08) ジャンプ用定義名: ハイフン等を Excel 構文に合わせてサニタイズ（UTF-8 等を含むファイル名で Names.Add 失敗を防止）。失敗ログに name/ref を付与。
  - 1.4.20 (2026-04-05) 完了通知を DONE 進捗 pickle（show_done_dialog / done_items / seq）に集約し done_then_merge 二重経路を解消。_progress_write に単調 seq。progress の req_dict に excel_rect・done_delay_ms。ERROR/OVER_LIMIT も seq 付与。
  - 1.4.19 (2026-04-06) req_path 修正、運用・診断ログ（phase/ms/req）、READY_UI 経過ログ、merge_csv の WaitForm 救済、致命的エラー時 notify_wait_form_ready。結合処理に prep/write 区間ログ。
  - 1.4.18 (2026-03-09) オートフィットをセル数→行数基準に変更。定数は core_cst.AUTOFIT_MAX_ROWS に統一。
"""
from __future__ import annotations

from contextlib import nullcontext

__version__ = "1.4.29"
import os
import re
import threading
import unicodedata
import time

from pathlib import Path
from typing import Any
from core.core_log import get_diag_logger, get_logger
from core.core_cursor import (
    notify_ui_ready,
    notify_wait_form_ready,
    progress_dialog_wait_cursor_on,
)
from core.core_progress_wait import wait_after_progress_done
from core.csv_tool_progress_ui import enrich_progress_req_dict
from core import core_cst as cst
from svc.svc_csv_ld import (  # noqa: E402
    _capture_book_attach_keys,
    _resolve_book_and_sheet,
)

try:
    from core import core_xlc as xlc
except Exception:  # pragma: no cover
    xlc = None  # type: ignore

try:
    from core import core_stat
except Exception:  # pragma: no cover
    core_stat = None  # type: ignore

from ui_qt.ipc_file import get_ipc_root, get_request_dir, read_pickle, write_pickle
from svc.svc_host import ensure_ui_server

logger = get_logger(__name__)
_mg_diag = get_diag_logger("hc_csv_tool.diag.csv_mg")


def _elapsed_ms(since: float) -> int:
    return max(0, int((time.perf_counter() - since) * 1000))


def _mg_trace(fmt: str, *args: object) -> None:
    try:
        _mg_diag.info(fmt, *args)
    except Exception:
        pass


# ======================================================================================
# PHASE LOG (工程ログ) - 恒久的な計測基盤
#   - 対処療法ではなく「計測」。停止点/詰まり点をログだけで即時特定するための共通出力。
#   - merge_csv() の開始で _PHASE_T0_PERF をセットし、以降はどこから呼ばれても同じ基準で計測する。
# ======================================================================================
_PHASE_T0_PERF: float | None = None


def _phase(tag: str, **kv: Any) -> None:
    """工程ログ（P*-*）のフック。

    以前はここで [PHASE] ログを出力していたが、
    現在はログノイズ抑制のため出力を行わない。
    （呼び出し側との互換性のため関数自体は残す）
    """
    global _PHASE_T0_PERF
    if _PHASE_T0_PERF is None:
        return
    # 計測用の基準時刻だけ維持し、ログは出さない
    _ = tag, kv, _PHASE_T0_PERF


def _submit_request_dict(req_dict: dict[str, Any]) -> Path:
    """ui_server への要求を request_dir に pickle で投げる（最小依存）。"""
    req_dir = get_request_dir()
    req_dir.mkdir(parents=True, exist_ok=True)
    ts_ms = int(time.time() * 1000)
    req_path = req_dir / f"req_{ts_ms}_{os.getpid()}_{threading.get_ident()}.pkl"
    write_pickle(req_path, req_dict)
    return req_path


# Excel 最大行数（超過時は結合中止して警告表示。1048000 は運用上の上限）
EXCEL_MAX_ROWS = 1048000

# Progress (shared with UI)
def _progress_path(sheet_id: str) -> Path:
    root = get_ipc_root()
    d = root / "progress"
    d.mkdir(parents=True, exist_ok=True)
    return d / f"progress_{sheet_id}.pkl"


def _progress_write(path: Path, obj: dict[str, Any]) -> None:
    try:
        write_pickle(path, obj)
    except Exception:
        pass


# 進捗ファイルごとに単調増加する seq（UI 側の古い更新無視と整合）
_PROGRESS_SEQ: dict[str, int] = {}
_PROGRESS_CLOSE_ACK_TIMEOUT_SEC = 3.0
_PROGRESS_CLOSE_ACK_POLL_SEC = 0.03


def _progress_key(path: Path) -> str:
    try:
        return str(path.resolve())
    except Exception:
        return str(path)


def _progress_closed_ack_path(sheet_id: str) -> Path:
    d = Path(get_ipc_root()) / "progress"
    d.mkdir(parents=True, exist_ok=True)
    return d / f"progress_csv_mg_closed_{sheet_id}.pkl"


def _wait_progress_closed_ack(path: Path | None, timeout_sec: float = _PROGRESS_CLOSE_ACK_TIMEOUT_SEC) -> None:
    if path is None:
        return
    t0 = time.perf_counter()
    while True:
        try:
            if path.exists():
                return
        except Exception:
            return
        if (time.perf_counter() - t0) >= max(0.05, float(timeout_sec)):
            logger.info("[CSV_MG] progress close ack timeout: %s", str(path))
            return
        time.sleep(_PROGRESS_CLOSE_ACK_POLL_SEC)


def _progress_write_monotonic(path: Path, obj: dict[str, Any]) -> None:
    key = _progress_key(path)
    n = _PROGRESS_SEQ.get(key, -1) + 1
    _PROGRESS_SEQ[key] = n
    merged = dict(obj)
    merged["seq"] = n
    _progress_write(path, merged)


def _get_window_rect(hwnd: int) -> tuple[int, int, int, int] | None:
    """Excel 等の HWND 矩形を取得（UI の excel_rect 用）。csv_ld と同様の最小実装。"""
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


def _submit_progress_ui(
    parent_hwnd: int,
    sheet_id: str,
    progress_path: Path,
    phase_total: int,
    *,
    progress_closed_path: Path | None = None,
) -> None:
    """ui_server へ progress ウィンドウ表示を要求する（モデルレス、結果は捨てる）。"""
    try:
        ph = int(parent_hwnd or 0)
        excel_rect = _get_window_rect(ph)
        req_dir = get_request_dir()
        res_dir = Path(get_ipc_root()) / "result"
        try:
            res_dir.mkdir(parents=True, exist_ok=True)
        except Exception:
            pass
        ts_ms = int(time.time() * 1000)
        result_path = str(res_dir / f"res_progress_{ts_ms}_{os.getpid()}.pkl")
        req_inner: dict[str, Any] = enrich_progress_req_dict(
            {
                "action": "progress",
                "progress_path": str(progress_path),
                "phase_total": int(phase_total),
                "excel_lock": True,
            },
            done_delay_ms=1400,
            no_native_window=True,
        )
        if progress_closed_path is not None:
            cp = str(progress_closed_path).strip()
            if cp:
                req_inner["progress_closed_path"] = cp
        if excel_rect is not None:
            req_inner["excel_rect"] = list(excel_rect)
        payload = {
            "parent_hwnd": ph,
            "result_path": result_path,
            "ready_path": "",
            "sheet_id": str(sheet_id),
            "log_path": "",
            "action": "progress",
            "module": "ui_qt.ui_csv_mg",
            "req_dict": req_inner,
        }
        req_path = req_dir / f"req_{ts_ms}_{os.getpid()}.pkl"
        write_pickle(req_path, payload)
    except Exception:
        pass


def _submit_done_ui(
    parent_hwnd: int, sheet_id: str, done_items: list[dict[str, Any]]
) -> None:
    """完了通知を ui_csv_mg の done_then_merge で依頼する（フォールバック用）。

    CSV 結合の成功フローでは ProgressDialog 経由（DONE pickle の show_done_dialog）に統一した。
    progress_path が無い等の例外時のみこの経路を使う。

    done_items:
      [{"no":1, "name":"a.csv", "rows":10}, ...]
    """
    try:
        req_dir = get_request_dir()
        res_dir = Path(get_ipc_root()) / "results"
        try:
            res_dir.mkdir(parents=True, exist_ok=True)
        except Exception:
            pass
        ts_ms = int(time.time() * 1000)
        result_path = str(res_dir / f"res_done_then_merge_{ts_ms}_{os.getpid()}.pkl")
        req_dict = {
            "action": "done_then_merge",
            "items": done_items,
            "clear_table": True,
        }
        er_done = _get_window_rect(int(parent_hwnd or 0))
        if er_done is not None:
            req_dict["excel_rect"] = list(er_done)
        payload = {
            "parent_hwnd": int(parent_hwnd),
            "result_path": result_path,
            "ready_path": "",
            "sheet_id": str(sheet_id),
            "log_path": "",
            "action": "done_then_merge",
            "module": "ui_qt.ui_csv_mg",
            "req_dict": req_dict,
        }
        req_path = req_dir / f"req_{ts_ms}_{os.getpid()}.pkl"
        write_pickle(req_path, payload)
    except Exception:
        pass


# ======================================================================================
# UIサーバ起動（多重起動防止）
# ======================================================================================

_UI_SERVER_MUTEX_NAME = "hc_qt_ui_server_single_instance"
_UI_SERVER_REL_PATH = os.path.join("ui_qt", "ui_server.py")


def _resolve_ui_server_path() -> str:
    here = os.path.abspath(os.path.dirname(__file__))  # .../svc
    root = os.path.normpath(os.path.join(here, ".."))  # project root
    return os.path.join(root, _UI_SERVER_REL_PATH)


def ensure_qt_ui_server() -> None:
    """Qt UIサーバを起動（未起動なら起動）する。

    注意:
      - 起動/生存判定は svc.hc_host に一本化する（mutex/起動判定の不整合を防ぐ）。
      - 本モジュール側では多重起動防止や spawn を行わない。
    """
    _phase("P1-0", step="ensure_ui_server_begin")
    ensure_ui_server()
    _phase("P1-1", step="ensure_ui_server_done")


# ======================================================================================
# READY_UI / RESULT 監視
# ======================================================================================


def _watch_ready(ready_path: str, sheet_id: str, t_enter0: float) -> None:
    """
    READY_UI 監視（非同期）。
    READY_UI を受けたら、単発COM通知で砂時計OFF + 保険タイマ停止 + WaitForm 解除を要求する。
    """
    p = Path(ready_path)
    while True:
        # サーバが atomic write する前提だが、念のため size>0 で読む
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
                    "[CSV_MG] phase=ready_ui elapsed_ms=%s sheet_id=%s",
                    _elapsed_ms(t_enter0),
                    sheet_id or "",
                )
                _mg_trace(
                    "[CSV_MG_TRACE] phase=ready_ui sheet_id=%s elapsed_since_enter_ms=%s wall_perf_s=%.6f",
                    sheet_id or "",
                    _elapsed_ms(t_enter0),
                    time.perf_counter(),
                )
            return

        time.sleep(0.05)


def _shutdown_flag_exists() -> bool:
    """Excel終了/アドイン終了要求が出ているか（best-effort）"""
    try:
        return (Path(get_ipc_root()) / "control" / "shutdown.flag").exists()
    except Exception:
        return False


def _progress_write_merge_sheet_error(progress_path: Path, sheet_id: str) -> None:
    """COM 切れ等で結合先シートを解決できないとき。"""
    detail = (
        "対象シートに接続できませんでした。\n"
        "Excel を前面にしたうえで、もう一度 CSV 結合を実行してください。"
    )
    if sheet_id:
        detail += f"\n（GUID={sheet_id}）"
    _progress_write_monotonic(
        progress_path,
        {
            "status": "ERROR",
            "phase": "CSV結合を開始できません",
            "msg": detail,
            "detail": detail,
            "pct": 0,
        },
    )


def _resolve_merge_workbook_sheet(
    sheet_guid: str,
    parent_hwnd: int,
    attach_keys: tuple[int, str, str] | None,
    workbook_name: str,
) -> tuple[Any | None, Any | None]:
    """結合 UI 待ち後に HWND 固定で Workbook/Sheet を再解決する。"""
    sid = str(sheet_guid or "").strip()
    ak = attach_keys or (0, "", "")
    if sid:
        wb, sht = _resolve_book_and_sheet(
            None,
            sid,
            parent_hwnd,
            attach_keys=ak,
        )
        if wb is not None and sht is not None:
            logger.info(
                "[CSV_MG] phase=sheet_resolved sheet_id=%s hwnd=%s",
                sid,
                int(parent_hwnd or ak[0] or 0),
            )
        return wb, sht

    try:
        from svc.svc_server import _attach_book

        hwnd = int(parent_hwnd or ak[0] or 0)
        wb = _attach_book(
            excel_hwnd=hwnd,
            book_fullname=str(ak[1] or ""),
            book_name=str(ak[2] or workbook_name or ""),
        )
        sht = wb.sheets.active if wb is not None else None
        if wb is not None and sht is not None:
            logger.info("[CSV_MG] phase=sheet_resolved_active hwnd=%s", hwnd)
        return wb, sht
    except Exception as ex:
        logger.warning(
            "[CSV_MG] workbook attach failed hwnd=%s ex=%r",
            int(parent_hwnd or ak[0] or 0),
            ex,
        )
        return None, None


def _watch_result(
    result_path: str,
    workbook_name: str,
    sheet_id: str,
    t_enter0: float = 0.0,
    *,
    parent_hwnd: int = 0,
    attach_keys: tuple[int, str, str] | None = None,
) -> None:
    """
    UI結果(res)監視（非同期）。
    OK の場合のみ結合処理を実行する。
    """
    p = Path(result_path)
    while True:
        if _shutdown_flag_exists():
            return
        if p.exists() and p.stat().st_size > 0:
            try:
                d = read_pickle(p)
            except Exception:
                time.sleep(0.05)
                continue

            status = str(d.get("status", "")).strip().upper()
            files = list(d.get("files") or [])

            # If UI failed, log details into hc_csv.log so triage can be done without UI logs/pkl.
            if status != "OK":
                msg = str(d.get("message", "") or "").strip()
                tb = str(d.get("traceback", "") or "").strip()
                if msg:
                    logger.error("[UI_ERROR] %s", msg)
                if tb:
                    for ln in tb.splitlines():
                        logger.error("[UI_TRACE] %s", ln)

            if status == "OK" and files:
                logger.info(
                    "[CSV_MG] phase=result_ok file_count=%s elapsed_ms=%s sheet_id=%s",
                    len(files),
                    _elapsed_ms(t_enter0) if t_enter0 else 0,
                    sheet_id or "",
                )
                _mg_trace(
                    "[CSV_MG_TRACE] phase=result_ok file_count=%s elapsed_since_enter_ms=%s wall_perf_s=%.6f",
                    len(files),
                    _elapsed_ms(t_enter0) if t_enter0 else 0,
                    time.perf_counter(),
                )
                progress_path = (
                    Path(get_ipc_root())
                    / "progress"
                    / f"progress_{int(time.time()*1000)}_{os.getpid()}.pkl"
                )
                try:
                    progress_path.parent.mkdir(parents=True, exist_ok=True)
                except Exception:
                    pass
                progress_closed_path = _progress_closed_ack_path(sheet_id or "csv_mg")
                try:
                    progress_closed_path.unlink(missing_ok=True)
                except Exception:
                    pass
                try:
                    progress_dialog_wait_cursor_on(str(sheet_id or "progress"))
                except Exception:
                    pass
                _progress_write_monotonic(
                    progress_path,
                    {
                        "status": "RUN",
                        "phase_i": 1,
                        "phase": "開始",
                        "done": 0,
                        "total": 0,
                        "pct": 0,
                    },
                )
                try:
                    if int(parent_hwnd or 0):
                        _submit_progress_ui(
                            parent_hwnd=parent_hwnd,
                            sheet_id=sheet_id,
                            progress_path=progress_path,
                            phase_total=4,
                            progress_closed_path=progress_closed_path,
                        )
                except Exception:
                    pass
                mode = str(d.get("radio") or d.get("mode") or "mode_append").strip()
                _merge_files_to_sheet(
                    workbook_name=workbook_name,
                    sheet_guid=sheet_id,
                    paths=files,
                    progress_path=progress_path,
                    mode=mode,
                    parent_hwnd=parent_hwnd,
                    progress_closed_path=progress_closed_path,
                    attach_keys=attach_keys,
                )
            else:
                logger.info(
                    "[CSV_MG] phase=result_cancel status=%s elapsed_ms=%s sheet_id=%s",
                    status,
                    _elapsed_ms(t_enter0) if t_enter0 else 0,
                    sheet_id or "",
                )
                _mg_trace(
                    "[CSV_MG_TRACE] phase=result_cancel status=%s elapsed_since_enter_ms=%s wall_perf_s=%.6f",
                    status,
                    _elapsed_ms(t_enter0) if t_enter0 else 0,
                    time.perf_counter(),
                )
            return
        time.sleep(0.05)


# ======================================================================================
# CSV結合（旧版ロジックの最小復活：pandas + xlwings）
# ======================================================================================


def _read_csv_df(path: str):
    import pandas as pd

    # 1行目はヘッダ（結合対象から除外）
    for enc in ("utf-8-sig", "utf-8", "cp932"):
        try:
            return pd.read_csv(path, encoding=enc, dtype=str, header=0)
        except Exception:
            continue
    return None


def _read_csv_header_and_data(path: str) -> tuple[list[Any], list[list[Any]]]:
    """CSVを読み、1行目をヘッダ、2行目以降をデータとして返す。(header_row, data_rows)"""
    df = _read_csv_df(path)
    if df is None or df.empty:
        return [], []
    header_row = list(df.columns.astype(str))
    data_rows = df.fillna("").values.tolist()
    return header_row, data_rows


def _excel_defined_name_lead_ok(ch: str) -> bool:
    """Excel（日本語 UI）の定義名: 先頭は英字・かな・漢字・アンダースコア等（数字不可）。"""
    if not ch:
        return False
    if ch == "_":
        return True
    if ch.isdigit():
        return False
    try:
        return bool(unicodedata.category(ch).startswith("L"))
    except (TypeError, ValueError):
        return False


def _sanitize_excel_defined_name(raw: str) -> str:
    """定義名用に記号を落とし、先頭が規則に合わないときは先頭に _ を付ける。"""
    s = (raw or "").strip()
    s = s.replace(" ", "_").replace("-", "_")
    s = re.sub(r"[^\w\u3000-\u303f\u3040-\u309f\u30a0-\u30ff.]", "_", s, flags=re.UNICODE)
    s = re.sub(r"_+", "_", s).strip("_")
    if not s:
        s = "Range"
    if not _excel_defined_name_lead_ok(s[0]):
        s = f"_{s}"
    return s[:255]


def _merge_files_to_sheet(
    workbook_name: str,
    sheet_guid: str,
    paths: list[str],
    progress_path: Path | None = None,
    mode: str = "mode_append",
    *,
    parent_hwnd: int = 0,
    progress_closed_path: Path | None = None,
    attach_keys: tuple[int, str, str] | None = None,
) -> None:
    """指定されたCSV群を、指定シートの最終行へ追記する（pandas + xlwings）。

    - mode_append: 最初のファイルのみヘッダを追加し結合
    - mode_replace: 全ファイルでヘッダ付きデータで結合
    - mode_preview: 全ファイルのヘッダなしで結合
    - シートにデータがなければ1行目から、データがあれば最終行の次から追加
    - 進捗は progress_path(pkl) が指定されたときのみ出力する
    - スレッドから呼ばれるため、COM初期化/解除をこの関数内で行う
    """
    try:
        import pythoncom  # noqa: WPS433

        pythoncom.CoInitialize()
    except Exception:
        pythoncom = None  # type: ignore

    try:
        t_merge_exec0 = time.perf_counter()
        mode_key = (mode or "mode_append").strip().lower()
        if mode_key not in ("mode_append", "mode_replace", "mode_preview"):
            mode_key = "mode_append"

        tables: list[tuple[str, list[list[Any]], int, int]] = []
        done_items: list[dict[str, Any]] = []
        total_rows = 0
        seq = 0
        for file_idx, p in enumerate(paths):
            seq += 1
            p_str = os.path.abspath(str(p))
            name = os.path.basename(p_str)
            header_row, data_rows = _read_csv_header_and_data(p_str)

            rows = 0
            cols = 0
            values_to_write: list[list[Any]] = []

            if header_row or data_rows:
                cols = len(header_row) if header_row else (len(data_rows[0]) if data_rows else 0)
                if mode_key == "mode_append":
                    if file_idx == 0:
                        values_to_write = [header_row] + data_rows if header_row else data_rows
                    else:
                        values_to_write = data_rows
                elif mode_key == "mode_replace":
                    values_to_write = ([header_row] + data_rows) if header_row else data_rows
                else:
                    values_to_write = data_rows
                rows = len(values_to_write)
                if rows and not cols and values_to_write:
                    cols = len(values_to_write[0])

            done_items.append({"no": seq, "name": name, "rows": int(rows)})

            if rows > 0 and cols > 0:
                tables.append((name, values_to_write, rows, cols))
                total_rows += rows

        prep_ms = _elapsed_ms(t_merge_exec0)
        logger.info(
            "[CSV_MG] phase=merge_prep_done files=%s total_rows=%s prep_ms=%s",
            len(paths),
            total_rows,
            prep_ms,
        )
        _mg_trace(
            "[CSV_MG_TRACE] phase=merge_prep_done prep_ms=%s wall_perf_s=%.6f",
            prep_ms,
            time.perf_counter(),
        )

        if progress_path is not None:
            _progress_write_monotonic(
                progress_path,
                {
                    "status": "RUN",
                    "phase_i": 2,
                    "phase": "読込/準備",
                    "done": 0,
                    "total": total_rows,
                    "pct": 0,
                },
            )

        _phase("P3-0", step="resolve_workbook_sheet")
        wb, sht = _resolve_merge_workbook_sheet(
            sheet_guid,
            parent_hwnd,
            attach_keys,
            workbook_name,
        )
        if wb is None or sht is None:
            logger.error(
                "[CSV_MG] 結合先シート解決失敗 sheet_guid=%s hwnd=%s",
                sheet_guid or "",
                parent_hwnd,
            )
            if progress_path is not None:
                _progress_write_merge_sheet_error(progress_path, str(sheet_guid or ""))
            return

        # シートにデータがなければ1行目から、あれば最終データの次の行から
        _phase("P3-7", step="detect_last_row")
        last_cell = sht.range("A" + str(sht.cells.last_cell.row)).end("up")
        last_row = int(last_cell.row)
        try:
            a1_val = sht.range("A1").value
            sheet_empty = last_row <= 1 and (a1_val is None or str(a1_val).strip() == "")
        except Exception:
            sheet_empty = last_row <= 1
        start_row = 1 if sheet_empty else (last_row + 1)
        _phase("P3-8", step="start_row_ready", start_row=start_row, sheet_empty=sheet_empty)

        # 最大行数監視: シート既存行数＋結合総行数が上限を超える場合は警告して結合画面へ戻る（ケース4: 警告OKで結合画面再表示・テーブル維持）
        if last_row + total_rows > EXCEL_MAX_ROWS:
            if progress_path is not None:
                req_dir = get_request_dir()
                ipc_root = Path(get_ipc_root())
                res_dir = ipc_root / "results"
                ready_dir = ipc_root / "ready"
                try:
                    res_dir.mkdir(parents=True, exist_ok=True)
                    ready_dir.mkdir(parents=True, exist_ok=True)
                except Exception:
                    pass
                ts_ms = int(time.time() * 1000)
                return_res = str(res_dir / f"res_{sheet_guid}_{ts_ms}.pkl")
                return_ready = str(ready_dir / f"ready_{sheet_guid}_{ts_ms}.pkl")
                parent_hwnd2 = int(parent_hwnd or 0)
                return_merge_payload = {
                    "parent_hwnd": parent_hwnd2,
                    "sheet_id": str(sheet_guid),
                    "result_path": return_res,
                    "ready_path": return_ready,
                    "action": "csv_mg",
                    "module": "ui_qt.ui_csv_mg",
                    "req_dict": {
                        "action": "csv_mg",
                        "initial_files": list(paths),
                        "clear_table": False,
                    },
                }
                _progress_write_monotonic(
                    progress_path,
                    {
                        "status": "OVER_LIMIT",
                        "msg": (
                            f"Excelの最大行数（{EXCEL_MAX_ROWS:,}行）を超えるため結合できません。\n"
                            f"シートの現在の最終行: {last_row:,}行、今回の結合行数: {total_rows:,}行です。"
                        ),
                        "return_merge_payload": return_merge_payload,
                        "request_dir": str(req_dir),
                    },
                )
            logger.warning(
                "[CSV_MG] 最大行数超過 last_row=%s total_rows=%s max=%s",
                last_row, total_rows, EXCEL_MAX_ROWS,
            )
            return

        cur_row = start_row
        done_rows = 0
        max_col = 0
        sht_name = (getattr(sht, "name", None) or getattr(sht, "Name", None) or "Sheet1")

        # 一括書込みは表示停止で高速化（ScreenUpdating 復帰は DONE 後に遅延）
        t_excel_write0 = time.perf_counter()
        done_ser = [
            {
                "no": int(it.get("no", 0) or 0),
                "name": str(it.get("name", "") or ""),
                "rows": int(it.get("rows", 0) or 0),
            }
            for it in done_items
        ]
        with (xlc.suspend_sheet_updates(sht, restore_on_exit=False) if xlc else nullcontext()):
            for file_name, values, rows, cols in tables:
                block_start_row = cur_row
                max_col = max(max_col, int(cols)) if cols else max_col
                # chunked write (adaptive by column count)
                target_cells = int(
                    getattr(cst, "SVC_EXCEL_WRITE_TARGET_CELLS", 50000) or 50000
                )
                min_rows = int(getattr(cst, "SVC_EXCEL_WRITE_MIN_ROWS", 100) or 100)
                max_rows = int(getattr(cst, "SVC_EXCEL_WRITE_MAX_ROWS", 5000) or 5000)
                rows_per_chunk = max(
                    min_rows, min(max_rows, max(1, target_cells // max(int(cols), 1)))
                )

                if progress_path is not None:
                    pct = int(done_rows * 100 / max(total_rows, 1))
                    _progress_write_monotonic(
                        progress_path,
                        {
                            "status": "RUN",
                            "phase_i": 3,
                            "phase": "CSVファイル結合中",
                            "done": done_rows,
                            "total": total_rows,
                            "pct": pct,
                            "current_file": file_name,
                        },
                    )

                for i0 in range(0, len(values), rows_per_chunk):
                    if _shutdown_flag_exists():
                        if progress_path is not None:
                            _progress_write_monotonic(
                                progress_path,
                                {
                                    "status": "CANCEL",
                                    "phase_i": 3,
                                    "phase": "CSVファイル結合中",
                                    "done": done_rows,
                                    "total": total_rows,
                                },
                            )
                        return

                    chunk = values[i0 : i0 + rows_per_chunk]
                    sht.range((cur_row + i0, 1)).value = chunk

                    done_rows += len(chunk)
                    if progress_path is not None:
                        pct = int(done_rows * 100 / max(total_rows, 1))
                        _progress_write_monotonic(
                            progress_path,
                            {
                                "status": "RUN",
                                "phase_i": 3,
                                "phase": "CSVファイル結合中",
                                "done": done_rows,
                                "total": total_rows,
                                "pct": pct,
                                "current_file": file_name,
                            },
                        )

                # セルコメント: シート名-追加行番号 ファイル名 \n 追加行数
                # ジャンプ移動先: シート名 追加行番 ファイル名(拡張子除く)
                try:
                    first_cell = sht.range((block_start_row, 1))
                    comment_text = f"{sht_name}-{block_start_row} {file_name}\n{rows}行追加"
                    if hasattr(first_cell, "api") and first_cell.api is not None:
                        try:
                            existing = first_cell.api.Comment
                            if existing is not None:
                                existing.Delete()
                        except Exception:
                            pass
                        first_cell.api.AddComment(comment_text)
                except Exception:
                    pass
                try:
                    file_name_no_ext = os.path.splitext(file_name)[0] or file_name
                    name_raw = f"{sht_name} {block_start_row} {file_name_no_ext}"
                    name_safe = _sanitize_excel_defined_name(name_raw)
                    ref = f"='{sht_name}'!$A${block_start_row}"
                    existing_names = [n.name for n in wb.names]
                    base_name = name_safe
                    name_safe = base_name
                    idx = 0
                    while name_safe in existing_names:
                        idx += 1
                        name_safe = _sanitize_excel_defined_name(f"{base_name}_{idx}")[:255]
                    # ジャンプの移動先登録: COM API で確実に追加（xlwings names.add が効かない環境対策）
                    api = wb.api
                    api.Names.Add(Name=name_safe, RefersTo=ref)
                except Exception as e:
                    logger.warning(
                        "[CSV_MG] 名前追加失敗 row=%s name=%s ref=%s err=%s",
                        block_start_row,
                        locals().get("name_safe", ""),
                        locals().get("ref", ""),
                        e,
                    )

                cur_row += rows

            # 有効領域外をクリア（UsedRange の拡大防止）
            if xlc and max_col > 0 and cur_row > start_row:
                try:
                    xlc.clear_used_range_overflow(sht, cur_row - 1, max_col)
                except Exception:
                    pass

            # 出力完了後: 列幅オートフィット（行数が core_cst.AUTOFIT_MAX_ROWS 超過時はスキップ）
            if max_col > 0 and cur_row > start_row:
                last_row = cur_row - 1
                autofit_rows = last_row - start_row + 1
                max_af_rows = int(getattr(cst, "AUTOFIT_MAX_ROWS", 100000) or 100000)
                if progress_path is not None:
                    _progress_write_monotonic(
                        progress_path,
                        {
                            "status": "RUN",
                            "phase_i": 3,
                            "phase": "列幅調整中",
                            "done": total_rows,
                            "total": total_rows,
                            "pct": 99,
                        },
                    )
                if autofit_rows <= max_af_rows:
                    try:
                        rng = sht.range((start_row, 1), (last_row, max_col))
                        rng.columns.autofit()
                    except Exception:
                        pass

            # HC_STATUS_INFO とステータスバー
            if core_stat is not None:
                try:
                    sht_name_stat = getattr(sht, "name", None) or getattr(sht, "Name", None) or "Sheet1"
                    file_names_stat = [
                        (os.path.splitext(str(it.get("name", "")).strip())[0] or str(it.get("name", "")).strip())
                        for it in done_items
                        if it.get("name")
                    ]
                    if file_names_stat:
                        new_part = "＋".join(file_names_stat)
                        existing = (core_stat.get_status_info(sht) or "").strip()
                        if not existing:
                            value = f"{sht_name_stat}：{new_part}"
                        else:
                            value = existing + "＋" + new_part
                        core_stat.set_status_info(sht, value)
                        try:
                            app.api.StatusBar = value
                        except Exception:
                            pass
                except Exception:
                    pass

            if progress_path is not None:
                _phase("P4-9", step="done_progress_pickle", items=len(done_ser))
                _progress_write_monotonic(
                    progress_path,
                    {
                        "status": "DONE",
                        "phase_i": 4,
                        "phase": "完了",
                        "done": total_rows,
                        "total": total_rows,
                        "pct": 100,
                        "show_done_dialog": True,
                        "done_items": done_ser,
                    },
                )
                wait_after_progress_done(min_sec=1.0)

        excel_write_ms = _elapsed_ms(t_excel_write0)
        logger.info(
            "[CSV_MG] phase=merge_excel_write_done excel_write_ms=%s tables=%s",
            excel_write_ms,
            len(tables),
        )
        _mg_trace(
            "[CSV_MG_TRACE] phase=merge_excel_write_done excel_write_ms=%s wall_perf_s=%.6f",
            excel_write_ms,
            time.perf_counter(),
        )

        # 運用ログ: 完了（結合ファイル名を付与。長い場合は先頭数件＋他N件に省略）
        file_names = [str(it.get("name", "")).strip() for it in done_items if it.get("name")]
        if len(file_names) > 5:
            files_str = ", ".join(file_names[:3]) + f" 他{len(file_names) - 3}件"
        else:
            files_str = ", ".join(file_names) if file_names else ""
        logger.info(
            "[CSV_MG] 完了 追加行数=%s ファイル数=%s シート=%s ファイル=%s",
            total_rows,
            len(paths),
            getattr(sht, "name", "?"),
            files_str,
        )
        parent_hwnd2 = int(parent_hwnd or 0)
        if progress_path is not None:
            _wait_progress_closed_ack(progress_closed_path)
        elif parent_hwnd2:
            _phase("P4-9", step="done_popup_fallback", items=len(done_ser))
            _submit_done_ui(
                parent_hwnd=parent_hwnd2, sheet_id=sheet_guid, done_items=done_ser
            )

    except Exception as ex:
        logger.error("[CSV_MG] 結合失敗: %s", ex, exc_info=True)
        if progress_path is not None:
            _progress_write_monotonic(
                progress_path,
                {"status": "ERROR", "phase_i": 4, "phase": "エラー", "detail": str(ex)},
            )
    finally:
        if xlc is not None:
            try:
                _sht_restore = sht
            except NameError:
                pass
            else:
                try:
                    xlc.restore_screen_updating(_sht_restore)
                except Exception:
                    pass
        if progress_path is not None:
            try:
                _PROGRESS_SEQ.pop(_progress_key(progress_path), None)
            except Exception:
                pass
        try:
            if pythoncom is not None:
                pythoncom.CoUninitialize()
        except Exception:
            pass


# ======================================================================================
# Public API（VBA → hc_main → svc）
# ======================================================================================


def merge_csv(book: Any, sheet_id: str = "") -> None:
    """
    Method Name : merge_csv
    Arguments   : book (xlwings.Book), sheet_id (GUID)
    Return      : None

    概要:
      1) UIサーバが起動していなければ起動
      2) req投入（action="csv_mg"）
      3) READY_UI / RESULT を非同期監視して処理
      4) 結果を待機し、OKなら結合処理まで完了してから return（Excel操作を抑止）
    """
    # ==========================================================================
    # フェーズログ（恒久）セクション
    # ==========================================================================
    # PHASE: merge_csv 工程ログ基準時刻をセット
    global _PHASE_T0_PERF
    _PHASE_T0_PERF = time.perf_counter()

    _phase("P0-0", step="merge_csv_enter", sheet_id=sheet_id)
    t_enter0 = _PHASE_T0_PERF or time.perf_counter()
    logger.info(
        "[CSV_MG] 開始 sheet_id=%s svc_pid=%s",
        sheet_id or "",
        os.getpid(),
    )
    _mg_trace(
        "[CSV_MG_TRACE] phase=enter sheet_id=%s svc_pid=%s wall_perf_s=%.6f",
        sheet_id or "",
        os.getpid(),
        time.perf_counter(),
    )

    if book is None:
        logger.error("[CSV_MG] book=None のため中断（WaitForm 解除を試行）")
        try:
            notify_wait_form_ready(book=book)
        except Exception:
            pass
        return

    try:
        ensure_qt_ui_server()
        logger.info(
            "[CSV_MG] phase=after_ensure_ui_server elapsed_ms=%s",
            _elapsed_ms(t_enter0),
        )
        _mg_trace(
            "[CSV_MG_TRACE] phase=after_ensure_ui_server elapsed_ms=%s wall_perf_s=%.6f",
            _elapsed_ms(t_enter0),
            time.perf_counter(),
        )

        parent_hwnd = 0
        try:
            parent_hwnd = int(getattr(book.app, "hwnd", 0))
            os.environ["HC_EXCEL_HWND"] = str(parent_hwnd)
        except Exception as e:
            parent_hwnd = 0
            logger.warning("[CSV_MG] Excel HWND 取得失敗: %s", e)
        if not parent_hwnd:
            logger.warning("[CSV_MG] parent_hwnd=0")

        ipc_root = Path(get_ipc_root())
        res_path = str(
            ipc_root / "results" / f"res_{sheet_id}_{int(time.time()*1000)}.pkl"
        )
        ready_path = str(
            ipc_root / "ready" / f"ready_{sheet_id}_{int(time.time()*1000)}.pkl"
        )

        Path(res_path).parent.mkdir(parents=True, exist_ok=True)
        Path(ready_path).parent.mkdir(parents=True, exist_ok=True)

        # EOFError 対策：残骸があれば消す
        try:
            Path(res_path).unlink(missing_ok=True)
            Path(ready_path).unlink(missing_ok=True)
        except Exception:
            pass

        req_dict = {
            "parent_hwnd": int(parent_hwnd),
            "result_path": res_path,
            "ready_path": ready_path,
            "sheet_id": str(sheet_id or ""),
            "log_path": "",
            "action": "csv_mg",
            "module": "ui_qt.ui_csv_mg",
        }
        er_main = _get_window_rect(int(parent_hwnd or 0))
        if er_main is not None:
            req_dict["excel_rect"] = list(er_main)

        _phase("P2-0", step="submit_request_begin")

        req_path = _submit_request_dict(req_dict)
        _phase("P2-1", step="submit_request_done", req=str(req_path))
        logger.info(
            "[CSV_MG] ui_ipc ok req=%s sheet_id=%s hwnd=%s",
            req_path.name,
            sheet_id or "",
            parent_hwnd,
        )
        _mg_trace(
            "[CSV_MG_TRACE] ui_ipc ok req=%s sheet_id=%s hwnd=%s elapsed_ms=%s wall_perf_s=%.6f",
            req_path.name,
            sheet_id or "",
            parent_hwnd,
            _elapsed_ms(t_enter0),
            time.perf_counter(),
        )

        # READY/UI結果を待機（VBAを固める＝Excel操作を抑止し、途中終了で残留しない設計）
        # NOTE:
        #   - UIは別プロセス(ui_server)なので、この待機でUIが固まることはない
        #   - merge処理をバックグラウンドに残さない（Excel終了時にExcel/Pythonが残留する問題を抑止）

        # READY通知（カーソルガード解除等）はスレッドで待つ（VBA側タイマーを止めるため）
        th_ready = threading.Thread(
            target=_watch_ready,
            args=(ready_path, str(sheet_id or ""), t_enter0),
            name=f"ready_watch_{sheet_id}",
            daemon=True,
        )
        _phase("P2-3", step="ready_watch_start")
        th_ready.start()

        attach_keys = _capture_book_attach_keys(book)
        wb_name = ""
        try:
            wb_name = str(getattr(book, "name", "") or getattr(book, "Name", ""))
        except Exception:
            wb_name = ""

        # 結果を同期で待つ（OKならこのスレッドでmergeまで完了）
        _phase("P2-4", step="result_watch_begin", res=str(res_path))
        _watch_result(
            res_path,
            wb_name,
            str(sheet_id or ""),
            t_enter0,
            parent_hwnd=parent_hwnd,
            attach_keys=attach_keys,
        )
        _phase("P2-5", step="result_watch_done")
        logger.info(
            "[CSV_MG] phase=merge_csv_flow_done elapsed_ms=%s sheet_id=%s",
            _elapsed_ms(t_enter0),
            sheet_id or "",
        )
        _mg_trace(
            "[CSV_MG_TRACE] phase=merge_csv_flow_done elapsed_ms=%s wall_perf_s=%.6f",
            _elapsed_ms(t_enter0),
            time.perf_counter(),
        )

    except Exception as ex:
        logger.error("[CSV_MG] 致命的エラー: %s", ex, exc_info=True)
        try:
            notify_wait_form_ready(book=book)
        except Exception:
            pass
