# -*- coding: utf-8 -*-
"""
Python: 3.12
Module: svc/svc_row_dl.py
Created: 2026-03-06
Updated: 2026-05-05
Version: 1.2.1
Purpose:
  空白行削除（Excel UsedRange）。進捗→一覧確認→削除確定時のみ物理削除。画面は ui_row_dl + ui_row_dl.json。
History (latest 3):
  - 1.2.1 (2026-05-05) 進捗→サブ画面順序を ACK 待ちで保証。progress_closed_path を進捗UIへ渡し、進捗クローズ後に preview/done を表示（sleep 依存を廃止）。
  - 1.2.0 (2026-04-10) UsedRange 列開始の整合、進捗キャンセル、確認モーダル、Interactive 維持、完了待機。
  - 1.1.0 (2026-04-06) HC_LOG_PERF: [ROW_DL_PERF]。診断: [ROW_DL_TRACE]。
  - 1.0.0 (2026-03-11) hc_row_dl から分離。core_xlc/get_excel_context_from_hwnd、進捗IPC、完了通知。
  - 初出 (2026-03-06) 計画に基づく svc_row_dl + ui_row_dl 新規作成。
"""
from __future__ import annotations

import os
import sys
import threading
import time
from pathlib import Path
from typing import Any, Callable, Optional

# プロジェクトルートをパスに追加（Excel アドインから呼ばれるため）
_path_svc = os.path.dirname(os.path.abspath(__file__))
_root = os.path.dirname(_path_svc)
if _root not in sys.path:
    sys.path.insert(0, _root)

from core.core_log import get_diag_logger, get_logger, get_perf_logger  # noqa: E402
from ui_qt.ipc_file import get_ipc_root, get_request_dir, read_pickle, write_pickle  # noqa: E402

logger = get_logger(__name__)
_row_diag = get_diag_logger("hc_csv_tool.diag.row_dl")
_perf = get_perf_logger("svc.svc_row_dl.perf")
__version__ = "1.2.1"


class _RowDlCancelled(Exception):
    """進捗キャンセルで走査を打ち切る。"""


_CANCEL_POLL_INTERVAL_SEC = 0.2
_PROGRESS_CLOSE_ACK_TIMEOUT_SEC = 3.0
_PROGRESS_CLOSE_ACK_POLL_SEC = 0.03


def _elapsed_ms(since: float) -> int:
    return max(0, int((time.perf_counter() - since) * 1000))


def _row_trace(phase: str, t0: float, **kv: object) -> None:
    try:
        if kv:
            _row_diag.info(
                "[ROW_DL_TRACE] phase=%s cumulative_ms=%d %s",
                phase,
                _elapsed_ms(t0),
                " ".join("%s=%s" % (k, v) for k, v in kv.items()),
            )
        else:
            _row_diag.info("[ROW_DL_TRACE] phase=%s cumulative_ms=%d", phase, _elapsed_ms(t0))
    except Exception:
        pass


def _perf_row(phase: str, t0: float, **kv: object) -> None:
    try:
        if kv:
            _perf.info(
                "[ROW_DL_PERF] phase=%s cumulative_ms=%d %s",
                phase,
                _elapsed_ms(t0),
                " ".join("%s=%s" % (k, v) for k, v in kv.items()),
            )
        else:
            _perf.info("[ROW_DL_PERF] phase=%s cumulative_ms=%d", phase, _elapsed_ms(t0))
    except Exception:
        pass

try:
    from core import core_cst as cst
except Exception:
    cst = None  # type: ignore

try:
    from core import core_xlc as core_xlc_mod
except Exception:
    core_xlc_mod = None  # type: ignore


def _status_bar_save(book: Any) -> str:
    """
    現在の Excel ステータスバー文言を退避する。
    処理後に _status_bar_restore で復元するために使用する。
    """
    try:
        return str(book.app.api.StatusBar or "")
    except Exception:
        return ""


def _status_bar_set(book: Any, msg: str) -> None:
    """
    Excel のステータスバーに指定メッセージを表示する。
    処理中・完了・エラー時のユーザーへのフィードバック用。
    """
    try:
        book.app.api.DisplayStatusBar = True
        book.app.api.StatusBar = str(msg)
    except Exception:
        pass


def _status_bar_restore(book: Any, saved: str) -> None:
    """
    ステータスバーを _status_bar_save で退避した文言に戻す。
    処理終了時（正常・異常問わず）に必ず呼ぶ。
    """
    try:
        book.app.api.StatusBar = saved
    except Exception:
        pass


def _cfg() -> dict[str, Any]:
    """
    空白行削除用の画面・メッセージ設定を config/ui_row_dl.json から読み込む。
    読込失敗時は UiConfigLoadError が発生する（救済なし）。
    """
    if cst is None:
        return {}
    return cst.get_ui_config_from_file_required("row_dl")


def _msg(cfg: dict[str, Any], key: str, **fmt: Any) -> str:
    """
    設定の MESSAGES からキーに対応する文言を取得し、任意でフォーマットする。
    """
    m = (cfg.get("MESSAGES") or {}).get(key) or key
    try:
        return str(m).format(**fmt)
    except Exception:
        return str(m)


def _progress_path(sheet_id: str) -> Path:
    """
    進捗状態を書き出す Pickle ファイルのパスを返す。
    ui_server 側の進捗ダイアログがこのファイルをポーリングして表示を更新する。
    """
    d = Path(get_ipc_root()) / "progress"
    d.mkdir(parents=True, exist_ok=True)
    return d / f"progress_row_dl_{sheet_id}.pkl"


def _cancel_request_path(sheet_id: str) -> Path:
    d = Path(get_ipc_root()) / "progress"
    d.mkdir(parents=True, exist_ok=True)
    return d / f"cancel_req_row_dl_{sheet_id}.pkl"


def _progress_closed_ack_path(sheet_id: str) -> Path:
    """進捗ダイアログが閉じたことを示す ACK ファイル。"""
    d = Path(get_ipc_root()) / "progress"
    d.mkdir(parents=True, exist_ok=True)
    return d / f"progress_row_dl_closed_{sheet_id}.pkl"


def _wait_progress_closed_ack(path: Optional[Path], timeout_sec: float = _PROGRESS_CLOSE_ACK_TIMEOUT_SEC) -> None:
    """
    進捗画面のクローズACKを短時間待つ。
    タイムアウト時はログのみ残し先へ進む（処理全体を止めない）。
    """
    if path is None:
        return
    p = path
    t0 = time.perf_counter()
    while True:
        try:
            if p.exists():
                return
        except Exception:
            return
        if (time.perf_counter() - t0) >= max(0.05, float(timeout_sec)):
            logger.info("[ROW_DL] progress close ack timeout: %s", str(p))
            return
        time.sleep(_PROGRESS_CLOSE_ACK_POLL_SEC)


def _reset_cancel_path(path: Path) -> None:
    try:
        path.unlink(missing_ok=True)  # type: ignore[call-arg]
    except TypeError:
        try:
            if path.exists():
                path.unlink()
        except Exception:
            pass
    except Exception:
        pass


def _cancel_requested(path: Path) -> bool:
    try:
        d = read_pickle(path)
        return isinstance(d, dict) and bool(d.get("cancel"))
    except Exception:
        return False


def _wait_ui_dispatch_result(result_path: Path, timeout_sec: float = 120.0) -> dict[str, Any] | None:
    t0 = time.time()
    while (time.time() - t0) < timeout_sec:
        if result_path.exists() and result_path.stat().st_size > 0:
            try:
                return read_pickle(result_path)
            except Exception:
                pass
        time.sleep(0.05)
    return None


def _progress_write(path: Path, obj: dict[str, Any]) -> None:
    """
    進捗情報を Pickle で path に書き出す。
    seq が未指定の場合は既存ファイルの seq をインクリメントして順序を保証する。
    """
    try:
        from ui_qt.ipc_file import read_pickle

        obj = dict(obj)
        if "seq" not in obj:
            try:
                prev = read_pickle(path)
                seq = int(prev.get("seq", -1)) + 1 if isinstance(prev, dict) else 0
            except Exception:
                seq = 0
            obj["seq"] = seq
        write_pickle(path, obj)
    except Exception:
        pass


def _get_window_rect(hwnd: int) -> tuple[int, int, int, int] | None:
    """
    Win32 API でウィンドウのクライアント外枠（left, top, right, bottom）を取得する。
    進捗ダイアログを Excel ウィンドウ付近に表示する際の基準に使う。
    """
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
    cancel_request_path: Path | None = None,
    progress_closed_path: Path | None = None,
) -> None:
    """
    UI サーバに進捗画面表示を依頼する。req_*.pkl に payload を書き、ui_server が ui_row_dl.create_dialog を呼ぶ。
    """
    try:
        from svc.svc_host import ensure_ui_server

        ensure_ui_server()
        excel_rect = _get_window_rect(int(parent_hwnd or 0))
        res_dir = Path(get_ipc_root()) / "result"
        res_dir.mkdir(parents=True, exist_ok=True)
        ts_ms = int(time.time() * 1000)
        result_path = str(res_dir / f"res_progress_row_dl_{ts_ms}_{os.getpid()}.pkl")
        req_dict: dict[str, Any] = {
            "action": "progress",
            "progress_path": str(progress_path),
            "phase_total": int(phase_total),
            "excel_lock": False,
            "no_native_window": True,
        }
        if cancel_request_path is not None:
            cr = str(cancel_request_path).strip()
            if cr:
                req_dict["cancel_request_path"] = cr
        if progress_closed_path is not None:
            cp = str(progress_closed_path).strip()
            if cp:
                req_dict["progress_closed_path"] = cp
        if excel_rect is not None:
            req_dict["excel_rect"] = list(excel_rect)
        payload = {
            "parent_hwnd": int(parent_hwnd),
            "result_path": result_path,
            "ready_path": "",
            "sheet_id": str(sheet_id),
            "log_path": "",
            "action": "progress",
            "module": "ui_qt.ui_row_dl",
            "req_dict": req_dict,
        }
        req_path = get_request_dir() / f"req_{ts_ms}_{os.getpid()}_{threading.get_ident()}.pkl"
        write_pickle(req_path, payload)
        logger.info("[ROW_DL] progress UI request: %s", req_path)
    except Exception as exc:
        logger.warning("[ROW_DL] progress UI request failed: %s", exc)


def _submit_done_ui(parent_hwnd: int, sheet_id: str, message: str, title: str = "空白行削除") -> Path | None:
    """
    完了通知をモーダルで表示するため ui_server に依頼する。
    SCREENS.DONE の設定に従い、アイコン・中央表示・OK ボタンで閉じる。
    戻り値は ui_server が結果を書き込む result_path（待機用）。失敗時は None。
    """
    try:
        from svc.svc_host import ensure_ui_server

        ensure_ui_server()
        ts_ms = int(time.time() * 1000)
        res_dir = Path(get_ipc_root()) / "result"
        res_dir.mkdir(parents=True, exist_ok=True)
        result_path = res_dir / f"res_row_dl_done_{ts_ms}_{os.getpid()}.pkl"
        req_dict: dict[str, Any] = {
            "action": "row_dl_done",
            "modeless": False,
            "title": str(title),
            "message": str(message),
        }
        er_done = _get_window_rect(int(parent_hwnd or 0))
        if er_done is not None:
            req_dict["excel_rect"] = list(er_done)
        payload = {
            "parent_hwnd": int(parent_hwnd),
            "result_path": str(result_path),
            "ready_path": "",
            "sheet_id": str(sheet_id),
            "log_path": "",
            "action": "row_dl",
            "module": "ui_qt.ui_row_dl",
            "req_dict": req_dict,
        }
        req_path = get_request_dir() / f"req_{ts_ms}_{os.getpid()}_done.pkl"
        write_pickle(req_path, payload)
        logger.info("[ROW_DL] done UI request: %s", req_path)
        return result_path
    except Exception as exc:
        logger.warning("[ROW_DL] done UI request failed: %s", exc)
        return None


def _submit_preview_ui(
    parent_hwnd: int,
    sheet_id: str,
    *,
    kind: str,
    items: list[Any],
    cfg: dict[str, Any],
) -> Path | None:
    """空欄一覧の確認（削除／キャンセル）または空欄なし（OK）のモーダル。result_path を返す。"""
    try:
        from svc.svc_host import ensure_ui_server

        ensure_ui_server()
        ts_ms = int(time.time() * 1000)
        res_dir = Path(get_ipc_root()) / "result"
        res_dir.mkdir(parents=True, exist_ok=True)
        result_path = res_dir / f"res_row_dl_preview_{ts_ms}_{os.getpid()}.pkl"
        preview_cfg = (cfg.get("SCREENS") or {}).get("PREVIEW") or {}
        title = str(preview_cfg.get("TITLE") or "空白行削除")
        req_dict: dict[str, Any] = {
            "action": "row_dl_preview",
            "modeless": False,
            "preview_kind": str(kind),
            "items": list(items),
            "title": title,
        }
        er = _get_window_rect(int(parent_hwnd or 0))
        if er is not None:
            req_dict["excel_rect"] = list(er)
        payload = {
            "parent_hwnd": int(parent_hwnd),
            "result_path": str(result_path),
            "ready_path": "",
            "sheet_id": str(sheet_id),
            "log_path": "",
            "action": "row_dl",
            "module": "ui_qt.ui_row_dl",
            "req_dict": req_dict,
        }
        req_path = get_request_dir() / f"req_row_dl_preview_{ts_ms}_{os.getpid()}.pkl"
        write_pickle(req_path, payload)
        logger.info("[ROW_DL] preview UI request: %s", req_path)
        return result_path
    except Exception as exc:
        logger.warning("[ROW_DL] preview UI request failed: %s", exc)
        return None


def _normalize_2d(raw: Any, yn: int, xn: int) -> list[list[Any]]:
    """
    xlwings の Range.value で得た値を、yn×xn の 2 次元リストに正規化する。
    単一セル・1 行・複数行のいずれでも、欠けている要素は None で埋める。
    """
    if yn <= 0 or xn <= 0:
        return []
    if yn == 1 and xn == 1:
        return [[raw]]
    if yn == 1:
        row = raw if isinstance(raw, list) else [raw]
        return [row[:xn] + [None] * max(0, xn - len(row))]
    out: list[list[Any]] = []
    if not isinstance(raw, list):
        return [[None] * xn for _ in range(yn)]
    for r in range(yn):
        row = raw[r] if r < len(raw) else None
        if row is None:
            out.append([None] * xn)
        elif isinstance(row, list):
            out.append((row + [None] * xn)[:xn])
        else:
            out.append([row] + [None] * (xn - 1) if xn > 1 else [row])
    return out


def _read_sheet_matrix(
    ptr_s: Any,
    y1: int,
    x1: int,
    yn: int,
    xn: int,
    on_pct: Callable[[int, str, str], None],
    cfg: dict[str, Any],
    *,
    cancel_path: Path | None = None,
) -> Optional[list[list[Any]]]:
    """
    シートの指定範囲 (y1,x1) から yn×xn をチャンク単位で読み、2 次元リストで返す。
    読込中は on_pct で進捗コールバックを呼ぶ。失敗時は None。
    cancel_path があればチャンク境界でキャンセル要求を見て _RowDlCancelled を送出する。
    """
    msg_scan = _msg(cfg, "PHASE_SCAN")
    custom = (cfg.get("MESSAGES") or {}).get("PROGRESS_CUSTOM_SCAN") or "走査中"
    chunk_rows = max(200, min(5000, yn))
    acc: list[list[Any]] = []
    last_poll = 0.0

    def _poll_cancel() -> None:
        nonlocal last_poll
        if cancel_path is None:
            return
        now = time.monotonic()
        if now - last_poll < _CANCEL_POLL_INTERVAL_SEC:
            return
        last_poll = now
        if _cancel_requested(cancel_path):
            raise _RowDlCancelled()

    try:
        for r0 in range(0, yn, chunk_rows):
            _poll_cancel()
            r1 = min(r0 + chunk_rows, yn)
            pct = int(5 + (r1 / max(yn, 1)) * 35)
            on_pct(pct, msg_scan, custom)
            rng = ptr_s.range((y1 + r0, x1), (y1 + r1 - 1, x1 + xn - 1))
            part = rng.value
            sub = _normalize_2d(part, r1 - r0, xn)
            acc.extend(sub)
        return acc if len(acc) == yn else None
    except _RowDlCancelled:
        raise
    except Exception:
        return None


def _sheet_id_resolve(ptr_s: Any, sheet_id: str) -> str:
    """
    進捗用のシート識別子を返す。
    sheet_id が空の場合はシートの HC_GUID_B64 を取得し、無ければオブジェクト id ベースのフォールバックを使う。
    """
    s = str(sheet_id or "").strip()
    if s:
        return s
    if core_xlc_mod is not None:
        try:
            g = core_xlc_mod.get_sheet_prop(ptr_s, "HC_GUID_B64")
            if g:
                return str(g)[:48]
        except Exception:
            pass
    return f"row_dl_{abs(id(ptr_s))}"


def delete_empty_rows(target_hwnd: Optional[int] = None, sheet_id: str = "") -> None:
    """
    【概要】
        指定 HWND の Excel ブック・アクティブシートの使用範囲内で、空白行を検出して物理削除する。
    【補足】
        進捗・完了通知は ui_server 経由で ui_row_dl に依頼する。設定は config/ui_row_dl.json。
    """
    t_flow = time.perf_counter()
    _perf_row("enter", t_flow)
    _row_trace("enter", t_flow)

    if core_xlc_mod is None:
        logger.error("[ROW_DL] core_xlc not available")
        _perf_row("abort_no_core_xlc", t_flow)
        _row_trace("abort_no_core_xlc", t_flow)
        return
    ctx = core_xlc_mod.get_excel_context_from_hwnd(int(target_hwnd or 0), sheet_id)
    if ctx is None:
        logger.error("[ROW_DL] Excel context not available (xlwings + HWND)")
        _perf_row("abort_no_context", t_flow)
        _row_trace("abort_no_context", t_flow)
        return

    ptr_a, ptr_w, ptr_s, ph = ctx
    logger.info("[ROW_DL] 開始")
    _perf_row("after_context", t_flow, hwnd=ph)
    _row_trace("after_context", t_flow, hwnd=ph)
    cfg = _cfg()
    saved_status = _status_bar_save(ptr_w)  # 終了時に必ず復元
    sid = _sheet_id_resolve(ptr_s, sheet_id)
    prog_path = _progress_path(sid)
    seq = [0]  # 進捗の表示順序用

    def _upd(pct: int, phase: str, cur: str) -> None:
        seq[0] += 1
        _progress_write(
            prog_path,
            {
                "status": "RUN",
                "phase": phase,
                "msg": phase,
                "current_file": cur,
                "pct": max(0, min(100, pct)),
                "done": pct,
                "total": 100,
                "seq": seq[0],
            },
        )

    cancel_path = _cancel_request_path(sid)
    progress_closed_path = _progress_closed_ack_path(sid)
    _reset_cancel_path(progress_closed_path)
    interactive_locked = False
    progress_closed_waited = [False]

    def _ensure_progress_closed_before_modal() -> None:
        if progress_closed_waited[0]:
            return
        _wait_progress_closed_ack(progress_closed_path)
        progress_closed_waited[0] = True

    try:
        api_used = ptr_s.used_range
        val_r_orig = int(api_used.row)
        val_c_orig = int(api_used.column)
        val_r_num = int(api_used.rows.count)
        val_c_num = int(api_used.columns.count)
        if val_r_num < 1 or val_c_num < 1:
            logger.info("[ROW_DL] 使用範囲なし")
            _perf_row("early_no_used_range", t_flow)
            _row_trace("early_no_used_range", t_flow)
            return

        logger.info(
            "[ROW_DL] 使用範囲 row=%s col=%s rows=%s cols=%s",
            val_r_orig,
            val_c_orig,
            val_r_num,
            val_c_num,
        )
        _perf_row("after_used_range", t_flow, rows=val_r_num, cols=val_c_num)
        _row_trace("after_used_range", t_flow, rows=val_r_num, cols=val_c_num)

        ptr_a.api.Interactive = False
        interactive_locked = True

        _reset_cancel_path(cancel_path)
        _progress_write(
            prog_path,
            {
                "status": "RUN",
                "phase": _msg(cfg, "PHASE_SCAN"),
                "pct": 0,
                "done": 0,
                "total": 100,
                "seq": 0,
            },
        )
        _submit_progress_ui(
            ph,
            sid,
            prog_path,
            2,
            cancel_request_path=cancel_path,
            progress_closed_path=progress_closed_path,
        )
        _perf_row("after_progress_ui_submit", t_flow)
        _row_trace("after_progress_ui_submit", t_flow)

        try:
            arr = _read_sheet_matrix(
                ptr_s,
                val_r_orig,
                val_c_orig,
                val_r_num,
                val_c_num,
                _upd,
                cfg,
                cancel_path=cancel_path,
            )
        except _RowDlCancelled:
            logger.info("[ROW_DL] 走査キャンセル")
            _perf_row("user_cancel_scan", t_flow)
            _row_trace("user_cancel_scan", t_flow)
            _progress_write(prog_path, {"status": "DONE", "seq": 999})
            return

        if arr is None:
            logger.warning("[ROW_DL] 読込失敗")
            _perf_row("abort_matrix_read_failed", t_flow)
            _row_trace("abort_matrix_read_failed", t_flow)
            _progress_write(prog_path, {"status": "DONE", "seq": 999})
            _status_bar_set(ptr_w, _msg(cfg, "ERROR_PREFIX"))
            done_cfg = (cfg.get("SCREENS") or {}).get("DONE") or {}
            _ensure_progress_closed_before_modal()
            rp_err = _submit_done_ui(
                ph, sid, _msg(cfg, "ERROR_PREFIX"), str(done_cfg.get("TITLE") or "空白行削除")
            )
            if rp_err is not None:
                _wait_ui_dispatch_result(rp_err)
            return

        _perf_row("after_matrix_read", t_flow, rows=val_r_num, cols=val_c_num)
        _row_trace("after_matrix_read", t_flow, rows=val_r_num, cols=val_c_num)

        list_del_apis: list[Any] = []
        list_del_row_indices: list[int] = []
        for i_idx in range(len(arr)):
            it_vals = arr[i_idx]
            abs_y = val_r_orig + i_idx
            bool_is_empty = True
            for v_c in it_vals:
                if v_c is not None:
                    str_cell_v = str(v_c)
                    if str_cell_v.strip() != "":
                        bool_is_empty = False
                        break
            if bool_is_empty:
                str_row_addr = f"{abs_y}:{abs_y}"
                ptr_range_obj = ptr_s.range(str_row_addr)
                list_del_apis.append(ptr_range_obj.api)
                list_del_row_indices.append(abs_y)

        val_count_del = len(list_del_apis)
        if val_count_del >= 1:
            logger.info("[ROW_DL] 空白行 件数=%s 位置(行)=%s", val_count_del, list_del_row_indices)

        _progress_write(prog_path, {"status": "DONE", "seq": 999})
        _ensure_progress_closed_before_modal()

        _done_screen = (cfg.get("SCREENS") or {}).get("DONE") or {}
        done_title = str(_done_screen.get("TITLE") or "空白行削除")

        if val_count_del == 0:
            logger.info("[ROW_DL] 空白行なし")
            _perf_row("early_no_blank_rows", t_flow)
            _row_trace("early_no_blank_rows", t_flow)
            _status_bar_set(ptr_w, _msg(cfg, "STATUS_NONE"))
            rp0 = _submit_preview_ui(ph, sid, kind="none", items=[], cfg=cfg)
            if rp0 is not None:
                _wait_ui_dispatch_result(rp0)
            return

        _perf_row("after_scan_blank_rows", t_flow, delete_count=val_count_del)
        _row_trace("after_scan_blank_rows", t_flow, delete_count=val_count_del)

        rp_prev = _submit_preview_ui(
            ph, sid, kind="confirm", items=list_del_row_indices, cfg=cfg
        )
        if rp_prev is None:
            return
        res_prev = _wait_ui_dispatch_result(rp_prev)
        choice = (res_prev or {}).get("choice") if isinstance(res_prev, dict) else None
        if choice != "delete":
            logger.info("[ROW_DL] 削除を中止（確認: %s）", choice)
            _perf_row("user_cancel_confirm", t_flow, choice=choice)
            _row_trace("user_cancel_confirm", t_flow, choice=choice)
            return

        try:
            from svc.svc_undo import save_undo_snapshot

            save_undo_snapshot(ptr_w, sheet_id=sheet_id, target_hwnd=ph, excel_hwnd=ph)
        except Exception as e:
            logger.warning("[ROW_DL] save_undo_snapshot failed (undo unavailable): %s", e)

        from core import core_xlc

        with core_xlc.suspend_sheet_updates(ptr_s):
            for api_it in reversed(list_del_apis):
                api_it.Delete()

        positions_str = ", ".join(str(i) for i in list_del_row_indices)
        logger.info("[ROW_DL] 完了 抹消=%s 位置(行)=%s", val_count_del, list_del_row_indices)
        _status_bar_set(ptr_w, _msg(cfg, "STATUS_DONE", count=val_count_del, positions=positions_str))
        rp_done = _submit_done_ui(
            ph,
            sid,
            _msg(cfg, "STATUS_DONE", count=val_count_del, positions=positions_str),
            done_title,
        )
        if rp_done is not None:
            _wait_ui_dispatch_result(rp_done)
        _perf_row("after_delete_done", t_flow, deleted=val_count_del)
        _row_trace("after_delete_done", t_flow, deleted=val_count_del)

    except Exception as ex:
        logger.exception("[ROW_DL] %s", ex)
        try:
            _status_bar_set(ptr_w, f"{_msg(cfg, 'ERROR_PREFIX')}: {ex}")
        except Exception:
            pass
    finally:
        if interactive_locked:
            try:
                ptr_a.api.Interactive = True
            except Exception:
                pass
        try:
            from ui_qt.ui_common import enable_excel_window

            enable_excel_window(ph, True)
        except Exception:
            pass
        try:
            _progress_write(prog_path, {"status": "DONE", "seq": 999})
        except Exception:
            pass
        try:
            _status_bar_restore(ptr_w, saved_status)
        except Exception:
            pass
        try:
            from core import core_w32

            core_w32.bring_to_front(ph)
        except Exception:
            pass
        _perf_row("flow_end", t_flow)
        _row_trace("flow_end", t_flow)
