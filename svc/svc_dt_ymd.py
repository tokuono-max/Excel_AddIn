# -*- coding: utf-8 -*-
"""
Python: 3.12
Module: svc/svc_dt_ymd.py
Created: 2026-03-06
Updated: 2026-04-06
Version: 1.1.2
Purpose:
  日付変換（選択範囲を YYYY/MM/DD 形式へ一括変換）。処理は本モジュール、画面は ui_qt.ui_dt_ymd + config/ui_dt_ymd.json。
History (latest 3):
  - 1.1.0 (2026-04-06) HC_LOG_PERF: [DT_YMD_PERF]。診断: [DT_YMD_TRACE]。
  - 1.0.0 (2026-03-18) hc_dt_ymd から分離。core_xlc/get_excel_context_from_hwnd、進捗IPC、完了通知。
  - 初出 (2026-03-06) 計画に基づく svc_dt_ymd + ui_dt_ymd 新規作成。
"""
from __future__ import annotations

import os
import re
import sys
import threading
import time
import unicodedata
from pathlib import Path
from typing import Any, Callable, Optional

import pandas as pd

# プロジェクトルートをパスに追加（Excel アドインから呼ばれるため）
_path_svc = os.path.dirname(os.path.abspath(__file__))
_root = os.path.dirname(_path_svc)
if _root not in sys.path:
    sys.path.insert(0, _root)

from core.core_log import get_diag_logger, get_logger, get_perf_logger  # noqa: E402
from ui_qt.ipc_file import get_ipc_root, get_request_dir, write_pickle  # noqa: E402

logger = get_logger(__name__)
_dt_ymd_diag = get_diag_logger("hc_csv_tool.diag.dt_ymd")
_perf = get_perf_logger("svc.svc_dt_ymd.perf")
__version__ = "1.1.2"

_PROGRESS_CLOSE_ACK_TIMEOUT_SEC = 3.0
_PROGRESS_CLOSE_ACK_POLL_SEC = 0.03


def _elapsed_ms(since: float) -> int:
    return max(0, int((time.perf_counter() - since) * 1000))


def _dt_ymd_trace(phase: str, t0: float, **kv: object) -> None:
    try:
        if kv:
            _dt_ymd_diag.info(
                "[DT_YMD_TRACE] phase=%s cumulative_ms=%d %s",
                phase,
                _elapsed_ms(t0),
                " ".join("%s=%s" % (k, v) for k, v in kv.items()),
            )
        else:
            _dt_ymd_diag.info("[DT_YMD_TRACE] phase=%s cumulative_ms=%d", phase, _elapsed_ms(t0))
    except Exception:
        pass


def _perf_dt_ymd(phase: str, t0: float, **kv: object) -> None:
    try:
        if kv:
            _perf.info(
                "[DT_YMD_PERF] phase=%s cumulative_ms=%d %s",
                phase,
                _elapsed_ms(t0),
                " ".join("%s=%s" % (k, v) for k, v in kv.items()),
            )
        else:
            _perf.info("[DT_YMD_PERF] phase=%s cumulative_ms=%d", phase, _elapsed_ms(t0))
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
    日付変換用の画面・メッセージ設定を config/ui_dt_ymd.json から読み込む。
    読込失敗時は UiConfigLoadError が発生する（救済なし）。
    """
    if cst is None:
        return {}
    return cst.get_ui_config_from_file_required("dt_ymd")


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
    return d / f"progress_dt_ymd_{sheet_id}.pkl"


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


def _progress_closed_ack_path(sheet_id: str) -> Path:
    d = Path(get_ipc_root()) / "progress"
    d.mkdir(parents=True, exist_ok=True)
    return d / f"progress_dt_ymd_closed_{sheet_id}.pkl"


def _wait_progress_closed_ack(path: Optional[Path], timeout_sec: float = _PROGRESS_CLOSE_ACK_TIMEOUT_SEC) -> None:
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
            logger.info("[DT_YMD] progress close ack timeout: %s", str(p))
            return
        time.sleep(_PROGRESS_CLOSE_ACK_POLL_SEC)


def _submit_progress_ui(
    parent_hwnd: int,
    sheet_id: str,
    progress_path: Path,
    phase_total: int,
    *,
    progress_closed_path: Path | None = None,
) -> None:
    """
    UI サーバに進捗画面表示を依頼する。req_*.pkl に payload を書き、ui_server が ui_dt_ymd.create_dialog を呼ぶ。
    """
    try:
        from svc.svc_host import ensure_ui_server

        ensure_ui_server()
        excel_rect = _get_window_rect(int(parent_hwnd or 0))
        res_dir = Path(get_ipc_root()) / "result"
        res_dir.mkdir(parents=True, exist_ok=True)
        ts_ms = int(time.time() * 1000)
        result_path = str(res_dir / f"res_progress_dt_ymd_{ts_ms}_{os.getpid()}.pkl")
        req_dict: dict[str, Any] = {
            "action": "progress",
            "progress_path": str(progress_path),
            "phase_total": int(phase_total),
            "excel_lock": False,
            "no_native_window": True,
        }
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
            "module": "ui_qt.ui_dt_ymd",
            "req_dict": req_dict,
        }
        req_path = get_request_dir() / f"req_{ts_ms}_{os.getpid()}_{threading.get_ident()}.pkl"
        write_pickle(req_path, payload)
        logger.info("[DT_YMD] progress UI request: %s", req_path)
    except Exception as exc:
        logger.warning("[DT_YMD] progress UI request failed: %s", exc)


def _submit_done_ui(parent_hwnd: int, sheet_id: str, message: str, title: str = "日付変換") -> None:
    """
    完了通知をモーダルで表示するため ui_server に依頼する。
    SCREENS.DONE の設定に従い、アイコン・中央表示・OK ボタンで閉じる。
    """
    try:
        from svc.svc_host import ensure_ui_server

        ensure_ui_server()
        ts_ms = int(time.time() * 1000)
        res_dir = Path(get_ipc_root()) / "result"
        res_dir.mkdir(parents=True, exist_ok=True)
        result_path = str(res_dir / f"res_dt_ymd_done_{ts_ms}_{os.getpid()}.pkl")
        req_dict: dict[str, Any] = {
            "action": "dt_ymd_done",
            "modeless": False,
            "title": str(title),
            "message": str(message),
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
            "action": "dt_ymd",
            "module": "ui_qt.ui_dt_ymd",
            "req_dict": req_dict,
        }
        req_path = get_request_dir() / f"req_{ts_ms}_{os.getpid()}_done.pkl"
        write_pickle(req_path, payload)
        logger.info("[DT_YMD] done UI request: %s", req_path)
    except Exception as exc:
        logger.warning("[DT_YMD] done UI request failed: %s", exc)


def _submit_warning_ui(parent_hwnd: int, sheet_id: str, message: str, title: str = "日付変換") -> None:
    """
    ワーニング通知をモーダルで表示するため ui_server に依頼する。
    SCREENS.WARNING の設定に従い、Warning アイコン・メッセージ・OK ボタンで閉じる。
    """
    try:
        from svc.svc_host import ensure_ui_server

        ensure_ui_server()
        ts_ms = int(time.time() * 1000)
        res_dir = Path(get_ipc_root()) / "result"
        res_dir.mkdir(parents=True, exist_ok=True)
        result_path = str(res_dir / f"res_dt_ymd_warning_{ts_ms}_{os.getpid()}.pkl")
        req_dict: dict[str, Any] = {
            "action": "dt_ymd_warning",
            "modeless": False,
            "title": str(title),
            "message": str(message),
        }
        er_w = _get_window_rect(int(parent_hwnd or 0))
        if er_w is not None:
            req_dict["excel_rect"] = list(er_w)
        payload = {
            "parent_hwnd": int(parent_hwnd),
            "result_path": result_path,
            "ready_path": "",
            "sheet_id": str(sheet_id),
            "log_path": "",
            "action": "dt_ymd",
            "module": "ui_qt.ui_dt_ymd",
            "req_dict": req_dict,
        }
        req_path = get_request_dir() / f"req_{ts_ms}_{os.getpid()}_warning.pkl"
        write_pickle(req_path, payload)
        logger.info("[DT_YMD] warning UI request: %s", req_path)
    except Exception as exc:
        logger.warning("[DT_YMD] warning UI request failed: %s", exc)


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


def _normalize_date_text(v: Any) -> Any:
    """
    日付文字列の軽い正規化を行い、to_datetime で解釈しやすい形に寄せる。
    変換不能値は後段で元値を保持するため、ここでは文字整形のみを担う。
    """
    if v is None:
        return v
    try:
        if pd.isna(v):
            return v
    except Exception:
        pass

    s = str(v)
    if not s:
        return s

    s = unicodedata.normalize("NFKC", s)
    s = s.replace("\u3000", " ").strip()
    s = re.sub(r"\s+", " ", s)
    s = re.sub(r"[（(][^）)]*[）)]", "", s)
    s = re.sub(r"\s*年\s*", "/", s)
    s = re.sub(r"\s*月\s*", "/", s)
    s = re.sub(r"\s*日\s*", "", s)
    s = s.replace(".", "/").replace("-", "/")
    s = re.sub(r"/{2,}", "/", s)
    return s.strip()


def _parse_datetime_with_normalized_fallback(ser_col: pd.Series) -> pd.Series:
    """
    既存の to_datetime 判定を優先し、失敗分のみ正規化文字列で再判定する。
    """
    ser_dt = pd.to_datetime(ser_col, errors="coerce")
    mask_failed = ser_dt.isna()
    if bool(mask_failed.any()):
        ser_norm = ser_col.map(_normalize_date_text)
        ser_dt_norm = pd.to_datetime(ser_norm, errors="coerce")
        ser_dt = ser_dt.where(~mask_failed, ser_dt_norm)
    return ser_dt


def _read_sheet_matrix(
    ptr_s: Any,
    y1: int,
    x1: int,
    yn: int,
    xn: int,
    on_pct: Callable[[int, str, str], None],
    cfg: dict[str, Any],
) -> Optional[list[list[Any]]]:
    """
    シートの指定範囲 (y1,x1) から yn×xn をチャンク単位で読み、2 次元リストで返す。
    読込中は on_pct で進捗コールバックを呼ぶ。失敗時は None。
    """
    msg_read = _msg(cfg, "PHASE_READ")
    custom = (cfg.get("MESSAGES") or {}).get("PROGRESS_CUSTOM_READ") or "読込中"
    chunk_rows = max(200, min(5000, yn))
    acc: list[list[Any]] = []
    try:
        for r0 in range(0, yn, chunk_rows):
            r1 = min(r0 + chunk_rows, yn)
            pct = int(5 + (r1 / max(yn, 1)) * 35)
            on_pct(pct, msg_read, custom)
            rng = ptr_s.range((y1 + r0, x1), (y1 + r1 - 1, x1 + xn - 1))
            part = rng.value
            sub = _normalize_2d(part, r1 - r0, xn)
            acc.extend(sub)
        return acc if len(acc) == yn else None
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
    return f"dt_ymd_{abs(id(ptr_s))}"


def convert_date_ymd(target_hwnd: Optional[int] = None, sheet_id: str = "") -> None:
    """
    【概要】
        指定 HWND の Excel ブック・シートにおいて、選択範囲の日付を YYYY/MM/DD 形式へ一括変換する。
    【補足】
        進捗・完了通知は ui_server 経由で ui_dt_ymd に依頼する。設定は config/ui_dt_ymd.json。
    """
    t_flow = time.perf_counter()
    _perf_dt_ymd("enter", t_flow)
    _dt_ymd_trace("enter", t_flow)

    if core_xlc_mod is None:
        logger.error("[DT_YMD] core_xlc not available")
        _perf_dt_ymd("abort_no_core_xlc", t_flow)
        _dt_ymd_trace("abort_no_core_xlc", t_flow)
        return
    ctx = core_xlc_mod.get_excel_context_from_hwnd(int(target_hwnd or 0), sheet_id)
    if ctx is None:
        logger.error("[DT_YMD] Excel context not available (xlwings + HWND)")
        _perf_dt_ymd("abort_no_context", t_flow)
        _dt_ymd_trace("abort_no_context", t_flow)
        return

    ptr_a, ptr_w, ptr_s, ph = ctx
    logger.info("[DT_YMD] 開始")
    _perf_dt_ymd("after_context", t_flow, hwnd=ph)
    _dt_ymd_trace("after_context", t_flow, hwnd=ph)
    cfg = _cfg()
    saved_status = _status_bar_save(ptr_w)  # 終了時に必ず復元
    sid = _sheet_id_resolve(ptr_s, sheet_id)
    prog_path = _progress_path(sid)
    progress_closed_path = _progress_closed_ack_path(sid)
    progress_close_waited = False
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

    try:
        # 選択範囲を取得。無い場合はステータスバー＋完了ダイアログで終了
        ptr_sel = ptr_a.selection
        if ptr_sel is None:
            logger.info("[DT_YMD] 選択範囲なし")
            _perf_dt_ymd("early_no_selection", t_flow)
            _dt_ymd_trace("early_no_selection", t_flow)
            _status_bar_set(ptr_w, _msg(cfg, "NO_SELECTION"))
            done_cfg = (cfg.get("SCREENS") or {}).get("DONE") or {}
            _submit_done_ui(ph, sid, _msg(cfg, "NO_SELECTION"), str(done_cfg.get("TITLE") or "日付変換"))
            return

        areas: list[tuple[int, int, int, int]] = []
        try:
            api_areas = ptr_sel.api.Areas
            n_areas = int(api_areas.Count)
            for i in range(1, n_areas + 1):
                ar = api_areas.Item(i)
                y1_i = int(ar.Row)
                x1_i = int(ar.Column)
                yn_i = int(ar.Rows.Count)
                xn_i = int(ar.Columns.Count)
                if yn_i >= 1 and xn_i >= 1:
                    areas.append((y1_i, x1_i, yn_i, xn_i))
        except Exception:
            pass
        if not areas:
            val_y1 = int(ptr_sel.row)
            val_x1 = int(ptr_sel.column)
            val_yn = int(ptr_sel.rows.count)
            val_xn = int(ptr_sel.columns.count)
            if val_yn >= 1 and val_xn >= 1:
                areas.append((val_y1, val_x1, val_yn, val_xn))
        total_sel_rows = int(sum(a[2] for a in areas))
        total_sel_cols = int(sum(a[3] for a in areas))
        logger.info(
            "[DT_YMD] 選択範囲 areas=%s rows=%s cols=%s",
            len(areas),
            total_sel_rows,
            total_sel_cols,
        )
        _perf_dt_ymd("after_selection", t_flow, areas=len(areas), rows=total_sel_rows, cols=total_sel_cols)
        _dt_ymd_trace("after_selection", t_flow, areas=len(areas), rows=total_sel_rows, cols=total_sel_cols)
        if not areas:
            _perf_dt_ymd("early_empty_selection", t_flow)
            _dt_ymd_trace("early_empty_selection", t_flow)
            _status_bar_set(ptr_w, _msg(cfg, "NO_SELECTION"))
            done_cfg = (cfg.get("SCREENS") or {}).get("DONE") or {}
            _submit_done_ui(ph, sid, _msg(cfg, "NO_SELECTION"), str(done_cfg.get("TITLE") or "日付変換"))
            return

        # 進捗画面を表示し、UI サーバに進捗表示を依頼（3 フェーズ: 読込・解析・書込）
        _progress_write(
            prog_path,
            {
                "status": "RUN",
                "phase": _msg(cfg, "PHASE_READ"),
                "pct": 0,
                "done": 0,
                "total": 100,
                "seq": 0,
            },
        )
        try:
            progress_closed_path.unlink(missing_ok=True)
        except Exception:
            pass
        _submit_progress_ui(ph, sid, prog_path, 3, progress_closed_path=progress_closed_path)
        _perf_dt_ymd("after_progress_ui_submit", t_flow)
        _dt_ymd_trace("after_progress_ui_submit", t_flow)

        val_success_n = 0
        non_empty_count = 0
        processed_rows = 0
        converted_areas: list[tuple[int, int, list[list[object]]]] = []
        for a_idx, (val_y1, val_x1, val_yn, val_xn) in enumerate(areas, start=1):
            pct_read = int(5 + ((a_idx - 1) / max(1, len(areas))) * 30)
            _upd(pct_read, _msg(cfg, "PHASE_READ"), f"area {a_idx}/{len(areas)}")
            arr = _read_sheet_matrix(ptr_s, val_y1, val_x1, val_yn, val_xn, _upd, cfg)
            if arr is None:
                logger.warning("[DT_YMD] 読込失敗 area=%s", a_idx)
                _perf_dt_ymd("abort_matrix_read_failed", t_flow, area=a_idx)
                _dt_ymd_trace("abort_matrix_read_failed", t_flow, area=a_idx)
                _progress_write(prog_path, {"status": "DONE", "seq": 999})
                if not progress_close_waited:
                    _wait_progress_closed_ack(progress_closed_path)
                    progress_close_waited = True
                _status_bar_set(ptr_w, _msg(cfg, "ERROR_PREFIX"))
                done_cfg = (cfg.get("SCREENS") or {}).get("DONE") or {}
                _submit_done_ui(ph, sid, _msg(cfg, "ERROR_PREFIX"), str(done_cfg.get("TITLE") or "日付変換"))
                return
            _upd(
                int(35 + (a_idx / max(1, len(areas))) * 15),
                _msg(cfg, "PHASE_ANALYZE"),
                (cfg.get("MESSAGES") or {}).get("PROGRESS_CUSTOM_ANALYZE") or "演算中",
            )
            df_worker = pd.DataFrame(arr)
            for j_idx in range(val_xn):
                ser_col = df_worker.iloc[:, j_idx]
                non_empty_count += int((ser_col.notna() & (ser_col.astype(str).str.strip() != "")).sum())
            for j_idx in range(val_xn):
                ser_col = df_worker.iloc[:, j_idx]
                ser_dt = _parse_datetime_with_normalized_fallback(ser_col)
                ser_fmt = ser_dt.dt.strftime("%Y/%m/%d")
                ser_final = ser_fmt.fillna(ser_col)
                df_worker.iloc[:, j_idx] = ser_final
                ser_mask = ser_dt.notnull()
                val_success_n += int(ser_mask.sum())
            converted_areas.append((val_y1, val_x1, df_worker.values.tolist()))
            processed_rows += val_yn

        _perf_dt_ymd("after_matrix_read", t_flow, rows=processed_rows, cols=total_sel_cols)
        _dt_ymd_trace("after_matrix_read", t_flow, rows=processed_rows, cols=total_sel_cols)
        _perf_dt_ymd("after_analyze", t_flow, success_n=val_success_n, non_empty=non_empty_count)
        _dt_ymd_trace("after_analyze", t_flow, success_n=val_success_n, non_empty=non_empty_count)

        # 日付形式チェック: 空でないセルが1つ以上あるが、いずれも日付でない場合はワーニングのみで変換しない
        if non_empty_count >= 1 and val_success_n == 0:
            _progress_write(prog_path, {"status": "DONE", "seq": 999})
            if not progress_close_waited:
                _wait_progress_closed_ack(progress_closed_path)
                progress_close_waited = True
            _status_bar_set(ptr_w, _msg(cfg, "WARNING_NOT_DATE"))
            done_cfg = (cfg.get("SCREENS") or {}).get("DONE") or {}
            _submit_warning_ui(ph, sid, _msg(cfg, "WARNING_NOT_DATE"), str(done_cfg.get("TITLE") or "日付変換"))
            logger.warning("[DT_YMD] 運用ログ 日付形式なし 走査=%s 非日付セルのみ", processed_rows)
            _perf_dt_ymd("early_warning_not_date", t_flow)
            _dt_ymd_trace("early_warning_not_date", t_flow)
            return

        # 書込フェーズ: Interactive 停止、write_chunk で Excel に反映（進捗 50〜100%）
        _upd(50, _msg(cfg, "PHASE_WRITE"), (cfg.get("MESSAGES") or {}).get("PROGRESS_CUSTOM_WRITE") or "書込中")
        try:
            from svc.svc_undo import save_undo_snapshot

            save_undo_snapshot(ptr_w, sheet_id=sheet_id, target_hwnd=ph, excel_hwnd=ph)
        except Exception as e:
            logger.warning("[DT_YMD] save_undo_snapshot failed (undo unavailable): %s", e)
        ptr_a.api.Interactive = False
        try:
            from core import core_xlc

            total_rows_all = max(1, sum(len(a[2]) for a in converted_areas))
            written_rows = 0
            for y1_i, x1_i, arr_i in converted_areas:
                total_rows_i = len(arr_i)

                def _write_progress_cb(done: int) -> None:
                    if total_rows_i > 0:
                        seq[0] += 1
                        pct = int(50 + ((written_rows + done) / total_rows_all) * 50)
                        _progress_write(
                            prog_path,
                            {
                                "status": "RUN",
                                "phase": _msg(cfg, "PHASE_WRITE"),
                                "pct": min(100, pct),
                                "done": pct,
                                "total": 100,
                                "seq": seq[0],
                            },
                        )

                core_xlc.write_chunk(
                    ptr_s,
                    y1_i,
                    x1_i,
                    arr_i,
                    progress_cb=_write_progress_cb,
                )
                written_rows += total_rows_i
        finally:
            try:
                ptr_a.api.Interactive = True
            except Exception:
                pass

        _perf_dt_ymd("after_write_chunk", t_flow, success_n=val_success_n)
        _dt_ymd_trace("after_write_chunk", t_flow, success_n=val_success_n)

        _progress_write(prog_path, {"status": "DONE", "seq": 999})
        if not progress_close_waited:
            _wait_progress_closed_ack(progress_closed_path)
            progress_close_waited = True
        logger.info("[DT_YMD] 完了 走査=%s 変換件数=%s", processed_rows, val_success_n)
        logger.info("[DT_YMD] 運用ログ 日付変換 走査=%s 変換件数=%s", processed_rows, val_success_n)
        _status_bar_set(ptr_w, _msg(cfg, "STATUS_DONE", scanned=processed_rows, count=val_success_n))
        done_cfg = (cfg.get("SCREENS") or {}).get("DONE") or {}
        _submit_done_ui(
            ph,
            sid,
            _msg(cfg, "STATUS_DONE", scanned=processed_rows, count=val_success_n),
            str(done_cfg.get("TITLE") or "日付変換"),
        )
        _perf_dt_ymd("after_done_ui", t_flow, success_n=val_success_n)
        _dt_ymd_trace("after_done_ui", t_flow, success_n=val_success_n)

    except Exception as ex:
        logger.exception("[DT_YMD] %s", ex)
        try:
            _status_bar_set(ptr_w, f"{_msg(cfg, 'ERROR_PREFIX')}: {ex}")
        except Exception:
            pass
    finally:
        try:
            ptr_a.api.Interactive = True
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
        _perf_dt_ymd("flow_end", t_flow)
        _dt_ymd_trace("flow_end", t_flow)
