# -*- coding: utf-8 -*-
"""一括実行の compute フェーズ（COM なし）。完了後 spill し svc_server へ batch_write を依頼する。"""
from __future__ import annotations

import os
import pickle
import re
import time
from pathlib import Path
from typing import Any

from core.core_log import get_data_agg_diag_logger, get_logger
from svc.data_agg_batch_spill import batch_spill_dir, cleanup_batch_spill, write_batch_spill
from svc.data_agg_cancel import (
    DataAggCancelled,
    batch_cancel_scope,
    cancel_request_path_data_agg_batch,
    clear_batch_worker_pid,
    log_cancel_detected,
    make_cancel_check,
    register_batch_worker_pid,
    reset_cancel_path,
)

logger = get_logger(__name__)
_agg_diag = get_data_agg_diag_logger()


def _submit_batch_write_svc_request(
    parent_hwnd: int,
    sheet_id: str,
    payload: dict[str, Any],
) -> None:
    """compute ワーカーから svc_server へ batch_write を非同期投入する。"""
    from svc.svc_host import ensure_svc_server  # noqa: WPS433

    from svc.svc_data_agg import _require_ipc_root  # noqa: WPS433

    ensure_svc_server()
    ipc_root = _require_ipc_root()
    req_dir = ipc_root / "svc_requests"
    req_dir.mkdir(parents=True, exist_ok=True)
    ts_ms = int(time.time() * 1000)
    req_path = req_dir / ("svc_req_%s_%s.pkl" % (ts_ms, os.getpid()))
    req = {
        "action": "data_agg",
        "args": [],
        "kwargs": {
            "excel_hwnd": int(parent_hwnd or 0),
            "sheet_id": str(sheet_id or ""),
            "payload": payload,
        },
    }
    req_path.write_bytes(pickle.dumps(req, protocol=pickle.HIGHEST_PROTOCOL))
    logger.info(
        "[DATA_AGG] batch_write svc_req submitted sheet_id=%s spill=%s req=%s",
        sheet_id,
        payload.get("spill_dir"),
        req_path.name,
    )


def _dispatch_batch_write(
    parent_hwnd: int,
    sheet_id: str,
    *,
    spill_dir: Path,
    batch_run_id: str,
    notify_parent: bool,
    prog_path: str,
    cancel_path: str,
    meta_extra: dict[str, Any] | None = None,
) -> None:
    payload: dict[str, Any] = {
        "action": "batch_write",
        "spill_dir": str(spill_dir),
        "batch_run_id": str(batch_run_id or ""),
        "notify_parent_dialog": bool(notify_parent),
        "prog_path": str(prog_path or ""),
        "cancel_request_path": str(cancel_path or ""),
    }
    if meta_extra:
        payload.update(meta_extra)
    _submit_batch_write_svc_request(parent_hwnd, sheet_id, payload)


def run_batch_compute(parent_hwnd: int, sheet_id: str, payload: dict[str, Any]) -> None:
    """
    short_runner 上で実行: scan + compute + sort + spill → svc batch_write。
    register_batch_worker_pid は自 PID（compute）のみ。Excel COM は触らない。
    """
    from svc import svc_data_agg_scenario as scenario_mod  # noqa: WPS433
    from svc import svc_data_agg_scan as scan_mod  # noqa: WPS433
    from svc import svc_data_agg_write as write_mod  # noqa: WPS433
    from svc.svc_data_agg import (  # noqa: WPS433
        _batch_hook_progress_lines,
        _batch_hook_resolve_current_file,
        _batch_progress_pct_from_hook,
        _clear_active_batch_run_if_current,
        _excel_options_log_summary,
        _get_config,
        _read_active_batch_run_id,
        _require_ipc_root,
        _submit_progress_ui,
        compute_batch_table_rows,
        write_pickle,
    )
    from svc.data_agg_cancel import batch_cancel_tombstone_blocks  # noqa: WPS433

    t_batch_wall = time.perf_counter()
    batch_start_ts_ms = int(time.time() * 1000)
    notify_parent = bool(payload.get("notify_parent_dialog", False))
    batch_run_id = str(payload.get("batch_run_id") or "").strip()

    def _dlog(msg: str, *args: Any) -> None:
        try:
            _agg_diag.info("[DATA_AGG_DIAG] batch_compute " + msg, *args)
        except Exception:
            pass

    ipc_root: Path | None = None
    try:
        ipc_root = _require_ipc_root()
    except Exception:
        ipc_root = None

    def _finish_compute_only(msg: str, *, ok: bool, spill_path: Path | None = None) -> None:
        if ipc_root is not None and batch_run_id:
            try:
                _clear_active_batch_run_if_current(sheet_id, ipc_root, batch_run_id)
            except Exception:
                pass
        if spill_path is not None:
            try:
                cleanup_batch_spill(spill_path)
            except Exception:
                pass

    scenario_path_user = str(payload.get("scenario_path") or "").strip()
    scenario_snapshot_path = str(payload.get("scenario_snapshot_path") or "").strip()

    if ipc_root is not None and batch_run_id:
        try:
            if batch_cancel_tombstone_blocks(sheet_id, ipc_root, batch_run_id):
                _dlog("stale_skip run_id=%s reason=cancel_tombstone", batch_run_id)
                if scenario_snapshot_path:
                    try:
                        Path(scenario_snapshot_path).unlink(missing_ok=True)
                    except OSError:
                        pass
                return
        except Exception:
            pass

    if ipc_root is not None and batch_run_id:
        active_run_id = _read_active_batch_run_id(sheet_id, ipc_root)
        if active_run_id and active_run_id != batch_run_id:
            _dlog("stale_skip run_id=%s active=%s", batch_run_id, active_run_id)
            if scenario_snapshot_path:
                try:
                    Path(scenario_snapshot_path).unlink(missing_ok=True)
                except OSError:
                    pass
            return

    load_path = scenario_snapshot_path or scenario_path_user
    if scenario_snapshot_path and not Path(scenario_snapshot_path).is_file():
        if batch_run_id:
            _dlog("stale_skip reason=missing_snapshot")
            try:
                if ipc_root is not None and batch_run_id:
                    _clear_active_batch_run_if_current(sheet_id, ipc_root, batch_run_id)
            except Exception:
                pass
            return
        if scenario_path_user:
            load_path = scenario_path_user
        else:
            return

    if not load_path:
        _dlog("abort reason=no_scenario_path")
        return

    scenario_path_log = scenario_path_user or scenario_snapshot_path
    scenario_id_fallback = str(Path(load_path).stem)
    _dlog(
        "enter hwnd=%s sheet_id=%s load_path=%s",
        parent_hwnd,
        sheet_id,
        load_path,
    )

    try:
        data = scenario_mod.load_scenario(load_path)
    except Exception as e:
        logger.warning("[DATA_AGG] シナリオ読込失敗: %s", e)
        _tms = int((time.perf_counter() - t_batch_wall) * 1000)
        if ipc_root is not None and batch_run_id:
            spill = batch_spill_dir(ipc_root, sheet_id, batch_run_id)
            write_batch_spill(
                spill,
                [],
                [],
                {
                    "abort": True,
                    "abort_phase": "load",
                    "error": str(e),
                    "user_msg": "シナリオの読込に失敗しました: %s" % e,
                    "scenario_id": scenario_id_fallback,
                    "scenario_path_log": scenario_path_log,
                    "batch_run_id": batch_run_id,
                    "batch_start_ts_ms": batch_start_ts_ms,
                    "compute_ms": _tms,
                    "notify_parent": notify_parent,
                },
            )
            try:
                clear_batch_worker_pid(sheet_id, ipc_root)
            except Exception:
                pass
            _dispatch_batch_write(
                parent_hwnd,
                sheet_id,
                spill_dir=spill,
                batch_run_id=batch_run_id,
                notify_parent=notify_parent,
                prog_path="",
                cancel_path="",
            )
        return
    finally:
        if scenario_snapshot_path:
            try:
                Path(scenario_snapshot_path).unlink(missing_ok=True)
            except OSError:
                pass

    errs = scenario_mod.validate_scenario(data)
    stem_user = Path(scenario_path_user).stem if scenario_path_user.strip() else ""
    id_in_json = str(data.get("id") or "").strip()
    scenario_id = (stem_user or id_in_json or scenario_id_fallback).strip() or scenario_id_fallback

    if errs:
        _tms = int((time.perf_counter() - t_batch_wall) * 1000)
        if ipc_root is not None and batch_run_id:
            spill = batch_spill_dir(ipc_root, sheet_id, batch_run_id)
            write_batch_spill(
                spill,
                [],
                [],
                {
                    "abort": True,
                    "abort_phase": "validate",
                    "error": "; ".join(str(x) for x in errs[:5]),
                    "user_msg": "シナリオの検証エラー:\n" + "\n".join(errs[:5]),
                    "scenario_id": scenario_id,
                    "scenario_path_log": scenario_path_log,
                    "excel_write_summary": _excel_options_log_summary(data.get("excel_options")),
                    "batch_run_id": batch_run_id,
                    "batch_start_ts_ms": batch_start_ts_ms,
                    "compute_ms": _tms,
                    "notify_parent": notify_parent,
                },
            )
            try:
                clear_batch_worker_pid(sheet_id, ipc_root)
            except Exception:
                pass
            _dispatch_batch_write(
                parent_hwnd,
                sheet_id,
                spill_dir=spill,
                batch_run_id=batch_run_id,
                notify_parent=notify_parent,
                prog_path="",
                cancel_path="",
            )
        return

    items = data.get("items") or []
    if not items:
        _tms = int((time.perf_counter() - t_batch_wall) * 1000)
        if ipc_root is not None and batch_run_id:
            spill = batch_spill_dir(ipc_root, sheet_id, batch_run_id)
            write_batch_spill(
                spill,
                [],
                [],
                {
                    "abort": True,
                    "abort_phase": "no_items",
                    "error": "no_items",
                    "user_msg": "項目が定義されていません。",
                    "scenario_id": scenario_id,
                    "scenario_path_log": scenario_path_log,
                    "excel_write_summary": _excel_options_log_summary(data.get("excel_options")),
                    "batch_run_id": batch_run_id,
                    "batch_start_ts_ms": batch_start_ts_ms,
                    "compute_ms": _tms,
                    "notify_parent": notify_parent,
                },
            )
            try:
                clear_batch_worker_pid(sheet_id, ipc_root)
            except Exception:
                pass
            _dispatch_batch_write(
                parent_hwnd,
                sheet_id,
                spill_dir=spill,
                batch_run_id=batch_run_id,
                notify_parent=notify_parent,
                prog_path="",
                cancel_path="",
            )
        return

    if ipc_root is None:
        logger.error("[DATA_AGG] batch_compute IPC root unavailable")
        return

    cancel_path = cancel_request_path_data_agg_batch(sheet_id, ipc_root)
    reset_cancel_path(cancel_path)
    cancel_check = make_cancel_check(cancel_path, min_interval_sec=0.0)
    register_batch_worker_pid(sheet_id, ipc_root)

    prog_path = ipc_root / "progress" / (
        "data_agg_batch_%s_%s.pkl" % (os.getpid(), int(time.time() * 1000))
    )
    try:
        prog_path.unlink(missing_ok=True)
    except OSError:
        pass
    prog_seq = [0]
    prog_last_pct = [2]
    cfg_msgs = (_get_config().get("MESSAGES") or {})

    def _prog_write(**kw: Any) -> None:
        if write_pickle is None:
            return
        prog_seq[0] += 1
        phase = str(kw.get("phase", "") or "")
        d: dict[str, Any] = {
            "status": str(kw.get("status", "RUN")),
            "seq": prog_seq[0],
            "pct": int(max(0, min(100, int(kw.get("pct", 5) or 5)))),
            "phase": phase,
            "phase_i": int(kw.get("phase_i", 0) or 0),
            "phase_total": 4,
            "msg": phase,
            "show_done_dialog": False,
        }
        if kw.get("done") is not None:
            d["done"] = kw["done"]
        if kw.get("total") is not None:
            d["total"] = kw["total"]
        cf = str(kw.get("current_file", "") or "").strip()
        if cf:
            d["current_file"] = cf
        dt = str(kw.get("detail", "") or "").strip()
        if dt:
            d["detail"] = dt
        try:
            prog_path.parent.mkdir(parents=True, exist_ok=True)
            write_pickle(prog_path, d)
        except Exception:
            pass

    def _prog_cancel() -> None:
        _prog_write(
            status="CANCEL",
            pct=max(prog_last_pct[0], 5),
            phase="中止",
            phase_i=4,
            done=prog_last_pct[0],
            total=100,
        )

    def _submit_abort_write(
        *,
        phase: str,
        files_n: int = 0,
        compute_ms: int | None = None,
        event_log_rows: list[list[Any]] | None = None,
        error: str = "cancelled",
        user_msg: str = "",
    ) -> None:
        _tms = int((time.perf_counter() - t_batch_wall) * 1000)
        spill = batch_spill_dir(ipc_root, sheet_id, batch_run_id)
        msg = user_msg or str(cfg_msgs.get("STATUS_CANCEL") or "一括実行を中止しました。").strip()
        write_batch_spill(
            spill,
            [],
            [],
            {
                "abort": True,
                "abort_phase": phase,
                "error": error,
                "user_msg": msg,
                "scenario_id": scenario_id,
                "scenario_path_log": scenario_path_log,
                "files_n": int(files_n or 0),
                "compute_ms": compute_ms if compute_ms is not None else _tms,
                "event_log_rows": event_log_rows or [],
                "excel_write_summary": _excel_options_log_summary(data.get("excel_options")),
                "batch_run_id": batch_run_id,
                "batch_start_ts_ms": batch_start_ts_ms,
                "notify_parent": notify_parent,
            },
        )
        _prog_cancel()
        try:
            clear_batch_worker_pid(sheet_id, ipc_root)
        except Exception:
            pass
        _dispatch_batch_write(
            parent_hwnd,
            sheet_id,
            spill_dir=spill,
            batch_run_id=batch_run_id,
            notify_parent=notify_parent,
            prog_path=str(prog_path),
            cancel_path=str(cancel_path),
        )

    scan_cfg = data.get("scan") or {}
    start_path = scan_cfg.get("start_path") or "."
    ext_t = tuple(scan_cfg.get("extensions") or [".xlsx", ".xlsm", ".csv"])
    kw = scan_cfg.get("keyword") or ""
    rec = bool(scan_cfg.get("recursive"))

    file_paths_holder: list[list[Path]] = [[]]

    def _batch_hook(sub: int, suffix: str, *rest: Any) -> None:
        fps = file_paths_holder[0]
        nf_l = max(len(fps), 1)
        ni_l = max(len(items), 1)
        fi_kw = int(rest[0]) if len(rest) >= 1 else None
        nf_kw = int(rest[1]) if len(rest) >= 2 else None
        if cancel_check is not None:
            cancel_check(force=True)
        raw = _batch_progress_pct_from_hook(
            sub,
            suffix,
            nf_l,
            ni_l,
            file_index=fi_kw,
            n_files_total=nf_kw,
        )
        prog_last_pct[0] = max(prog_last_pct[0], min(92, raw))
        pi = min(4, max(1, int(sub) - 3))
        phase_txt, detail_txt = _batch_hook_progress_lines(sub, suffix)
        cf = _batch_hook_resolve_current_file(str(suffix or ""), fi_kw, fps)
        prog_kw: dict[str, Any] = {
            "pct": prog_last_pct[0],
            "phase": phase_txt[:120],
            "phase_i": pi,
            "done": prog_last_pct[0],
            "total": 100,
        }
        if detail_txt:
            prog_kw["detail"] = detail_txt
        elif cf:
            prog_kw["current_file"] = cf
        _prog_write(**prog_kw)

    _prog_write(
        pct=2,
        phase=str(cfg_msgs.get("PHASE_SCAN") or "フォルダを走査中..."),
        phase_i=0,
        done=0,
        total=100,
    )
    _submit_progress_ui(
        parent_hwnd,
        sheet_id,
        str(prog_path),
        phase_total=4,
        extra_req={
            "done_delay_ms": 450,
            "progress_poll_ms": 90,
            "cancel_request_path": str(cancel_path),
            "data_agg_batch_notify_parent": notify_parent,
            "data_agg_batch_scenario_id": str(scenario_id),
            "data_agg_batch_scenario_path": str(scenario_path_log),
        },
    )

    with batch_cancel_scope(cancel_check):
        try:
            file_paths = scan_mod.scan_folder(
                start_path,
                recursive=rec,
                extensions=ext_t,
                keyword=kw,
                cancel_check=cancel_check,
            )
        except DataAggCancelled:
            log_cancel_detected(
                sheet_id=sheet_id,
                phase="scan",
                ipc_root=ipc_root,
            )
            _submit_abort_write(phase="scan")
            return
        file_paths_holder[0] = list(file_paths)

    if not file_paths:
        _tms = int((time.perf_counter() - t_batch_wall) * 1000)
        spill = batch_spill_dir(ipc_root, sheet_id, batch_run_id)
        write_batch_spill(
            spill,
            [],
            [],
            {
                "abort": True,
                "abort_phase": "zero_files",
                "error": "zero_files_after_scan",
                "user_msg": "対象ファイルが 0 件でした。",
                "scenario_id": scenario_id,
                "scenario_path_log": scenario_path_log,
                "files_n": 0,
                "compute_ms": _tms,
                "excel_write_summary": _excel_options_log_summary(data.get("excel_options")),
                "batch_run_id": batch_run_id,
                "batch_start_ts_ms": batch_start_ts_ms,
                "notify_parent": notify_parent,
            },
        )
        try:
            clear_batch_worker_pid(sheet_id, ipc_root)
        except Exception:
            pass
        _dispatch_batch_write(
            parent_hwnd,
            sheet_id,
            spill_dir=spill,
            batch_run_id=batch_run_id,
            notify_parent=notify_parent,
            prog_path=str(prog_path),
            cancel_path=str(cancel_path),
        )
        return

    data = dict(data)
    data["id"] = scenario_id
    event_log_rows: list[list[Any]] = []
    t_compute = time.perf_counter()
    try:
        with batch_cancel_scope(cancel_check):
            headers, table_rows, event_log_rows, join_events_total = compute_batch_table_rows(
                data,
                file_paths,
                probe_caller="excel_batch_submit",
                progress_hook=_batch_hook,
                cancel_check=cancel_check,
            )
    except DataAggCancelled:
        dt_compute_ms = int((time.perf_counter() - t_compute) * 1000)
        log_cancel_detected(
            sheet_id=sheet_id,
            phase="compute",
            files_n=len(file_paths),
            ipc_root=ipc_root,
        )
        _submit_abort_write(
            phase="compute",
            files_n=len(file_paths),
            compute_ms=dt_compute_ms,
            event_log_rows=event_log_rows,
        )
        return
    except Exception as e:
        logger.exception("[DATA_AGG] compute_batch_table_rows failed: %s", e)
        dt_compute_ms = int((time.perf_counter() - t_compute) * 1000)
        _tms = int((time.perf_counter() - t_batch_wall) * 1000)
        spill = batch_spill_dir(ipc_root, sheet_id, batch_run_id)
        write_batch_spill(
            spill,
            [],
            [],
            {
                "abort": True,
                "abort_phase": "compute",
                "error": str(e),
                "user_msg": "集約計算中にエラーが発生しました: %s" % e,
                "scenario_id": scenario_id,
                "scenario_path_log": scenario_path_log,
                "files_n": len(file_paths),
                "compute_ms": dt_compute_ms,
                "event_log_rows": event_log_rows,
                "excel_write_summary": _excel_options_log_summary(data.get("excel_options")),
                "batch_run_id": batch_run_id,
                "batch_start_ts_ms": batch_start_ts_ms,
                "notify_parent": notify_parent,
            },
        )
        try:
            clear_batch_worker_pid(sheet_id, ipc_root)
        except Exception:
            pass
        _dispatch_batch_write(
            parent_hwnd,
            sheet_id,
            spill_dir=spill,
            batch_run_id=batch_run_id,
            notify_parent=notify_parent,
            prog_path=str(prog_path),
            cancel_path=str(cancel_path),
        )
        return

    dt_compute_ms = int((time.perf_counter() - t_compute) * 1000)
    excel_opts = scenario_mod.normalize_excel_options(data.get("excel_options"))
    table_rows = write_mod.sort_table_rows_for_excel_options(headers, table_rows, excel_opts)

    if cancel_check is not None:
        try:
            cancel_check(force=True)
        except DataAggCancelled:
            log_cancel_detected(
                sheet_id=sheet_id,
                phase="pre_write",
                files_n=len(file_paths),
                ipc_root=ipc_root,
            )
            _submit_abort_write(
                phase="pre_write",
                files_n=len(file_paths),
                compute_ms=dt_compute_ms,
                event_log_rows=event_log_rows,
            )
            return

    from svc.svc_data_agg import _resolve_match_keys_to_headers  # noqa: WPS433

    match_cols = _resolve_match_keys_to_headers(data.get("match_keys") or [], items, headers)
    column_modes: list[str] = []
    for it in items:
        if not isinstance(it, dict):
            column_modes.append("fill_in")
            continue
        lin_i = scenario_mod.infer_item_lineage(it.get("sources") or [])
        if lin_i == "__mixed__":
            lin_i = None
        column_modes.append(
            scenario_mod.normalize_item_write_mode(it.get("write_mode"), lineage=lin_i)
        )

    spill = batch_spill_dir(ipc_root, sheet_id, batch_run_id)
    write_batch_spill(
        spill,
        headers,
        table_rows,
        {
            "abort": False,
            "scenario_id": scenario_id,
            "scenario_path_log": scenario_path_log,
            "files_n": len(file_paths),
            "join_events_total": join_events_total,
            "compute_ms": dt_compute_ms,
            "event_log_rows": event_log_rows,
            "excel_opts": excel_opts,
            "column_modes": column_modes,
            "match_cols": match_cols,
            "items_n": len(items),
            "batch_run_id": batch_run_id,
            "batch_start_ts_ms": batch_start_ts_ms,
            "notify_parent": notify_parent,
            "headers": headers,
        },
    )
    _dlog(
        "compute_ok spill=%s header_count=%s row_count=%s compute_ms=%s",
        spill,
        len(headers),
        len(table_rows),
        dt_compute_ms,
    )
    prog_last_pct[0] = max(prog_last_pct[0], 93)
    _prog_write(
        pct=93,
        phase="マスターへ書き込み",
        phase_i=4,
        done=93,
        total=100,
    )
    try:
        clear_batch_worker_pid(sheet_id, ipc_root)
    except Exception:
        pass
    _dispatch_batch_write(
        parent_hwnd,
        sheet_id,
        spill_dir=spill,
        batch_run_id=batch_run_id,
        notify_parent=notify_parent,
        prog_path=str(prog_path),
        cancel_path=str(cancel_path),
    )
    _finish_compute_only("", ok=True, spill_path=None)
