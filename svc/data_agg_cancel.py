# -*- coding: utf-8 -*-
"""本番データ集約一括の協調キャンセル（IPC pickle フラグ）とワーカー強制終了。"""
from __future__ import annotations

import contextlib
import contextvars
import json
import os
import subprocess
import sys
import time
from concurrent.futures import Future, ThreadPoolExecutor
from pathlib import Path
from typing import Any, Callable, Iterable, Iterator, Optional

_CANCEL_POLL_INTERVAL_SEC = 0.05
_COOP_WAIT_INITIAL_MS = 400
_COOP_WAIT_EXTENDED_MS = 3000
_active_cancel: contextvars.ContextVar[Callable[..., None] | None] = contextvars.ContextVar(
    "data_agg_batch_cancel_check", default=None
)


class DataAggCancelled(Exception):
    """操作者が進捗画面から一括実行を中止した。"""


def cancel_request_path_data_agg_batch(sheet_id: str, ipc_root: Path) -> Path:
    d = ipc_root / "progress"
    d.mkdir(parents=True, exist_ok=True)
    sid = str(sheet_id or "").strip() or "default"
    return d / ("cancel_req_data_agg_batch_%s.pkl" % sid)


def cancel_request_path_data_agg_master_debug(ipc_root: Path, *, token: str = "") -> Path:
    """マスタデバッグ進捗の協調キャンセル（本番一括の強制終了経路とは別名）。"""
    d = ipc_root / "progress"
    d.mkdir(parents=True, exist_ok=True)
    tok = str(token or "").strip()
    if not tok:
        tok = "%s" % int(time.time() * 1000)
    return d / ("cancel_req_data_agg_master_debug_%s.pkl" % tok)


def reset_cancel_path(path: Path) -> None:
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


def cancel_requested(path: Path) -> bool:
    try:
        from ui_qt.ipc_file import read_pickle  # noqa: WPS433

        d = read_pickle(path)
        return isinstance(d, dict) and bool(d.get("cancel"))
    except Exception:
        return False


def make_cancel_check(
    path: Path | None,
    *,
    min_interval_sec: float = _CANCEL_POLL_INTERVAL_SEC,
) -> Callable[..., None] | None:
    """
    キャンセル pickle をポールし、要求時は DataAggCancelled を送出する。
    force=True のときは間隔制限なし。min_interval_sec=0 で毎回ポール（本番一括向け）。
    """
    if path is None:
        return None
    last_poll: list[float] = [0.0]
    interval = max(0.0, float(min_interval_sec))

    def _check(*, force: bool = False) -> None:
        now = time.monotonic()
        if not force and interval > 0 and (now - last_poll[0]) < interval:
            return
        last_poll[0] = now
        if cancel_requested(path):
            raise DataAggCancelled()

    return _check


@contextlib.contextmanager
def batch_cancel_scope(
    check: Optional[Callable[..., None]],
) -> Iterator[None]:
    """本番一括の走査〜集約中、深い処理から poll_active_cancel できるようにする。"""
    token = _active_cancel.set(check)
    try:
        yield
    finally:
        _active_cancel.reset(token)


def poll_active_cancel(*, force: bool = False) -> None:
    """スコープ内の cancel_check を実行する。"""
    chk = _active_cancel.get()
    if chk is not None:
        chk(force=force)


def poll_active_cancel_every(index: int, *, stride: int = 16, force: bool = False) -> None:
    if stride < 1:
        stride = 1
    if index % stride != 0:
        return
    poll_active_cancel(force=force)


def abort_pending_futures(
    futures: Iterable[Future[Any]],
    *,
    executor: ThreadPoolExecutor | None = None,
    wait: bool = False,
) -> None:
    """
    協調キャンセル時に未開始 future をキャンセルし、必要なら executor を即時シャットダウンする。
    実行中ワーカーは止められないが、待ち行列への投入は止める（完走時の正しさは不問）。
    """
    for fut in futures:
        try:
            fut.cancel()
        except Exception:
            pass
    if executor is None:
        return
    try:
        executor.shutdown(wait=wait, cancel_futures=True)
    except TypeError:
        try:
            executor.shutdown(wait=wait)
        except Exception:
            pass
    except Exception:
        pass


def batch_coop_cancel_detected_path(sheet_id: str, ipc_root: Path) -> Path:
    d = ipc_root / "progress"
    d.mkdir(parents=True, exist_ok=True)
    sid = str(sheet_id or "").strip() or "default"
    return d / ("data_agg_batch_coop_cancel_%s.pkl" % sid)


def write_batch_coop_cancel_detected(
    sheet_id: str,
    ipc_root: Path,
    *,
    phase: str = "",
    files_n: int = 0,
) -> None:
    """UI 側が協調キャンセル完了を待てるよう、ワーカー検知時刻を IPC に記録する。"""
    try:
        from ui_qt.ipc_file import write_pickle  # noqa: WPS433

        write_pickle(
            batch_coop_cancel_detected_path(sheet_id, ipc_root),
            {
                "sheet_id": str(sheet_id or "").strip(),
                "phase": str(phase or "").strip(),
                "files_n": int(files_n or 0),
                "ts_ms": int(time.time() * 1000),
            },
        )
    except Exception:
        pass


def clear_batch_coop_cancel_detected(sheet_id: str, ipc_root: Path) -> None:
    try:
        batch_coop_cancel_detected_path(sheet_id, ipc_root).unlink(missing_ok=True)  # type: ignore[call-arg]
    except TypeError:
        try:
            p = batch_coop_cancel_detected_path(sheet_id, ipc_root)
            if p.exists():
                p.unlink()
        except Exception:
            pass
    except Exception:
        pass


def coop_cancel_detected(sheet_id: str, ipc_root: Path) -> bool:
    try:
        return batch_coop_cancel_detected_path(sheet_id, ipc_root).is_file()
    except Exception:
        return False


def log_cancel_detected(
    *,
    sheet_id: str,
    phase: str,
    files_n: int = 0,
    ipc_root: Path | None = None,
) -> None:
    """ワーカーが中止フラグを検知したとき（hc_csv.log / hc_csv_diag.log）。"""
    if ipc_root is not None:
        write_batch_coop_cancel_detected(
            sheet_id,
            ipc_root,
            phase=phase,
            files_n=files_n,
        )
    try:
        from core.core_log import get_data_agg_diag_logger, get_logger  # noqa: WPS433

        get_logger(__name__).info(
            "[DATA_AGG] batch cancel detected sheet_id=%s phase=%s files_n=%s",
            sheet_id,
            phase,
            files_n,
        )
        get_data_agg_diag_logger().info(
            "[DATA_AGG_DIAG] batch_run cancel detected sheet_id=%s phase=%s files_n=%s",
            sheet_id,
            phase,
            files_n,
        )
    except Exception:
        pass


def delete_output_sheet_if_any(sheet: Any) -> None:
    """中止時に作成済みの新規出力シートを削除する（失敗は握りつぶす）。"""
    if sheet is None:
        return
    try:
        sheet.delete()
    except Exception:
        pass


def batch_active_path(sheet_id: str, ipc_root: Path) -> Path:
    d = ipc_root / "progress"
    d.mkdir(parents=True, exist_ok=True)
    sid = str(sheet_id or "").strip() or "default"
    return d / ("data_agg_batch_active_%s.pkl" % sid)


def batch_cancel_tombstone_path(sheet_id: str, ipc_root: Path) -> Path:
    d = ipc_root / "progress"
    d.mkdir(parents=True, exist_ok=True)
    sid = str(sheet_id or "").strip() or "default"
    return d / ("data_agg_batch_cancelled_%s.pkl" % sid)


def clear_batch_active_run(sheet_id: str, ipc_root: Path) -> None:
    try:
        batch_active_path(sheet_id, ipc_root).unlink(missing_ok=True)  # type: ignore[call-arg]
    except TypeError:
        try:
            p = batch_active_path(sheet_id, ipc_root)
            if p.exists():
                p.unlink()
        except Exception:
            pass
    except Exception:
        pass


def clear_batch_cancel_tombstone(sheet_id: str, ipc_root: Path) -> None:
    try:
        batch_cancel_tombstone_path(sheet_id, ipc_root).unlink(missing_ok=True)  # type: ignore[call-arg]
    except TypeError:
        try:
            p = batch_cancel_tombstone_path(sheet_id, ipc_root)
            if p.exists():
                p.unlink()
        except Exception:
            pass
    except Exception:
        pass


def write_batch_cancel_tombstone(
    sheet_id: str,
    ipc_root: Path,
    *,
    run_id: str = "",
) -> None:
    """強制中止後の svc_req 再実行を防ぐ tombstone。"""
    try:
        from ui_qt.ipc_file import write_pickle  # noqa: WPS433

        write_pickle(
            batch_cancel_tombstone_path(sheet_id, ipc_root),
            {
                "run_id": str(run_id or "").strip(),
                "ts_ms": int(time.time() * 1000),
            },
        )
    except Exception:
        pass


def batch_cancel_tombstone_blocks(
    sheet_id: str,
    ipc_root: Path,
    run_id: str = "",
) -> bool:
    """中止 tombstone が残っていれば同一 run_id（または run_id 未指定 replay）をブロック。"""
    try:
        from ui_qt.ipc_file import read_pickle  # noqa: WPS433

        d = read_pickle(batch_cancel_tombstone_path(sheet_id, ipc_root))
        if not isinstance(d, dict):
            return False
        t_run = str(d.get("run_id") or "").strip()
        rid = str(run_id or "").strip()
        if t_run and rid:
            return t_run == rid
        if t_run and not rid:
            return True
        return bool(t_run or d.get("ts_ms"))
    except Exception:
        return False


def _svc_req_is_data_agg_batch_for_sheet(req_path: Path, sheet_id: str) -> bool:
    sid = str(sheet_id or "").strip()
    if not sid:
        return False
    try:
        from ui_qt.ipc_file import read_pickle  # noqa: WPS433

        req = read_pickle(req_path)
        if not isinstance(req, dict):
            return False
        if str(req.get("action") or "").strip() != "data_agg":
            return False
        kwargs = req.get("kwargs")
        if not isinstance(kwargs, dict):
            return False
        if str(kwargs.get("sheet_id") or "").strip() != sid:
            return False
        payload = kwargs.get("payload")
        if not isinstance(payload, dict):
            return False
        act = str(payload.get("action") or "").strip()
        return act in ("batch_run", "batch_write")
    except Exception:
        return False


def purge_pending_data_agg_batch_svc_requests(ipc_root: Path, sheet_id: str) -> int:
    """未処理の data_agg batch_run svc_req を削除（svc_server 強制終了時の残骸対策）。"""
    req_dir = ipc_root / "svc_requests"
    if not req_dir.is_dir():
        return 0
    removed = 0
    for p in sorted(req_dir.glob("svc_req_*.pkl")):
        if not _svc_req_is_data_agg_batch_for_sheet(p, sheet_id):
            continue
        try:
            p.unlink(missing_ok=True)  # type: ignore[call-arg]
            removed += 1
        except TypeError:
            try:
                if p.exists():
                    p.unlink()
                    removed += 1
            except Exception:
                pass
        except Exception:
            pass
    return removed


def batch_worker_pid_path(sheet_id: str, ipc_root: Path) -> Path:
    d = ipc_root / "progress"
    d.mkdir(parents=True, exist_ok=True)
    sid = str(sheet_id or "").strip() or "default"
    return d / ("data_agg_batch_worker_%s.pid" % sid)


def register_batch_worker_pid(sheet_id: str, ipc_root: Path) -> None:
    """一括 compute ワーカー（short_runner / svc 子プロセス）の PID を IPC に登録。"""
    p = batch_worker_pid_path(sheet_id, ipc_root)
    try:
        p.write_text(str(os.getpid()), encoding="ascii")
    except Exception:
        pass


def clear_batch_worker_pid(sheet_id: str, ipc_root: Path) -> None:
    try:
        batch_worker_pid_path(sheet_id, ipc_root).unlink(missing_ok=True)  # type: ignore[call-arg]
    except TypeError:
        try:
            p = batch_worker_pid_path(sheet_id, ipc_root)
            if p.exists():
                p.unlink()
        except Exception:
            pass
    except Exception:
        pass
    clear_batch_coop_cancel_detected(sheet_id, ipc_root)


def sheet_id_from_cancel_path(cancel_path: Path | str) -> str:
    name = Path(cancel_path).name
    prefix = "cancel_req_data_agg_batch_"
    if not name.startswith(prefix) or not name.endswith(".pkl"):
        return ""
    return name[len(prefix) : -4]


def read_batch_worker_pid(sheet_id: str, ipc_root: Path) -> int | None:
    p = batch_worker_pid_path(sheet_id, ipc_root)
    if not p.is_file():
        return None
    try:
        pid = int(p.read_text(encoding="ascii").strip())
    except (OSError, ValueError):
        return None
    return pid if pid > 0 else None


def terminate_pid_tree(pid: int) -> bool:
    """Windows: taskkill /F /T。UI 自身の PID は終了しない。"""
    if pid <= 0 or pid == os.getpid():
        return False
    if sys.platform == "win32":
        flags = 0
        if hasattr(subprocess, "CREATE_NO_WINDOW"):
            flags = subprocess.CREATE_NO_WINDOW
        try:
            cp = subprocess.run(
                ["taskkill", "/F", "/T", "/PID", str(pid)],
                capture_output=True,
                creationflags=flags,
                timeout=15.0,
            )
            return cp.returncode == 0
        except Exception:
            return False
    try:
        import signal

        os.kill(pid, signal.SIGTERM)
        return True
    except Exception:
        return False


def terminate_batch_worker(sheet_id: str, ipc_root: Path) -> bool:
    pid = read_batch_worker_pid(sheet_id, ipc_root)
    if pid is None:
        return False
    ok = terminate_pid_tree(pid)
    try:
        clear_batch_worker_pid(sheet_id, ipc_root)
    except Exception:
        pass
    return ok


def wait_batch_worker_exit(sheet_id: str, ipc_root: Path, *, timeout_ms: int) -> bool:
    """PID ファイルが消える（ワーカー自然終了）まで短時間待つ。"""
    to_ms = max(0, int(timeout_ms))
    if to_ms <= 0:
        return read_batch_worker_pid(sheet_id, ipc_root) is None
    t0 = time.monotonic()
    while True:
        if read_batch_worker_pid(sheet_id, ipc_root) is None:
            return True
        if (time.monotonic() - t0) * 1000 >= to_ms:
            return False
        time.sleep(0.05)


def wait_batch_worker_exit_adaptive(
    sheet_id: str,
    ipc_root: Path,
    *,
    initial_ms: int = _COOP_WAIT_INITIAL_MS,
    extended_ms: int = _COOP_WAIT_EXTENDED_MS,
) -> tuple[bool, bool]:
    """
    ワーカー PID 消失を待つ。協調キャンセル検知マーカーがあれば extended_ms まで延長する。

    Returns:
        (exited, coop_detected_during_wait)
    """
    init_ms = max(0, int(initial_ms))
    ext_ms = max(init_ms, int(extended_ms))
    t0 = time.monotonic()
    coop_seen = False
    while True:
        if read_batch_worker_pid(sheet_id, ipc_root) is None:
            return True, coop_seen
        if coop_cancel_detected(sheet_id, ipc_root):
            coop_seen = True
        elapsed_ms = (time.monotonic() - t0) * 1000
        limit_ms = ext_ms if coop_seen else init_ms
        if elapsed_ms >= limit_ms:
            return False, coop_seen
        time.sleep(0.05)


def write_progress_cancel_status(progress_path: Path) -> None:
    """協調中止と併用: 進捗 pickle を即 CANCEL にする（UI タイマー待ちを短くする）。"""
    from ui_qt.ipc_file import read_pickle, write_pickle  # noqa: WPS433

    cur: dict[str, Any] = {}
    try:
        raw = read_pickle(progress_path)
        if isinstance(raw, dict):
            cur = raw
    except Exception:
        pass
    seq = int(cur.get("seq") or 0) + 1000
    write_pickle(
        progress_path,
        {
            "status": "CANCEL",
            "seq": seq,
            "pct": int(cur.get("pct") or 5),
            "phase": "中止",
            "phase_i": 4,
            "phase_total": int(cur.get("phase_total") or 4),
            "msg": "中止",
            "show_done_dialog": False,
            "done": cur.get("done"),
            "total": cur.get("total"),
        },
    )


def _iter_recent_event_rows(book: Any, *, lookback_rows: int = 20) -> list[list[Any]]:
    try:
        from svc import svc_data_agg_write as write_mod  # noqa: WPS433

        ws = write_mod._locate_event_log_sheet(book)  # type: ignore[attr-defined]
        if ws is None:
            return []
        ur = getattr(ws, "used_range", None)
        lc = getattr(ur, "last_cell", None) if ur is not None else None
        last_r = int(getattr(lc, "row", 0) or 0)
        if last_r < 2:
            return []
        start_r = max(2, last_r - max(1, int(lookback_rows)) + 1)
        n_col = len(getattr(write_mod, "EVENT_LOG_HEADERS", []) or []) or 9
        raw = ws.range((start_r, 1), (last_r, n_col)).value
        if raw is None:
            return []
        if isinstance(raw, (list, tuple)) and raw and not isinstance(raw[0], (list, tuple)):
            return [list(raw)]
        if isinstance(raw, (list, tuple)):
            return [list(r) if isinstance(r, (list, tuple)) else [r] for r in raw]
    except Exception:
        return []
    return []


def _event_log_row_kind_sid_path_detail(
    row: list[Any],
) -> tuple[str, str, str, str]:
    """レポート行から区分・シナリオID・対象パス・詳細を取る（列追加前後両対応）。"""
    n = len(row)
    # 新: 記録日時, 処理時間, 出力行数, 区分, 書込み方式, 出力シート名, シナリオID, 対象パス, 詳細
    if n >= 9:
        return (
            str(row[3] or "").strip(),
            str(row[6] or "").strip(),
            str(row[7] or "").strip(),
            str(row[8] or ""),
        )
    # 旧（処理時間あり・出力行数なし）: …, 処理時間, 区分, …, 詳細
    return (
        str(row[2] if n >= 3 else "").strip(),
        str(row[5] if n >= 6 else "").strip(),
        str(row[6] if n >= 7 else "").strip(),
        str(row[7] if n >= 8 else ""),
    )


def _has_recent_cancel_summary(book: Any, *, scenario_id: str, scenario_path: str) -> bool:
    sid = str(scenario_id or "").strip()
    sp = str(scenario_path or "").strip()
    for row in reversed(_iter_recent_event_rows(book)):
        kind, row_sid, row_sp, detail_s = _event_log_row_kind_sid_path_detail(row)
        if kind != "一括実行・中止":
            continue
        if sid and row_sid and sid != row_sid:
            continue
        if sp and row_sp and sp != row_sp:
            continue
        if detail_s:
            try:
                d = json.loads(detail_s)
                if str(d.get("結果") or "").strip() == "中止":
                    return True
            except Exception:
                pass
        return True
    return False


def append_cancel_event_log_from_ui(
    *,
    parent_hwnd: int,
    sheet_id: str,
    scenario_id: str = "",
    scenario_path: str = "",
) -> bool:
    """強制終了時の補完: UI 側からイベントログへ「一括実行・中止」を追記。"""
    if int(parent_hwnd or 0) <= 0:
        return False
    try:
        from core.core_xlc import get_excel_context_from_hwnd  # noqa: WPS433
        from svc import svc_data_agg_write as write_mod  # noqa: WPS433
    except Exception:
        return False
    try:
        ctx = get_excel_context_from_hwnd(int(parent_hwnd), str(sheet_id or ""))
        if not ctx:
            return False
        _app, book, _sheet, _hwnd = ctx
        if _has_recent_cancel_summary(
            book,
            scenario_id=str(scenario_id or ""),
            scenario_path=str(scenario_path or ""),
        ):
            return True
        row = write_mod.format_batch_run_summary_row(
            str(scenario_id or ""),
            str(scenario_path or ""),
            ok=False,
            error="cancelled",
        )
        write_mod.append_event_log_rows(book, [row])
        return True
    except Exception:
        return False


def _resolve_cancel_ipc_root(
    cancel_path: Path | str,
    ipc_root: Path | str | None,
) -> Path:
    if ipc_root:
        return Path(ipc_root)
    try:
        from core import core_env  # noqa: WPS433

        raw = core_env.ipc_dir_raw()
        if raw:
            return Path(raw)
    except Exception:
        pass
    return Path(cancel_path).parent.parent


def run_data_agg_batch_force_terminate_no_com(
    *,
    cancel_path: Path | str,
    progress_path: Path | str | None = None,
    ipc_root: Path | str | None = None,
    notify_parent: bool = False,
    write_progress_cancel: bool = True,
    cooperative_wait_ms: int = _COOP_WAIT_INITIAL_MS,
    cooperative_wait_extended_ms: int = _COOP_WAIT_EXTENDED_MS,
) -> dict[str, Any]:
    """
    一括キャンセルのファイル／プロセス処理（COM なし）。

    UI スレッドを塞がないよう進捗ダイアログからバックグラウンドで呼ぶ。
    Excel COM（イベントログ追記・Interactive 復元）は呼び出し側が UI スレッドで行う。
    """
    sid = sheet_id_from_cancel_path(cancel_path)
    empty: dict[str, Any] = {
        "ok": False,
        "sheet_id": "",
        "terminated": False,
        "exited_cooperatively": False,
        "coop_detected": False,
        "need_event_log": False,
        "purged_svc_req": 0,
        "run_id": "",
    }
    if not sid:
        return empty
    root = _resolve_cancel_ipc_root(cancel_path, ipc_root)
    if write_progress_cancel and progress_path:
        try:
            write_progress_cancel_status(Path(progress_path))
        except Exception:
            pass
    exited_cooperatively, coop_detected = wait_batch_worker_exit_adaptive(
        sid,
        root,
        initial_ms=int(cooperative_wait_ms or _COOP_WAIT_INITIAL_MS),
        extended_ms=int(cooperative_wait_extended_ms or _COOP_WAIT_EXTENDED_MS),
    )
    terminated = False
    if not exited_cooperatively:
        terminated = terminate_batch_worker(sid, root)
    if terminated:
        try:
            from svc.svc_host import ensure_svc_server  # noqa: WPS433

            ensure_svc_server()
        except Exception:
            pass
    active_run_id = ""
    try:
        from ui_qt.ipc_file import read_pickle  # noqa: WPS433

        active = read_pickle(batch_active_path(sid, root))
        if isinstance(active, dict):
            active_run_id = str(active.get("run_id") or "").strip()
    except Exception:
        pass
    write_batch_cancel_tombstone(sid, root, run_id=active_run_id)
    clear_batch_active_run(sid, root)
    purged = purge_pending_data_agg_batch_svc_requests(root, sid)
    if notify_parent:
        try:
            from ui_qt.ipc_file import write_batch_done_notify  # noqa: WPS433

            msg = "一括実行を中止しました。"
            try:
                from core.core_cst import resolve_config_file_path  # noqa: WPS433
                import json

                cfg_path = resolve_config_file_path("ui_data_agg.json")
                cfg = json.loads(cfg_path.read_text(encoding="utf-8"))
                msg = str((cfg.get("MESSAGES") or {}).get("STATUS_CANCEL") or msg).strip()
            except Exception:
                pass
            write_batch_done_notify(sid, "データ集約", msg, ok=False)
        except Exception:
            pass
    clear_batch_coop_cancel_detected(sid, root)
    need_event_log = bool(terminated or exited_cooperatively)
    try:
        from core.core_log import get_logger  # noqa: WPS433

        get_logger(__name__).info(
            "[DATA_AGG] batch force terminate from UI sheet_id=%s terminated=%s "
            "cooperative_done=%s coop_detected=%s wait_ms=%s wait_ext_ms=%s "
            "progress=%s purged_svc_req=%s run_id=%s com=deferred",
            sid,
            terminated,
            exited_cooperatively,
            coop_detected,
            int(cooperative_wait_ms or _COOP_WAIT_INITIAL_MS),
            int(cooperative_wait_extended_ms or _COOP_WAIT_EXTENDED_MS),
            bool(progress_path),
            purged,
            active_run_id or "-",
        )
    except Exception:
        pass
    return {
        "ok": True,
        "sheet_id": sid,
        "terminated": terminated,
        "exited_cooperatively": exited_cooperatively,
        "coop_detected": coop_detected,
        "need_event_log": need_event_log,
        "purged_svc_req": purged,
        "run_id": active_run_id,
    }


def force_data_agg_batch_cancel_from_ui(
    *,
    cancel_path: Path | str,
    progress_path: Path | str | None = None,
    ipc_root: Path | str | None = None,
    notify_parent: bool = False,
    parent_hwnd: int = 0,
    scenario_id: str = "",
    scenario_path: str = "",
    cooperative_wait_ms: int = _COOP_WAIT_INITIAL_MS,
    cooperative_wait_extended_ms: int = _COOP_WAIT_EXTENDED_MS,
    apply_com_side_effects: bool = True,
) -> bool:
    """
    進捗のキャンセル押下互換 API: 協調待ちのうえ compute ワーカーを強制終了する。

    既定では Win32 即時解除 →（wait/kill・ファイル処理）→ COM 復元／イベントログ。
    進捗ダイアログ本番経路は run_data_agg_batch_force_terminate_no_com を非同期実行し、
    COM は UI 側で遅延適用する（本関数の同期呼び出しでは UI を塞ぎ得る）。
    """
    sid = sheet_id_from_cancel_path(cancel_path)
    if not sid:
        return False
    ph = int(parent_hwnd or 0)
    try:
        from core.excel_host_restore import unlock_excel_host_window  # noqa: WPS433

        unlock_excel_host_window(ph)
    except Exception:
        pass
    result = run_data_agg_batch_force_terminate_no_com(
        cancel_path=cancel_path,
        progress_path=progress_path,
        ipc_root=ipc_root,
        notify_parent=notify_parent,
        write_progress_cancel=True,
        cooperative_wait_ms=cooperative_wait_ms,
        cooperative_wait_extended_ms=cooperative_wait_extended_ms,
    )
    terminated = bool(result.get("terminated"))
    if apply_com_side_effects:
        try:
            from core.excel_host_restore import restore_excel_host_ui_state  # noqa: WPS433

            restore_excel_host_ui_state(ph, sid, com=True)
        except Exception:
            pass
        if result.get("need_event_log"):
            try:
                append_cancel_event_log_from_ui(
                    parent_hwnd=ph,
                    sheet_id=sid,
                    scenario_id=str(scenario_id or ""),
                    scenario_path=str(scenario_path or ""),
                )
            except Exception:
                pass
        try:
            from core.excel_host_restore import restore_excel_host_ui_state  # noqa: WPS433

            restore_excel_host_ui_state(ph, sid, com=True)
        except Exception:
            pass
    return terminated
