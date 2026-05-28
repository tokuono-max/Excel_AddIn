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
from pathlib import Path
from typing import Any, Callable, Iterator, Optional

_CANCEL_POLL_INTERVAL_SEC = 0.05
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


def log_cancel_detected(*, sheet_id: str, phase: str, files_n: int = 0) -> None:
    """ワーカーが中止フラグを検知したとき（hc_csv.log / hc_csv_diag.log）。"""
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
        raw = ws.range((start_r, 1), (last_r, 8)).value
        if raw is None:
            return []
        if isinstance(raw, (list, tuple)) and raw and not isinstance(raw[0], (list, tuple)):
            return [list(raw)]
        if isinstance(raw, (list, tuple)):
            return [list(r) if isinstance(r, (list, tuple)) else [r] for r in raw]
    except Exception:
        return []
    return []


def _has_recent_cancel_summary(book: Any, *, scenario_id: str, scenario_path: str) -> bool:
    sid = str(scenario_id or "").strip()
    sp = str(scenario_path or "").strip()
    for row in reversed(_iter_recent_event_rows(book)):
        kind = str(row[2] if len(row) >= 3 else "").strip()
        if kind != "一括実行・中止":
            continue
        row_sid = str(row[5] if len(row) >= 6 else "").strip()
        row_sp = str(row[6] if len(row) >= 7 else "").strip()
        if sid and row_sid and sid != row_sid:
            continue
        if sp and row_sp and sp != row_sp:
            continue
        if len(row) >= 8:
            try:
                d = json.loads(str(row[7] or ""))
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


def force_data_agg_batch_cancel_from_ui(
    *,
    cancel_path: Path | str,
    progress_path: Path | str | None = None,
    ipc_root: Path | str | None = None,
    notify_parent: bool = False,
    parent_hwnd: int = 0,
    scenario_id: str = "",
    scenario_path: str = "",
    cooperative_wait_ms: int = 900,
) -> bool:
    """
    進捗のキャンセル押下: 協調フラグに加え compute ワーカーを強制終了する。
    openpyxl 等の長時間ブロック中でも UI から即座に止める。
    """
    sid = sheet_id_from_cancel_path(cancel_path)
    if not sid:
        return False
    root: Path | None = None
    if ipc_root:
        root = Path(ipc_root)
    else:
        try:
            from core import core_env  # noqa: WPS433

            raw = core_env.ipc_dir_raw()
            if raw:
                root = Path(raw)
        except Exception:
            pass
        if root is None:
            root = Path(cancel_path).parent.parent
    exited_cooperatively = wait_batch_worker_exit(
        sid, root, timeout_ms=int(cooperative_wait_ms or 0)
    )
    terminated = False
    if not exited_cooperatively:
        terminated = terminate_batch_worker(sid, root)
    if progress_path:
        try:
            write_progress_cancel_status(Path(progress_path))
        except Exception:
            pass
    if terminated:
        try:
            append_cancel_event_log_from_ui(
                parent_hwnd=int(parent_hwnd or 0),
                sheet_id=sid,
                scenario_id=str(scenario_id or ""),
                scenario_path=str(scenario_path or ""),
            )
        except Exception:
            pass
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
    try:
        from core.core_log import get_logger  # noqa: WPS433

        get_logger(__name__).info(
            "[DATA_AGG] batch force terminate from UI sheet_id=%s terminated=%s "
            "cooperative_done=%s wait_ms=%s progress=%s",
            sid,
            terminated,
            exited_cooperatively,
            int(cooperative_wait_ms or 0),
            bool(progress_path),
        )
    except Exception:
        pass
    return terminated
