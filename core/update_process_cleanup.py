# -*- coding: utf-8 -*-
"""Release packaged HC child processes / mutexes before overwriting app\\bin during updates."""

from __future__ import annotations

import csv
import ctypes
import io
import os
import subprocess
import time
from collections.abc import Callable
from pathlib import Path
from typing import Any

_SYNCHRONIZE = 0x00100000
_MUTEX_UI = "Global\\HC_QT_UI_SERVER"
_MUTEX_SVC = "Global\\HC_SVC_SERVER"
_MUTEX_MAIN = "Global\\HC_MAIN_RUNNER"
_MUTEX_MAIN_LEGACY = "Global\\HC_BRIDGE_RUNNER"

# Short-lived / UI / svc / main (order matters for typical dependency)
_HC_TASKKILL_ORDER = (
    "hc_xlwings_short_runner.exe",
    "hc_ui_server.exe",
    "hc_svc_server.exe",
    "hc_main.exe",
)


def _is_mutex_exists(name: str) -> bool:
    try:
        kernel32 = ctypes.windll.kernel32  # type: ignore[attr-defined]
        handle = kernel32.OpenMutexW(_SYNCHRONIZE, False, name)
        if handle:
            kernel32.CloseHandle(handle)
            return True
        return False
    except Exception:
        return False


def mutex_snapshot() -> dict[str, bool]:
    return {
        "main": _is_mutex_exists(_MUTEX_MAIN),
        "main_legacy": _is_mutex_exists(_MUTEX_MAIN_LEGACY),
        "svc": _is_mutex_exists(_MUTEX_SVC),
        "ui": _is_mutex_exists(_MUTEX_UI),
    }


def is_hc_svc_server_process() -> bool:
    """True when this process is the packaged svc_server executable."""
    if os.name != "nt":
        return False
    return _current_process_image_name_nt() == "hc_svc_server.exe"


def mutex_blocks_pending_apply(snap: dict[str, bool], *, relax_svc_self: bool = False) -> bool:
    """
    Return True if apply_pending_update must wait or abort for HC mutexes.

    When relax_svc_self is True (ribbon「すぐに更新」from hc_svc_server + defer to hc_updater),
    the svc mutex held by this process is ignored so preparation / spawn can proceed.
    """
    if snap.get("main") or snap.get("main_legacy") or snap.get("ui"):
        return True
    if snap.get("svc") and not relax_svc_self:
        return True
    return False


def should_relax_svc_mutex_for_interactive_defer(pending: dict[str, Any]) -> bool:
    """Interactive single-confirm bin apply from svc_server: defer via hc_updater, not inline bin."""
    if not bool(pending.get("skip_apply_confirm")):
        return False
    if os.environ.get("CSV_TOOL_APPLY_PENDING_INLINE_BIN") == "1":
        return False
    return is_hc_svc_server_process()


def wait_mutex_clear(timeout_sec: int, poll_sec: float = 0.5) -> tuple[bool, dict[str, bool]]:
    t0 = time.time()
    last = mutex_snapshot()
    while time.time() - t0 < max(1, int(timeout_sec)):
        last = mutex_snapshot()
        if not any(last.values()):
            return True, last
        time.sleep(max(0.05, float(poll_sec)))
    return False, last


def request_packaged_shutdown_flags() -> None:
    try:
        from ui_qt import ipc_file

        ipc_file.write_shutdown_flag()
    except Exception:
        pass
    try:
        from svc.svc_host import _write_svc_shutdown_flag

        _write_svc_shutdown_flag()
    except Exception:
        pass


def taskkill_other_hc_updater_processes(log_append: Callable[[str], None]) -> None:
    """End other ``hc_updater.exe`` processes (not the current PID). Needed when the worker
    runs ``hc_updater.py`` via Python while a stale ``hc_updater.exe`` still holds the binary."""
    if os.name != "nt":
        return
    cf = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    my_pid = os.getpid()
    try:
        cp = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq hc_updater.exe", "/FO", "CSV", "/NH"],
            capture_output=True,
            text=True,
            check=False,
            creationflags=cf,
        )
    except OSError as e:
        log_append(f"pre_apply_kill: hc_updater_tasklist err={e!s}")
        return
    pids: list[int] = []
    for row in csv.reader(io.StringIO(cp.stdout or "")):
        if len(row) < 2:
            continue
        name = row[0].strip().strip('"').lower()
        if name != "hc_updater.exe":
            continue
        try:
            pids.append(int(row[1].strip('"')))
        except ValueError:
            continue
    killed_any = False
    for pid in pids:
        if pid == my_pid:
            continue
        log_append(f"pre_apply_kill: taskkill /F /PID {pid} (hc_updater.exe)")
        subprocess.run(
            ["taskkill", "/F", "/PID", str(pid)],
            capture_output=True,
            text=True,
            check=False,
            creationflags=cf,
        )
        killed_any = True
    if killed_any:
        time.sleep(0.5)


def probe_tasklist_line() -> str:
    if os.name != "nt":
        return "tasklist=unsupported"
    try:
        cp = subprocess.run(
            ["tasklist"],
            capture_output=True,
            text=True,
            check=False,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        out = (cp.stdout or "").lower()
        flags = {
            "hc_main": "hc_main.exe" in out,
            "hc_svc_server": "hc_svc_server.exe" in out,
            "hc_ui_server": "hc_ui_server.exe" in out,
            "hc_xlwings_short_runner": "hc_xlwings_short_runner.exe" in out,
            "hc_updater": "hc_updater.exe" in out,
            "excel": "excel.exe" in out,
        }
        return "tasklist_rc={rc} running={flags}".format(rc=cp.returncode, flags=flags)
    except Exception as e:
        return f"tasklist_probe_failed={type(e).__name__}: {e}"


def _cfg_skip_kill(cfg: dict[str, Any]) -> bool:
    # Emergency opt-out: skips all taskkill in ensure_packaged_children_stopped (may leave file locks).
    v = cfg.get("BOOTSTRAP_SKIP_PROCESS_KILL", 0)
    if isinstance(v, bool):
        return v
    return str(v).strip().lower() in {"1", "true", "yes", "on"}


def _cfg_grace_sec(cfg: dict[str, Any]) -> float:
    try:
        return max(0.0, float(cfg.get("BOOTSTRAP_PRE_APPLY_GRACE_SEC", 3)))
    except Exception:
        return 3.0


def _cfg_mutex_wait(cfg: dict[str, Any]) -> int:
    try:
        return max(5, int(cfg.get("BOOTSTRAP_MUTEX_WAIT_SEC", 20)))
    except Exception:
        return 20


def _current_process_image_name_nt() -> str:
    """Lowercase basename of this process executable (GetModuleFileNameW)."""
    if os.name != "nt":
        return ""
    try:
        buf = ctypes.create_unicode_buffer(32768)
        n = ctypes.windll.kernel32.GetModuleFileNameW(None, buf, len(buf))  # type: ignore[attr-defined]
        if not n:
            return ""
        return Path(buf.value).name.lower()
    except Exception:
        return ""


def _taskkill_hc_children(log_append: Callable[[str], None]) -> None:
    if os.name != "nt":
        return
    cf = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    self_im = _current_process_image_name_nt()
    for im in _HC_TASKKILL_ORDER:
        if self_im and im.lower() == self_im:
            log_append(f"pre_apply_kill: skip_taskkill self pid={os.getpid()} image={im}")
            continue
        log_append(f"pre_apply_kill: taskkill /F /T /IM {im}")
        cp = subprocess.run(
            ["taskkill", "/F", "/T", "/IM", im],
            capture_output=True,
            text=True,
            check=False,
            creationflags=cf,
        )
        if cp.returncode not in (0, 128):
            err = (cp.stderr or cp.stdout or "").strip()[:400]
            log_append(f"pre_apply_kill: taskkill rc={cp.returncode} detail={err!r}")


def ensure_packaged_children_stopped(
    log_append: Callable[[str], None],
    cfg: dict[str, Any],
    *,
    phase: str,
    force_taskkill: bool,
) -> None:
    """Write shutdown flags, wait for HC mutexes to clear, optional grace, optional taskkill."""
    log_append(f"pre_apply: phase={phase} probe_begin {probe_tasklist_line()}")
    request_packaged_shutdown_flags()
    wait_sec = _cfg_mutex_wait(cfg)
    ok, snap = wait_mutex_clear(timeout_sec=wait_sec, poll_sec=0.5)
    log_append(f"pre_apply: phase={phase} mutex_wait ok={ok} state={snap}")
    grace = _cfg_grace_sec(cfg)
    if grace > 0:
        time.sleep(grace)
    log_append(f"pre_apply: phase={phase} after_grace {probe_tasklist_line()}")
    if not force_taskkill:
        return
    if _cfg_skip_kill(cfg):
        log_append(f"pre_apply: phase={phase} skip_taskkill (BOOTSTRAP_SKIP_PROCESS_KILL)")
        return
    _taskkill_hc_children(log_append)
    time.sleep(1.0)
    log_append(f"pre_apply: phase={phase} probe_after_kill {probe_tasklist_line()}")
