# -*- coding: utf-8 -*-
"""Update staging cleanup: payload dir, single-generation full_prev archives, stale gate locks."""

from __future__ import annotations

import json
import shutil
import time
from collections.abc import Callable
from pathlib import Path

from core.startup_session_gate import (
    INIT_BRIDGE_SUPPRESS_AFTER_FULL_SEC,
    IN_PROGRESS_SUPPRESS_SEC,
)

_ARCHIVE_FULL_REL = Path("update") / "archive" / "full"
_LOCKS_REL = Path("update") / "locks"
_PAYLOAD_REL = Path("update") / "payload"
_STARTUP_GATE_GLOB = "startup_excel_*.json"


def cleanup_update_payload_dir(install_root: Path, log: Callable[[str], None] | None = None) -> None:
    payload = install_root / _PAYLOAD_REL
    if not payload.is_dir():
        return
    try:
        shutil.rmtree(payload, ignore_errors=True)
        if log:
            log(f"update_housekeeping: removed payload path={payload}")
    except Exception as e:
        if log:
            log(f"update_housekeeping: payload rmtree err={type(e).__name__}: {e}")


def sweep_full_prev_to_single_generation(install_root: Path, log: Callable[[str], None] | None = None) -> None:
    """Keep retain.json zip_path if valid, else newest full_prev_*.zip; remove other full_prev_*.zip."""
    archive_dir = install_root / _ARCHIVE_FULL_REL
    if not archive_dir.is_dir():
        return
    retain_path = archive_dir / "retain.json"
    keep_key: str | None = None
    try:
        if retain_path.is_file():
            meta = json.loads(retain_path.read_text(encoding="utf-8-sig"))
            if isinstance(meta, dict):
                zp = str(meta.get("zip_path") or "").strip()
                if zp:
                    p = Path(zp)
                    if p.is_file():
                        keep_key = str(p.resolve()).lower()
    except Exception:
        pass
    if keep_key is None:
        cands = sorted(archive_dir.glob("full_prev_*.zip"), key=lambda p: p.stat().st_mtime, reverse=True)
        if not cands:
            if log:
                log("update_housekeeping: sweep_full_prev no candidates")
            return
        try:
            keep_key = str(cands[0].resolve()).lower()
        except Exception:
            keep_key = str(cands[0]).lower()
    removed = 0
    for fp in archive_dir.glob("full_prev_*.zip"):
        try:
            key = str(fp.resolve()).lower()
        except Exception:
            key = str(fp).lower()
        if key == keep_key:
            continue
        try:
            fp.unlink(missing_ok=True)
            removed += 1
        except OSError as e:
            if log:
                log(f"update_housekeeping: sweep unlink failed path={fp} err={type(e).__name__}: {e}")
    if log:
        log(f"update_housekeeping: sweep_full_prev removed={removed}")


def _startup_gate_lock_stale(
    state: dict,
    *,
    fp: Path,
    now_ts: float,
) -> bool:
    """True when startup_excel_*.json is no longer needed for duplicate-suppress."""
    if state.get("in_progress"):
        started = float(state.get("started_mono") or 0)
        if started > 0 and (now_ts - started) < IN_PROGRESS_SUPPRESS_SEC:
            return False
        return True
    completed = float(state.get("completed_mono") or 0)
    if completed > 0:
        return (now_ts - completed) >= INIT_BRIDGE_SUPPRESS_AFTER_FULL_SEC
    try:
        return (now_ts - fp.stat().st_mtime) >= IN_PROGRESS_SUPPRESS_SEC
    except OSError:
        return True


def sweep_stale_startup_excel_gate_locks(
    install_root: Path,
    log: Callable[[str], None] | None = None,
    *,
    keep_hwnd: int | None = None,
    now_ts: float | None = None,
) -> None:
    """Remove stale startup_excel_{hwnd}.json gate files (closed Excel / completed startup)."""
    locks_dir = install_root / _LOCKS_REL
    if not locks_dir.is_dir():
        return
    now = time.time() if now_ts is None else float(now_ts)
    keep_name = (
        f"startup_excel_{int(keep_hwnd)}.json"
        if keep_hwnd is not None and int(keep_hwnd) > 0
        else None
    )
    removed = 0
    for fp in locks_dir.glob(_STARTUP_GATE_GLOB):
        if keep_name and fp.name == keep_name:
            continue
        try:
            raw = json.loads(fp.read_text(encoding="utf-8-sig"))
            state = raw if isinstance(raw, dict) else {}
        except Exception:
            state = {}
        if not _startup_gate_lock_stale(state, fp=fp, now_ts=now):
            continue
        try:
            fp.unlink(missing_ok=True)
            removed += 1
        except OSError as e:
            if log:
                log(
                    "update_housekeeping: sweep_startup_gate unlink failed "
                    f"path={fp} err={type(e).__name__}: {e}"
                )
    for fp in locks_dir.glob(f"{_STARTUP_GATE_GLOB}.new"):
        try:
            fp.unlink(missing_ok=True)
            removed += 1
        except OSError as e:
            if log:
                log(
                    "update_housekeeping: sweep_startup_gate tmp unlink failed "
                    f"path={fp} err={type(e).__name__}: {e}"
                )
    if log:
        log(f"update_housekeeping: sweep_startup_gate removed={removed}")


def run_startup_housekeeping(
    install_root: Path,
    log: Callable[[str], None] | None = None,
    *,
    keep_gate_hwnd: int | None = None,
) -> None:
    """Startup path when no pending update: stale gate locks + single full_prev generation."""
    if not install_root.is_dir():
        return
    sweep_stale_startup_excel_gate_locks(
        install_root,
        log,
        keep_hwnd=keep_gate_hwnd,
    )
    sweep_full_prev_to_single_generation(install_root, log)


def post_deferred_bin_success_housekeeping(
    install_root: Path,
    *,
    log: Callable[[str], None] | None = None,
) -> None:
    """After hc_updater successfully applied bin (defer path): clear payload + single full_prev."""
    if not install_root.is_dir():
        return
    cleanup_update_payload_dir(install_root, log)
    sweep_stale_startup_excel_gate_locks(install_root, log)
    sweep_full_prev_to_single_generation(install_root, log)
