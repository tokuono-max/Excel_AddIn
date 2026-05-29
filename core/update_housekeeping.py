# -*- coding: utf-8 -*-
"""Update staging cleanup: payload dir, single-generation full_prev archives."""

from __future__ import annotations

import json
import shutil
from collections.abc import Callable
from pathlib import Path

_ARCHIVE_FULL_REL = Path("update") / "archive" / "full"
_PAYLOAD_REL = Path("update") / "payload"


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


def post_deferred_bin_success_housekeeping(
    install_root: Path,
    *,
    log: Callable[[str], None] | None = None,
) -> None:
    """After hc_updater successfully applied bin (defer path): clear payload + single full_prev."""
    if not install_root.is_dir():
        return
    cleanup_update_payload_dir(install_root, log)
    sweep_full_prev_to_single_generation(install_root, log)
