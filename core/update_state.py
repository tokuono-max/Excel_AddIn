from __future__ import annotations

import json
import os
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any


RUNTIME_DEFAULTS: dict[str, Any] = {
    "BOOTSTRAP_APPLY_TIMEOUT_SEC": 120,
    "PATCH_RETRY_IN_RUN_MAX": 3,
    "PATCH_RETRY_WAIT_SEC_1": 2,
    "PATCH_RETRY_WAIT_SEC_2": 5,
    "UPDATER_EXCEL_WAIT_TIMEOUT_SEC": 600,
    "BOOTSTRAP_MUTEX_WAIT_SEC": 20,
    "BOOTSTRAP_PRE_APPLY_GRACE_SEC": 3,
    "BOOTSTRAP_SKIP_PROCESS_KILL": 0,
}


@dataclass
class UpdatePaths:
    install_root: Path
    update_root: Path
    payload_root: Path
    pending_path: Path
    lock_path: Path
    log_path: Path


def build_paths(install_root: Path) -> UpdatePaths:
    update_root = install_root / "update"
    return UpdatePaths(
        install_root=install_root,
        update_root=update_root,
        payload_root=update_root / "payload",
        pending_path=update_root / "pending.json",
        lock_path=update_root / "locks" / "apply.lock",
        log_path=update_root / "logs" / "bootstrap_update.log",
    )


def load_runtime_config(install_root: Path) -> dict[str, Any]:
    cfg = dict(RUNTIME_DEFAULTS)
    path = install_root / "config" / "update_runtime.json"
    try:
        raw = json.loads(path.read_text(encoding="utf-8-sig"))
    except Exception:
        return cfg
    if not isinstance(raw, dict):
        return cfg
    for k in cfg.keys():
        if k in raw:
            cfg[k] = raw[k]
    for k in (
        "BOOTSTRAP_APPLY_TIMEOUT_SEC",
        "PATCH_RETRY_IN_RUN_MAX",
        "PATCH_RETRY_WAIT_SEC_1",
        "PATCH_RETRY_WAIT_SEC_2",
        "UPDATER_EXCEL_WAIT_TIMEOUT_SEC",
    ):
        try:
            if k == "UPDATER_EXCEL_WAIT_TIMEOUT_SEC":
                cfg[k] = max(60, int(cfg[k]))
            else:
                cfg[k] = max(0, int(cfg[k]))
        except Exception:
            cfg[k] = int(RUNTIME_DEFAULTS[k])
    try:
        cfg["BOOTSTRAP_MUTEX_WAIT_SEC"] = max(5, int(cfg.get("BOOTSTRAP_MUTEX_WAIT_SEC", 20)))
    except Exception:
        cfg["BOOTSTRAP_MUTEX_WAIT_SEC"] = 20
    try:
        cfg["BOOTSTRAP_PRE_APPLY_GRACE_SEC"] = max(0.0, float(cfg.get("BOOTSTRAP_PRE_APPLY_GRACE_SEC", 3)))
    except Exception:
        cfg["BOOTSTRAP_PRE_APPLY_GRACE_SEC"] = 3.0
    return cfg


def read_pending(paths: UpdatePaths) -> dict[str, Any] | None:
    try:
        raw = json.loads(paths.pending_path.read_text(encoding="utf-8-sig"))
        return raw if isinstance(raw, dict) else None
    except Exception:
        return None


def write_pending(paths: UpdatePaths, pending: dict[str, Any]) -> None:
    paths.update_root.mkdir(parents=True, exist_ok=True)
    fd, tmp_name = tempfile.mkstemp(prefix="pending_", suffix=".json", dir=str(paths.update_root))
    os.close(fd)
    tmp = Path(tmp_name)
    try:
        tmp.write_text(json.dumps(pending, ensure_ascii=False, indent=2), encoding="utf-8")
        tmp.replace(paths.pending_path)
    finally:
        try:
            tmp.unlink(missing_ok=True)
        except Exception:
            pass


def clear_pending(paths: UpdatePaths) -> None:
    try:
        paths.pending_path.unlink(missing_ok=True)
    except Exception:
        pass
