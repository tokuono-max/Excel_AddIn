# -*- coding: utf-8 -*-
"""Windows: replace files under app\\bin while this process still has them mapped."""

from __future__ import annotations

import os
import shutil
import sys
import time
from collections.abc import Callable, Iterable
from pathlib import Path

# Nuitka / CPython runtime files commonly loaded from app\\bin (see tools/nuitka legacy promote script).
RUNTIME_KNOWN_BASENAMES = frozenset(
    {
        "hc_updater.exe",
        "python.exe",
        "pythonw.exe",
        "python312.dll",
        "python3.dll",
        "libcrypto-3.dll",
        "libssl-3.dll",
        "sqlite3.dll",
    }
)

_SIDECAR_SUFFIX = ".was_running_{pid}"


def process_bin_dir() -> Path | None:
    """Directory of sys.executable when resolvable."""
    try:
        return Path(sys.executable).resolve().parent
    except OSError:
        return None


def collect_self_sidecar_dst_paths(*, extra: Iterable[Path] | None = None) -> set[Path]:
    """Paths this process may keep open under its executable directory."""
    out: set[Path] = set()
    if extra is not None:
        for p in extra:
            try:
                if p.is_file():
                    out.add(p.resolve())
            except OSError:
                pass
    pbin = process_bin_dir()
    if pbin is None:
        return out
    try:
        pbin_res = pbin.resolve()
    except OSError:
        return out
    try:
        exe = Path(sys.executable).resolve()
    except OSError:
        return out
    if exe.parent != pbin_res:
        return out
    out.add(exe)
    for name in RUNTIME_KNOWN_BASENAMES:
        cand = pbin_res / name
        try:
            if cand.is_file():
                out.add(cand.resolve())
        except OSError:
            pass
    return out


def sidecar_path_for(dst: Path) -> Path:
    return dst.with_name(f"{dst.name}{_SIDECAR_SUFFIX.format(pid=os.getpid())}")


def replace_via_sidecar(
    src: Path,
    dst: Path,
    *,
    log: Callable[[str], None] | None = None,
    label: str = "in_use_replace",
) -> None:
    """
    Rename dst aside then copy src -> dst (Windows in-use exe/dll pattern).
    Running process keeps using the image via the sidecar path.
    """
    side = sidecar_path_for(dst)
    if log is not None:
        log(f"{label}: begin dst={dst}")
    try:
        if side.is_file():
            side.unlink(missing_ok=True)
    except OSError:
        pass
    if dst.is_file():
        os.replace(dst, side)
    try:
        shutil.copy2(src, dst)
    except Exception:
        if side.is_file() and not dst.is_file():
            try:
                os.replace(side, dst)
            except OSError:
                pass
        raise
    if log is not None:
        log(f"{label}: ok stale_sidecar={side.name if side.is_file() else '-'}")
    try:
        if side.is_file():
            side.unlink(missing_ok=True)
    except OSError:
        if log is not None:
            log(f"{label}: stale_unlink_deferred path={side}")


def copy_file_with_sharing_fallback(
    src: Path,
    dst: Path,
    *,
    proactive_sidecar: set[Path] | None = None,
    log: Callable[[str], None] | None = None,
    rel_label: str = "",
    max_attempts: int = 3,
) -> None:
    """copy2 with WinError-32 sidecar replace when dst is under this process app\\bin."""
    proactive = proactive_sidecar or set()
    use_proactive = False
    if dst.is_file():
        try:
            use_proactive = dst.resolve() in proactive
        except OSError:
            use_proactive = False
    if use_proactive:
        replace_via_sidecar(src, dst, log=log, label="sidecar_proactive")
        return
    last_err: OSError | None = None
    for attempt in range(max_attempts):
        try:
            shutil.copy2(src, dst)
            return
        except OSError as e:
            last_err = e
            wno = int(getattr(e, "winerror", 0) or 0)
            if wno == 32 and dst.is_file():
                replace_via_sidecar(
                    src,
                    dst,
                    log=log,
                    label=f"sidecar_win32 rel={rel_label or dst.name}",
                )
                return
            if wno == 32 and attempt < max_attempts - 1:
                if log is not None:
                    log(
                        "copy_retry_sharing attempt={a}/{m} rel={r}".format(
                            a=attempt + 1,
                            m=max_attempts,
                            r=rel_label or dst.name,
                        )
                    )
                time.sleep(1.0 if attempt == 0 else 2.0)
                continue
            raise
    if last_err is not None:
        raise last_err


def cleanup_stale_sidecar_files(
    bin_dir: Path,
    log: Callable[[str], None] | None = None,
) -> None:
    """Remove *.was_running_* sidecars left from prior updates."""
    if not bin_dir.is_dir():
        return
    for p in bin_dir.glob(f"*.was_running_*"):
        try:
            p.unlink(missing_ok=True)
        except OSError as e:
            if log is not None:
                log(f"sidecar_cleanup skip path={p} err={type(e).__name__}: {e}")
