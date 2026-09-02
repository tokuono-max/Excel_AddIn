# -*- coding: utf-8 -*-
"""ネットワーク（UNC / マップドドライブ）パス判定。"""
from __future__ import annotations

import os
from pathlib import Path

from svc.data_agg_path_norm import PathLike

_DRIVE_REMOTE = 4


def path_class(file_path: PathLike) -> str:
    """local / unc / mapped / unknown"""
    try:
        raw = str(file_path)
        if raw.startswith("\\\\") or raw.startswith("//"):
            return "unc"
        if os.name != "nt":
            return "local"
        try:
            resolved = Path(file_path).resolve()
        except OSError:
            resolved = Path(file_path)
        drive = os.path.splitdrive(str(resolved))[0]
        if not drive:
            return "local"
        import ctypes

        dt = int(ctypes.windll.kernel32.GetDriveTypeW(drive + "\\"))
        if dt == _DRIVE_REMOTE:
            return "mapped"
        return "local"
    except Exception:
        return "unknown"


def path_is_network(file_path: PathLike) -> bool:
    return path_class(file_path) in ("unc", "mapped")
