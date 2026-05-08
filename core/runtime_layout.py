# -*- coding: utf-8 -*-
"""Packaged EXE deployment vs development layout (HC_INSTALL_ROOT / HC_PACKAGED_DEPLOYMENT)."""

from __future__ import annotations

import os
from pathlib import Path

from core import core_env

ENV_INSTALL_ROOT = "HC_INSTALL_ROOT"
ENV_PACKAGED = "HC_PACKAGED_DEPLOYMENT"


def install_root() -> Path | None:
    """Installer sets HC_INSTALL_ROOT to the deployment tree (e.g. CSV_Tool root)."""
    raw = (core_env.get(ENV_INSTALL_ROOT) or "").strip()
    if not raw:
        return None
    try:
        p = Path(raw).resolve()
    except OSError:
        return None
    return p if p.is_dir() else None


def packaged_spawn_requested() -> bool:
    """True when installer / user env declares packaged mode (spawn EXEs under app\\)."""
    return core_env.truthy(os.environ.get(ENV_PACKAGED), empty_means_false=True)


def use_packaged_server_commands() -> bool:
    """True when packaged mode is on and bridge EXE exists under app\\bin\\."""
    if not packaged_spawn_requested():
        return False
    root = install_root()
    if root is None:
        return False
    exe = root / "app" / "bin" / "hc_main.exe"
    return exe.is_file()


def packaged_app_exe(filename: str) -> Path | None:
    """e.g. filename=hc_main.exe -> .../app/bin/hc_main.exe"""
    root = install_root()
    if root is None:
        return None
    p = (root / "app" / "bin" / filename).resolve()
    return p if p.is_file() else None


def env_with_packaged_dll_search_path(
    env: dict[str, str],
    install_root: Path | None,
) -> dict[str, str]:
    """子プロセス用 env の PATH 先頭に ``app\\bin`` とインストールルートを足す（Windows のみ）。

    Nuitka 製 ``hc_*.exe`` は起動直後に ``python312.dll`` 等をロードする。単一 ``app\\bin``
    に集約した配布では、PATH 先頭に ``app\\bin`` を足せば EXE 横以外からも解決しやすい。
    """
    if install_root is None or os.name != "nt":
        return env
    app_bin = install_root / "app" / "bin"
    prefixes: list[str] = []
    if app_bin.is_dir():
        prefixes.append(str(app_bin.resolve()))
    prefixes.append(str(install_root.resolve()))
    if not prefixes:
        return env
    out = dict(env)
    prev = out.get("PATH", "")
    out["PATH"] = os.pathsep.join(prefixes + ([prev] if prev else []))
    return out


def runtime_project_root(fallback_from_file: str) -> Path:
    """Project / install root: HC_INSTALL_ROOT if set, else parent of svc_host (dev)."""
    r = install_root()
    if r is not None:
        return r
    return Path(fallback_from_file).resolve().parent.parent
