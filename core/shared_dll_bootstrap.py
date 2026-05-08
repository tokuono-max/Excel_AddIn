# -*- coding: utf-8 -*-
"""Windows DLL 探索パス登録（``os.add_dll_directory``）。

配布では Nuitka 成果物を **単一の** ``app\\bin\\`` に集約する。各 ``hc_*.exe`` は同じ bin 直下に
依存 DLL・PySide6 等がある想定で、必要に応じて bin 配下のサブフォルダを探索へ追加する。

開発時は venv の ``python.exe`` 直下に bin 相当のレイアウトは無いため、多くの場合は何もしない。
"""

from __future__ import annotations

import os
import sys
from pathlib import Path


def ensure_shared_dll_search_path_for_layout(exe_parent: Path | str) -> Path | None:
    """配布レイアウト向けフック（互換名）。``app\\bin`` 単一ディレクトリでは EXE 直下で DLL が解決されることが多く、多くの場合は何もしない。

    `exe_parent` は凍結時 ``sys.executable`` の親（例: ``...\\app\\bin``）。
    """
    if os.name != "nt":
        return None
    _ = Path(exe_parent).resolve()
    return None


def ensure_shared_dll_search_path_next_to_executable() -> Path | None:
    """`sys.executable` の親を `exe_parent` として `ensure_shared_dll_search_path_for_layout` を呼ぶ。"""
    return ensure_shared_dll_search_path_for_layout(Path(sys.executable).resolve().parent)


def ensure_ui_server_windows_dll_search_paths() -> None:
    """`hc_ui_server.exe` 起動直後・PySide6 より前に呼ぶ。

    ``app\\bin`` に PySide6 / shiboken6 が並ぶ前提で、探索パスに bin とそのサブフォルダを足す。
    """
    if os.name != "nt":
        return
    add = getattr(os, "add_dll_directory", None)
    base = Path(sys.executable).resolve().parent

    candidates: list[Path] = []
    candidates.append(base)
    candidates.extend(
        [
            base / "shiboken6",
            base / "PySide6",
        ]
    )
    libqt = base / "PySide6" / "lib"
    if libqt.is_dir():
        candidates.append(libqt)

    seen: set[str] = set()
    path_prefix: list[str] = []
    for d in candidates:
        if not d.is_dir():
            continue
        p = str(d.resolve())
        if p in seen:
            continue
        seen.add(p)
        path_prefix.append(p)
        if add is not None:
            try:
                add(p)
            except OSError:
                pass

    if path_prefix:
        prev = os.environ.get("PATH", "")
        os.environ["PATH"] = os.pathsep.join(path_prefix) + (os.pathsep + prev if prev else "")
