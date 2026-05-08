# -*- coding: utf-8 -*-
"""
VERSION.txt（製品版の1行）を解決する。ヘルプ・版表示などで共有。
優先: HC_INSTALL_ROOT → core_sys.get_app_path() → リポジトリルート。
"""

from __future__ import annotations

import os
from pathlib import Path


def _project_root() -> Path:
    return Path(__file__).resolve().parent.parent


def candidate_version_txt_paths(*, project_root: Path | None = None) -> list[Path]:
    """重複を除いた探索順の VERSION.txt パス。"""
    out: list[Path] = []
    ir = (os.environ.get("HC_INSTALL_ROOT") or "").strip()
    if ir:
        out.append(Path(ir) / "VERSION.txt")
    try:
        from core import core_sys as cs

        app = cs.get_app_path()
        if app:
            out.append(Path(app) / "VERSION.txt")
    except Exception:
        pass
    root = project_root if project_root is not None else _project_root()
    out.append(root / "VERSION.txt")

    seen: set[str] = set()
    uniq: list[Path] = []
    for p in out:
        key = str(p)
        if key not in seen:
            seen.add(key)
            uniq.append(p)
    return uniq


def read_product_version_line() -> str:
    """VERSION.txt の先頭非空行。見つからなければ空文字。"""
    for p in candidate_version_txt_paths():
        if not p.is_file():
            continue
        try:
            lines = p.read_text(encoding="utf-8-sig").splitlines()
            if lines:
                v = lines[0].strip()
                if v:
                    return v
        except OSError:
            continue
    return ""
