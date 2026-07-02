# -*- coding: utf-8 -*-
"""Excel ブック登録済み HWND の IPC 記録（リボン fast path で register_book COM を省略）。"""

from __future__ import annotations

import json
import os
import time
from pathlib import Path
from typing import Any

_REGISTRY_FILE = "registered_excel_books.json"


def _control_dir() -> Path:
    from ui_qt import ipc_file

    return Path(str(ipc_file.get_ipc_root())) / "control"


def _read_registry() -> dict[str, Any]:
    path = _control_dir() / _REGISTRY_FILE
    try:
        if not path.is_file():
            return {}
        raw = json.loads(path.read_text(encoding="utf-8-sig"))
        return raw if isinstance(raw, dict) else {}
    except Exception:
        return {}


def _write_registry(data: dict[str, Any]) -> None:
    path = _control_dir() / _REGISTRY_FILE
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(".json.new")
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(tmp, path)


def mark_excel_book_registered(hwnd: int) -> None:
    """register_book 成功後に呼ぶ。同一 HWND の 2 回目以降リボンで COM 登録を省略可能にする。"""
    ph = int(hwnd or 0)
    if ph <= 0:
        return
    reg = _read_registry()
    reg[str(ph)] = {
        "hwnd": ph,
        "marked_mono": time.monotonic(),
        "marked_at": time.time(),
    }
    _write_registry(reg)


def should_skip_register_book_com(hwnd: int) -> bool:
    """常駐ホスト生存時、登録済みかつウィンドウ生存なら register_book を省略してよい。"""
    ph = int(hwnd or 0)
    if ph <= 0:
        return False
    try:
        from core.core_w32 import is_window

        if not is_window(ph):
            return False
    except Exception:
        return False
    return str(ph) in _read_registry()
