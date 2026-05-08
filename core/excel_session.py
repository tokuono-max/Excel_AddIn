# -*- coding: utf-8 -*-
"""
VBA xlwings RunPython や他モジュールから、モジュール名をソースに載せずに呼ぶ薄い入口。
実体は core.ribbon_invoke（ルート hc_main.py は常駐ブリッジ専用）。
"""
from __future__ import annotations

from typing import Any, Optional

from core.ribbon_invoke import clear_registry as _clear_registry
from core.ribbon_invoke import invoke as _invoke
from core.ribbon_invoke import register_book as _register_book


def clear_internal_registry() -> None:
    """Python 側のブック参照レジストリを空にする（TerminatePython 用）。"""
    _clear_registry()


def register_book(target_hwnd: Optional[int] = None) -> None:
    """起動時ブック登録（svc_host.excel_startup 用）。"""
    _register_book(target_hwnd=target_hwnd)


def invoke_action(
    action: str,
    target_hwnd: Optional[int],
    sheet_id: str,
    **kwargs: Any,
) -> None:
    """ribbon_invoke.invoke と同等（短寿命 RunPython・子プロセスからの単一入口）。"""
    _invoke(action=action, target_hwnd=target_hwnd, sheet_id=sheet_id, **kwargs)
