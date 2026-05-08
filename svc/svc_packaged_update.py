# -*- coding: utf-8 -*-
"""svc action: packaged update check."""

from __future__ import annotations

from core.packaged_update import check_for_updates_interactive


def check_for_updates(
    target_hwnd: int = 0,
    sheet_id: str = "",
    **_kwargs: object,
) -> None:
    """Bridge/svc 経路の更新確認エントリ。"""
    sid = str(sheet_id or "").strip() or "_"
    check_for_updates_interactive(
        "ribbon",
        owner_hwnd=int(target_hwnd or 0),
        sheet_id=sid,
    )
