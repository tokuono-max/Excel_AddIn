# -*- coding: utf-8 -*-
"""Excel ホストの UI 状態復元（データ集約キャンセル／強制終了後）。"""
from __future__ import annotations

from typing import Any

from core.core_log import get_logger
from core.excel_perf_mode import ensure_excel_events_enabled

logger = get_logger(__name__)

_XL_CURSOR_DEFAULT = -4143


def restore_excel_host_ui_state(parent_hwnd: int, sheet_id: str = "") -> bool:
    """Interactive / ScreenUpdating / CommandBars / 子 HWND をベストエフォートで復元する。

    svc ワーカー強制終了で suspend_sheet_updates の finally が走らない場合の救済。
    """
    ph = int(parent_hwnd or 0)
    if ph <= 0:
        return False
    restored = False
    try:
        from ui_qt.ui_common import enable_excel_window  # noqa: WPS433

        enable_excel_window(ph, True)
        restored = True
    except Exception as ex:
        logger.debug("[EXCEL_RESTORE] enable_excel_window failed hwnd=%s ex=%r", ph, ex)
    try:
        from core.core_xlc import (  # noqa: WPS433
            excel_try_set_main_commandbars_enabled,
            get_excel_context_from_hwnd,
        )

        ctx = get_excel_context_from_hwnd(ph, str(sheet_id or ""))
        if ctx:
            app, *_rest = ctx
            excel_try_set_main_commandbars_enabled(app, True)
            api = getattr(app, "api", None)
            if api is not None:
                for attr, val in (
                    ("Interactive", True),
                    ("ScreenUpdating", True),
                    ("EnableEvents", True),
                ):
                    try:
                        setattr(api, attr, val)
                        restored = True
                    except Exception:
                        pass
                try:
                    api.Cursor = _XL_CURSOR_DEFAULT
                    restored = True
                except Exception:
                    pass
    except Exception as ex:
        logger.warning(
            "[EXCEL_RESTORE] COM restore failed hwnd=%s sheet_id=%r ex=%r",
            ph,
            sheet_id,
            ex,
        )
    if restored:
        logger.info(
            "[EXCEL_RESTORE] restored hwnd=%s sheet_id=%r",
            ph,
            str(sheet_id or "") or "-",
        )
    return restored


def restore_excel_host_after_operation(
    parent_hwnd: int,
    sheet_id: str = "",
    app: Any = None,
) -> None:
    """svc 処理終了時: EnableEvents 復帰とホスト UI 復元をベストエフォートで行う。"""
    ensure_excel_events_enabled(app)
    restore_excel_host_ui_state(int(parent_hwnd or 0), sheet_id)
