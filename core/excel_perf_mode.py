# -*- coding: utf-8 -*-
"""Excel 高速化用の画面更新・イベント抑止と、終了時の EnableEvents 復帰。"""
from __future__ import annotations

from typing import Any

from core.core_log import get_logger

logger = get_logger(__name__)

_PERF_MODE_SAVED: dict[int, dict[str, Any]] = {}

_XL_CALC_MANUAL = -4135
_XL_CALC_AUTO = -4105


def ensure_excel_events_enabled(app: Any) -> None:
    """Application.EnableEvents を True に戻す（シート切替イベント復帰用）。"""
    if app is None:
        return
    try:
        api = getattr(app, "api", None) or app
        api.EnableEvents = True
    except Exception as ex:
        logger.debug("[EXCEL_PERF] ensure_excel_events_enabled failed: %r", ex)


def set_excel_performance_mode(app: Any, on: bool, *, disable_events: bool = True) -> None:
    """Excel の画面更新・計算・（任意で）イベントを抑止/復帰。

    off 時は保存値に関わらず EnableEvents=True を強制する。
    """
    if app is None:
        return
    try:
        api = getattr(app, "api", None) or app
        key = id(api)
        if on:
            saved: dict[str, Any] = {
                "ScreenUpdating": api.ScreenUpdating,
                "Calculation": api.Calculation,
                "DisplayAlerts": api.DisplayAlerts,
            }
            if disable_events:
                saved["EnableEvents"] = api.EnableEvents
            _PERF_MODE_SAVED[key] = saved
            api.ScreenUpdating = False
            api.Calculation = _XL_CALC_MANUAL
            if disable_events:
                api.EnableEvents = False
            api.DisplayAlerts = False
        else:
            saved_prev = _PERF_MODE_SAVED.pop(key, None)
            if saved_prev is not None:
                api.ScreenUpdating = saved_prev.get("ScreenUpdating", True)
                api.Calculation = saved_prev.get("Calculation", _XL_CALC_AUTO)
                api.DisplayAlerts = saved_prev.get("DisplayAlerts", True)
            else:
                api.ScreenUpdating = True
                api.Calculation = _XL_CALC_AUTO
            ensure_excel_events_enabled(app)
    except Exception as ex:
        logger.debug("[EXCEL_PERF] set_excel_performance_mode(on=%s) failed: %r", on, ex)
        if not on:
            ensure_excel_events_enabled(app)
