# -*- coding: utf-8 -*-
from __future__ import annotations

from unittest.mock import MagicMock

from core.excel_perf_mode import ensure_excel_events_enabled, set_excel_performance_mode


def test_set_excel_performance_mode_off_forces_enable_events_true() -> None:
    api = MagicMock()
    api.ScreenUpdating = True
    api.Calculation = -4105
    api.DisplayAlerts = True
    api.EnableEvents = False
    app = MagicMock()
    app.api = api

    set_excel_performance_mode(app, True)
    assert api.EnableEvents is False

    api.EnableEvents = False
    set_excel_performance_mode(app, False)
    assert api.EnableEvents is True


def test_set_excel_performance_mode_off_without_saved_still_enables_events() -> None:
    api = MagicMock()
    api.EnableEvents = False
    app = MagicMock()
    app.api = api

    set_excel_performance_mode(app, False)
    assert api.EnableEvents is True


def test_ensure_excel_events_enabled() -> None:
    api = MagicMock()
    api.EnableEvents = False
    app = MagicMock()
    app.api = api

    ensure_excel_events_enabled(app)
    assert api.EnableEvents is True


def test_set_excel_performance_mode_disable_events_false_does_not_turn_off_on_enter() -> None:
    api = MagicMock()
    api.EnableEvents = True
    app = MagicMock()
    app.api = api

    set_excel_performance_mode(app, True, disable_events=False)
    assert api.EnableEvents is True

    api.EnableEvents = False
    set_excel_performance_mode(app, False, disable_events=False)
    assert api.EnableEvents is True
