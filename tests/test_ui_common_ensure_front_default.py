# -*- coding: utf-8 -*-
"""ui_common prepare / ensure_front の既定 bring_excel_first テスト。"""
from __future__ import annotations

from unittest.mock import MagicMock, patch


def test_prepare_dialog_default_bring_excel_first_false() -> None:
    from ui_qt.ui_common import prepare_dialog_excel_center_before_show

    w = MagicMock()
    w.property.side_effect = lambda key: False if key == "_hc_prepare_skip_ensure_front" else None
    calls: list[bool] = []

    with (
        patch("ui_qt.ui_common._set_owner_hwnd"),
        patch("ui_qt.ui_common.center_on_excel"),
        patch("ui_qt.ui_common.log_ui_fg_phase"),
        patch(
            "ui_qt.ui_common.ensure_front",
            side_effect=lambda _w, _ph, *, bring_excel_first=False, _ff_retry=0: calls.append(
                bool(bring_excel_first)
            ),
        ),
    ):
        prepare_dialog_excel_center_before_show(w, 123)

    assert calls == [False]


def test_ensure_front_strengthens_when_not_bring_excel_first() -> None:
    from ui_qt.ui_common import ensure_front

    w = MagicMock()
    w.winId.return_value = 555
    w.isVisible.return_value = True
    w.property.return_value = False

    with (
        patch("ui_qt.ui_common._reapply_win32_owner_if_missing"),
        patch("ui_qt.ui_common.log_ui_fg_phase"),
        patch("ui_qt.ui_common._ff_diag"),
        patch("ui_qt.ui_common._ff_diag_ensure_front_snapshot"),
        patch("ui_qt.ui_common._trace"),
        patch("ui_qt.ui_common._trace_widget_rect"),
        patch("ctypes.windll.user32.SetWindowPos"),
        patch("ui_qt.ui_common._w32") as mock_w32,
    ):
        mock_w32.set_foreground_window_result.return_value = True
        ensure_front(w, 100, bring_excel_first=False)

    mock_w32.set_foreground_window_attach_input.assert_called_once_with(555)
    mock_w32.nudge_top_level_to_foreground.assert_called_once_with(555)


def test_prepare_dialog_honors_bring_excel_first_property() -> None:
    from ui_qt.ui_common import prepare_dialog_excel_center_before_show

    w = MagicMock()

    def _prop(key: str) -> object:
        if key == "_hc_prepare_skip_ensure_front":
            return False
        if key == "_hc_ensure_front_bring_excel_first":
            return True
        return None

    w.property.side_effect = _prop
    calls: list[bool] = []

    with (
        patch("ui_qt.ui_common._set_owner_hwnd"),
        patch("ui_qt.ui_common.center_on_excel"),
        patch("ui_qt.ui_common.log_ui_fg_phase"),
        patch(
            "ui_qt.ui_common.ensure_front",
            side_effect=lambda _w, _ph, *, bring_excel_first=False, _ff_retry=0: calls.append(
                bool(bring_excel_first)
            ),
        ),
    ):
        prepare_dialog_excel_center_before_show(w, 123)

    assert calls == [True]


def test_want_bring_excel_first_explicit_config() -> None:
    from ui_qt.ui_common import want_bring_excel_first_while_modal

    assert want_bring_excel_first_while_modal({"BRING_EXCEL_FIRST": True}) is True
    assert want_bring_excel_first_while_modal({"BRING_EXCEL_FIRST": False}) is False


def test_want_bring_excel_first_topmost_defaults_false() -> None:
    from ui_qt.ui_common import want_bring_excel_first_while_modal

    assert want_bring_excel_first_while_modal({"TOPMOST": True}) is False
    assert want_bring_excel_first_while_modal({"ALWAYS_IN_FRONT_OF_EXCEL": True}) is False


def test_want_bring_excel_first_follows_excel_lock_when_not_topmost() -> None:
    from ui_qt.ui_common import want_bring_excel_first_while_modal

    assert want_bring_excel_first_while_modal({"EXCEL_LOCK": True}) is True
    assert want_bring_excel_first_while_modal({"EXCEL_LOCK": False}) is False
