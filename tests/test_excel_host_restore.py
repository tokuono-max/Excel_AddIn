# -*- coding: utf-8 -*-
from __future__ import annotations

from unittest.mock import MagicMock, patch

from core.excel_host_restore import restore_excel_host_ui_state


def test_restore_excel_host_ui_state_no_hwnd() -> None:
    assert restore_excel_host_ui_state(0, "sid") is False


def test_restore_excel_host_ui_state_com_restore() -> None:
    api = MagicMock()
    app = MagicMock()
    app.api = api
    fake_ui_common = MagicMock()
    with patch.dict("sys.modules", {"ui_qt.ui_common": fake_ui_common}):
        with patch(
            "core.core_xlc.get_excel_context_from_hwnd",
            return_value=(app, MagicMock(), MagicMock(), 100),
        ):
            with patch(
                "core.core_xlc.excel_try_set_main_commandbars_enabled",
            ) as mock_bars:
                assert restore_excel_host_ui_state(100, "sheet1") is True
    fake_ui_common.enable_excel_window.assert_called_once_with(100, True)
    mock_bars.assert_called_once_with(app, True)
    assert api.Interactive is True
    assert api.ScreenUpdating is True
    assert api.EnableEvents is True
