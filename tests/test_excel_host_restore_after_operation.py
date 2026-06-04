# -*- coding: utf-8 -*-
from __future__ import annotations

from unittest.mock import MagicMock, patch

from core.excel_host_restore import restore_excel_host_after_operation


def test_restore_excel_host_after_operation_calls_both() -> None:
    app = MagicMock()
    with patch("core.excel_host_restore.ensure_excel_events_enabled") as mock_events:
        with patch("core.excel_host_restore.restore_excel_host_ui_state", return_value=True) as mock_ui:
            restore_excel_host_after_operation(100, "sid1", app)
    mock_events.assert_called_once_with(app)
    mock_ui.assert_called_once_with(100, "sid1")
