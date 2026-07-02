# -*- coding: utf-8 -*-
"""core_cursor のログ間引き・UI_READY 1行化。"""
from __future__ import annotations

import sys
from pathlib import Path
from unittest.mock import MagicMock

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from core import core_cursor  # noqa: E402


def test_log_cursor_wait_on_ok_throttles_info(monkeypatch) -> None:
    core_cursor._CURSOR_ON_LAST_LOG_MONO.clear()
    log = MagicMock()
    monkeypatch.setattr(core_cursor, "_LOG", log)
    monkeypatch.setattr(core_cursor.time, "monotonic", lambda: 100.0)

    core_cursor._log_cursor_wait_on_ok("Main.ForceCursorOn", "progress")
    log.info.assert_called_once()
    log.debug.assert_not_called()

    log.reset_mock()
    monkeypatch.setattr(core_cursor.time, "monotonic", lambda: 101.0)
    core_cursor._log_cursor_wait_on_ok("Main.ForceCursorOn", "progress")
    log.info.assert_not_called()
    log.debug.assert_called_once()


def test_finish_logs_single_info_line(monkeypatch) -> None:
    log = MagicMock()
    monkeypatch.setattr(core_cursor, "_LOG", log)
    result = core_cursor._finish(0.0, True, True, "", "my_reason")
    assert result.ok
    log.info.assert_called_once()
    assert "UI_READY: done ok" in str(log.info.call_args)
    assert "my_reason" in str(log.info.call_args)


def test_finish_failure_logs_error(monkeypatch) -> None:
    log = MagicMock()
    monkeypatch.setattr(core_cursor, "_LOG", log)
    result = core_cursor._finish(0.0, False, False, "boom", "fail_reason")
    assert not result.ok
    log.error.assert_called_once()
    assert "UI_READY: done ng" in str(log.error.call_args)
