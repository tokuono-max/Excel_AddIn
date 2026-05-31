# -*- coding: utf-8 -*-
"""ui_notification_sound のユニットテスト。"""
from __future__ import annotations

import pytest

from ui_qt import ui_notification_sound as uns


def test_notification_kind_from_icon() -> None:
    assert uns.notification_kind_from_icon("Critical") == "error"
    assert uns.notification_kind_from_icon("Warning") == "info"
    assert uns.notification_kind_from_icon("Information") == "info"


def test_play_notification_sound_respects_master_disable(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    called: list[int] = []

    def _fake_beep(code: int) -> None:
        called.append(code)

    monkeypatch.setenv("HC_NOTIFICATION_SOUND", "0")
    monkeypatch.setattr(uns.ctypes.windll.user32, "MessageBeep", _fake_beep)
    uns.play_notification_sound("done")
    assert called == []


def test_play_notification_sound_done(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    called: list[int] = []

    def _fake_beep(code: int) -> None:
        called.append(code)

    monkeypatch.delenv("HC_NOTIFICATION_SOUND", raising=False)
    monkeypatch.setattr(uns.ctypes.windll.user32, "MessageBeep", _fake_beep)
    uns.play_notification_sound("done")
    assert called == [uns._MB_OK]


def test_play_notification_on_widget_custom_kind(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    played: list[str] = []
    monkeypatch.setattr(uns, "play_notification_sound", lambda k: played.append(k))

    class _W:
        _hc_notification_sound_kind = "error"

    uns.play_notification_on_widget(_W())
    assert played == ["error"]
