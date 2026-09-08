# -*- coding: utf-8 -*-
"""一括完了ポーリング: active 消失と notify 書込の順序ずれ猶予。"""
from __future__ import annotations

import time
from pathlib import Path
from types import SimpleNamespace

import pytest

from ui_qt import ui_data_agg as ui_mod

def _host(*, locked: bool = True) -> SimpleNamespace:
    timer = SimpleNamespace(stopped=False)

    def stop() -> None:
        timer.stopped = True

    timer.stop = stop  # type: ignore[attr-defined]
    h = SimpleNamespace(
        _sheet_id="sheet-a",
        _batch_poll_sheet_id="sheet-a",
        _batch_poll_run_id="run-1",
        _batch_poll_deadline=time.time() + 3600.0,
        _batch_poll_active_gone_deadline=0.0,
        _batch_ui_locked=locked,
        _batch_poll_timer=timer,
        _released=False,
    )

    def _batch_active_pickle_present() -> bool:
        return False

    def _release_batch_ui_lock() -> None:
        h._batch_ui_locked = False
        h._released = True

    h._batch_active_pickle_present = _batch_active_pickle_present  # type: ignore[attr-defined]
    h._release_batch_ui_lock = _release_batch_ui_lock  # type: ignore[attr-defined]
    return h


def test_poll_keeps_timer_during_grace_when_active_gone_without_notify(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    host = _host()
    monkeypatch.setattr(ui_mod, "try_read_batch_done_notify", lambda _sid: None)

    ui_mod._DataAggMainWindow._on_batch_done_poll_tick(host)

    assert host._released is True
    assert host._batch_poll_timer.stopped is False
    assert host._batch_poll_active_gone_deadline > time.time()


def test_poll_shows_done_during_grace_after_active_gone(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    host = _host()
    host._batch_poll_active_gone_deadline = time.time() + 15.0
    notices: list[tuple[str, str]] = []

    monkeypatch.setattr(
        ui_mod,
        "try_read_batch_done_notify",
        lambda _sid: {
            "title": "データ集約",
            "message": "完了しました。",
            "ok": True,
            "run_id": "run-1",
        },
    )
    monkeypatch.setattr(ui_mod, "delete_batch_done_notify", lambda _sid: None)
    monkeypatch.setattr(
        ui_mod,
        "show_done_notice",
        lambda _parent, title, msg: notices.append((title, msg)),
    )

    ui_mod._DataAggMainWindow._on_batch_done_poll_tick(host)

    assert notices
    assert host._batch_poll_timer.stopped is True
    assert host._batch_poll_active_gone_deadline == 0.0


def test_poll_stops_after_grace_expires_without_notify(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    host = _host(locked=False)
    host._batch_poll_active_gone_deadline = time.time() - 0.1
    monkeypatch.setattr(ui_mod, "try_read_batch_done_notify", lambda _sid: None)

    ui_mod._DataAggMainWindow._on_batch_done_poll_tick(host)

    assert host._batch_poll_timer.stopped is True


def test_finish_helpers_notify_before_clear_active() -> None:
    """_finish / _finish_write は完了通知を active 削除より先に書く。"""
    src = (
        Path(__file__).resolve().parents[1] / "svc" / "svc_data_agg.py"
    ).read_text(encoding="utf-8")
    for marker in ("def _finish(", "def _finish_write("):
        i = src.index(marker)
        chunk = src[i : i + 1500]
        notify_i = chunk.index("_batch_done_notify(")
        clear_i = chunk.index("_clear_active_batch_run_if_current(")
        assert notify_i < clear_i, marker
