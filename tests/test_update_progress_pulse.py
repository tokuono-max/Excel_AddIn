# -*- coding: utf-8 -*-
"""更新進捗 UI pump（A/B 対策）のユニットテスト。"""
from __future__ import annotations

import sys
import threading
import time
from pathlib import Path
from typing import Any

import pytest

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from bootstrap import update_bootstrap as ub  # noqa: E402
from core import update_process_cleanup as upc  # noqa: E402


class _FakeUi:
    def __init__(self, *, active: bool = True) -> None:
        self.active = active
        self.calls: list[tuple[str, str, float]] = []

    def set(self, title: str, message: str, progress: float) -> None:
        self.calls.append((title, message, progress))


def test_progress_pulse_uses_external_when_ui_inactive() -> None:
    ui = _FakeUi(active=False)
    external: list[tuple[str, str, float]] = []
    ub.set_external_progress_pulse(
        lambda t, m, p: external.append((t, m, p)),
    )
    try:
        ub._progress_pulse(ui, "準備中", "msg", 8.0)
    finally:
        ub.clear_external_progress_pulse()
    assert ui.calls == []
    assert external == [("準備中", "msg", 8.0)]


def test_progress_pulse_prefers_active_ui_over_external() -> None:
    ui = _FakeUi(active=True)
    external: list[tuple[str, str, float]] = []
    ub.set_external_progress_pulse(
        lambda t, m, p: external.append((t, m, p)),
    )
    try:
        ub._progress_pulse(ui, "t", "m", 1.0)
    finally:
        ub.clear_external_progress_pulse()
    assert ui.calls == [("t", "m", 1.0)]
    assert external == []


def test_pulse_while_blocking_pumps_during_slow_work() -> None:
    ui = _FakeUi(active=True)
    started = threading.Event()

    def _slow() -> str:
        started.set()
        time.sleep(0.35)
        return "ok"

    t0 = time.perf_counter()
    out = ub._pulse_while_blocking(ui, "準備中", "構築中", 8, _slow, poll_sec=0.05)
    elapsed = time.perf_counter() - t0
    assert out == "ok"
    assert started.is_set()
    assert len(ui.calls) >= 2
    assert elapsed >= 0.3


def test_sleep_with_ui_pulse_calls_callback() -> None:
    calls: list[int] = []

    def _pulse() -> None:
        calls.append(1)

    t0 = time.perf_counter()
    upc.sleep_with_ui_pulse(0.35, poll_sec=0.08, ui_pulse=_pulse)
    elapsed = time.perf_counter() - t0
    assert len(calls) >= 2
    assert elapsed >= 0.3


def test_ensure_packaged_children_stopped_pulses_during_grace(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    pulses: list[str] = []

    monkeypatch.setattr(upc, "request_packaged_shutdown_flags", lambda: None)
    monkeypatch.setattr(upc, "wait_mutex_clear", lambda *_a, **_k: (True, {}))
    monkeypatch.setattr(upc, "_taskkill_hc_children", lambda *_a, **_k: None)

    def _pulse() -> None:
        pulses.append("x")

    upc.ensure_packaged_children_stopped(
        lambda _m: None,
        {"BOOTSTRAP_PRE_APPLY_GRACE_SEC": 0.25, "BOOTSTRAP_MUTEX_WAIT_SEC": 5},
        phase="test",
        force_taskkill=True,
        ui_pulse=_pulse,
    )
    assert len(pulses) >= 2


def test_hc_updater_registers_external_pulse_during_apply_pending(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch,
) -> None:
    import hc_updater as hu  # noqa: E402

    install = tmp_path / "inst"
    install.mkdir()
    (install / "config").mkdir()
    (install / "config" / "ui_update_check.json").write_text("{}", encoding="utf-8")
    job_path = tmp_path / "job.json"
    raw: dict[str, Any] = {
        "InstallRoot": str(install),
        "LogPath": str(install / "hc_update.log"),
        "InlineBin": False,
        "UiReadyPath": str(tmp_path / "ready.json"),
    }
    job_path.write_text("{}", encoding="utf-8")

    registered: list[bool] = []
    cleared: list[bool] = []

    class _Ui:
        def __init__(self, messages: dict[str, str] | None = None) -> None:
            _ = messages

        @property
        def active(self) -> bool:
            return True

        def set(self, *_a: Any, **_k: Any) -> None:
            pass

    monkeypatch.setattr(hu, "_ProgressUi", _Ui)
    monkeypatch.setattr(hu, "_load_update_messages_for_install", lambda _r: {})
    monkeypatch.setattr(hu, "_write_ui_ready_marker", lambda *_a, **_k: None)

    original_set = ub.set_external_progress_pulse

    def _track_set(fn: Any) -> None:
        registered.append(True)
        original_set(fn)

    monkeypatch.setattr(ub, "set_external_progress_pulse", _track_set)
    monkeypatch.setattr(ub, "clear_external_progress_pulse", lambda: cleared.append(True))
    monkeypatch.setattr(
        "bootstrap.update_bootstrap.apply_pending_update",
        lambda _r: {"ok": True, "deferred_inline_bin_apply": False},
    )

    hu._run_apply_pending_job(job_path, raw)
    assert registered == [True]
    assert cleared == [True]
