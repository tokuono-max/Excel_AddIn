# -*- coding: utf-8 -*-
from __future__ import annotations

import json
import time
from pathlib import Path

import pytest

from core import startup_session_gate as gate


@pytest.fixture
def install_root(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Path:
    root = tmp_path / "CSV_Tool"
    (root / "update" / "locks").mkdir(parents=True)
    monkeypatch.setenv("HC_INSTALL_ROOT", str(root))
    monkeypatch.setenv("HC_PACKAGED_DEPLOYMENT", "1")
    monkeypatch.delenv(gate.ENV_DISABLE_GATE, raising=False)
    return root


def _write_gate(root: Path, hwnd: int, state: dict) -> None:
    p = root / "update" / "locks" / f"startup_excel_{hwnd}.json"
    p.write_text(json.dumps(state, ensure_ascii=False), encoding="utf-8")


def test_skip_when_register_in_progress(install_root: Path) -> None:
    hwnd = 1001
    _write_gate(
        install_root,
        hwnd,
        {"in_progress": True, "started_mono": time.time(), "first_prefix": "startup_full"},
    )
    d = gate.evaluate_startup_ui_gate(hwnd, "init_bridge")
    assert d.skip_update_ui is True
    assert d.reason == "startup_register_in_progress"


def test_skip_init_bridge_after_startup_full(install_root: Path) -> None:
    hwnd = 1002
    _write_gate(
        install_root,
        hwnd,
        {
            "in_progress": False,
            "completed_mono": time.time(),
            "first_prefix": "startup_full",
        },
    )
    d = gate.evaluate_startup_ui_gate(hwnd, "init_bridge")
    assert d.skip_update_ui is True
    assert d.reason == "duplicate_init_bridge_after_startup_full"


def test_gate_context_marks_completed(install_root: Path) -> None:
    hwnd = 1003
    with gate.excel_startup_ui_gate(hwnd, "startup_full") as d:
        assert d.skip_update_ui is False
    p = install_root / "update" / "locks" / f"startup_excel_{hwnd}.json"
    state = json.loads(p.read_text(encoding="utf-8"))
    assert state.get("in_progress") is False
    assert state.get("completed_mono")
