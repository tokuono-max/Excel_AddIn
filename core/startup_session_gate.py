# -*- coding: utf-8 -*-
"""Excel 起動シーケンスの二重 RunPython 時に、更新 UI（bootstrap / 新版確認）の重複表示を抑止する。"""

from __future__ import annotations

import json
import os
import time
from contextlib import contextmanager
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterator

from core.runtime_layout import install_root

# 1 回目が更新ダイアログ待ちでブロック中に 2 回目が走る猶予（秒）
IN_PROGRESS_SUPPRESS_SEC = 180
# Workbook_Open 成功直後の InitPythonServer（init_bridge）抑止（秒）
INIT_BRIDGE_SUPPRESS_AFTER_FULL_SEC = 120

ENV_DISABLE_GATE = "HC_STARTUP_UI_GATE_DISABLE"


@dataclass(frozen=True)
class StartupUiGateDecision:
    skip_update_ui: bool
    reason: str = ""


def _gate_disabled() -> bool:
    v = os.environ.get(ENV_DISABLE_GATE)
    if v is None:
        return False
    return str(v).strip().lower() in ("1", "true", "yes", "on")


def _gate_path(root: Path, hwnd: int) -> Path:
    return root / "update" / "locks" / f"startup_excel_{int(hwnd or 0)}.json"


def _read_state(path: Path) -> dict[str, Any]:
    try:
        if not path.is_file():
            return {}
        raw = json.loads(path.read_text(encoding="utf-8-sig"))
        return raw if isinstance(raw, dict) else {}
    except Exception:
        return {}


def _write_state(path: Path, state: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".new")
    tmp.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(tmp, path)


def evaluate_startup_ui_gate(hwnd: int, perf_prefix: str) -> StartupUiGateDecision:
    """更新 UI をスキップすべきか（二重 RunPython / 遅延 InitPythonServer）。"""
    if _gate_disabled():
        return StartupUiGateDecision(skip_update_ui=False)
    if int(hwnd or 0) <= 0:
        return StartupUiGateDecision(skip_update_ui=False)

    root = install_root()
    if root is None or not root.is_dir():
        return StartupUiGateDecision(skip_update_ui=False)

    path = _gate_path(root, hwnd)
    state = _read_state(path)
    now = time.time()
    prefix = str(perf_prefix or "").strip()

    if state.get("in_progress"):
        started = float(state.get("started_mono") or 0)
        if started > 0 and (now - started) < IN_PROGRESS_SUPPRESS_SEC:
            return StartupUiGateDecision(
                skip_update_ui=True,
                reason="startup_register_in_progress",
            )

    if prefix == "init_bridge":
        completed = float(state.get("completed_mono") or 0)
        first = str(state.get("first_prefix") or "").strip()
        if (
            completed > 0
            and first == "startup_full"
            and (now - completed) < INIT_BRIDGE_SUPPRESS_AFTER_FULL_SEC
        ):
            return StartupUiGateDecision(
                skip_update_ui=True,
                reason="duplicate_init_bridge_after_startup_full",
            )
        if completed > 0 and (now - completed) < 15:
            return StartupUiGateDecision(
                skip_update_ui=True,
                reason="duplicate_init_bridge_recent",
            )

    return StartupUiGateDecision(skip_update_ui=False)


def mark_startup_ui_begin(hwnd: int, perf_prefix: str) -> None:
    root = install_root()
    if root is None or not root.is_dir() or int(hwnd or 0) <= 0:
        return
    path = _gate_path(root, hwnd)
    state = _read_state(path)
    now = time.time()
    state.update(
        {
            "hwnd": int(hwnd),
            "pid": os.getpid(),
            "first_prefix": str(state.get("first_prefix") or perf_prefix or "").strip(),
            "last_prefix": str(perf_prefix or "").strip(),
            "in_progress": True,
            "started_mono": now,
        }
    )
    _write_state(path, state)


def mark_startup_ui_end(hwnd: int, perf_prefix: str) -> None:
    root = install_root()
    if root is None or not root.is_dir() or int(hwnd or 0) <= 0:
        return
    path = _gate_path(root, hwnd)
    state = _read_state(path)
    now = time.time()
    state.update(
        {
            "hwnd": int(hwnd),
            "pid": os.getpid(),
            "last_prefix": str(perf_prefix or "").strip(),
            "in_progress": False,
            "completed_mono": now,
        }
    )
    if not str(state.get("first_prefix") or "").strip():
        state["first_prefix"] = str(perf_prefix or "").strip()
    _write_state(path, state)


@contextmanager
def excel_startup_ui_gate(
    hwnd: int, perf_prefix: str
) -> Iterator[StartupUiGateDecision]:
    decision = evaluate_startup_ui_gate(hwnd, perf_prefix)
    if decision.skip_update_ui:
        yield decision
        return
    mark_startup_ui_begin(hwnd, perf_prefix)
    try:
        yield decision
    finally:
        mark_startup_ui_end(hwnd, perf_prefix)
