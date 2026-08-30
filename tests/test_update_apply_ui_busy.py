# -*- coding: utf-8 -*-
"""更新適用: zip コピー BUSY 先出し・apply_pending の進捗 UI 順序。"""
from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any

import pytest

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from bootstrap import update_bootstrap as ub  # noqa: E402
from core import packaged_update as pu  # noqa: E402
from core.update_state import build_paths, write_pending  # noqa: E402


def test_queue_pending_bin_update_with_busy_ui_runs_under_busy(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    order: list[str] = []

    def _fake_busy(message: str, fn: Any, **kwargs: Any) -> tuple[bool, str]:
        order.append(f"busy:{message[:8]}")
        return fn()

    monkeypatch.setattr(pu, "_run_with_update_busy_ui", _fake_busy)
    monkeypatch.setattr(
        pu,
        "_queue_pending_bin_update",
        lambda *a, **kw: (order.append("queue"), (True, "queued"))[1],
    )

    ok, msg = pu._queue_pending_bin_update_with_busy_ui(
        {},
        source="test",
        require_admin=False,
        skip_apply_confirm=True,
        owner_hwnd=100,
        sheet_id="s1",
    )
    assert ok is True
    assert msg == "queued"
    assert order[0].startswith("busy:")
    assert order[1] == "queue"


class _FakeProgressUi:
    active = True
    cancelled = False

    def set(self, *_a: Any, **_k: Any) -> None:
        pass

    def close(self) -> None:
        pass


def _minimal_install(tmp_path: Path) -> Path:
    install = tmp_path / "app"
    cfg_dir = install / "config"
    cfg_dir.mkdir(parents=True)
    (install / "VERSION.txt").write_text("1.0.0.0\n", encoding="utf-8")
    ui_cfg = {
        "MESSAGES": {
            "UPDATER_PHASE_WAIT_TITLE": "Excel の終了を待っています",
            "UPDATER_PHASE_WAIT_MESSAGE": "すべての Excel を閉じてください。",
            "UPDATER_PHASE_STOP_PROCESSES_MESSAGE": "関連プロセスを終了しています…",
            "PROGRESS_PREPARE_TITLE": "準備中",
            "PROGRESS_PREPARE_MSG": "更新に必要なファイルを用意しています。",
        }
    }
    (cfg_dir / "ui_update_check.json").write_text(
        json.dumps(ui_cfg, ensure_ascii=False), encoding="utf-8"
    )
    return install


def test_apply_pending_creates_progress_ui_before_mutex(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    install = _minimal_install(tmp_path)
    paths = build_paths(install)
    write_pending(
        paths,
        {
            "schema_version": 2,
            "apply_scope": "bin_only",
            "skip_apply_confirm": True,
            "mode": "patch",
            "target_bin_version": "1.0.0.1",
            "catalog_path": "",
            "state": "downloaded",
        },
    )
    events: list[str] = []

    monkeypatch.setattr(
        ub,
        "_ProgressUi",
        lambda: (events.append("ui_create"), _FakeProgressUi())[1],
    )

    def _mutex_blocks(*_a: Any, **_k: Any) -> bool:
        events.append("mutex_check")
        return True

    monkeypatch.setattr(ub, "mutex_blocks_pending_apply", _mutex_blocks)
    monkeypatch.setattr(ub, "ensure_packaged_children_stopped", lambda *_a, **_k: None)

    res = ub._apply_pending_update_impl(install)
    assert res.get("ok") is False
    assert "blocked_by_running_process" in str(res.get("error") or "")
    assert events.index("ui_create") < events.index("mutex_check")


def test_apply_pending_confirm_before_progress_ui(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    install = _minimal_install(tmp_path)
    paths = build_paths(install)
    write_pending(
        paths,
        {
            "schema_version": 2,
            "apply_scope": "bin_only",
            "mode": "patch",
            "target_bin_version": "1.0.0.1",
            "catalog_path": "",
            "state": "downloaded",
        },
    )
    events: list[str] = []

    monkeypatch.setattr(
        ub,
        "_confirm_pending_apply_before_progress",
        lambda *_a, **_k: (events.append("confirm"), False)[1],
    )
    monkeypatch.setattr(
        ub,
        "_ProgressUi",
        lambda: (events.append("ui_create"), _FakeProgressUi())[1],
    )

    res = ub._apply_pending_update_impl(install)
    assert res.get("deferred") is True
    assert events == ["confirm"]
    assert "ui_create" not in events
