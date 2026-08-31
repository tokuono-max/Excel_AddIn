# -*- coding: utf-8 -*-
"""hc_updater apply_pending ジョブ（continuous defer）のユニットテスト。"""
from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any

import pytest

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

import hc_updater as hu  # noqa: E402
from bootstrap import update_bootstrap as ub  # noqa: E402


class _FakeProgressUi:
    active = True
    cancelled = False
    sets: list[tuple[str, str, int]] = []

    def __init__(self, messages: dict[str, str] | None = None) -> None:
        _ = messages

    def set(self, title: str, message: str, progress: int) -> None:
        self.sets.append((title, message, progress))

    def close(self) -> None:
        pass


def test_write_ui_ready_marker_writes_ready_flag(tmp_path: Path) -> None:
    ready = tmp_path / "ready.json"
    hu._write_ui_ready_marker(ready, ui_active=True)
    raw = json.loads(ready.read_text(encoding="utf-8"))
    assert raw.get("ready") is True
    assert raw.get("ui_active") is True


def test_run_apply_pending_job_continuous_hands_off_to_bin_apply(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch,
) -> None:
    install = tmp_path / "inst"
    (install / "config").mkdir(parents=True)
    (install / "config" / "ui_update_check.json").write_text(
        json.dumps(
            {
                "MESSAGES": {
                    "PROGRESS_PREPARE_TITLE": "準備中",
                    "PROGRESS_PREPARE_MSG": "用意しています",
                    "PROGRESS_DEFER_DONE_TITLE": "準備完了",
                    "PROGRESS_DEFER_DONE_TEMPLATE": "Excel を終了してください",
                }
            }
        ),
        encoding="utf-8",
    )
    worker_zip = install / "worker.zip"
    worker_zip.write_bytes(b"worker")
    ui_ready = tmp_path / "ui_ready.json"
    job_path = tmp_path / "job.json"
    log_path = install / "logs" / "hc_update.log"
    raw = {
        "JobType": "apply_pending",
        "InstallRoot": str(install),
        "LogPath": str(log_path),
        "Source": "test",
        "InlineBin": False,
        "UiReadyPath": str(ui_ready),
    }
    job_path.write_text(json.dumps(raw), encoding="utf-8")

    deferred = {
        "ok": True,
        "applied": False,
        "deferred_inline_bin_apply": True,
        "worker_zip_path": str(worker_zip),
        "worker_zip_sha": "abc",
        "worker_apply_mode": "patch",
        "target_bin_version": "1.0.0.1",
        "display_version": "1.0.0.1",
    }
    bin_apply_calls: list[Any] = []

    monkeypatch.setattr(hu, "_ProgressUi", _FakeProgressUi)
    monkeypatch.setattr(ub, "apply_pending_update", lambda _r: deferred)
    monkeypatch.setattr(
        hu,
        "_run_bin_apply_job",
        lambda job, ui, msgs, **kw: bin_apply_calls.append((job.apply_mode, kw)) or 0,
    )

    rc = hu._run_apply_pending_job(job_path, raw)

    assert rc == 0
    assert len(bin_apply_calls) == 1
    assert bin_apply_calls[0][0] == "patch"
    assert ui_ready.is_file()
    assert json.loads(ui_ready.read_text(encoding="utf-8")).get("ready") is True
    assert "continuous_bin_apply done" in log_path.read_text(encoding="utf-8")


def test_run_apply_pending_job_inline_bin_skips_continuous_handoff(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch,
) -> None:
    install = tmp_path / "inst"
    install.mkdir()
    job_path = tmp_path / "job.json"
    log_path = install / "hc_update.log"
    raw = {
        "InstallRoot": str(install),
        "LogPath": str(log_path),
        "InlineBin": True,
    }
    job_path.write_text(json.dumps(raw), encoding="utf-8")

    bin_apply_called: list[str] = []

    monkeypatch.setattr(ub, "apply_pending_update", lambda _r: {"ok": True, "applied": True})
    monkeypatch.setattr(
        hu,
        "_run_bin_apply_job",
        lambda *_a, **_k: bin_apply_called.append("called") or 0,
    )

    rc = hu._run_apply_pending_job(job_path, raw)

    assert rc == 0
    assert bin_apply_called == []
