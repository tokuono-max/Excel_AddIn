# -*- coding: utf-8 -*-
"""更新適用: apply_pending の進捗 UI 順序・fast queue。"""
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


def test_queue_pending_bin_update_fast_skips_local_copy(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    install = tmp_path / "inst"
    install.mkdir()
    patch_src = install / "patch_src.zip"
    patch_src.write_bytes(b"patch")
    st = {
        "latest_bin_version": "1.0.0.1",
        "catalog_path": "",
        "bin_zip_path": str(patch_src),
        "bin_apply_mode": "patch",
        "bin_zip_sha256_expected": "",
        "bin_full_zip_path": "",
        "bin_full_zip_sha256_expected": "",
    }
    copies: list[str] = []
    monkeypatch.setattr(pu, "_install_root", lambda: install)
    monkeypatch.setattr(
        pu,
        "_copy_payload_to_local",
        lambda local, src: copies.append(str(local)),
    )
    ok, _msg = pu._queue_pending_bin_update(
        st,
        source="test",
        skip_apply_confirm=True,
        copy_payload=False,
    )
    assert ok is True
    assert copies == []
    pending = build_paths(install).pending_path.read_text(encoding="utf-8")
    assert '"local_path": ""' in pending.replace(" ", "") or '"local_path":""' in pending.replace(" ", "")


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
            "skip_apply_confirm": False,
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
    monkeypatch.setattr(
        ub,
        "_confirm_pending_apply_before_progress",
        lambda *_a, **_k: True,
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


def test_show_bin_update_prompt_wraps_apply_start_in_busy_ui(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    install = tmp_path / "inst"
    install.mkdir()
    patch_src = install / "patch.zip"
    patch_src.write_bytes(b"patch")
    st = {
        "installed_bin": "1.0.0.0",
        "latest_bin_version": "1.0.0.1",
        "bin_zip_path": str(patch_src),
        "bin_apply_mode": "patch",
    }
    events: list[str] = []

    monkeypatch.setattr(pu, "_install_root", lambda: install)
    monkeypatch.setattr(pu, "_message_box", lambda *_a, **_k: pu.IDYES)
    monkeypatch.setattr(
        pu,
        "_update_check_busy_begin",
        lambda *_a, **_k: events.append("busy_begin"),
    )
    monkeypatch.setattr(
        pu,
        "_update_check_busy_end",
        lambda *_a, **_k: events.append("busy_end"),
    )
    monkeypatch.setattr(
        pu,
        "_start_interactive_bin_apply_from_status",
        lambda *_a, **_k: (events.append("start_apply"), {"ok": True, "pending_path": "p"})[1],
    )

    pu._show_bin_update_prompt(st, owner_hwnd=12345, sheet_id="s1")
    assert events == ["busy_begin", "start_apply", "busy_end"]


def test_run_with_update_busy_ui_reuse_skips_second_begin(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    events: list[str] = []

    monkeypatch.setattr(
        pu,
        "_update_check_busy_begin",
        lambda *_a, **_k: events.append("begin"),
    )
    monkeypatch.setattr(
        pu,
        "_update_check_busy_end",
        lambda *_a, **_k: events.append("end"),
    )
    monkeypatch.setattr(pu, "_UPDATE_BUSY_ACTIVE", True)

    def _fn() -> str:
        return "ok"

    out = pu._run_with_update_busy_ui(
        "msg",
        _fn,
        owner_hwnd=100,
        sheet_id="s",
        reuse_busy=True,
    )
    assert out == "ok"
    assert events == []


def test_resolve_update_busy_owner_hwnd_from_env(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from core import core_env

    monkeypatch.delenv(core_env.ENV_EXCEL_HWND, raising=False)
    assert pu._resolve_update_busy_owner_hwnd(0) == 0
    monkeypatch.setenv(core_env.ENV_EXCEL_HWND, "12345")
    assert pu._resolve_update_busy_owner_hwnd(0) == 12345
    assert pu._resolve_update_busy_owner_hwnd(999) == 999


def test_run_with_update_busy_ui_zero_hwnd_still_shows_busy(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from core import core_env

    events: list[int] = []
    monkeypatch.setenv(core_env.ENV_EXCEL_HWND, "555")
    monkeypatch.setattr(pu, "_UPDATE_BUSY_ACTIVE", False)
    monkeypatch.setattr(
        pu,
        "_update_check_busy_begin",
        lambda *_a, **kw: events.append(int(kw.get("owner_hwnd") or 0)),
    )
    monkeypatch.setattr(pu, "_update_check_busy_end", lambda *_a, **_k: None)

    pu._run_with_update_busy_ui("m", lambda: 1, owner_hwnd=0, sheet_id="s")
    assert events == [555]


def test_run_with_update_busy_ui_zero_hwnd_without_env_still_dispatches(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from core import core_env

    calls: list[int] = []
    monkeypatch.delenv(core_env.ENV_EXCEL_HWND, raising=False)
    monkeypatch.setattr(pu, "_UPDATE_BUSY_ACTIVE", False)
    monkeypatch.setattr(
        pu,
        "_update_check_busy_begin",
        lambda *_a, **kw: calls.append(int(kw.get("owner_hwnd") or 0)),
    )
    monkeypatch.setattr(pu, "_update_check_busy_end", lambda *_a, **_k: None)
    pu._run_with_update_busy_ui("m", lambda: None, owner_hwnd=0)
    assert calls == [0]


def test_wait_for_updater_ui_ready_detects_marker(tmp_path: Path) -> None:
    ready = tmp_path / "ready.json"
    assert pu.wait_for_updater_ui_ready(ready, timeout_sec=0.2) is False
    ready.write_text(json.dumps({"ready": True}), encoding="utf-8")
    assert pu.wait_for_updater_ui_ready(ready, timeout_sec=1.0) is True


def test_wait_for_updater_ui_ready_ignores_file_without_ready_flag(tmp_path: Path) -> None:
    ready = tmp_path / "ready.json"
    ready.write_text("{}", encoding="utf-8")
    assert pu.wait_for_updater_ui_ready(ready, timeout_sec=0.2) is False


def test_run_with_update_busy_ui_refreshes_message_when_startup_busy_active(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    events: list[str] = []

    monkeypatch.setattr(
        pu,
        "_update_check_busy_begin",
        lambda *_a, **_k: events.append("begin"),
    )
    monkeypatch.setattr(
        pu,
        "_update_check_busy_end",
        lambda *_a, **_k: events.append("end"),
    )
    monkeypatch.setattr(pu, "_UPDATE_BUSY_ACTIVE", True)

    pu._run_with_update_busy_ui(
        "prepare-msg",
        lambda: None,
        owner_hwnd=100,
        sheet_id="s",
        reuse_busy=False,
    )
    assert events == ["end", "begin"]


def test_spawn_apply_pending_job_includes_ui_ready_path(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    install = tmp_path / "inst"
    install.mkdir()
    (install / "update" / "tmp").mkdir(parents=True)
    from core.update_state import build_paths, write_pending

    write_pending(
        build_paths(install),
        {
            "schema_version": 2,
            "apply_scope": "bin_only",
            "mode": "patch",
            "target_bin_version": "1.0.0.1",
            "catalog_path": "",
            "state": "downloaded",
        },
    )
    launched: list[list[str]] = []

    def _popen(cmd: Any, **_k: Any) -> Any:
        launched.append(list(cmd))
        class _P:
            pid = 999

        return _P()

    monkeypatch.setattr(pu, "_resolve_hc_updater_argv", lambda *_a, **_k: ["hc_updater.exe", "--job"])
    monkeypatch.setattr(pu.subprocess, "Popen", _popen)
    ok, job_path = pu.spawn_apply_pending_via_hc_updater(install, source="test")
    assert ok is True
    raw = json.loads(Path(job_path).read_text(encoding="utf-8"))
    assert raw.get("UiReadyPath")
    assert pu.ui_ready_path_from_apply_pending_job(job_path) == Path(str(raw["UiReadyPath"]))
