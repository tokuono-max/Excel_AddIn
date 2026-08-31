# -*- coding: utf-8 -*-
"""更新 UX: 主経路・起動シーケンスの結合テスト（mock ベース）。"""
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
from core.runtime_layout import ENV_PACKAGED  # noqa: E402
from core.update_state import build_paths, read_pending  # noqa: E402


def _install_with_zip(tmp_path: Path) -> tuple[Path, Path]:
    install = tmp_path / "inst"
    cfg = install / "config"
    cfg.mkdir(parents=True)
    (install / "VERSION.txt").write_text("1.0.0.0\n", encoding="utf-8")
    (cfg / "ui_update_check.json").write_text("{}", encoding="utf-8")
    zip_path = install / "patch.zip"
    zip_path.write_bytes(b"patch")
    return install, zip_path


def _status_with_zip(zip_path: Path) -> dict[str, Any]:
    return {
        "needs_bin_update": True,
        "installed_bin": "1.0.0.0",
        "latest_bin_version": "1.0.0.1",
        "bin_zip_path": str(zip_path),
        "bin_apply_mode": "patch",
        "catalog_path": "",
    }


def test_start_interactive_bin_apply_from_status_success_waits_ui_ready(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    install, zip_path = _install_with_zip(tmp_path)
    ready = tmp_path / "ui_ready.json"
    job_path = tmp_path / "job.json"
    job_path.write_text(
        json.dumps({"UiReadyPath": str(ready)}),
        encoding="utf-8",
    )

    monkeypatch.setattr(pu, "_install_root", lambda: install)
    monkeypatch.setattr(pu, "_resolve_bin_apply_paths_light", lambda _st: None)
    monkeypatch.setattr(pu, "_resolve_install_scope", lambda _r: "current")
    monkeypatch.setattr(pu, "_resolve_require_admin_interactive_apply", lambda _r, _s: False)
    monkeypatch.setattr(pu, "_append_update_log", lambda *_a, **_k: None)
    monkeypatch.setattr(
        pu,
        "spawn_apply_pending_via_hc_updater",
        lambda *_a, **_k: (True, str(job_path)),
    )
    wait_calls: list[Path] = []

    def _wait(path: Path, **kw: Any) -> bool:
        wait_calls.append(path)
        ready.write_text(json.dumps({"ready": True}), encoding="utf-8")
        return True

    monkeypatch.setattr(pu, "wait_for_updater_ui_ready", _wait)

    res = pu._start_interactive_bin_apply_from_status(
        _status_with_zip(zip_path),
        source="interactive_confirm",
    )

    assert res.get("ok") is True
    assert read_pending(build_paths(install)) is not None
    assert wait_calls == [ready]
    assert not ready.is_file()


def test_start_interactive_bin_apply_from_status_spawn_failure_clears_pending(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    install, zip_path = _install_with_zip(tmp_path)

    monkeypatch.setattr(pu, "_install_root", lambda: install)
    monkeypatch.setattr(pu, "_resolve_bin_apply_paths_light", lambda _st: None)
    monkeypatch.setattr(pu, "_resolve_install_scope", lambda _r: "current")
    monkeypatch.setattr(pu, "_resolve_require_admin_interactive_apply", lambda _r, _s: False)
    monkeypatch.setattr(pu, "_append_update_log", lambda *_a, **_k: None)
    monkeypatch.setattr(
        pu,
        "spawn_apply_pending_via_hc_updater",
        lambda *_a, **_k: (False, "spawn failed"),
    )

    res = pu._start_interactive_bin_apply_from_status(
        _status_with_zip(zip_path),
        source="test",
    )

    assert res.get("ok") is False
    assert res.get("error_kind") == "spawn"
    assert read_pending(build_paths(install)) is None


def test_start_interactive_bin_apply_from_status_ui_ready_timeout_still_ok(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    install, zip_path = _install_with_zip(tmp_path)
    ready = tmp_path / "ui_ready.json"
    job_path = tmp_path / "job.json"
    job_path.write_text(json.dumps({"UiReadyPath": str(ready)}), encoding="utf-8")

    monkeypatch.setattr(pu, "_install_root", lambda: install)
    monkeypatch.setattr(pu, "_resolve_bin_apply_paths_light", lambda _st: None)
    monkeypatch.setattr(pu, "_resolve_install_scope", lambda _r: "current")
    monkeypatch.setattr(pu, "_resolve_require_admin_interactive_apply", lambda _r, _s: False)
    monkeypatch.setattr(pu, "_append_update_log", lambda *_a, **_k: None)
    monkeypatch.setattr(
        pu,
        "spawn_apply_pending_via_hc_updater",
        lambda *_a, **_k: (True, str(job_path)),
    )
    monkeypatch.setattr(pu, "wait_for_updater_ui_ready", lambda *_a, **_k: False)

    res = pu._start_interactive_bin_apply_from_status(
        _status_with_zip(zip_path),
        source="interactive_confirm",
    )

    assert res.get("ok") is True


def test_start_interactive_bin_apply_from_status_invalid_root(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(pu, "_install_root", lambda: None)
    res = pu._start_interactive_bin_apply_from_status(
        {"bin_zip_path": "/missing.zip"},
        source="test",
    )
    assert res.get("ok") is False
    assert res.get("error_kind") == "invalid"


def test_spawn_interactive_apply_pending_and_wait_no_pending_returns_false(
    tmp_path: Path,
) -> None:
    install = tmp_path / "empty"
    install.mkdir()
    assert pu._spawn_interactive_apply_pending_and_wait(install, source="test") is False


def test_run_excel_startup_update_sequence_uses_busy_when_packaged(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    calls: list[str] = []

    monkeypatch.setenv(ENV_PACKAGED, "1")
    monkeypatch.setattr(
        pu,
        "_run_with_update_busy_ui",
        lambda msg, fn, **kw: (calls.append(str(msg)), fn())[1],
    )
    monkeypatch.setattr(
        pu,
        "maybe_check_updates_on_startup",
        lambda **kw: calls.append(f"startup_check reuse={kw.get('reuse_busy')}"),
    )

    order: list[str] = []

    pu.run_excel_startup_update_sequence(
        owner_hwnd=100,
        sheet_id="s",
        bootstrap_apply_fn=lambda: order.append("bootstrap"),
        ensure_bridge_fn=lambda: order.append("bridge"),
        register_book_fn=lambda: order.append("register"),
    )

    assert order == ["bootstrap", "bridge", "register"]
    assert any("startup_check reuse=True" in c for c in calls)
    assert len(calls) >= 2


def test_run_excel_startup_update_sequence_skips_busy_when_not_packaged(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.delenv(ENV_PACKAGED, raising=False)
    busy_calls: list[str] = []
    order: list[str] = []

    monkeypatch.setattr(
        pu,
        "_run_with_update_busy_ui",
        lambda *_a, **_k: busy_calls.append("busy"),
    )
    monkeypatch.setattr(
        pu,
        "maybe_check_updates_on_startup",
        lambda **_k: order.append("check"),
    )

    pu.run_excel_startup_update_sequence(
        owner_hwnd=0,
        sheet_id="_",
        bootstrap_apply_fn=lambda: order.append("bootstrap"),
        ensure_bridge_fn=lambda: order.append("bridge"),
        register_book_fn=lambda: order.append("register"),
    )

    assert order == ["bootstrap", "bridge", "register"]
    assert busy_calls == []
    assert "check" not in order


def test_maybe_check_updates_on_startup_reuse_busy_skips_check_busy(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch,
) -> None:
    install = tmp_path / "inst"
    install.mkdir()
    (install / "config").mkdir()
    (install / "config" / "ui_update_check.json").write_text("{}", encoding="utf-8")

    monkeypatch.setenv(ENV_PACKAGED, "1")
    pu._skip_startup_version_check_this_launch = False
    monkeypatch.setattr(pu, "_install_root", lambda: install)
    monkeypatch.setattr(pu, "maybe_show_updater_result_from_previous_run", lambda *_a, **_k: None)
    monkeypatch.setattr(pu, "discard_bin_apply_success_marker_if_present", lambda: None)
    monkeypatch.setattr(
        pu,
        "check_for_updates",
        lambda **_k: {"ok": True, "needs_bin_update": False},
    )
    nested_busy: list[str] = []
    monkeypatch.setattr(
        pu,
        "_run_with_update_busy_ui",
        lambda *_a, **_k: nested_busy.append("called"),
    )
    applied: list[dict[str, Any]] = []

    monkeypatch.setattr(
        pu,
        "_apply_startup_update_check_result",
        lambda st, **kw: applied.append(st),
    )

    pu.maybe_check_updates_on_startup(owner_hwnd=1, sheet_id="s", reuse_busy=True)

    assert nested_busy == []
    assert applied and applied[0].get("needs_bin_update") is False


def test_maybe_check_updates_on_startup_without_reuse_wraps_check_in_busy(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch,
) -> None:
    install = tmp_path / "inst"
    install.mkdir()
    (install / "config").mkdir()
    (install / "config" / "ui_update_check.json").write_text("{}", encoding="utf-8")

    monkeypatch.setenv(ENV_PACKAGED, "1")
    pu._skip_startup_version_check_this_launch = False
    monkeypatch.setattr(pu, "_install_root", lambda: install)
    monkeypatch.setattr(pu, "maybe_show_updater_result_from_previous_run", lambda *_a, **_k: None)
    monkeypatch.setattr(pu, "discard_bin_apply_success_marker_if_present", lambda: None)

    st_out = {"ok": True, "needs_bin_update": False}

    def _check(**_k: Any) -> dict[str, Any]:
        return st_out

    monkeypatch.setattr(pu, "check_for_updates", _check)
    busy_msgs: list[str] = []

    def _busy(msg: str, fn: Any, **kw: Any) -> dict[str, Any]:
        busy_msgs.append(msg)
        return fn()

    monkeypatch.setattr(pu, "_run_with_update_busy_ui", _busy)
    monkeypatch.setattr(pu, "_apply_startup_update_check_result", lambda *_a, **_k: None)

    pu.maybe_check_updates_on_startup(owner_hwnd=1, sheet_id="s", reuse_busy=False)

    assert len(busy_msgs) == 1
    assert "更新" in busy_msgs[0] or "確認" in busy_msgs[0]


def test_ui_ready_path_from_apply_pending_job_reads_payload(tmp_path: Path) -> None:
    job = tmp_path / "job.json"
    ready = tmp_path / "ready.json"
    job.write_text(json.dumps({"UiReadyPath": str(ready)}), encoding="utf-8")
    assert pu.ui_ready_path_from_apply_pending_job(job) == ready
    assert pu.ui_ready_path_from_apply_pending_job(tmp_path / "missing.json") is None
