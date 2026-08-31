# -*- coding: utf-8 -*-
"""apply_pending: ribbon defer (patch) と hc_updater inline (full) の版アップ経路シミュレーション。"""
from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any
from unittest.mock import patch

import pytest

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from bootstrap import update_bootstrap as ub  # noqa: E402
from core import packaged_update as pu  # noqa: E402
from core import update_process_cleanup as upc  # noqa: E402
from core.update_state import build_paths, read_pending, write_pending  # noqa: E402


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
    (install / "VERSION.txt").write_text("1.1.9.5\n", encoding="utf-8")
    ui_cfg = {
        "MESSAGES": {
            "UPDATER_PHASE_WAIT_TITLE": "Excel の終了を待っています",
            "UPDATER_PHASE_WAIT_MESSAGE": "すべての Excel を閉じてください。",
            "UPDATER_PHASE_STOP_PROCESSES_MESSAGE": "関連プロセスを終了しています…",
            "PROGRESS_PREPARE_TITLE": "準備中",
            "PROGRESS_PREPARE_MSG": "更新に必要なファイルを用意しています。",
            "PROGRESS_PREPARE_MSG_PATCH_BUILD": "差分パッケージを構築しています。",
            "PROGRESS_DEFER_DONE_TITLE": "準備完了",
            "PROGRESS_DEFER_DONE_TEMPLATE": "Excel をすべて終了してください。",
            "PROGRESS_INLINE_DONE_TITLE": "完了",
            "PROGRESS_INLINE_DONE_MSG": "更新が完了しました。",
        }
    }
    (cfg_dir / "ui_update_check.json").write_text(
        json.dumps(ui_cfg, ensure_ascii=False), encoding="utf-8"
    )
    return install


_SVC_ONLY_MUTEX = {"main": False, "main_legacy": False, "svc": True, "ui": False}
_CLEAR_MUTEX = {"main": False, "main_legacy": False, "svc": False, "ui": False}


def _patch_pending(install: Path, *, skip_apply_confirm: bool = True) -> Path:
    paths = build_paths(install)
    patch_zip = install / "patch_src.zip"
    patch_zip.write_bytes(b"patch")
    write_pending(
        paths,
        {
            "schema_version": 2,
            "apply_scope": "bin+bootstrap",
            "skip_apply_confirm": skip_apply_confirm,
            "mode": "patch",
            "target_bin_version": "1.1.10",
            "catalog_path": "",
            "state": "downloaded",
            "patch": {"zip_path": str(patch_zip), "sha256": ""},
        },
    )
    return patch_zip


def _full_pending(install: Path) -> Path:
    paths = build_paths(install)
    full_zip = install / "full_src.zip"
    full_zip.write_bytes(b"full")
    write_pending(
        paths,
        {
            "schema_version": 2,
            "apply_scope": "bin+bootstrap",
            "mode": "full",
            "target_bin_version": "1.1.10",
            "catalog_path": "",
            "state": "downloaded",
            "full": {"zip_path": str(full_zip), "sha256": "abc"},
        },
    )
    return full_zip


def test_ribbon_patch_defer_legacy_spawn_blocked_in_svc(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """svc 内 in-process bin defer（レガシー 2 本 spawn）は拒否し pending を残す。"""
    install = _minimal_install(tmp_path)
    patch_zip = _patch_pending(install)
    paths = build_paths(install)
    worker_zip = install / "worker.zip"
    worker_zip.write_bytes(b"worker")

    monkeypatch.setattr(ub, "_ProgressUi", _FakeProgressUi)
    monkeypatch.setattr(ub, "_try_apply_bootstrap_swap", lambda *_a, **_k: (None, None))
    monkeypatch.setattr(ub, "mutex_snapshot", lambda: dict(_SVC_ONLY_MUTEX))
    monkeypatch.setattr(ub, "ensure_packaged_children_stopped", lambda *_a, **_k: None)
    monkeypatch.setattr(
        ub,
        "_resolve_payload",
        lambda *_a, **_k: (patch_zip, ""),
    )
    monkeypatch.setattr(
        ub,
        "materialize_manifest_patch_zip",
        lambda **_k: (worker_zip, None, {"changed": 1}, None),
    )
    monkeypatch.setattr(
        ub,
        "_persist_zip_for_hc_updater",
        lambda _z, _p: worker_zip,
    )
    spawn_calls: list[str] = []

    def _spawn(_root: Path, **kwargs: Any) -> None:
        spawn_calls.append(str(kwargs.get("apply_mode") or ""))

    monkeypatch.setattr(pu, "spawn_hc_updater_for_pending_bin_apply", _spawn)

    with patch.object(upc, "is_hc_svc_server_process", return_value=True):
        res = ub._apply_pending_update_impl(install)

    assert res.get("ok") is False
    assert res.get("error_code") == "E_LEGACY_DEFER_UNSUPPORTED"
    assert "legacy_defer_spawn_blocked" in paths.log_path.read_text(encoding="utf-8")
    assert spawn_calls == []
    assert read_pending(paths) is not None


def test_hc_updater_inline_full_apply_with_clear_mutex(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """hc_updater 経路 (full, inline): mutex 解放後はその場で bin 適用する。"""
    install = _minimal_install(tmp_path)
    full_zip = _full_pending(install)
    paths = build_paths(install)
    applied: list[str] = []

    monkeypatch.setenv("CSV_TOOL_APPLY_PENDING_INLINE_BIN", "1")
    monkeypatch.setattr(ub, "_ProgressUi", _FakeProgressUi)
    monkeypatch.setattr(
        ub,
        "_confirm_pending_apply_before_progress",
        lambda *_a, **_k: True,
    )
    monkeypatch.setattr(ub, "_try_apply_bootstrap_swap", lambda *_a, **_k: (None, None))
    monkeypatch.setattr(ub, "mutex_snapshot", lambda: dict(_CLEAR_MUTEX))
    monkeypatch.setattr(ub, "ensure_packaged_children_stopped", lambda *_a, **_k: None)
    monkeypatch.setattr(
        ub,
        "_resolve_payload",
        lambda *_a, **_k: (full_zip, "abc"),
    )

    def _apply_zip(_install: Path, _zip: Path, *_a: Any, **_k: Any) -> None:
        applied.append(_zip.name)

    monkeypatch.setattr(ub, "_apply_zip", _apply_zip)

    with patch.object(upc, "is_hc_svc_server_process", return_value=False):
        res = ub._apply_pending_update_impl(install)

    assert res.get("ok") is True
    assert res.get("applied") is True
    assert not res.get("deferred_to_updater")
    assert applied == ["full_src.zip"]
    assert read_pending(paths) is None

    log_text = paths.log_path.read_text(encoding="utf-8")
    assert "relax_svc_self=True" not in log_text
    assert "defer_bin_to_updater=False" in log_text
    assert "apply_result=success" in log_text


def test_ribbon_patch_without_relax_still_blocks_on_svc_mutex(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """2 段階確認（skip_apply_confirm なし）かつ svc mutex 残存時はブロックする。"""
    install = _minimal_install(tmp_path)
    _patch_pending(install, skip_apply_confirm=False)

    monkeypatch.setattr(ub, "_ProgressUi", _FakeProgressUi)
    monkeypatch.setattr(ub, "mutex_snapshot", lambda: dict(_SVC_ONLY_MUTEX))
    monkeypatch.setattr(ub, "ensure_packaged_children_stopped", lambda *_a, **_k: None)
    monkeypatch.setattr(
        ub,
        "_confirm_pending_apply_before_progress",
        lambda *_a, **_k: True,
    )

    with patch.object(upc, "is_hc_svc_server_process", return_value=False):
        with patch.object(upc, "is_hc_updater_process", return_value=False):
            res = ub._apply_pending_update_impl(install)

    assert res.get("ok") is False
    assert "blocked_by_running_process" in str(res.get("error") or "")


def test_hc_updater_defer_skips_mutex_gate_with_single_confirm(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """hc_updater + skip_apply_confirm: mutex 待ちを省略し defer 準備へ進む。"""
    install = _minimal_install(tmp_path)
    patch_zip = _patch_pending(install, skip_apply_confirm=True)
    paths = build_paths(install)
    worker_zip = install / "worker.zip"
    worker_zip.write_bytes(b"worker")

    monkeypatch.setattr(ub, "_ProgressUi", _FakeProgressUi)
    monkeypatch.setattr(ub, "_try_apply_bootstrap_swap", lambda *_a, **_k: (None, None))
    monkeypatch.setattr(
        ub,
        "_resolve_payload",
        lambda *_a, **_k: (patch_zip, ""),
    )
    monkeypatch.setattr(
        ub,
        "materialize_manifest_patch_zip",
        lambda **_k: (worker_zip, None, {"changed": 1}, None),
    )
    monkeypatch.setattr(
        ub,
        "_persist_zip_for_hc_updater",
        lambda _z, _p: worker_zip,
    )

    with patch.object(upc, "is_hc_svc_server_process", return_value=False):
        with patch.object(upc, "is_hc_updater_process", return_value=True):
            res = ub._apply_pending_update_impl(install)

    assert res.get("ok") is True
    assert res.get("deferred_to_updater") is True
    assert res.get("deferred_inline_bin_apply") is True
    assert str(res.get("worker_zip_path") or "").endswith("worker.zip")
    log_text = paths.log_path.read_text(encoding="utf-8")
    assert "skip_mutex_gate=True" in log_text
    assert "mutex_busy graceful_shutdown" not in log_text
    assert "deferred_inline_bin_apply" in log_text
