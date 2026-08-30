# -*- coding: utf-8 -*-
"""更新確認: light check と prepare 分割・sha256 回数の回帰。"""
from __future__ import annotations

import hashlib
import json
import sys
from pathlib import Path
from typing import Any

import pytest

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from core import packaged_update as pu  # noqa: E402


def _sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    h.update(path.read_bytes())
    return h.hexdigest().lower()


def _make_catalog_tree(tmp_path: Path) -> tuple[Path, Path, dict[str, Any]]:
    install = tmp_path / "install"
    install.mkdir()
    (install / "VERSION.txt").write_text("1.0.0.0\n", encoding="utf-8")
    deploy = tmp_path / "deploy"
    sub = deploy / "releases" / "1.0.1"
    sub.mkdir(parents=True)
    patch_zip = sub / "bin_1.0.0_1.0.1_d.zip"
    patch_zip.write_text("patch-payload", encoding="utf-8")
    patch_sha = _sha256_file(patch_zip)
    full_zip = sub / "bin_1.0.1_full.zip"
    full_zip.write_bytes(b"full-payload-bytes")
    full_sha = _sha256_file(full_zip)
    cat = deploy / "catalog.json"
    catalog_obj = {
        "schema_version": 3,
        "set_version": "1.0.1",
        "bin": {
            "latest_version": "1.0.1.0",
            "full": {
                "relative_path": "releases/1.0.1/bin_1.0.1_full.zip",
                "sha256": full_sha,
            },
            "patch": {
                "relative_path": "releases/1.0.1/bin_1.0.0_1.0.1_d.zip",
                "sha256": patch_sha,
                "from_min_version": "1.0.0.0",
                "from_max_version": "1.0.0.0",
            },
        },
    }
    cat.write_text(json.dumps(catalog_obj), encoding="utf-8")
    cfg_dir = install / "config"
    cfg_dir.mkdir()
    (cfg_dir / "catalog_path.txt").write_text(str(cat.resolve()) + "\n", encoding="utf-8")
    return install, cat, catalog_obj


def test_hint_bin_apply_mode_patch_without_sha256(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    _install, cat, catalog_obj = _make_catalog_tree(tmp_path)

    def _boom(_p: Path) -> str:
        raise AssertionError("sha256 must not run in hint")

    monkeypatch.setattr(pu, "_sha256_file", _boom)
    hint = pu._hint_bin_apply_mode(catalog_obj, cat, "1.0.0.0")
    assert hint == "patch"


def test_check_for_updates_light_skips_sha256_and_zip_paths(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    install, _cat, _catalog_obj = _make_catalog_tree(tmp_path)
    monkeypatch.setenv("HC_INSTALL_ROOT", str(install))
    sha_calls: list[str] = []

    def _counting_sha(p: Path) -> str:
        sha_calls.append(str(p))
        return _sha256_file(p)

    monkeypatch.setattr(pu, "_sha256_file", _counting_sha)
    st = pu.check_for_updates(source="test", notify_offline=False)
    assert st["ok"] is True
    assert st["needs_bin_update"] is True
    assert st.get("bin_zip_path") is None
    assert st.get("bin_update_prepare_error") is None
    assert st.get("bin_apply_mode") in ("patch", "full", None)
    assert sha_calls == []


def test_prepare_bin_update_status_hashes_once_for_patch(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    install, cat, _catalog_obj = _make_catalog_tree(tmp_path)
    monkeypatch.setenv("HC_INSTALL_ROOT", str(install))
    sha_calls: list[str] = []

    def _counting_sha(p: Path) -> str:
        sha_calls.append(str(p))
        return _sha256_file(p)

    monkeypatch.setattr(pu, "_sha256_file", _counting_sha)
    st: dict[str, Any] = {
        "needs_bin_update": True,
        "catalog_path": str(cat),
        "installed_bin": "1.0.0.0",
        "latest_bin_version": "1.0.1.0",
    }
    pu.prepare_bin_update_status(st)
    assert st.get("bin_apply_mode") == "patch"
    assert st.get("bin_zip_path")
    assert st.get("bin_update_prepare_error") is None
    assert len(sha_calls) == 1
    assert "bin_1.0.0_1.0.1_d.zip" in sha_calls[0].replace("\\", "/")


def test_prepare_bin_update_status_full_only_when_patch_ineligible(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    install, cat, _catalog_obj = _make_catalog_tree(tmp_path)
    monkeypatch.setenv("HC_INSTALL_ROOT", str(install))
    sha_calls: list[str] = []

    def _counting_sha(p: Path) -> str:
        sha_calls.append(str(p))
        return _sha256_file(p)

    monkeypatch.setattr(pu, "_sha256_file", _counting_sha)
    st: dict[str, Any] = {
        "needs_bin_update": True,
        "catalog_path": str(cat),
        "installed_bin": "1.0.1.0",
        "latest_bin_version": "1.0.1.0",
    }
    pu.prepare_bin_update_status(st)
    assert st.get("bin_apply_mode") == "full"
    assert len(sha_calls) == 1
    assert "bin_1.0.1_full.zip" in sha_calls[0].replace("\\", "/")


def test_resolve_bin_full_zip_path_skips_sha256(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    _install, cat, catalog_obj = _make_catalog_tree(tmp_path)

    def _boom(_p: Path) -> str:
        raise AssertionError("sha256 must not run in resolve-only")

    monkeypatch.setattr(pu, "_sha256_file", _boom)
    zp, zsha, err = pu._resolve_bin_full_zip_path(catalog_obj, cat)
    assert err is None
    assert zp is not None
    assert zsha
