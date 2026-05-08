# -*- coding: utf-8 -*-
"""Tests for bin patch vs full selection in core.packaged_update."""

from __future__ import annotations

import hashlib
import json
import shutil
import zipfile
from pathlib import Path

import pytest

from core.packaged_update import (
    _catalog_bootstrap_latest,
    _materialize_patch_zip_for_worker,
    _parse_bootstrap_triplet,
    _patch_meta_eligible,
    _prepare_bin_apply,
    needs_bootstrap_update_bool,
)


def test_needs_bootstrap_update_bool() -> None:
    assert needs_bootstrap_update_bool("1.0.0", "1.0.1") is True
    assert needs_bootstrap_update_bool("1.0.1", "1.0.0") is False
    assert needs_bootstrap_update_bool(None, "1.0.0") is True
    assert needs_bootstrap_update_bool("1.0.0", "") is False


def test_patch_meta_eligible_exact_bounds() -> None:
    p = {"from_min_version": "1.0.0", "from_max_version": "1.0.0"}
    assert _patch_meta_eligible("1.0.0", p) is True
    assert _patch_meta_eligible("1.0.1", p) is False


def test_patch_meta_eligible_range() -> None:
    p = {"from_min_version": "1.1.0", "from_max_version": "1.3.99"}
    assert _patch_meta_eligible("1.2.0", p) is True
    assert _patch_meta_eligible("1.1.0", p) is True
    assert _patch_meta_eligible("1.3.99", p) is True
    assert _patch_meta_eligible("1.0.9", p) is False
    assert _patch_meta_eligible("1.4.0", p) is False


def test_patch_meta_eligible_requires_at_least_one_bound() -> None:
    assert _patch_meta_eligible("1.0.0", {}) is False
    assert _patch_meta_eligible("1.0.0", {"from_min_version": "", "from_max_version": ""}) is False


def test_patch_meta_eligible_min_only() -> None:
    p = {"from_min_version": "2.0.0", "from_max_version": ""}
    assert _patch_meta_eligible("2.0.0", p) is True
    assert _patch_meta_eligible("1.9.9", p) is False


def test_patch_meta_eligible_max_only() -> None:
    p = {"from_min_version": "", "from_max_version": "1.5.0"}
    assert _patch_meta_eligible("1.0.0", p) is True
    assert _patch_meta_eligible("1.5.0", p) is True
    assert _patch_meta_eligible("1.5.1", p) is False


def test_parse_bootstrap_triplet() -> None:
    assert _parse_bootstrap_triplet("1.2.3") == (1, 2, 3)
    assert _parse_bootstrap_triplet("01.02.003") == (1, 2, 3)
    assert _parse_bootstrap_triplet("1.2") is None
    assert _parse_bootstrap_triplet("1.2.3.4") is None


def test_catalog_bootstrap_latest_optional() -> None:
    assert _catalog_bootstrap_latest({}) is None
    assert _catalog_bootstrap_latest({"bootstrap": {}}) is None
    assert _catalog_bootstrap_latest({"bootstrap": {"latest_version": "1.2.3"}}) == "1.2.3"


def _sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    h.update(path.read_bytes())
    return h.hexdigest().lower()


def test_prepare_bin_prefers_patch_when_eligible(tmp_path: Path) -> None:
    """When patch zip exists and sha matches, mode is patch."""
    deploy = tmp_path / "deploy"
    sub = deploy / "releases" / "1.0.1"
    sub.mkdir(parents=True)
    patch_zip = sub / "bin_1.0.0_1.0.1_d.zip"
    patch_zip.write_text("patch-payload", encoding="utf-8")
    sha = _sha256_file(patch_zip)
    full_zip = sub / "bin_1.0.1_full.zip"
    full_zip.write_bytes(b"x")
    full_sha = _sha256_file(full_zip)

    cat = deploy / "catalog.json"
    catalog_obj = {
        "schema_version": 3,
        "bin": {
            "latest_version": "1.0.1",
            "full": {
                "relative_path": "releases/1.0.1/bin_1.0.1_full.zip",
                "sha256": full_sha,
            },
            "patch": {
                "relative_path": "releases/1.0.1/bin_1.0.0_1.0.1_d.zip",
                "sha256": sha,
                "from_min_version": "1.0.0",
                "from_max_version": "1.0.0",
            },
        },
    }
    cat.write_text(json.dumps(catalog_obj), encoding="utf-8")

    mode, zp, zsha, err = _prepare_bin_apply(catalog_obj, cat, "1.0.0", None)
    assert err is None
    assert mode == "patch"
    assert zp == patch_zip.resolve()
    assert zsha == sha


def test_prepare_bin_falls_back_to_full_when_patch_not_eligible(tmp_path: Path) -> None:
    deploy = tmp_path / "deploy"
    sub = deploy / "releases" / "1.0.1"
    sub.mkdir(parents=True)
    full_zip = sub / "bin_1.0.1_full.zip"
    full_zip.write_bytes(b"full-bytes-here")
    full_sha = _sha256_file(full_zip)
    patch_zip = sub / "bin_1.0.0_1.0.1_d.zip"
    patch_zip.write_text("p", encoding="utf-8")
    psha = _sha256_file(patch_zip)

    cat = deploy / "catalog.json"
    catalog_obj = {
        "bin": {
            "latest_version": "1.0.1",
            "full": {"relative_path": "releases/1.0.1/bin_1.0.1_full.zip", "sha256": full_sha},
            "patch": {
                "relative_path": "releases/1.0.1/bin_1.0.0_1.0.1_d.zip",
                "sha256": psha,
                "from_min_version": "1.0.0",
                "from_max_version": "1.0.0",
            },
        },
    }
    cat.write_text(json.dumps(catalog_obj), encoding="utf-8")

    mode, zp, zsha, err = _prepare_bin_apply(catalog_obj, cat, "1.0.1", None)
    assert err is None
    assert mode == "full"
    assert zp == full_zip.resolve()
    assert zsha == full_sha


def test_prepare_bin_patch_sha_mismatch_falls_back_to_full(tmp_path: Path) -> None:
    deploy = tmp_path / "deploy"
    sub = deploy / "releases" / "1.0.1"
    sub.mkdir(parents=True)
    full_zip = sub / "bin_1.0.1_full.zip"
    full_zip.write_bytes(b"full2")
    full_sha = _sha256_file(full_zip)
    patch_zip = sub / "bin_d.zip"
    patch_zip.write_text("p2", encoding="utf-8")

    cat = deploy / "catalog.json"
    catalog_obj = {
        "bin": {
            "latest_version": "1.0.1",
            "full": {"relative_path": "releases/1.0.1/bin_1.0.1_full.zip", "sha256": full_sha},
            "patch": {
                "relative_path": "releases/1.0.1/bin_d.zip",
                "sha256": "0" * 64,
                "from_min_version": "1.0.0",
                "from_max_version": "1.0.0",
            },
        },
    }
    cat.write_text(json.dumps(catalog_obj), encoding="utf-8")

    mode, zp, _, err = _prepare_bin_apply(catalog_obj, cat, "1.0.0", None)
    assert err is None
    assert mode == "full"
    assert zp == full_zip.resolve()


def test_materialize_patch_manifest_builds_worker_zip(tmp_path: Path) -> None:
    pytest.importorskip("bsdiff4", reason="bsdiff4 is required")
    import bsdiff4

    install_root = tmp_path / "install"
    old_fp = install_root / "app" / "bin" / "a.dll"
    old_fp.parent.mkdir(parents=True, exist_ok=True)
    old_bytes = b"HELLO_OLD"
    new_bytes = b"HELLO_NEW_CONTENT"
    old_fp.write_bytes(old_bytes)

    patch_zip = tmp_path / "patch.zip"
    addin_data = b"addin-new"
    old_sha = hashlib.sha256(old_bytes).hexdigest().lower()
    new_sha = hashlib.sha256(new_bytes).hexdigest().lower()
    patch_bytes = bsdiff4.diff(old_bytes, new_bytes)
    manifest = {
        "version": 1,
        "patch_format": "bsdiff4-manifest-v1",
        "base_version": "1.0.0",
        "target_version": "1.0.1",
        "entries": [
            {
                "path": "app/bin/a.dll",
                "op": "bsdiff",
                "old_sha256": old_sha,
                "new_sha256": new_sha,
                "patch_file": "patches/a.bsdiff",
            },
            {
                "path": "addin/addin.xlam",
                "op": "copy",
                "new_sha256": hashlib.sha256(addin_data).hexdigest().lower(),
                "file": "files/addin/addin.xlam",
            },
        ],
    }
    with zipfile.ZipFile(patch_zip, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("manifest.json", json.dumps(manifest, ensure_ascii=False))
        zf.writestr("patches/a.bsdiff", patch_bytes)
        zf.writestr("files/addin/addin.xlam", addin_data)

    out_zip, cleanup_dir, stats, err = _materialize_patch_zip_for_worker(
        install_root=install_root, patch_zip=patch_zip, target_bin_version="1.0.1"
    )
    try:
        assert err is None
        assert out_zip != patch_zip
        assert out_zip.is_file()
        assert cleanup_dir.is_dir()
        assert isinstance(stats, dict)
        assert stats["bsdiff"] == 1
        assert stats["copy"] == 1
        assert stats["target_version"] == "1.0.1"

        with zipfile.ZipFile(out_zip, "r") as zf:
            assert zf.read("app/bin/a.dll") == new_bytes
            assert zf.read("addin/addin.xlam") == addin_data
            assert zf.read("VERSION.txt").decode("utf-8").strip() == "1.0.1"
    finally:
        shutil.rmtree(cleanup_dir, ignore_errors=True)


def test_materialize_patch_legacy_keeps_original_zip(tmp_path: Path) -> None:
    install_root = tmp_path / "install"
    install_root.mkdir(parents=True, exist_ok=True)
    patch_zip = tmp_path / "legacy_patch.zip"
    with zipfile.ZipFile(patch_zip, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("app/bin/a.exe", b"x")

    out_zip, cleanup_dir, stats, err = _materialize_patch_zip_for_worker(
        install_root=install_root, patch_zip=patch_zip, target_bin_version="1.0.1"
    )
    assert err is None
    assert out_zip == patch_zip
    assert str(cleanup_dir) in (".", "")
    assert stats is None
