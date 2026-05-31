# -*- coding: utf-8 -*-
from __future__ import annotations

import sys
from pathlib import Path
from unittest.mock import patch

_ROOT = Path(__file__).resolve().parents[1]
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from core import win_running_file_replace as wr


def test_collect_self_sidecar_includes_runtime_dlls(tmp_path: Path) -> None:
    bin_dir = tmp_path / "app" / "bin"
    bin_dir.mkdir(parents=True)
    exe = bin_dir / "python.exe"
    exe.write_bytes(b"MZ")
    (bin_dir / "libcrypto-3.dll").write_bytes(b"dll")
    with patch.object(sys, "executable", str(exe)):
        paths = wr.collect_self_sidecar_dst_paths()
    assert exe.resolve() in paths
    assert (bin_dir / "libcrypto-3.dll").resolve() in paths


def test_replace_via_sidecar_updates_target(tmp_path: Path) -> None:
    dst = tmp_path / "libcrypto-3.dll"
    dst.write_bytes(b"old")
    src = tmp_path / "new.dll"
    src.write_bytes(b"new")
    wr.replace_via_sidecar(src, dst)
    assert dst.read_bytes() == b"new"


def test_copy_file_proactive_sidecar_for_runtime_dll(tmp_path: Path) -> None:
    bin_dir = tmp_path / "app" / "bin"
    bin_dir.mkdir(parents=True)
    exe = bin_dir / "python.exe"
    exe.write_bytes(b"MZ")
    dst = bin_dir / "libcrypto-3.dll"
    dst.write_bytes(b"old")
    src = tmp_path / "src" / "libcrypto-3.dll"
    src.parent.mkdir()
    src.write_bytes(b"new")
    with patch.object(sys, "executable", str(exe)):
        proactive = wr.collect_self_sidecar_dst_paths()
    wr.copy_file_with_sharing_fallback(src, dst, proactive_sidecar=proactive)
    assert dst.read_bytes() == b"new"


def test_copy_file_win32_fallback_uses_sidecar(tmp_path: Path) -> None:
    dst = tmp_path / "target.dll"
    dst.write_bytes(b"old")
    src = tmp_path / "source.dll"
    src.write_bytes(b"new")
    calls = {"n": 0}

    def fake_copy2(s: Path, d: Path) -> None:
        calls["n"] += 1
        if calls["n"] == 1:
            err = OSError(13, "locked")
            err.winerror = 32  # type: ignore[attr-defined]
            raise err
        dst.write_bytes(src.read_bytes())

    with patch.object(wr.shutil, "copy2", side_effect=fake_copy2):
        wr.copy_file_with_sharing_fallback(src, dst)
    assert dst.read_bytes() == b"new"


def test_cleanup_stale_sidecar_files(tmp_path: Path) -> None:
    bin_dir = tmp_path / "bin"
    bin_dir.mkdir()
    stale = bin_dir / "hc_updater.exe.was_running_999"
    stale.write_bytes(b"x")
    wr.cleanup_stale_sidecar_files(bin_dir)
    assert not stale.is_file()
