# -*- coding: utf-8 -*-
"""hc_updater full 適用（merge 経路）のユニットテスト。"""

from __future__ import annotations

import sys
from pathlib import Path
from unittest.mock import patch

_ROOT = Path(__file__).resolve().parents[1]
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from hc_updater import (
    _apply_delete_list,
    _copy_merge_tree,
    _mirror_tree,
    _running_from_tree_root,
)


def test_copy_merge_tree_updates_without_removing_dst_root(tmp_path: Path) -> None:
    dst = tmp_path / "app" / "bin"
    dst.mkdir(parents=True)
    (dst / "only_in_dst.txt").write_text("keep", encoding="utf-8")
    (dst / "hc_main.exe").write_text("old", encoding="utf-8")

    src = tmp_path / "src_bin"
    src.mkdir()
    (src / "hc_main.exe").write_text("new", encoding="utf-8")
    (src / "new_file.dll").write_text("dll", encoding="utf-8")

    log = tmp_path / "hc_update.log"
    _copy_merge_tree(src, dst, log)

    assert (dst / "only_in_dst.txt").read_text(encoding="utf-8") == "keep"
    assert (dst / "hc_main.exe").read_text(encoding="utf-8") == "new"
    assert (dst / "new_file.dll").read_text(encoding="utf-8") == "dll"


def test_apply_delete_list_removes_paths_under_install_root(tmp_path: Path) -> None:
    install = tmp_path / "inst"
    obsolete = install / "app" / "bin" / "obsolete.dll"
    obsolete.parent.mkdir(parents=True)
    obsolete.write_text("x", encoding="utf-8")
    keep = install / "app" / "bin" / "keep.exe"
    keep.write_text("y", encoding="utf-8")

    dl = tmp_path / "__delete_list.txt"
    dl.write_text("app/bin/obsolete.dll\n", encoding="utf-8")
    _apply_delete_list(install, dl)

    assert not obsolete.is_file()
    assert keep.is_file()


def test_running_from_tree_root_detects_executable_parent(tmp_path: Path) -> None:
    bin_dir = tmp_path / "app" / "bin"
    bin_dir.mkdir(parents=True)
    fake_exe = bin_dir / "hc_updater.exe"
    fake_exe.write_bytes(b"MZ")
    with patch.object(sys, "executable", str(fake_exe)):
        assert _running_from_tree_root(bin_dir) is True
    with patch.object(sys, "executable", str(tmp_path / "other" / "tool.exe")):
        assert _running_from_tree_root(bin_dir) is False


def test_mirror_tree_refuses_when_running_from_dst(tmp_path: Path) -> None:
    bin_dir = tmp_path / "app" / "bin"
    bin_dir.mkdir(parents=True)
    fake_exe = bin_dir / "hc_updater.exe"
    fake_exe.write_bytes(b"MZ")
    src = tmp_path / "src"
    src.mkdir()
    (src / "a.txt").write_text("a", encoding="utf-8")
    log = tmp_path / "log.txt"

    with patch.object(sys, "executable", str(fake_exe)):
        try:
            _mirror_tree(src, bin_dir, log)
            raised = False
        except RuntimeError as e:
            raised = True
            assert "mirror refused" in str(e)
    assert raised
