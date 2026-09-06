# -*- coding: utf-8 -*-
from __future__ import annotations

import json
import sys
from pathlib import Path

_root = Path(__file__).resolve().parents[1]
if str(_root / "tools") not in sys.path:
    sys.path.insert(0, str(_root / "tools"))

from set_catalog_reinstall_flag import (
    apply_reinstall_flag,
    default_catalog_path,
    parse_flag,
    set_catalog_reinstall_flag,
)


def test_parse_flag() -> None:
    assert parse_flag("on") is True
    assert parse_flag("OFF") is False


def test_apply_only_sets_flag() -> None:
    cat = {"bin": {"latest_version": "1.1.11"}}
    out = apply_reinstall_flag(cat, enabled=True)
    assert out["bin"]["require_uninstall_reinstall"] is True
    assert "installer" not in out
    out = apply_reinstall_flag(out, enabled=False)
    assert out["bin"]["require_uninstall_reinstall"] is False


def test_set_catalog_file_roundtrip(tmp_path: Path) -> None:
    cat = tmp_path / "catalog.json"
    cat.write_text(
        json.dumps({"schema_version": 3, "bin": {"latest_version": "1.1.11"}}, ensure_ascii=False),
        encoding="utf-8",
    )
    set_catalog_reinstall_flag(tmp_path, enabled=True)
    data = json.loads(cat.read_text(encoding="utf-8"))
    assert data["bin"]["require_uninstall_reinstall"] is True
    assert "installer" not in data
    set_catalog_reinstall_flag(cat, enabled=False)
    data = json.loads(cat.read_text(encoding="utf-8"))
    assert data["bin"]["require_uninstall_reinstall"] is False


def test_default_catalog_path_uses_version_txt(tmp_path: Path) -> None:
    (tmp_path / "VERSION.txt").write_text("1.1.10.6\n", encoding="utf-8")
    dest = tmp_path / "dist" / "releases" / "1.1.10.6"
    dest.mkdir(parents=True)
    (dest / "catalog.json").write_text("{}", encoding="utf-8")
    got = default_catalog_path(tmp_path)
    assert got == dest / "catalog.json"
