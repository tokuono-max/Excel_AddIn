# -*- coding: utf-8 -*-
"""tools/release/full.bat must stay ASCII + CRLF (cmd-safe)."""
from __future__ import annotations

from pathlib import Path

from tools.release.assert_bat_ascii_crlf import check_bat

_ROOT = Path(__file__).resolve().parents[1]


def test_tools_release_full_bat_ascii_crlf() -> None:
    path = _ROOT / "tools" / "release" / "full.bat"
    assert path.is_file(), path
    errs = check_bat(path)
    assert not errs, errs


def test_release_full_bat_wrapper_ascii_crlf() -> None:
    path = _ROOT / "release" / "full.bat"
    assert path.is_file(), path
    errs = check_bat(path)
    assert not errs, errs
    text = path.read_text(encoding="ascii")
    assert "tools\\release\\full.bat" in text
