# -*- coding: utf-8 -*-
from __future__ import annotations

import sys
from pathlib import Path

import pytest

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from svc.svc_data_agg_extract import file_paths_for_source_extract  # noqa: E402


def _cell_src(pattern: str) -> dict:
    return {
        "type": "cell",
        "ui_scenario_source_v1": {
            "file_pattern": pattern,
            "file_name_rule": "含む",
        },
    }


def test_file_paths_for_source_filters_by_pattern() -> None:
    paths = [
        "C:/a/光特性履歴.xlsx",
        "C:/a/紐づけ履歴.xlsx",
        "C:/a/other.xlsm",
    ]
    out = file_paths_for_source_extract(paths, _cell_src("光特性"))
    assert out == ["C:/a/光特性履歴.xlsx"]
    out2 = file_paths_for_source_extract(paths, _cell_src("紐づけ"))
    assert out2 == ["C:/a/紐づけ履歴.xlsx"]


def test_file_paths_for_source_empty_pattern_keeps_all() -> None:
    paths = ["C:/a/one.xlsx", "C:/a/two.xlsx"]
    out = file_paths_for_source_extract(paths, _cell_src(""))
    assert out == paths


def test_file_paths_for_source_dedupes_preserving_order() -> None:
    paths = ["C:/a/dup.xlsx", "C:/a/dup.xlsx", "C:/a/光特性.xlsx"]
    out = file_paths_for_source_extract(paths, _cell_src("光特性"))
    assert out == ["C:/a/光特性.xlsx"]
