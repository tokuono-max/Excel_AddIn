# -*- coding: utf-8 -*-
from __future__ import annotations

import sys
from pathlib import Path

import pytest

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from svc.svc_data_agg_extract import (  # noqa: E402
    file_paths_for_source_extract,
    source_passes_file_name_filter,
)


def _cell_src(pattern: str, rule: str = "含む") -> dict:
    return {
        "type": "cell",
        "ui_scenario_source_v1": {
            "file_pattern": pattern,
            "file_name_rule": rule,
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


def test_file_pattern_multi_contains_or() -> None:
    """カンマ区切り「含む」はいずれかに該当（OR）。先頭/末尾カンマは無視。"""
    paths = [
        "C:/a/光特性履歴.xlsx",
        "C:/a/紐づけ履歴.xlsx",
        "C:/a/other.xlsm",
    ]
    out = file_paths_for_source_extract(paths, _cell_src(",光特性,紐づけ,"))
    assert out == ["C:/a/光特性履歴.xlsx", "C:/a/紐づけ履歴.xlsx"]


def test_file_pattern_multi_exact_or() -> None:
    paths = [
        "C:/a/光特性.xlsx",
        "C:/a/光特性履歴.xlsx",
        "C:/a/紐づけ.xlsx",
    ]
    out = file_paths_for_source_extract(paths, _cell_src("光特性,紐づけ", "完全一致"))
    assert out == ["C:/a/光特性.xlsx", "C:/a/紐づけ.xlsx"]


def test_file_pattern_multi_not_contains_and() -> None:
    """「含まない」はいずれのトークンも含まないファイルのみ。"""
    src = _cell_src("光特性,紐づけ", "含まない")
    assert source_passes_file_name_filter("C:/a/other.xlsx", src) is True
    assert source_passes_file_name_filter("C:/a/光特性履歴.xlsx", src) is False
    assert source_passes_file_name_filter("C:/a/紐づけ履歴.xlsx", src) is False


def test_file_pattern_only_commas_keeps_all() -> None:
    paths = ["C:/a/one.xlsx", "C:/a/two.xlsx"]
    out = file_paths_for_source_extract(paths, _cell_src(", , ,"))
    assert out == paths
