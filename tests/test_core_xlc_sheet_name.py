# -*- coding: utf-8 -*-
"""Excel シート名整形（core_xlc）のユニットテスト。"""
from __future__ import annotations

from core import core_xlc as xlc


def test_sanitize_excel_sheet_name_truncates_to_31() -> None:
    long_name = "あ" * 50
    assert len(xlc.sanitize_excel_sheet_name(long_name)) == 31


def test_sanitize_excel_sheet_name_strips_invalid_chars() -> None:
    assert xlc.sanitize_excel_sheet_name("  A/B[C]:x  ") == "ABCx"


def test_unique_excel_sheet_name_in_names_suffix_within_31() -> None:
    existing = {"ABCDEF"}
    out = xlc.unique_excel_sheet_name_in_names(existing, "ABCDEF")
    assert len(out) <= 31
    assert out not in existing
    assert out.startswith("ABCD")


def test_excel_sheet_name_for_split_part_reserves_suffix() -> None:
    base = "X" * 31
    name = xlc.excel_sheet_name_for_split_part(base, 12)
    assert len(name) == 31
    assert name.endswith("-12")


def test_excel_sheet_name_for_split_part_single_digit() -> None:
    base = "データファイル名がとても長い場合のテスト用"
    name = xlc.excel_sheet_name_for_split_part(base, 1)
    assert len(name) <= 31
    assert name.endswith("-1")
