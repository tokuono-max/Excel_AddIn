# -*- coding: utf-8 -*-
"""core_excel_text のユニットテスト。"""
from __future__ import annotations

from core.core_excel_text import as_excel_forced_text, matrix_as_excel_forced_text, scalar_to_text


def test_as_excel_forced_text_prefixes_non_empty() -> None:
    assert as_excel_forced_text("00123") == "'00123"
    assert as_excel_forced_text("2024-01-01 12:00:00") == "'2024-01-01 12:00:00"
    assert as_excel_forced_text("1000000000000000") == "'1000000000000000"
    assert as_excel_forced_text("hello") == "'hello"


def test_as_excel_forced_text_empty_and_already_frozen() -> None:
    assert as_excel_forced_text("") == ""
    assert as_excel_forced_text(None) == ""
    assert as_excel_forced_text("'123") == "'123"


def test_as_excel_forced_text_float_without_scientific() -> None:
    assert as_excel_forced_text(1000000000000000.0) == "'1000000000000000"


def test_scalar_to_text_keeps_short_decimal() -> None:
    assert scalar_to_text(2020.4) == "2020.4"
    assert as_excel_forced_text(2020.4) == "'2020.4"
    assert "03999" not in scalar_to_text(2020.4)


def test_scalar_to_text_avoids_scientific_for_tiny_float() -> None:
    s = scalar_to_text(1e-7)
    assert "e" not in s.lower()
    assert s.startswith("0.0000001")


def test_matrix_as_excel_forced_text() -> None:
    out = matrix_as_excel_forced_text([["a", ""], ["1", "2"]])
    assert out == [["'a", ""], ["'1", "'2"]]
