# -*- coding: utf-8 -*-
from __future__ import annotations

from core.excel_display_read import parse_excel_clipboard_tsv


def test_parse_single_cell() -> None:
    assert parse_excel_clipboard_tsv("hello") == [["hello"]]


def test_parse_tab_row() -> None:
    assert parse_excel_clipboard_tsv("a\tb\tc") == [["a", "b", "c"]]


def test_parse_multiline_crlf() -> None:
    text = "a\tb\r\nc\td"
    assert parse_excel_clipboard_tsv(text, expected_rows=2, expected_cols=2) == [
        ["a", "b"],
        ["c", "d"],
    ]


def test_parse_quoted_cell_with_tab() -> None:
    text = '"a\tinside"\tb'
    rows = parse_excel_clipboard_tsv(text)
    assert rows == [["a\tinside", "b"]]
