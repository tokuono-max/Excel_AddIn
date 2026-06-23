# -*- coding: utf-8 -*-
"""join_compare_display_key: 結合・照合比較用の表示本体文字列化。"""
from __future__ import annotations

import sys
from pathlib import Path

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from core.core_join_compare import join_compare_display_key  # noqa: E402


def test_strips_leading_excel_forced_text_apostrophe() -> None:
    assert join_compare_display_key("'20220527") == "20220527"
    assert join_compare_display_key("20220527") == "20220527"


def test_strips_whitespace_and_apostrophe() -> None:
    assert join_compare_display_key("  '20220527  ") == "20220527"


def test_date_slash_form_unchanged_except_apostrophe() -> None:
    assert join_compare_display_key("'2022/05/27") == "2022/05/27"
    assert join_compare_display_key("2022/05/27") == "2022/05/27"


def test_none_and_empty() -> None:
    assert join_compare_display_key(None) == ""
    assert join_compare_display_key("") == ""
    assert join_compare_display_key("   ") == ""


def test_float_integer_like() -> None:
    assert join_compare_display_key(20220527.0) == "20220527"


def test_does_not_normalize_date_formats() -> None:
    assert join_compare_display_key("'20220527") != join_compare_display_key("'2022/05/27")


def test_plain_text_with_apostrophe_prefix() -> None:
    assert join_compare_display_key("'ODN-164") == "ODN-164"


def test_row_key_apostrophe_consistency() -> None:
    from svc.svc_data_agg_write import _row_key  # noqa: WPS433

    row_a = ["'20220527", "x"]
    row_b = ["20220527", "y"]
    assert _row_key(row_a, [0]) == _row_key(row_b, [0])
