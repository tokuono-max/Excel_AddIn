# -*- coding: utf-8 -*-
"""core.core_value_shape の単体テスト。"""
from __future__ import annotations

import sys
from pathlib import Path

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from core.core_value_shape import (  # noqa: E402
    apply_value_shape,
    compile_shape_script,
    parse_and_apply_commands,
    tokenize_shape_script,
)


def test_tokenize_quoted_comma() -> None:
    assert tokenize_shape_script('rep,"a,b","c"') == ["rep", "a,b", "c"]


def test_tokenize_csv_escape_quote() -> None:
    assert tokenize_shape_script(r'rep,"say ""hi""","x"') == ['rep', 'say "hi"', "x"]


def test_tokenize_semicolon_command_boundary() -> None:
    assert tokenize_shape_script("trim;split,1") == ["trim", "split", "1"]
    assert apply_value_shape("line1\nline2", "trim;split,1") == "line1"


def test_tokenize_semicolon_inside_quotes() -> None:
    assert tokenize_shape_script('rep,"a;b","x"') == ["rep", "a;b", "x"]
    assert apply_value_shape("a;b", 'rep,"a;b","z"') == "z"


def test_rep_replaces_all() -> None:
    assert apply_value_shape("xaxa", 'rep,"a","b"') == "xbxb"


def test_rep_empty_search_noop() -> None:
    assert apply_value_shape("abc", 'rep,"","x"') == "abc"


def test_mid_one_based() -> None:
    assert apply_value_shape("abcdef", "mid,2,3") == "bcd"


def test_pipeline_trim_rep() -> None:
    assert apply_value_shape("  a,a  ", 'trim,rep,",",";"') == "a;a"


def test_compile_ok() -> None:
    ok, msg = compile_shape_script("trim,case,upper")
    assert ok and msg == ""


def test_compile_unknown() -> None:
    ok, msg = compile_shape_script("bogus")
    assert not ok
    assert "未知" in msg or "bogus" in msg


def test_parse_cut() -> None:
    assert parse_and_apply_commands("abcdef", tokenize_shape_script("cut,2,2")) == "adef"


def test_ins_position_one_based() -> None:
    assert parse_and_apply_commands("ab", tokenize_shape_script('ins,2,"Z"')) == "aZb"


def test_split_first_line() -> None:
    assert apply_value_shape("line1\nline2", "split,1") == "line1"


def test_split_crlf_second_line() -> None:
    assert apply_value_shape("a\r\nb", "split,2") == "b"


def test_split_out_of_range() -> None:
    assert apply_value_shape("only", "split,2") == ""


def test_split_result_has_no_newlines() -> None:
    assert "\n" not in apply_value_shape("x\ny", "split,1")
    assert "\r" not in apply_value_shape("x\ry", "split,1")


def test_compile_split() -> None:
    ok, msg = compile_shape_script("split,1")
    assert ok and msg == ""
