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
    evaluate_shape_expr,
    parse_and_apply_commands,
    tokenize_shape_script,
    validate_shape_expr_syntax,
)

_SAMPLE = "ABCDEFGHIJK"


def test_expr_left_pos_examples() -> None:
    assert apply_value_shape(_SAMPLE, 'left,pos("GH")') == "ABCDEFG"
    assert apply_value_shape(_SAMPLE, 'left,pos("GH")-1') == "ABCDEF"
    assert apply_value_shape(_SAMPLE, 'left,len()-pos("I")') == "AB"


def test_expr_len_literal() -> None:
    assert evaluate_shape_expr('len("GH")', _SAMPLE) == 2
    assert evaluate_shape_expr("len()", _SAMPLE) == len(_SAMPLE)


def test_expr_ins_before_after() -> None:
    assert apply_value_shape(_SAMPLE, 'ins,pos("G"),"123"') == "ABCDEF123GHIJK"
    assert apply_value_shape(_SAMPLE, 'ins,pos("G")+1,"123"') == "ABCDEFG123HIJK"
    assert apply_value_shape(_SAMPLE, 'ins,pos("GH")+len("GH"),"123"') == "ABCDEFGH123IJK"


def test_expr_mid_cut() -> None:
    assert apply_value_shape(_SAMPLE, 'mid,pos("E")+1,pos("I")-pos("E")-1') == "FGH"
    assert apply_value_shape(_SAMPLE, 'cut,pos("B"),pos("I")-pos("B")+1') == "AJK"


def test_expr_right_after_marker() -> None:
    assert apply_value_shape(_SAMPLE, 'right,len()-pos("G")') == "HIJK"


def test_expr_pos_not_found_skips_command() -> None:
    assert apply_value_shape(_SAMPLE, 'left,pos("ZZZ")') == _SAMPLE
    assert apply_value_shape(_SAMPLE, 'trim,left,pos("ZZZ")') == _SAMPLE


def test_expr_case_sensitive_pos() -> None:
    assert apply_value_shape("abcDef", 'left,pos("D")') == "abcD"
    assert apply_value_shape("abcDef", 'left,pos("d")') == "abcDef"


def test_compile_expr_ok() -> None:
    ok, msg = compile_shape_script('left,pos("GH")-1')
    assert ok and msg == ""


def test_compile_expr_bad() -> None:
    ok, msg = compile_shape_script("left,pos(")
    assert not ok
    assert "不正" in msg


def test_validate_expr_syntax() -> None:
    assert validate_shape_expr_syntax('pos("G")+1')[0]
    assert not validate_shape_expr_syntax("pos(G)")[0]


def test_tokenize_expr_arg() -> None:
    assert tokenize_shape_script('left,pos("GH")-1') == ["left", 'pos("GH")-1']


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


def test_left_basic() -> None:
    assert apply_value_shape("abcdef", "left,3") == "abc"
    assert apply_value_shape("abcdef", "Left,3") == "abc"


def test_right_basic() -> None:
    assert apply_value_shape("abcdef", "right,3") == "def"
    assert apply_value_shape("abcdef", "Right,2") == "ef"


def test_left_right_edges() -> None:
    assert apply_value_shape("ab", "left,0") == ""
    assert apply_value_shape("ab", "right,0") == ""
    assert apply_value_shape("ab", "left,9") == "ab"
    assert apply_value_shape("ab", "right,9") == "ab"
    assert apply_value_shape("", "left,3") == ""
    assert apply_value_shape("ab", "left,-1") == "ab"
    assert apply_value_shape("ab", "right,abc") == "ab"


def test_compile_left_right() -> None:
    ok, msg = compile_shape_script("left,3,right,2")
    assert ok and msg == ""
    ok2, msg2 = compile_shape_script("left")
    assert not ok2
    assert "不足" in msg2


def test_shape_step_apply() -> None:
    from core.core_value_shape import (
        apply_value_shape_for_test,
        apply_value_shape_step_for_test,
        shape_command_count,
    )

    sample = "  abc  "
    script = "trim,wide"
    assert shape_command_count(script) == 2
    r1, d1, e1 = apply_value_shape_step_for_test(sample, script, 1)
    assert e1 is None
    assert r1 == "abc"
    assert d1 == "trim,"
    r2, d2, e2 = apply_value_shape_step_for_test(sample, script, 2)
    assert e2 is None
    assert d2 == "trim,wide"
    r_all, err = apply_value_shape_for_test(sample, script)
    assert err is None
    assert r_all == r2


def test_shape_step_rep_quoted_display() -> None:
    from core.core_value_shape import apply_value_shape_step_for_test

    script = 'rep,"ー","",rep,"-",""'
    _, d1, e1 = apply_value_shape_step_for_test("x", script, 1)
    assert e1 is None
    assert d1 == 'rep,"ー","",'
    _, d2, e2 = apply_value_shape_step_for_test("x", script, 2)
    assert e2 is None
    assert d2 == 'rep,"ー","",rep,"-",""'


def test_shape_step_syntax_error() -> None:
    from core.core_value_shape import (
        apply_value_shape_for_test,
        shape_script_syntax_error_block,
    )

    _, err = apply_value_shape_for_test("x", "bogus")
    assert err

    script = 'rep,"ー","",rerp,"-",""'
    ok, msg, block = shape_script_syntax_error_block(script)
    assert not ok
    assert "rerp" in msg
    assert block == 'rerp,"-",""'
