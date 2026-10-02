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
    ok, msg = compile_shape_script("trim,upr")
    assert ok and msg == ""


def test_compile_unknown() -> None:
    ok, msg = compile_shape_script("bogus")
    assert not ok
    assert "未知" in msg or "bogus" in msg


def test_compile_case_removed() -> None:
    ok, msg = compile_shape_script("case(upper)")
    assert not ok
    assert "未知" in msg or "case" in msg


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


def test_paren_form_basic_and_legacy_mix() -> None:
    assert apply_value_shape("  abc  ", "trim()") == "abc"
    assert apply_value_shape("abcdef", "left(3)") == "abc"
    assert apply_value_shape("abcdef", 'rep("a","z")') == "zbcdef"
    assert apply_value_shape("  x  ", "trim();upr()") == "X"
    assert apply_value_shape("  x  ", "trim,upr()") == "X"
    assert apply_value_shape("a\nb", "split(1);trim") == "a"
    ok, msg = compile_shape_script('split(1);rep("a","b")')
    assert ok and msg == ""


def test_paren_form_numeric_expr() -> None:
    assert apply_value_shape(_SAMPLE, 'left(pos("GH"))') == "ABCDEFG"
    assert apply_value_shape(_SAMPLE, 'left(pos("GH")-1)') == "ABCDEF"
    assert apply_value_shape(_SAMPLE, "left(len())") == _SAMPLE
    assert apply_value_shape(_SAMPLE, 'ins(pos("G"),"X")') == "ABCDEFXGHIJK"
    ok, msg = compile_shape_script('left(pos("GH")-1)')
    assert ok and msg == ""


def test_join_and_me() -> None:
    assert apply_value_shape("ABC", 'join("X","Y")') == "XY"
    assert apply_value_shape("ABC", "join()") == ""
    assert apply_value_shape("ABC", 'join(me,"-","Z")') == "ABC-Z"
    assert apply_value_shape("ABCDEF", 'join(me,"-",left(3))') == "ABCDEF-ABC"
    assert apply_value_shape("hello", 'join(me,"/",upr())') == "hello/HELLO"
    # 旧形 join（次の裸コマンド手前まで）
    assert apply_value_shape("Q", 'join,"A","B",trim') == "AB"
    ok, msg = compile_shape_script('join(me,"-",left(3))')
    assert ok and msg == ""


def test_tokenize_paren_keeps_inner_commas() -> None:
    assert tokenize_shape_script('rep("a,b","c");trim') == ['rep("a,b","c")', "trim"]
    assert apply_value_shape("a,b", 'rep("a,b","c")') == "c"


def test_shape_step_paren_join() -> None:
    from core.core_value_shape import (
        apply_value_shape_step_for_test,
        shape_command_count,
    )

    script = 'trim();split(1);join(me,"-","x")'
    assert shape_command_count(script) == 3
    r1, d1, e1 = apply_value_shape_step_for_test("  a\nb  ", script, 1)
    assert e1 is None and r1 == "a\nb" and d1 == "trim();"
    r2, d2, e2 = apply_value_shape_step_for_test("  a\nb  ", script, 2)
    assert e2 is None and r2 == "a" and d2 == "trim();split(1);"
    r3, d3, e3 = apply_value_shape_step_for_test("  a\nb  ", script, 3)
    assert e3 is None and r3 == "a-x"
    assert d3 == 'trim();split(1);join(me,"-","x")'


def test_join_args_share_same_current() -> None:
    """A1: join 各引数は開始時点の同一現在値（引数間で現在値は更新しない）。"""
    assert apply_value_shape("  ab  ", 'join(trim(),"|",me)') == "ab|  ab  "
    assert apply_value_shape("ABCDEF", 'join(me,"-",left(3))') == "ABCDEF-ABC"


def test_zero_arity_extra_args_lenient_shortage_errors() -> None:
    """A2: 0引数への余分は寛容、不足はエラー。"""
    ok, msg = compile_shape_script('trim("x")')
    assert ok and msg == ""
    assert apply_value_shape(" ab ", 'trim("x")') == "ab"
    ok_w, _ = compile_shape_script("wide(1)")
    assert ok_w
    ok_s, msg_s = compile_shape_script("split()")
    assert not ok_s
    assert "不足" in msg_s
    assert apply_value_shape("a\nb", "split()") == "a\nb"


def test_bare_me_is_unknown_command() -> None:
    """A3: 先頭の裸 me は未知コマンド（専用文言は付けない）。"""
    ok, msg = compile_shape_script("me")
    assert not ok
    assert "未知" in msg
    assert apply_value_shape("ABC", "me") == "ABC"
    ok2, msg2 = compile_shape_script("trim,me,left,1")
    assert not ok2
    assert "未知" in msg2


def test_legacy_join_is_supported() -> None:
    """A4: 旧形 join,a,b,… は正式サポート（次の裸コマンド手前まで）。"""
    assert apply_value_shape("Q", 'join,"A","B",trim') == "AB"
    assert apply_value_shape("Q", 'join,"A","B",upr') == "AB"
    ok, msg = compile_shape_script('join,"A","B",trim')
    assert ok and msg == ""


def test_upr_low_tok_rep_count_lenstr() -> None:
    assert apply_value_shape("AbC", "upr()") == "ABC"
    assert apply_value_shape("AbC", "low()") == "abc"
    assert apply_value_shape("hello", "upr(me)") == "HELLO"
    assert apply_value_shape("hello", "upr(left(2))") == "HE"
    assert apply_value_shape("HELLO", "low(right(2))") == "lo"
    assert apply_value_shape("a_b_c", 'tok("_",2)') == "b"
    assert apply_value_shape("a_b_c", 'tok("_",1)') == "a"
    assert apply_value_shape("a_b_c", 'tok("_",9)') == ""
    assert apply_value_shape("a_b_c", 'tok("",2)') == "a_b_c"
    assert apply_value_shape("aaa", 'rep("a","A")') == "AAA"
    assert apply_value_shape("aaa", 'rep(1,"a","A")') == "Aaa"
    assert apply_value_shape("aaa", 'rep(2,"a","A")') == "AAa"
    assert apply_value_shape("aaa", "rep,1,a,A") == "Aaa"
    assert apply_value_shape("hello", "lenstr()") == "5"
    assert apply_value_shape("hello", "lenstr(me)") == "5"
    assert apply_value_shape("hello", "lenstr(left(2))") == "2"
    assert apply_value_shape("hello", 'join("L=",lenstr())') == "L=5"
    # 式の len は従来どおり
    assert apply_value_shape("AB", "left(len())") == "AB"
    ok, _ = compile_shape_script('upr();tok("_",1);rep(1,"a","b");lenstr()')
    assert ok


def test_ui_hint_documents_me_as_parameter_only() -> None:
    """HINT が合意した me／入れ子の言い回しと矛盾しないこと。"""
    import json
    from pathlib import Path

    cfg = json.loads(
        (Path(__file__).resolve().parents[1] / "config" / "ui_data_agg.json").read_text(
            encoding="utf-8"
        )
    )
    for block in ("DETAIL_CELL", "DETAIL_NAME"):
        h = cfg["SCREENS"]["SCENARIO_EDIT"][block]["VALUE_SHAPE_HINT_HTML"]
        assert "パラメータのみ" in h
        assert "直前コマンド結果" in h
        assert "同一現在値" in h
        assert "外側開始時の現在値" in h
        # 記述例は () 形が主、旧形も可
        assert "trim()" in h
        assert "rep(\"検索\",\"置換\")" in h
        assert "upr()" in h
        assert "tok(" in h
        assert "lenstr()" in h
        assert "case(upper)" not in h
        assert "left(数)" in h
        assert "旧形" in h
        assert "split,ブロック" not in h
        short = cfg["SCREENS"]["SCENARIO_EDIT"][block]["VALUE_SHAPE_HINT_SHORT_HTML"]
        assert "パラメータのみ" in short
        assert "() 形" in short and "旧形も可" in short
    tip = cfg["SCREENS"]["SCENARIO_EDIT"]["DSL_TEST"]["TIP_CMD_INPUT"]
    assert "パラメータのみ" in tip
    assert "直前コマンド結果" in tip
    assert "() 形を推奨" in tip
    assert "旧形" in tip
