# -*- coding: utf-8 -*-
"""主キースキップ一致文字パースと終端トリムの単体テスト。"""
from __future__ import annotations

import sys
from pathlib import Path

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from svc.data_agg_primary_end import (  # noqa: E402
    END_MODE_N_COUNT,
    END_MODE_UNTIL_EMPTY,
    END_MODE_UNTIL_LAST,
    apply_until_last_trim,
    effective_skip_primary_tokens,
    parse_skip_primary_match,
    primary_value_matches_skip_tokens,
    source_end_mode,
    source_keep_empty_primary_slots,
)


def test_parse_empty_and_spaces() -> None:
    assert parse_skip_primary_match("") == [""]
    assert parse_skip_primary_match("   ") == [""]
    assert parse_skip_primary_match(None) == [""]


def test_parse_leading_comma_includes_blank() -> None:
    assert parse_skip_primary_match(",A,b,c") == ["", "A", "b", "c"]


def test_parse_double_comma_and_space_token() -> None:
    assert parse_skip_primary_match("a,,b,c") == ["a", "", "b", "c"]
    assert parse_skip_primary_match("a, ,b") == ["a", "", "b"]


def test_parse_trailing_comma_ignores_blank() -> None:
    assert parse_skip_primary_match("a,b,c,") == ["a", "b", "c"]
    assert parse_skip_primary_match("a,b,c,  ") == ["a", "b", "c"]


def test_match_tokens() -> None:
    assert primary_value_matches_skip_tokens(None, [""])
    assert primary_value_matches_skip_tokens("  ", [""])
    assert primary_value_matches_skip_tokens("-", ["-", "なし"])
    assert not primary_value_matches_skip_tokens("x", ["-", "なし"])
    assert not primary_value_matches_skip_tokens(None, ["-", "なし"])


def test_effective_tokens_until_empty_strips_blank() -> None:
    src = {"skip_empty_primary": True, "skip_primary_match": ",A,-", "repeat_until_empty": True}
    assert effective_skip_primary_tokens(src) == ["A", "-"]
    src_n = {
        "skip_empty_primary": True,
        "skip_primary_match": "",
        "repeat_until_empty": False,
        "repeat_max": 3,
    }
    assert effective_skip_primary_tokens(src_n) == [""]


def test_source_end_mode() -> None:
    assert source_end_mode({"repeat_until_empty": True}) == END_MODE_UNTIL_EMPTY
    assert source_end_mode({"repeat_until_last": True, "repeat_until_empty": False}) == END_MODE_UNTIL_LAST
    assert source_end_mode({"repeat_until_empty": False, "repeat_max": 5}) == END_MODE_N_COUNT


def test_keep_empty_slots_until_last() -> None:
    assert source_keep_empty_primary_slots(
        {"repeat_until_last": True, "repeat_until_empty": False}
    )
    assert source_keep_empty_primary_slots(
        {"repeat_until_empty": False, "repeat_max": 3, "skip_empty_primary": True}
    )
    assert not source_keep_empty_primary_slots(
        {"repeat_until_empty": False, "repeat_max": 3, "skip_empty_primary": False}
    )


def test_trim_until_last() -> None:
    assert apply_until_last_trim(["a", None, "b", "", ""], until_last=True) == ["a", None, "b"]
    assert apply_until_last_trim(["", ""], until_last=True) == []
    assert apply_until_last_trim(["a", None], until_last=False) == ["a", None]
