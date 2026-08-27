# -*- coding: utf-8 -*-
"""シナリオ結果一覧: #n[項目名] の列展開（形式外行があっても項目列へ分ける）。"""
from __future__ import annotations

import sys
from pathlib import Path

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from ui_qt.ui_data_agg_debug import expand_hash_bracket_value_groups  # noqa: E402


def test_expand_hash_bracket_basic() -> None:
    got = expand_hash_bracket_value_groups(
        [
            "#1[製品コード] A",
            "#1[製品コード] B",
            "#2[出荷番号] X",
            "#2[出荷番号] Y",
        ]
    )
    assert got is not None
    assert [t for t, _ in got] == ["製品コード", "出荷番号"]
    assert got[0][1] == ["A", "B"]
    assert got[1][1] == ["X", "Y"]


def test_expand_hash_bracket_skips_bad_line_still_expands() -> None:
    """旧仕様は形式外1行で全体失敗→「連携キー」1列のまま。部分一致で展開する。"""
    got = expand_hash_bracket_value_groups(
        [
            "#1[製品コード] A",
            "…（以降省略・上限50件）",  # 上限省略行（# 無し）→ 旧実装だと全体失敗
            "#2[出荷番号] X",
            "#1[製品コード] B",
            "#2[出荷番号] Y",
        ]
    )
    assert got is not None
    assert [t for t, _ in got] == ["製品コード", "出荷番号"]
    assert got[0][1] == ["A", "B"]
    assert got[1][1] == ["X", "Y"]


def test_expand_hash_bracket_none_when_no_tags() -> None:
    assert expand_hash_bracket_value_groups(["ただの値", "もう一つ"]) is None
    assert expand_hash_bracket_value_groups([]) is None
