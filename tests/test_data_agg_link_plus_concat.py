# -*- coding: utf-8 -*-
"""連携キーの + 複数セル結合、およびセル内改行除去の抽出共通挙動。"""
from __future__ import annotations

import sys
from pathlib import Path
from unittest.mock import patch

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from svc.data_agg_value_post import (  # noqa: E402
    _coerce_cell_scalar_to_full_text,
    postprocess_cell_primary,
    postprocess_link_rule_value,
)
from svc.svc_data_agg_extract import (  # noqa: E402
    _extract_from_cell_rule,
    _extract_from_cell_rule_with_context,
    _split_plus_cell_refs,
)


def test_coerce_strips_cell_newlines() -> None:
    assert _coerce_cell_scalar_to_full_text("電\n源") == "電源"
    assert _coerce_cell_scalar_to_full_text("電\r\n源") == "電源"
    assert _coerce_cell_scalar_to_full_text("電\r源") == "電源"
    assert "電源" in postprocess_cell_primary("電\n源", {})
    assert "\n" not in postprocess_cell_primary("電\n源", {})
    assert "電源" in postprocess_link_rule_value("電\n源", {})


def test_postprocess_split_picks_line_before_newline_strip() -> None:
    """整形 DSL の split は改行除去より前に効き、採用行を返す。最終結果に改行は残さない。"""
    raw = "見出し\n値A"
    out1 = postprocess_cell_primary(raw, {"value_shape_script": "split,1"})
    out2 = postprocess_cell_primary(raw, {"value_shape_script": "split,2"})
    assert out1.lstrip("'") == "見出し"
    assert out2.lstrip("'") == "値A"
    assert "\n" not in out1 and "\r" not in out1
    assert "\n" not in out2 and "\r" not in out2
    link2 = postprocess_link_rule_value(raw, {"value_shape_script": "split,2"})
    assert link2.lstrip("'") == "値A"
    assert "\n" not in link2


def test_split_plus_cell_refs() -> None:
    assert _split_plus_cell_refs("D10") == ["D10"]
    assert _split_plus_cell_refs("D10+E10") == ["D10", "E10"]
    assert _split_plus_cell_refs(" D10 + E11 + F12 ") == ["D10", "E11", "F12"]
    assert _split_plus_cell_refs("D10++E10") == ["D10", "E10"]
    assert _split_plus_cell_refs("") == []
    assert _split_plus_cell_refs("+") == []


def test_extract_link_plus_preserves_newlines_for_split() -> None:
    """+ 連結でも改行は後処理まで残し、split で採用行を選べる。"""
    cells = {"D10": "A\nB", "E10": "C"}

    def _fake_extract(_path, sheet_name=None, cell_ref=None):  # noqa: ARG001
        return cells.get(str(cell_ref or "").upper())

    with patch("svc.svc_data_agg_extract.extract_cell", side_effect=_fake_extract):
        v = _extract_from_cell_rule(
            "dummy.xlsx",
            {"sheet_name": "S"},
            {
                "cell": "D10+E10",
                "mode": "セル座標",
                "value_shape_script": "split,2",
            },
            allow_plus_concat=True,
        )
    assert str(v).lstrip("'") == "BC"
    assert "\n" not in str(v)

    cells = {"D10": "AB", "E10": "CD", "F10": None, "G10": "XY"}

    def _fake_extract(_path, sheet_name=None, cell_ref=None):  # noqa: ARG001
        return cells.get(str(cell_ref or "").upper())

    src = {"sheet_name": "S"}
    with patch("svc.svc_data_agg_extract.extract_cell", side_effect=_fake_extract):
        v = _extract_from_cell_rule(
            "dummy.xlsx",
            src,
            {"cell": "D10+E10", "mode": "セル座標", "row": 0, "col": 0},
            allow_plus_concat=True,
        )
        assert "ABCD" in str(v)
        v_empty = _extract_from_cell_rule(
            "dummy.xlsx",
            src,
            {"cell": "D10+F10+G10", "mode": "セル座標"},
            allow_plus_concat=True,
        )
        assert "ABXY" in str(v_empty)


def test_extract_join_does_not_split_plus() -> None:
    """結合キーは allow_plus_concat=False のため D10+E10 を分割せず 1 参照として解決する。"""
    seen: list[str] = []

    def _fake_extract(_path, sheet_name=None, cell_ref=None):  # noqa: ARG001
        seen.append(str(cell_ref or ""))
        return "X"

    with patch("svc.svc_data_agg_extract.extract_cell", side_effect=_fake_extract):
        _extract_from_cell_rule(
            "dummy.xlsx",
            {"sheet_name": "S"},
            {"cell": "D10+E10", "mode": "セル座標"},
            allow_plus_concat=False,
        )
    assert len(seen) == 1
    # パース不能な「D10+E10」全体は A1 相当へフォールバック（複数セル結合しない）
    assert seen[0].upper() == "A1"


def test_extract_link_plus_shared_offset_per_iter() -> None:
    """rule_iter=1, row=1 → 各パートに +1 行（D10→D11, E10→E11）。"""
    seen: list[str] = []

    def _fake_extract(_path, sheet_name=None, cell_ref=None):  # noqa: ARG001
        ref = str(cell_ref or "").upper()
        seen.append(ref)
        return {"D11": "P", "E11": "Q"}.get(ref, "")

    with patch("svc.svc_data_agg_extract.extract_cell", side_effect=_fake_extract):
        v = _extract_from_cell_rule_with_context(
            "dummy.xlsx",
            {"sheet_name": "S"},
            {"cell": "D10+E10", "mode": "セル座標", "row": 1, "col": 0},
            {"rule_iter_index": 1, "iter_index": 1},
            allow_plus_concat=True,
        )
    assert "PQ" in str(v)
    assert seen == ["D11", "E11"]
