# -*- coding: utf-8 -*-
"""+ セル座標式（引用混在）の分割・検証・大文字化。"""
from __future__ import annotations

import sys
from pathlib import Path
from unittest.mock import patch

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from svc.data_agg_plus_cell_spec import (  # noqa: E402
    ascii_upper_plus_cell_spec,
    first_anchor_cell_ref,
    parse_plus_cell_spec,
    plus_spec_needs_concat,
    validate_plus_cell_spec,
)
from svc.svc_data_agg_extract import (  # noqa: E402
    _extract_from_cell_rule,
    _extract_from_cell_rule_with_context,
    extract_item_values,
)
from svc.svc_data_agg_scenario import validate_scenario  # noqa: E402


def test_parse_plain_plus() -> None:
    parts, err = parse_plus_cell_spec("D10+E10")
    assert err is None
    assert [(p.kind, p.value) for p in parts] == [
        ("cell", "D10"),
        ("cell", "E10"),
    ]
    assert first_anchor_cell_ref(parts) == "D10"
    assert plus_spec_needs_concat(parts) is True


def test_parse_quoted_literal_mix() -> None:
    parts, err = parse_plus_cell_spec('A1+"-"+B1')
    assert err is None
    assert [(p.kind, p.value) for p in parts] == [
        ("cell", "A1"),
        ("literal", "-"),
        ("cell", "B1"),
    ]
    parts2, err2 = parse_plus_cell_spec('"ABC"+A1+B1')
    assert err2 is None
    assert first_anchor_cell_ref(parts2) == "A1"
    parts3, err3 = parse_plus_cell_spec('A1+""""+B1')
    assert err3 is None
    assert parts3[1].kind == "literal" and parts3[1].value == '"'


def test_validate_requires_cell() -> None:
    assert validate_plus_cell_spec('"x"+"y"', empty_ok=True) is not None
    assert validate_plus_cell_spec("", empty_ok=True) is None
    assert validate_plus_cell_spec("", empty_ok=False) is not None
    assert validate_plus_cell_spec("A1+ZZ", empty_ok=True) is not None
    assert validate_plus_cell_spec("A1+B1", empty_ok=True) is None


def test_ascii_upper_preserves_quotes() -> None:
    assert ascii_upper_plus_cell_spec('a1+"ab"+b1') == 'A1+"ab"+B1'
    assert ascii_upper_plus_cell_spec("d10+d11") == "D10+D11"


def test_extract_link_plus_with_literal() -> None:
    cells = {"A1": "X", "B1": "Y"}

    def _fake(_path, sheet_name=None, cell_ref=None):  # noqa: ARG001
        return cells.get(str(cell_ref or "").upper())

    with patch("svc.svc_data_agg_extract.extract_cell", side_effect=_fake):
        v = _extract_from_cell_rule(
            "dummy.xlsx",
            {"sheet_name": "S"},
            {"cell": 'A1+"-"+B1', "mode": "セル座標"},
            allow_plus_concat=True,
        )
    assert "X-Y" in str(v)


def test_extract_join_plus_enabled() -> None:
    seen: list[str] = []

    def _fake(_path, sheet_name=None, cell_ref=None):  # noqa: ARG001
        seen.append(str(cell_ref or "").upper())
        return "Z"

    with patch("svc.svc_data_agg_extract.extract_cell", side_effect=_fake):
        v = _extract_from_cell_rule(
            "dummy.xlsx",
            {"sheet_name": "S"},
            {"cell": "D10+E10", "mode": "セル座標"},
            allow_plus_concat=True,
        )
    assert seen == ["D10", "E10"]
    assert "ZZ" in str(v)


def test_extract_fixed_mode_no_plus_split() -> None:
    with patch("svc.svc_data_agg_extract.extract_cell") as ex:
        v = _extract_from_cell_rule(
            "dummy.xlsx",
            {"sheet_name": "S"},
            {"cell": 'A1+"-"+B1', "mode": "固定値"},
            allow_plus_concat=True,
        )
        ex.assert_not_called()
    assert 'A1+"-"+B1' in str(v)


def test_resolve_fixed_text_value_csv_quotes() -> None:
    from svc.data_agg_plus_cell_spec import resolve_fixed_text_value

    assert resolve_fixed_text_value('"ABC"') == "ABC"
    assert resolve_fixed_text_value('""""') == '"'
    assert resolve_fixed_text_value('""') == ""
    assert resolve_fixed_text_value("PAST") == "PAST"
    assert resolve_fixed_text_value('A1+"-"+B1') == 'A1+"-"+B1'
    assert resolve_fixed_text_value('  "X"  ') == "X"
    assert resolve_fixed_text_value("  PAST  ", strip_unquoted=True) == "PAST"
    assert resolve_fixed_text_value("  PAST  ", strip_unquoted=False) == "  PAST  "


def test_extract_fixed_mode_strips_outer_quotes() -> None:
    with patch("svc.svc_data_agg_extract.extract_cell") as ex:
        v = _extract_from_cell_rule(
            "dummy.xlsx",
            {"sheet_name": "S"},
            {"cell": '"ABC"', "mode": "固定値"},
            allow_plus_concat=True,
        )
        ex.assert_not_called()
    assert str(v).lstrip("'") == "ABC"

    with patch("svc.svc_data_agg_extract.extract_cell") as ex:
        v2 = _extract_from_cell_rule(
            "dummy.xlsx",
            {"sheet_name": "S"},
            {"cell": '""""', "mode": "固定値"},
            allow_plus_concat=True,
        )
        ex.assert_not_called()
    assert str(v2).lstrip("'") == '"'


def test_name_extract_fixed_strips_outer_quotes() -> None:
    item = {
        "sources": [
            {
                "type": "name_extract",
                "source_type": "file_name",
                "search_condition": "contains",
                "search_text": "",
                "length_value": '"ODN-1"',
                "ui_scenario_source_v1": {"extract_mode": "fixed"},
            }
        ]
    }
    vals = extract_item_values(r"C:\tmp\any.xlsx", item)
    texts = [str(v).lstrip("'") for v in vals if v is not None]
    assert texts == ["ODN-1"]


def test_primary_plus_iterates_on_first_cell() -> None:
    cells = {
        "A1": "1",
        "B1": "a",
        "A2": "2",
        "B2": "b",
        "A3": "",
        "B3": "x",
    }

    def _fake(_path, sheet_name=None, cell_ref=None):  # noqa: ARG001
        return cells.get(str(cell_ref or "").upper())

    item = {
        "sources": [
            {
                "type": "cell",
                "sheet_name": "S",
                "cell_ref": 'A1+"-"+B1',
                "row_offset": 1,
                "col_offset": 0,
                "repeat_direction": "vertical",
                "repeat_until_empty": True,
                "repeat_max": 0,
            }
        ]
    }
    with patch("svc.svc_data_agg_extract.extract_cell", side_effect=_fake):
        vals = extract_item_values("dummy.xlsx", item)
    # A3 が空で停止。連結結果は後処理後
    texts = [str(v).lstrip("'") for v in vals if v is not None]
    assert texts == ["1-a", "2-b"]


def test_link_plus_offset_with_literal() -> None:
    seen: list[str] = []

    def _fake(_path, sheet_name=None, cell_ref=None):  # noqa: ARG001
        ref = str(cell_ref or "").upper()
        seen.append(ref)
        return {"A2": "P", "B2": "Q"}.get(ref, "")

    with patch("svc.svc_data_agg_extract.extract_cell", side_effect=_fake):
        v = _extract_from_cell_rule_with_context(
            "dummy.xlsx",
            {"sheet_name": "S"},
            {"cell": 'A1+"x"+B1', "mode": "セル座標", "row": 1, "col": 0},
            {"rule_iter_index": 1, "iter_index": 1},
            allow_plus_concat=True,
        )
    assert "PxQ" in str(v)
    assert seen == ["A2", "B2"]


def test_format_plus_cell_input_error() -> None:
    from svc.data_agg_plus_cell_spec import format_plus_cell_input_error

    assert (
        format_plus_cell_input_error(
            "セル座標が不正です: AM",
            field_label="セル座標/固定値",
            value="AM",
            key_kind="連携キー",
            key_index=6,
        )
        == "連携キー #6 セル座標/固定値「AM」入力異常：セル座標が不正です: AM"
    )
    assert (
        format_plus_cell_input_error(
            "セル座標を入力してください",
            field_label="セル座標(Excel方式)",
            value="",
            key_kind="主キー",
        )
        == "主キー セル座標(Excel方式)「（空欄）」入力異常：セル座標を入力してください"
    )


def test_format_grouped_validation_errors_by_scenario() -> None:
    from svc.data_agg_plus_cell_spec import (
        attach_scenario_to_cell_error,
        format_grouped_validation_errors,
        format_plus_cell_input_error,
    )

    d1 = format_plus_cell_input_error(
        "セル座標が不正です: AM",
        field_label="セル座標/固定値",
        value="AM",
        key_kind="連携キー",
        key_index=6,
    )
    d2 = format_plus_cell_input_error(
        "セル座標が不正です: ZZ",
        field_label="セル座標",
        value="ZZ",
        key_kind="結合キー",
        key_index=1,
    )
    msgs = [
        attach_scenario_to_cell_error("A案", d1),
        attach_scenario_to_cell_error("A案", d1),  # 同一箇所は1つ
        attach_scenario_to_cell_error("B案", d2),
    ]
    text = format_grouped_validation_errors(msgs)
    assert text == (
        "シナリオ「A案」\n"
        "・連携キー #6 セル座標/固定値「AM」入力異常：セル座標が不正です: AM\n"
        "シナリオ「B案」\n"
        "・結合キー #1 セル座標「ZZ」入力異常：セル座標が不正です: ZZ"
    )


def test_validate_scenario_cell_specs_optional_for_load_compat() -> None:
    """既存シナリオの列のみ座標（AM）は厳密検証で弾くが、読込用フラグでは通す。"""
    data = {
        "items": [
            {
                "id": "i1",
                "name": "項目1",
                "write_mode": "append",
                "sources": [
                    {
                        "type": "cell",
                        "scenario_name": "旧シナリオ",
                        "cell_ref": "A1",
                        "row_offset": 1,
                        "col_offset": 0,
                        "repeat_until_empty": True,
                        "ui_scenario_source_v1": {
                            "link_defs": [
                                {
                                    "item": "連携先",
                                    "cell": "AM",
                                    "mode": "セル座標",
                                    "row": 0,
                                    "col": 0,
                                }
                            ]
                        },
                    }
                ],
            }
        ]
    }
    strict = validate_scenario(data, check_cell_specs=True)
    assert any("シナリオ「旧シナリオ」" in e and "連携キー #1" in e and "「AM」" in e for e in strict)
    soft = validate_scenario(data, check_cell_specs=False)
    assert not any("AM" in e for e in soft)


def test_unique_validation_messages_keeps_first() -> None:
    from svc.data_agg_plus_cell_spec import unique_validation_messages

    msgs = [
        "連携キー #6 セル座標/固定値「AM」入力異常：x",
        "連携キー #6 セル座標/固定値「AM」入力異常：x",
        "結合キー #1 セル座標「ZZ」入力異常：y",
        "連携キー #6 セル座標/固定値「AM」入力異常：x",
    ]
    assert unique_validation_messages(msgs) == [
        "連携キー #6 セル座標/固定値「AM」入力異常：x",
        "結合キー #1 セル座標「ZZ」入力異常：y",
    ]
