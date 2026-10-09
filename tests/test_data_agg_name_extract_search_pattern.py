# -*- coding: utf-8 -*-
"""名前から取得の検索文字（名前パターン AND/OR）の実行時評価。"""
from __future__ import annotations

from pathlib import Path

from svc.svc_data_agg_extract import extract_from_name, name_extract_search_matches


def _src(
    *,
    search_text: str,
    search_condition: str = "include",
    source_type: str = "file_name",
) -> dict:
    return {
        "type": "name_extract",
        "source_type": source_type,
        "search_condition": search_condition,
        "search_text": search_text,
    }


def test_name_extract_or_pattern_on_stem(tmp_path: Path) -> None:
    p = tmp_path / "report_east.xlsx"
    p.write_text("x", encoding="utf-8")
    assert name_extract_search_matches(p, _src(search_text='"east"|"west"'))
    assert not name_extract_search_matches(p, _src(search_text='"north"|"west"'))


def test_name_extract_and_pattern_on_stem(tmp_path: Path) -> None:
    p = tmp_path / "ship_2024_final.xlsx"
    p.write_text("x", encoding="utf-8")
    assert name_extract_search_matches(p, _src(search_text='"2024"&"ship"'))
    assert not name_extract_search_matches(p, _src(search_text='"2024"&"draft"'))


def test_name_extract_exact_or(tmp_path: Path) -> None:
    p = tmp_path / "A.xlsx"
    p.write_text("x", encoding="utf-8")
    assert name_extract_search_matches(
        p, _src(search_text='"A"|"B"', search_condition="exact")
    )
    assert not name_extract_search_matches(
        p, _src(search_text='"B"|"C"', search_condition="exact")
    )


def test_name_extract_exclude_negates_expr(tmp_path: Path) -> None:
    p = tmp_path / "foo_bar.xlsx"
    p.write_text("x", encoding="utf-8")
    assert not name_extract_search_matches(
        p, _src(search_text='"foo"', search_condition="exclude")
    )
    assert name_extract_search_matches(
        p, _src(search_text='"zzz"', search_condition="exclude")
    )


def test_name_extract_dir_name_pattern(tmp_path: Path) -> None:
    d = tmp_path / "出荷_東"
    d.mkdir()
    p = d / "a.xlsx"
    p.write_text("x", encoding="utf-8")
    assert name_extract_search_matches(
        p,
        _src(
            search_text='"東"|"西"',
            source_type="dir_name",
        ),
    )


def test_name_extract_legacy_plain_still_matches_at_runtime(tmp_path: Path) -> None:
    """実行時は legacy（囲みなし）も解釈。UI／読込強制変換とは別。"""
    p = tmp_path / "B_file.xlsx"
    p.write_text("x", encoding="utf-8")
    assert name_extract_search_matches(p, _src(search_text="B"))


def test_extract_from_name_filter_uses_pattern(tmp_path: Path) -> None:
    p = tmp_path / "alpha_beta.xlsx"
    p.write_text("x", encoding="utf-8")
    assert (
        extract_from_name(
            p,
            source_type="file_name",
            search_condition="include",
            search_text='"alpha"|"gamma"',
            start_mode="head",
            length_mode="end",
        )
        == "alpha_beta"
    )
    assert (
        extract_from_name(
            p,
            source_type="file_name",
            search_condition="include",
            search_text='"gamma"|"delta"',
            start_mode="head",
            length_mode="end",
        )
        == ""
    )


def test_empty_search_text_no_filter(tmp_path: Path) -> None:
    p = tmp_path / "anything.xlsx"
    p.write_text("x", encoding="utf-8")
    assert name_extract_search_matches(
        p, _src(search_text="", search_condition="include")
    )
