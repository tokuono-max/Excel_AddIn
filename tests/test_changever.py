# -*- coding: utf-8 -*-
from __future__ import annotations

from pathlib import Path

from core.changever import (
    catalog_changever_relative_path,
    format_changever_block,
    parse_changever_sections,
    resolve_changever_path,
)


def test_parse_bracket_and_kind() -> None:
    text = (
        "[1.1.8.4]\n"
        "- バックアップ廃止\n"
        "\n"
        "[bootstrap 1.0.9]\n"
        "- ランナー更新\n"
        "[1.1.9.4]\n"
        "進捗で差分とフルを区別\n"
    )
    secs = parse_changever_sections(text)
    assert secs[0] == ("bin", "1.1.8.4", ["- バックアップ廃止"])
    assert secs[1] == ("bootstrap", "1.0.9", ["- ランナー更新"])
    assert secs[2][0] == "bin"
    assert secs[2][1] == "1.1.9.4"
    assert secs[2][2] == ["進捗で差分とフルを区別"]


def test_format_bin_range_only() -> None:
    text = (
        "[1.1.8.4]\n- old\n"
        "[1.1.9.4]\n- 履歴表示\n- 進捗文言\n"
        "[1.2.0.0]\n- future\n"
    )
    block = format_changever_block(
        text, kind="bin", installed="1.1.8.4", latest="1.1.9.4"
    )
    assert "1.1.9.4" in block
    assert "履歴表示" in block
    assert "old" not in block
    assert "future" not in block
    assert block.startswith("変更内容:")


def test_format_truncate_max_lines() -> None:
    lines = "\n".join(f"- item{i}" for i in range(1, 12))
    text = f"[1.1.9.4]\n{lines}\n"
    block = format_changever_block(
        text, kind="bin", installed="1.1.8.4", latest="1.1.9.4", max_lines=8
    )
    assert "（続きは CHANGEVER.txt）" in block
    assert "item9" not in block


def test_format_empty_when_file_blank() -> None:
    assert format_changever_block("", kind="bin", installed="1.0.0", latest="1.1.0") == ""


def test_resolve_default_and_catalog_relative(tmp_path: Path) -> None:
    cat = tmp_path / "catalog.json"
    cat.write_text("{}", encoding="utf-8")
    notes = tmp_path / "CHANGEVER.txt"
    notes.write_text("[1.0.0]\n- x\n", encoding="utf-8")
    p = resolve_changever_path(cat, {})
    assert p == notes
    named = tmp_path / "notes" / "ver.txt"
    named.parent.mkdir()
    named.write_text("[1.0.0]\n- y\n", encoding="utf-8")
    data = {"release_notes": {"relative_path": "notes/ver.txt"}}
    assert catalog_changever_relative_path(data) == "notes/ver.txt"
    assert resolve_changever_path(cat, data) == named
