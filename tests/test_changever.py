# -*- coding: utf-8 -*-
from __future__ import annotations

from core.changever import (
    format_ver_history_update_block,
    format_ver_history_viewer_text,
    parse_ver_history_sections,
    ver_history_block_for_update,
)


def _sample_cfg() -> dict:
    return {
        "VER_HISTORY": {
            "BIN": [
                {
                    "version": "1.1.9.4",
                    "items": ["履歴表示", "進捗文言"],
                },
                {
                    "version": "1.1.8.4",
                    "items": ["old"],
                },
                {
                    "version": "1.2.0.0",
                    "items": ["future"],
                },
            ],
            "BOOTSTRAP": [
                {
                    "version": "1.0.9",
                    "items": ["ランナー更新"],
                }
            ],
        },
        "MESSAGES": {"VER_HISTORY_EMPTY": "（空）"},
    }


def test_parse_ver_history_sections() -> None:
    secs = parse_ver_history_sections(_sample_cfg())
    assert secs[0] == ("bin", "1.1.9.4", ["履歴表示", "進捗文言"])
    assert secs[1] == ("bin", "1.1.8.4", ["old"])
    assert secs[2] == ("bin", "1.2.0.0", ["future"])
    assert secs[3] == ("bootstrap", "1.0.9", ["ランナー更新"])


def test_format_bin_range_only() -> None:
    block = format_ver_history_update_block(
        _sample_cfg(), kind="bin", installed="1.1.8.4", latest="1.1.9.4"
    )
    assert "1.1.9.4" in block
    assert "履歴表示" in block
    assert "old" not in block
    assert "future" not in block
    assert "ランナー更新" not in block
    assert block.startswith("変更内容:")


def test_format_shows_all_matching_lines() -> None:
    lines = [f"item{i}" for i in range(1, 14)]
    cfg = {"VER_HISTORY": {"BIN": [{"version": "1.1.9.4", "items": lines}]}}
    block = format_ver_history_update_block(
        cfg, kind="bin", installed="1.1.8.4", latest="1.1.9.4"
    )
    assert "item1" in block
    assert "item13" in block
    assert "（続きはヘルプの変更履歴）" not in block


def test_format_truncate_when_max_lines_set() -> None:
    lines = [f"item{i}" for i in range(1, 12)]
    cfg = {"VER_HISTORY": {"BIN": [{"version": "1.1.9.4", "items": lines}]}}
    block = format_ver_history_update_block(
        cfg,
        kind="bin",
        installed="1.1.8.4",
        latest="1.1.9.4",
        max_lines=8,
    )
    assert "（続きはヘルプの変更履歴）" in block
    assert "item9" not in block


def test_format_from_ui_help_json() -> None:
    import json
    from pathlib import Path

    path = Path(__file__).resolve().parents[1] / "config" / "ui_help.json"
    cfg = json.loads(path.read_text(encoding="utf-8"))
    block = format_ver_history_update_block(
        cfg, kind="bin", installed="1.1.8.4", latest="1.1.9.4"
    )
    assert "1.1.9.4" in block
    assert "複数セル結合" in block
    assert "blank singleton" in block
    assert "旧版バックアップ" not in block
    viewer = format_ver_history_viewer_text(cfg)
    assert "[CSV Tool 1.1.9.4]" in viewer
    assert "[bootstrap 1.0.9]" in viewer
    assert "旧版バックアップ" in viewer


def test_viewer_shows_all_kinds() -> None:
    text = format_ver_history_viewer_text(_sample_cfg())
    assert "[CSV Tool 1.1.9.4]" in text
    assert "[CSV Tool 1.1.8.4]" in text
    assert "[bootstrap 1.0.9]" in text
    assert "履歴表示" in text
    assert "ランナー更新" in text


def test_viewer_empty() -> None:
    assert "登録されていません" in format_ver_history_viewer_text({})


def test_update_block_empty() -> None:
    assert (
        format_ver_history_update_block(
            {}, kind="bin", installed="1.0.0", latest="1.1.0"
        )
        == ""
    )


def test_ver_history_block_for_update_uses_cfg() -> None:
    block = ver_history_block_for_update(
        kind="bin",
        installed="1.1.8.4",
        latest="1.1.9.4",
        cfg=_sample_cfg(),
    )
    assert "1.1.9.4" in block
