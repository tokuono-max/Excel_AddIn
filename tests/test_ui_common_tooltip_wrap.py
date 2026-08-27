# -*- coding: utf-8 -*-
"""ui_common ツールチップ折り返しの単体テスト。"""
from ui_qt.ui_common import (
    _TOOLTIP_WRAP_WIDTH,
    _normalize_tooltip_text,
    _wrap_tooltip_line,
)


def test_wrap_tooltip_line_short_unchanged():
    s = "短い説明文"
    assert _wrap_tooltip_line(s) == [s]


def test_wrap_tooltip_line_splits_long_japanese():
    s = "あ" * 60
    lines = _wrap_tooltip_line(s, width=45)
    assert len(lines) >= 2
    assert all(len(line) <= 45 for line in lines)
    assert "".join(lines) == s


def test_wrap_tooltip_line_prefers_punctuation_break():
    s = "これは長い説明文です。続きの説明がここに続きます。"
    lines = _wrap_tooltip_line(s, width=20)
    assert any(line.endswith("。") for line in lines[:-1])


def test_normalize_tooltip_preserves_explicit_newlines():
    raw = "1行目\\n2行目は短い"
    out = _normalize_tooltip_text(raw)
    assert "\n" in out
    assert "1行目" in out
    assert "2行目" in out


def test_normalize_tooltip_skips_html():
    html = "<b>太字</b>の長い説明" + ("あ" * 80)
    assert _normalize_tooltip_text(html) == html


def test_normalize_tooltip_wraps_each_paragraph():
    para = "あ" * (_TOOLTIP_WRAP_WIDTH + 10)
    raw = para + "\\n" + para
    out = _normalize_tooltip_text(raw)
    assert out.count("\n") >= 3
