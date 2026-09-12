# -*- coding: utf-8 -*-
"""#7: 加工チェックは完全一致＋旧別名。部分一致による誤適用を防ぐ。"""
from __future__ import annotations

from svc.data_agg_value_post import (
    PROCESS_CHECK_ALIASES_BY_SLOT,
    apply_check_labels,
)


def test_current_labels_apply() -> None:
    assert apply_check_labels("  ab  ", ["トリム"]) == "ab"
    assert apply_check_labels("ＡＢ", ["全角→半角"]) == "AB"
    assert apply_check_labels("44708", ["年月日変換"], raw=44708.0) == "2022/05/27"


def test_legacy_aliases_still_apply() -> None:
    assert apply_check_labels("Ａ", ["全角→半角（英数字・記号）"]) == "A"
    assert apply_check_labels("Ａ", ["半角変換"]) == "A"
    assert apply_check_labels("44708", ["日付変換"], raw=44708.0) == "2022/05/27"
    assert (
        apply_check_labels("44708", ["日付変換 (yyyy/mm/dd)"], raw=44708.0) == "2022/05/27"
    )


def test_substring_false_positive_rejected() -> None:
    # 旧部分一致なら誤ってトリム／日付が走る
    assert apply_check_labels("  x  ", ["事前トリム処理"]) == "  x  "
    assert apply_check_labels("44708", ["更新日付メモ"], raw=44708.0) == "44708"


def test_alias_slots_cover_canonical_and_legacy() -> None:
    assert "トリム" in PROCESS_CHECK_ALIASES_BY_SLOT[0]
    assert "全角→半角" in PROCESS_CHECK_ALIASES_BY_SLOT[1]
    assert "半角変換" in PROCESS_CHECK_ALIASES_BY_SLOT[1]
    assert "年月日変換" in PROCESS_CHECK_ALIASES_BY_SLOT[2]
    assert "日付変換" in PROCESS_CHECK_ALIASES_BY_SLOT[2]
