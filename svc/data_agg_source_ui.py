# -*- coding: utf-8 -*-
"""シナリオ取得ソースに紐づく UI 由来の保存ブロック（ファイルフィルタ・シート・セル・連携／結合定義等）の JSON キー。"""

from __future__ import annotations

from typing import Any

# 現行キー（保存の正）
SCENARIO_SOURCE_UI_KEY = "ui_scenario_source_v1"
# 旧キー（読み込みのみ後方互換・値は既存シナリオ JSON と一致させる必要あり）
SCENARIO_SOURCE_UI_KEY_LEGACY = "ui_scenario_proto_v1"


def source_ui_block(source: dict[str, Any]) -> dict[str, Any] | None:
    """取得ソース dict から UI 保存ブロックを取得。現行キーがあればそれを優先（空 dict も正当）。無い・非 dict ならレガシー。"""
    if SCENARIO_SOURCE_UI_KEY in source:
        v = source.get(SCENARIO_SOURCE_UI_KEY)
        if isinstance(v, dict):
            return v
    leg = source.get(SCENARIO_SOURCE_UI_KEY_LEGACY)
    return leg if isinstance(leg, dict) else None


def ensure_source_ui_block(source: dict[str, Any]) -> dict[str, Any]:
    """編集用にミュータブルな UI ブロックを返し、キャノニカルキーへ寄せてレガシーキーを除去する。"""
    p = source_ui_block(source)
    if p is None:
        p = {}
    source[SCENARIO_SOURCE_UI_KEY] = p
    source.pop(SCENARIO_SOURCE_UI_KEY_LEGACY, None)
    return p
