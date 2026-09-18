# -*- coding: utf-8 -*-
"""シナリオ取得ソースに紐づく UI 由来の保存ブロック（ファイルフィルタ・シート・セル・連携／結合定義等）の JSON キー。"""

from __future__ import annotations

from collections.abc import Sequence
from typing import Any

from svc.data_agg_sheet_resolve import parse_comma_separated_patterns

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


def item_source_file_patterns(item: dict[str, Any]) -> list[str]:
    """
    セル系 sources の file_pattern トークン（小文字）を重複なく列挙。

    カンマ区切りは parse_comma_separated_patterns（抽出フィルタと同じ）で分割する。
    横断判定のトークン比較・デバッグ表示用。厳密なファイル一致は item_file_filter_specs。
    """
    patterns: list[str] = []
    for src in item.get("sources") or []:
        if not isinstance(src, dict):
            continue
        if str(src.get("type") or "cell").strip().lower() != "cell":
            continue
        block = source_ui_block(src)
        if not isinstance(block, dict):
            continue
        for tok in parse_comma_separated_patterns(block.get("file_pattern")):
            p = tok.lower()
            if p and p not in patterns:
                patterns.append(p)
    return patterns


def item_file_filter_specs(item: dict[str, Any]) -> list[dict[str, str]]:
    """
    抽出と同じ file_pattern / file_name_rule を持つ制限ソースの仕様一覧。

    file_pattern トークンが空のソースは含めない（フィルタなし）。
    """
    specs: list[dict[str, str]] = []
    for src in item.get("sources") or []:
        if not isinstance(src, dict):
            continue
        if str(src.get("type") or "cell").strip().lower() != "cell":
            continue
        block = source_ui_block(src)
        if not isinstance(block, dict):
            continue
        raw_pat = str(block.get("file_pattern") or "")
        if not parse_comma_separated_patterns(raw_pat):
            continue
        rule = str(block.get("file_name_rule") or "含む").strip() or "含む"
        specs.append({"file_pattern": raw_pat, "file_name_rule": rule})
    return specs


def file_path_matches_filter_specs(
    file_path: str, specs: Sequence[dict[str, Any]] | None
) -> bool:
    """
    結合／出力行判定を抽出と同じ file_name フィルタで評価する。

    specs 空 → False。いずれかの spec が True → True。
    ブックキャッシュには触れない。
    """
    if not specs:
        return False
    from svc.svc_data_agg_extract import source_passes_file_name_filter

    for spec in specs:
        if not isinstance(spec, dict):
            continue
        src = {
            "type": "cell",
            "ui_scenario_source_v1": {
                "file_pattern": spec.get("file_pattern"),
                "file_name_rule": spec.get("file_name_rule") or "含む",
            },
        }
        if source_passes_file_name_filter(file_path, src):
            return True
    return False


def item_join_defs_list(item: dict[str, Any]) -> list[dict[str, Any]]:
    """セル系ソースの join_defs。キャッシュには触れない。"""
    for src in item.get("sources") or []:
        if isinstance(src, dict) and str(src.get("type") or "").strip().lower() == "cell":
            block = source_ui_block(src)
            if isinstance(block, dict):
                return [x for x in (block.get("join_defs") or []) if isinstance(x, dict)]
    return []


def item_link_defs_list(item: dict[str, Any]) -> list[dict[str, Any]]:
    """セル系ソースの link_defs。判定は従来どおり。キャッシュには触れない。"""
    for src in item.get("sources") or []:
        if isinstance(src, dict) and str(src.get("type") or "").strip().lower() == "cell":
            block = source_ui_block(src)
            if isinstance(block, dict):
                return [x for x in (block.get("link_defs") or []) if isinstance(x, dict)]
    return []


def join_search_targets_from_defs(join_defs: list[dict[str, Any]]) -> list[str]:
    targets: list[str] = []
    for jd in join_defs:
        name = str(jd.get("item") or "").strip()
        if name and name not in targets:
            targets.append(name)
    return targets


def join_comparison_side_items(
    host_item: dict[str, Any],
    items: list[dict[str, Any]],
    headers: list[str],
) -> list[dict[str, Any]]:
    """横断結合の比較列を供給する項目一覧。判定は従来どおり。キャッシュには触れない。"""
    join_targets = join_search_targets_from_defs(item_join_defs_list(host_item))
    if not join_targets:
        return []
    out: list[dict[str, Any]] = []
    for i, it in enumerate(items):
        if not isinstance(it, dict):
            continue
        h = headers[i] if i < len(headers) else str(it.get("name") or it.get("id") or "")
        h = str(h).strip()
        supplies = h in join_targets
        if not supplies:
            for src in it.get("sources") or []:
                if not isinstance(src, dict):
                    continue
                block = source_ui_block(src)
                if not isinstance(block, dict):
                    continue
                for ld in block.get("link_defs") or []:
                    if str(ld.get("item") or "").strip() in join_targets:
                        supplies = True
                        break
                if supplies:
                    break
        if supplies:
            out.append(it)
    return out


def join_comparison_side_file_filter_specs(
    host_item: dict[str, Any],
    items: list[dict[str, Any]],
    headers: list[str],
) -> list[dict[str, str]]:
    """横断結合の比較側項目の file フィルタ仕様。抽出ルールは変えない。"""
    specs: list[dict[str, str]] = []
    for it in join_comparison_side_items(host_item, items, headers):
        for spec in item_file_filter_specs(it):
            if spec not in specs:
                specs.append(spec)
    return specs


def join_comparison_side_file_patterns(
    host_item: dict[str, Any],
    items: list[dict[str, Any]],
    headers: list[str],
) -> list[str]:
    """横断結合の比較列を供給する項目の file_pattern（小文字・重複なし）。"""
    specs = join_comparison_side_file_filter_specs(host_item, items, headers)
    patterns: list[str] = []
    for spec in specs:
        for tok in parse_comma_separated_patterns(spec.get("file_pattern")):
            p = tok.lower()
            if p and p not in patterns:
                patterns.append(p)
    return patterns


def patterns_overlap(a: list[str], b: list[str]) -> bool:
    if not a or not b:
        return True
    for pa in a:
        for pb in b:
            if pa in pb or pb in pa:
                return True
    return False


def join_host_needs_cross_file_pool(
    host_item: dict[str, Any],
    items: list[dict[str, Any]],
    headers: list[str],
) -> bool:
    """
    結合ホストの file_pattern と、結合比較列を供給する項目の pattern が異なるとき True。

    headers は呼び出し互換のため受ける。判定には使わない。ブックキャッシュには触れない。
    """
    _ = headers
    join_targets = join_search_targets_from_defs(item_join_defs_list(host_item))
    if not join_targets:
        return False
    host_patterns = item_source_file_patterns(host_item)

    def _differs(other_patterns: list[str]) -> bool:
        return bool(
            host_patterns
            and other_patterns
            and not patterns_overlap(host_patterns, other_patterns)
        )

    for it in items:
        if not isinstance(it, dict):
            continue
        name = str(it.get("name") or it.get("id") or "").strip()
        if name not in join_targets:
            continue
        if _differs(item_source_file_patterns(it)):
            return True
    for target in join_targets:
        for it in items:
            if not isinstance(it, dict):
                continue
            for src in it.get("sources") or []:
                if not isinstance(src, dict):
                    continue
                block = source_ui_block(src)
                if not isinstance(block, dict):
                    continue
                for ld in block.get("link_defs") or []:
                    if str(ld.get("item") or "").strip() != target:
                        continue
                    if _differs(item_source_file_patterns(it)):
                        return True
    return False


def scenario_has_join_defs(items: list[dict[str, Any]]) -> bool:
    """いずれかのセル系ソースに join_defs があるか。判定は従来どおり。キャッシュには触れない。"""
    for it in items:
        if not isinstance(it, dict):
            continue
        for src in it.get("sources") or []:
            if not isinstance(src, dict):
                continue
            if (src.get("type") or "").strip().lower() != "cell":
                continue
            pb = source_ui_block(src)
            if isinstance(pb, dict) and (pb.get("join_defs") or []):
                return True
    return False


def collect_linked_and_join_targets(
    items: list[dict[str, Any]],
) -> tuple[set[str], set[str]]:
    """シナリオ内の連携先項目名・結合項目名。判定は従来どおり。キャッシュには触れない。"""
    linked_targets: set[str] = set()
    join_targets: set[str] = set()
    for item in items:
        for src in item.get("sources") or []:
            if not isinstance(src, dict):
                continue
            pb = source_ui_block(src)
            if not isinstance(pb, dict):
                continue
            for ld in pb.get("link_defs") or []:
                if isinstance(ld, dict):
                    nm = str(ld.get("item") or "").strip()
                    if nm:
                        linked_targets.add(nm)
            for jd in pb.get("join_defs") or []:
                if isinstance(jd, dict):
                    nm = str(jd.get("item") or "").strip()
                    if nm:
                        join_targets.add(nm)
            if (src.get("type") or "").strip().lower() == "name_extract":
                nm2 = str(pb.get("path_item") or "").strip()
                if nm2:
                    linked_targets.add(nm2)
    return linked_targets, join_targets


def item_sources_pass_file(item: dict[str, Any], file_path: str) -> bool:
    """項目の sources のいずれかがファイル名フィルタを通るか。照合は抽出と同じ関数。"""
    from svc.svc_data_agg_extract import source_passes_file_name_filter

    for src in item.get("sources") or []:
        if isinstance(src, dict) and source_passes_file_name_filter(file_path, src):
            return True
    return False
