# -*- coding: utf-8 -*-
"""シナリオのシート名条件（左端／完全一致／含む／含まない）の解決。"""

from __future__ import annotations

from collections.abc import Sequence
from pathlib import Path
from typing import Any, Literal

SheetRuleKind = Literal["left", "exact", "contains", "not_contains"]

SHEET_MISS_LABEL = "（該当なし）"


def classify_sheet_rule(rule: str | None) -> SheetRuleKind:
    """UI 文言／英語エイリアスを正規化する。空・不明は left（左端）扱い。"""
    r = str(rule or "").strip()
    if not r:
        return "left"
    rl = r.lower()
    if "左端" in r or rl in ("left", "leftmost"):
        return "left"
    # 「含まない」を先に判定（「含む」より長い／優先）
    if "含まない" in r or rl in ("exclude", "not_contains", "not-contains"):
        return "not_contains"
    if "含む" in r or rl in ("contains", "include"):
        return "contains"
    if "完全一致" in r or rl in ("exact", "equals"):
        return "exact"
    return "left"


def resolve_all_sheet_names_by_rule(
    sheetnames: Sequence[str],
    rule: str | None,
    pattern: str | None,
    *,
    case_sensitive: bool = True,
) -> list[str]:
    """
    ブックのシート名一覧から、条件に合うシート名を左端から順にすべて返す。

    - left: 先頭シートのみ（1 件）
    - exact: 完全一致（0〜1 件）
    - contains / not_contains: 一致するすべて（ブック上の左→右順）
    - pattern が空の exact/contains/not_contains: 空リスト
    """
    names = [str(x) for x in sheetnames if str(x).strip() != ""]
    if not names:
        return []
    kind = classify_sheet_rule(rule)
    if kind == "left":
        return [names[0]]
    sn = str(pattern or "").strip()
    if not sn:
        return []

    def _key(s: str) -> str:
        return s if case_sensitive else s.casefold()

    sn_k = _key(sn)
    if kind == "exact":
        return [x for x in names if _key(x) == sn_k]
    if kind == "contains":
        return [x for x in names if sn_k in _key(x)]
    if kind == "not_contains":
        return [x for x in names if sn_k not in _key(x)]
    return [names[0]]


def resolve_sheet_name_by_rule(
    sheetnames: Sequence[str],
    rule: str | None,
    pattern: str | None,
    *,
    case_sensitive: bool = True,
) -> str | None:
    """条件に合う最初のシート名。該当なしは None。"""
    matched = resolve_all_sheet_names_by_rule(
        sheetnames, rule, pattern, case_sensitive=case_sensitive
    )
    return matched[0] if matched else None


def list_workbook_sheet_names(file_path: str | Path) -> list[str] | None:
    """
    ブックのシート名一覧。
    CSV などシート概念なしは None。読取失敗は []。
    """
    p = Path(file_path)
    suffix = p.suffix.lower()
    if suffix == ".csv":
        return None
    if suffix == ".xls":
        from svc.data_agg_xls_io import list_xls_sheet_names, xls_reader_unavailable_message

        if xls_reader_unavailable_message():
            return []
        return list_xls_sheet_names(p)
    if suffix in (".xlsx", ".xlsm"):
        try:
            import openpyxl  # noqa: E402
        except Exception:
            return []
        try:
            wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
            names = list(wb.sheetnames or [])
            wb.close()
            return [str(x) for x in names if str(x).strip() != ""]
        except Exception:
            return []
    return []


def matching_sheets_for_cell_source(
    file_path: str | Path,
    src: dict[str, Any] | None,
) -> list[str] | None:
    """
    セル系ソースのシート名条件に合うシート一覧（左→右）。
    名前取得系・CSV などシート解決不要時は None。
    該当なし・読取失敗は []。
    """
    if not isinstance(src, dict):
        return None
    typ = str(src.get("type") or "cell").strip().lower()
    if typ in ("name_extract", "metadata", "meta", "filename"):
        return None
    from svc.data_agg_source_ui import source_ui_block

    sn = str(src.get("sheet_name") or "").strip()
    pb = source_ui_block(src) or {}
    rule = str(pb.get("sheet_rule") or "")
    names = list_workbook_sheet_names(file_path)
    if names is None:
        return None
    if not names:
        return []
    return resolve_all_sheet_names_by_rule(names, rule, sn)


def patch_item_sheet_exact(item: dict[str, Any], sheet_name: str) -> dict[str, Any]:
    """sources[0] の sheet_name を実名にし、sheet_rule を完全一致にする。"""
    import copy

    from svc.data_agg_source_ui import ensure_source_ui_block

    out = copy.deepcopy(item)
    out_s0 = (out.get("sources") or [None])[0]
    if isinstance(out_s0, dict):
        out_s0["sheet_name"] = sheet_name
        ensure_source_ui_block(out_s0)["sheet_rule"] = "完全一致"
    return out
