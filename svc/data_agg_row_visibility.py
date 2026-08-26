# -*- coding: utf-8 -*-
"""Excel 行の非表示（手動・オートフィルタ）判定。抽出の skip_hidden_rows 用。"""
from __future__ import annotations

import logging
from pathlib import Path
from typing import Any, Callable, Optional

logger = logging.getLogger(__name__)

_CACHE_KEY = "hidden_row_sets"  # frame[path][sheet_key] -> set[int] (0-based) | frozenset


def source_wants_skip_hidden_rows(src: dict[str, Any] | None) -> bool:
    """ソースで非表示・フィルタ行を走査から除外するか（既定 OFF）。"""
    if not isinstance(src, dict):
        return False
    v = src.get("skip_hidden_rows")
    if v is True or v == 1:
        return True
    if isinstance(v, str) and v.strip().lower() in ("1", "true", "yes", "on"):
        return True
    return False


def _sheet_cache_key(sheet_name: Optional[str]) -> str:
    return str(sheet_name or "").strip() or "__LEFTMOST__"


def _frame_hidden_cache() -> dict[str, Any] | None:
    try:
        from svc.svc_data_agg_extract import _xlsx_workbook_cache_top
    except Exception:
        return None
    frame = _xlsx_workbook_cache_top()
    if frame is None:
        return None
    return frame.setdefault(_CACHE_KEY, {})


def _load_hidden_rows_xlsx(path: Path, sheet_name: Optional[str]) -> set[int]:
    """openpyxl 通常ロードで row_dimensions.hidden を収集（0 始まり行）。"""
    try:
        import openpyxl
    except ImportError:
        return set()
    hidden: set[int] = set()
    try:
        wb = openpyxl.load_workbook(
            path, read_only=False, data_only=False, keep_links=False
        )
    except Exception as e:
        logger.debug("[DATA_AGG_HIDDEN] xlsx open 失敗 %s: %s", path, e)
        return set()
    try:
        names = list(getattr(wb, "sheetnames", None) or [])
        sn = str(sheet_name or "").strip()
        if sn and sn in names:
            ws = wb[sn]
        else:
            ws = wb.active
        dims = getattr(ws, "row_dimensions", None)
        if dims is None:
            return hidden
        for idx, dim in dims.items():
            try:
                r1 = int(idx)
            except (TypeError, ValueError):
                continue
            if bool(getattr(dim, "hidden", False)):
                if r1 >= 1:
                    hidden.add(r1 - 1)
    except Exception as e:
        logger.debug("[DATA_AGG_HIDDEN] xlsx dims 失敗 %s: %s", path, e)
    finally:
        try:
            wb.close()
        except Exception:
            pass
    return hidden


def _load_hidden_rows_xls(path: Path, sheet_name: Optional[str]) -> set[int]:
    """xlrd formatting_info=True で rowinfo_map.hidden を収集（0 始まり行）。"""
    try:
        import xlrd
    except ImportError:
        return set()
    hidden: set[int] = set()
    try:
        book = xlrd.open_workbook(
            str(path), on_demand=False, formatting_info=True
        )
    except Exception as e:
        logger.debug("[DATA_AGG_HIDDEN] xls open 失敗 %s: %s", path, e)
        return set()
    try:
        names = list(book.sheet_names() or [])
        sn = str(sheet_name or "").strip()
        if sn and sn in names:
            sheet = book.sheet_by_name(sn)
        else:
            sheet = book.sheet_by_index(0)
        info_map = getattr(sheet, "rowinfo_map", None) or {}
        for r, info in info_map.items():
            try:
                ri = int(r)
            except (TypeError, ValueError):
                continue
            if bool(getattr(info, "hidden", False)):
                hidden.add(ri)
    except Exception as e:
        logger.debug("[DATA_AGG_HIDDEN] xls rowinfo 失敗 %s: %s", path, e)
    finally:
        try:
            release = getattr(book, "release_resources", None)
            if callable(release):
                release()
        except Exception:
            pass
    return hidden


def get_hidden_excel_rows(
    path: Path | str,
    sheet_name: Optional[str] = None,
) -> set[int]:
    """
    非表示（手動・フィルタで隠れた）行の 0 始まり集合。
    .csv や判定不能時は空集合。xlsx_workbook_scope 内ではキャッシュする。
    """
    p = Path(path)
    suf = p.suffix.lower()
    if suf not in (".xlsx", ".xlsm", ".xls"):
        return set()
    sk = _sheet_cache_key(sheet_name)
    pk = str(p.resolve())
    cache = _frame_hidden_cache()
    if cache is not None:
        by_sheet = cache.setdefault(pk, {})
        hit = by_sheet.get(sk)
        if hit is not None:
            return set(hit)
    if suf in (".xlsx", ".xlsm"):
        hidden = _load_hidden_rows_xlsx(p, sheet_name)
    else:
        hidden = _load_hidden_rows_xls(p, sheet_name)
    if cache is not None:
        cache.setdefault(pk, {})[sk] = frozenset(hidden)
    return hidden


def make_row_hidden_predicate(
    path: Path | str,
    sheet_name: Optional[str],
    *,
    enabled: bool,
) -> Callable[[int], bool] | None:
    """
    Excel 行（0 始まり）が非表示なら True を返す述語。
    enabled=False または対象外形式なら None（スキップ無し）。
    """
    if not enabled:
        return None
    p = Path(path)
    if p.suffix.lower() not in (".xlsx", ".xlsm", ".xls"):
        return None
    hidden = get_hidden_excel_rows(p, sheet_name)
    if not hidden:
        return None

    def _is_hidden(excel_row_0: int) -> bool:
        try:
            return int(excel_row_0) in hidden
        except (TypeError, ValueError):
            return False

    return _is_hidden
