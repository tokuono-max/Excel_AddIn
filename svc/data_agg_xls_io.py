# -*- coding: utf-8 -*-
"""旧形式 Excel（.xls / BIFF）の読取。openpyxl 非対応のため xlrd を用いる。"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any, Optional

logger = logging.getLogger(__name__)

_XLS_BOOK_CACHE_KEY = "xls_books"
_XLS_MAT_CACHE_KEY = "xls_sheet_mats"

try:
    import xlrd as _xlrd_mod
except ImportError:
    _xlrd_mod = None


def is_xls_suffix(suffix: str) -> bool:
    return (suffix or "").lower() == ".xls"


def _import_xlrd() -> Any | None:
    if _xlrd_mod is None:
        logger.warning("[DATA_AGG_XLS] xlrd が利用できません（.xls 読取不可）")
    return _xlrd_mod


def xls_reader_unavailable_message() -> str | None:
    """xlrd 未導入時のユーザー向け文言。利用可なら None。"""
    if _xlrd_mod is None:
        return "（.xls読取不可: xlrd未導入）"
    return None


def open_xls_workbook(path: Path | str, *, on_demand: bool = True) -> Any | None:
    """xlrd Book を開く。失敗時 None。"""
    xlrd = _import_xlrd()
    if xlrd is None:
        return None
    p = Path(path)
    try:
        # formatting_info=False: 値取得向け。on_demand でシート単位遅延。
        return xlrd.open_workbook(str(p), on_demand=on_demand, formatting_info=False)
    except Exception as e:
        logger.warning("[DATA_AGG_XLS] open 失敗 %s: %s", p, e)
        return None


def close_xls_workbook(book: Any) -> None:
    if book is None:
        return
    try:
        release = getattr(book, "release_resources", None)
        if callable(release):
            release()
    except Exception:
        pass


def list_xls_sheet_names_from_book(book: Any) -> list[str]:
    """開済み xlrd Book からシート名一覧（open/close しない）。"""
    if book is None:
        return []
    try:
        return [str(x) for x in (book.sheet_names() or []) if str(x).strip() != ""]
    except Exception as e:
        logger.warning("[DATA_AGG_XLS] sheet_names(book) 失敗: %s", e)
        return []


def list_xls_sheet_names(path: Path | str) -> list[str]:
    """シート名を左→右順で返す。失敗時は空リスト。"""
    if _xlrd_mod is None:
        _import_xlrd()
        return []
    book = open_xls_workbook(path, on_demand=True)
    if book is None:
        return []
    try:
        names = [str(x) for x in (book.sheet_names() or []) if str(x).strip() != ""]
        return names
    except Exception as e:
        logger.warning("[DATA_AGG_XLS] sheet_names 失敗 %s: %s", path, e)
        return []
    finally:
        close_xls_workbook(book)


def _xls_book_from_scope(path: Path) -> Any | None:
    """xlsx_workbook_scope 相当フレームに xlrd Book をキャッシュする。"""
    try:
        from svc.svc_data_agg_extract import _xlsx_workbook_cache_top
    except Exception:
        return None
    frame = _xlsx_workbook_cache_top()
    if frame is None:
        return None
    books: dict[str, Any] = frame.setdefault(_XLS_BOOK_CACHE_KEY, {})
    key = str(path.resolve())
    hit = books.get(key)
    if hit is not None:
        return hit
    book = open_xls_workbook(path, on_demand=False)
    if book is None:
        return None
    books[key] = book
    return book


def resolve_xls_sheet(book: Any, sheet_name: Optional[str]) -> Any | None:
    """シートオブジェクト。名前一致が無ければ先頭。"""
    if book is None:
        return None
    try:
        names = list(book.sheet_names() or [])
    except Exception:
        return None
    if not names:
        return None
    sn = str(sheet_name or "").strip()
    if sn and sn in names:
        try:
            return book.sheet_by_name(sn)
        except Exception:
            pass
    try:
        return book.sheet_by_index(0)
    except Exception:
        return None


def _cell_value_from_sheet(sheet: Any, row: int, col: int, book: Any) -> Any:
    """0 始まり行列から Python 値へ。日付は datetime 化を試みる。"""
    if sheet is None or row < 0 or col < 0:
        return None
    try:
        if row >= int(sheet.nrows) or col >= int(sheet.ncols):
            return None
    except Exception:
        return None
    try:
        cell = sheet.cell(row, col)
    except Exception:
        return None
    xlrd = _import_xlrd()
    if xlrd is None:
        return getattr(cell, "value", None)
    ctype = getattr(cell, "ctype", None)
    val = getattr(cell, "value", None)
    if ctype == xlrd.XL_CELL_EMPTY or ctype == xlrd.XL_CELL_BLANK:
        return None
    if ctype == xlrd.XL_CELL_DATE:
        try:
            from svc.data_agg_excel_read import extract_read_scalar

            dt = xlrd.xldate_as_datetime(val, getattr(book, "datemode", 0))
            return extract_read_scalar(dt)
        except Exception:
            return val
    if ctype == xlrd.XL_CELL_BOOLEAN:
        return bool(val)
    if ctype == xlrd.XL_CELL_ERROR:
        return None
    # NUMBER / TEXT / others
    if isinstance(val, float) and val.is_integer():
        # 見た目整数の数値は int に寄せる（openpyxl data_only と近い）
        try:
            return int(val)
        except Exception:
            return val
    return val


def read_xls_cell(
    path: Path | str,
    sheet_name: Optional[str],
    cell_ref: str,
    *,
    book: Any | None = None,
) -> Any:
    """
    .xls から A1 セル値を返す。
    book を渡した場合はそれを使い（呼び出し側で close 責任）、未指定時はスコープキャッシュまたは都度 open。
    """
    from svc.svc_data_agg_extract import _parse_cell_ref

    p = Path(path)
    col, row = _parse_cell_ref(cell_ref)
    if col is None or row is None:
        return None
    owned = False
    wb = book
    if wb is None:
        wb = _xls_book_from_scope(p)
    if wb is None:
        wb = open_xls_workbook(p, on_demand=True)
        owned = wb is not None
    if wb is None:
        return None
    try:
        sheet = resolve_xls_sheet(wb, sheet_name)
        return _cell_value_from_sheet(sheet, row, col, wb)
    finally:
        if owned:
            close_xls_workbook(wb)


def materialize_xls_sheet_matrix(
    path: Path | str,
    sheet_name: Optional[str],
    *,
    book: Any | None = None,
) -> list[list[Any]]:
    """シート全セルを行リストにする（反復読取用）。"""
    p = Path(path)
    owned = False
    wb = book
    if wb is None:
        wb = _xls_book_from_scope(p)
    if wb is None:
        wb = open_xls_workbook(p, on_demand=False)
        owned = wb is not None
    if wb is None:
        return []
    try:
        sheet = resolve_xls_sheet(wb, sheet_name)
        if sheet is None:
            return []
        rows: list[list[Any]] = []
        nrows = int(getattr(sheet, "nrows", 0) or 0)
        ncols = int(getattr(sheet, "ncols", 0) or 0)
        for r in range(nrows):
            row_vals: list[Any] = []
            for c in range(ncols):
                row_vals.append(_cell_value_from_sheet(sheet, r, c, wb))
            rows.append(row_vals)
        return rows
    finally:
        if owned:
            close_xls_workbook(wb)


def get_xls_sheet_matrix_cached(
    path: Path,
    sheet_name: Optional[str],
) -> Optional[list[list[Any]]]:
    """スコープ内なら行列をキャッシュ。スコープ外は都度 materialize。"""
    try:
        from svc.svc_data_agg_extract import _xlsx_workbook_cache_top
    except Exception:
        return materialize_xls_sheet_matrix(path, sheet_name)

    frame = _xlsx_workbook_cache_top()
    if frame is None:
        return materialize_xls_sheet_matrix(path, sheet_name)
    mats: dict[Any, list[list[Any]]] = frame.setdefault(_XLS_MAT_CACHE_KEY, {})
    resolved_name = str(sheet_name or "").strip() or "__LEFTMOST__"
    key = (str(path.resolve()), resolved_name)
    hit = mats.get(key)
    if hit is not None:
        return hit
    book = _xls_book_from_scope(path)
    mat = materialize_xls_sheet_matrix(path, sheet_name, book=book)
    # 実シート名で再キー（左端解決後）
    if book is not None:
        sh = resolve_xls_sheet(book, sheet_name)
        try:
            title = str(getattr(sh, "name", "") or resolved_name)
        except Exception:
            title = resolved_name
        key2 = (str(path.resolve()), title)
        mats[key2] = mat
        if key2 != key:
            mats[key] = mat
    else:
        mats[key] = mat
    return mat


def read_xls_repeated_series(
    path: Path | str,
    sheet_name: Optional[str],
    *,
    base_col: int,
    base_row: int,
    row_step: int,
    col_step: int,
    limit: int,
    repeat_until_empty: bool,
    skip_row_hidden: Any = None,
    rule_iters_out: Any = None,
) -> list[Any]:
    """縦/横反復を行列から読む。"""
    from svc.svc_data_agg_extract import _matrix_cell_value, _read_repeated_series_from_matrix

    p = Path(path)
    mat = get_xls_sheet_matrix_cached(p, sheet_name)
    if not mat:
        return []
    if row_step == 0 and col_step == 0:
        if skip_row_hidden is not None and skip_row_hidden(base_row):
            return []
        v0 = _matrix_cell_value(mat, base_col, base_row)
        if repeat_until_empty and (v0 is None or v0 == ""):
            return []
        out = [v0] * max(0, limit)
        if rule_iters_out is not None:
            rule_iters_out.clear()
            rule_iters_out.extend([0] * len(out))
        return out
    return _read_repeated_series_from_matrix(
        mat,
        base_col=base_col,
        base_row=base_row,
        row_step=row_step,
        col_step=col_step,
        limit=limit,
        repeat_until_empty=repeat_until_empty,
        skip_row_hidden=skip_row_hidden,
        rule_iters_out=rule_iters_out,
    )
