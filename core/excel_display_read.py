# -*- coding: utf-8 -*-
"""Excel 範囲を画面上の表示文字列として読み取る（CSV 保存向け）。"""
from __future__ import annotations

import csv
import io
import os
from typing import Any

from core import core_env
from core.core_log import get_logger

logger = get_logger(__name__)


def use_display_text_for_csv_save() -> bool:
    """画面上の表示文字列で CSV 保存するか（既定 ON）。HC_CSV_SV_USE_VALUE_READ=1 で OFF。"""
    if os.name != "nt":
        return False
    if core_env.truthy(core_env.get("HC_CSV_SV_USE_VALUE_READ")):
        return False
    if core_env.truthy(core_env.get("HC_CSV_SV_USE_DISPLAY_TEXT")):
        return True
    if core_env.get("HC_CSV_SV_USE_DISPLAY_TEXT", "").strip().lower() in ("0", "false", "no"):
        return False
    return True


def parse_excel_clipboard_tsv(
    clip_text: str,
    *,
    expected_rows: int = 0,
    expected_cols: int = 0,
) -> list[list[str]]:
    """Excel のコピー（タブ区切り）を 2 次元の文字列リストにパースする。"""
    if not clip_text:
        return []
    normalized = clip_text.replace("\r\n", "\n").replace("\r", "\n")
    if normalized.endswith("\n"):
        normalized = normalized[:-1]
    if "\n" not in normalized and "\t" not in normalized:
        return [[normalized]]
    reader = csv.reader(
        io.StringIO(normalized),
        delimiter="\t",
        lineterminator="\n",
        quoting=csv.QUOTE_MINIMAL,
    )
    rows: list[list[str]] = []
    for row in reader:
        rows.append([str(c) if c is not None else "" for c in row])
    if not rows:
        return []
    if expected_cols > 0:
        for i, row in enumerate(rows):
            if len(row) < expected_cols:
                rows[i] = row + [""] * (expected_cols - len(row))
            elif len(row) > expected_cols:
                rows[i] = row[:expected_cols]
    if expected_rows > 0 and len(rows) != expected_rows:
        logger.debug(
            "[EXCEL_DISPLAY_READ] clipboard row count mismatch expected=%s got=%s",
            expected_rows,
            len(rows),
        )
    return rows


def read_range_display_text_matrix(
    sheet_ptr: Any,
    *,
    row_start: int,
    col_start: int,
    n_rows: int,
    n_cols: int,
) -> list[list[str]]:
    """1 範囲を Excel 表示どおりの文字列行列で読む（Copy → クリップボード）。"""
    if n_rows < 1 or n_cols < 1:
        return []
    end_row = row_start + n_rows - 1
    end_col = col_start + n_cols - 1
    rng = sheet_ptr.range((row_start, col_start), (end_row, end_col))
    app_api = None
    try:
        app_api = sheet_ptr.book.app.api
    except Exception:
        pass
    try:
        rng.api.Copy()
    except Exception as ex:
        raise RuntimeError(f"Excel Copy failed: {ex}") from ex
    try:
        clip = _read_clipboard_unicode_windows()
    finally:
        if app_api is not None:
            try:
                app_api.CutCopyMode = False
            except Exception:
                pass
    return parse_excel_clipboard_tsv(
        clip,
        expected_rows=n_rows,
        expected_cols=n_cols,
    )


def _read_clipboard_unicode_windows() -> str:
    import win32clipboard  # noqa: WPS433

    win32clipboard.OpenClipboard()
    try:
        if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_UNICODETEXT):
            data = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
            return str(data) if data is not None else ""
        return ""
    finally:
        win32clipboard.CloseClipboard()
