# -*- coding: utf-8 -*-
"""日付変換（dt_ymd / dt_hm）向けの選択範囲縮小・読込・変換・差分書込ヘルパ。"""
from __future__ import annotations

from typing import Any, Callable, Optional

import pandas as pd


def _cell_changed(a: Any, b: Any) -> bool:
    try:
        if pd.isna(a) and pd.isna(b):
            return False
    except Exception:
        pass
    return a != b


def _row_changed(row_o: list[Any], row_f: list[Any]) -> bool:
    cols = min(len(row_o), len(row_f))
    if cols <= 0:
        return False
    return any(_cell_changed(row_o[j], row_f[j]) for j in range(cols))


def snapshot_region_from_areas(
    areas: list[tuple[int, int, int, int]],
) -> tuple[int, int, int, int] | None:
    """選択 areas の外接矩形 (y1, x1, yn, xn) 1-based。Undo 部分スナップショット用。"""
    if not areas:
        return None
    y1 = min(int(a[0]) for a in areas)
    x1 = min(int(a[1]) for a in areas)
    y2 = max(int(a[0]) + int(a[2]) - 1 for a in areas)
    x2 = max(int(a[1]) + int(a[3]) - 1 for a in areas)
    yn = max(1, y2 - y1 + 1)
    xn = max(1, x2 - x1 + 1)
    return (y1, x1, yn, xn)


def trim_areas_to_used_range(
    ptr_s: Any,
    areas: list[tuple[int, int, int, int]],
) -> list[tuple[int, int, int, int]]:
    """各選択 Area と UsedRange の交差に縮小する（全列選択などの空行読込を避ける）。"""
    if not areas:
        return areas
    try:
        ur = ptr_s.used_range
        uy1 = int(ur.row)
        ux1 = int(ur.column)
        uyn = int(ur.rows.count)
        uxn = int(ur.columns.count)
        if uyn < 1 or uxn < 1:
            return areas
        uy2 = uy1 + uyn - 1
        ux2 = ux1 + uxn - 1
    except Exception:
        return areas

    out: list[tuple[int, int, int, int]] = []
    for y1, x1, yn, xn in areas:
        if yn < 1 or xn < 1:
            continue
        y2 = y1 + yn - 1
        x2 = x1 + xn - 1
        iy1 = max(y1, uy1)
        ix1 = max(x1, ux1)
        iy2 = min(y2, uy2)
        ix2 = min(x2, ux2)
        if iy1 <= iy2 and ix1 <= ix2:
            out.append((iy1, ix1, iy2 - iy1 + 1, ix2 - ix1 + 1))
    return out


def _read_chunk_row_count(yn: int) -> int:
    """COM 往復を減らすチャンク行数（大きい範囲は一括読込）。"""
    if yn <= 30_000:
        return yn
    return max(2_000, min(10_000, yn))


def read_sheet_matrix(
    ptr_s: Any,
    y1: int,
    x1: int,
    yn: int,
    xn: int,
    on_pct: Callable[[int, str, str], None],
    *,
    msg_read: str,
    custom_read: str,
    normalize_2d: Callable[[Any, int, int], list[list[Any]]],
) -> Optional[list[list[Any]]]:
    """シート範囲をチャンク単位で読み 2 次元リストで返す。失敗時 None。"""
    if yn <= 0 or xn <= 0:
        return []
    chunk_rows = _read_chunk_row_count(yn)
    acc: list[list[Any]] = []
    try:
        for r0 in range(0, yn, chunk_rows):
            r1 = min(r0 + chunk_rows, yn)
            pct = int(5 + (r1 / max(yn, 1)) * 35)
            on_pct(pct, msg_read, custom_read)
            rng = ptr_s.range((y1 + r0, x1), (y1 + r1 - 1, x1 + xn - 1))
            part = rng.value
            sub = normalize_2d(part, r1 - r0, xn)
            acc.extend(sub)
        return acc if len(acc) == yn else None
    except Exception:
        return None


def format_datetime_column(
    ser_col: pd.Series,
    ser_dt: pd.Series,
    fmt: str,
    shape_fallback: Callable[[Any], str],
) -> tuple[pd.Series, int, int]:
    """
    解析済み日時列を表示用文字列へ整形する（ベクトル化）。
    戻り値: (整形後 Series, 変換成功数, 非空セル数)
    """
    stripped = ser_col.astype(str).str.strip()
    non_empty = int((ser_col.notna() & (stripped != "")).sum())
    ser_fmt = ser_dt.dt.strftime(fmt)
    mask_ok = ser_dt.notna()
    success = int(mask_ok.sum())
    if not bool(mask_ok.any()):
        return ser_col.map(shape_fallback), success, non_empty
    if bool(mask_ok.all()):
        return ser_fmt.astype(object), success, non_empty
    out = ser_fmt.astype(object)
    fail_mask = ~mask_ok
    out.loc[fail_mask] = ser_col.loc[fail_mask].map(shape_fallback)
    return out, success, non_empty


def _group_contiguous_row_indices(indices: list[int]) -> list[tuple[int, int]]:
    """0 始まり行インデックスを連続ブロック (start, count) にまとめる。"""
    if not indices:
        return []
    sorted_idx = sorted(set(indices))
    groups: list[tuple[int, int]] = []
    start = sorted_idx[0]
    prev = sorted_idx[0]
    for idx in sorted_idx[1:]:
        if idx == prev + 1:
            prev = idx
            continue
        groups.append((start, prev - start + 1))
        start = idx
        prev = idx
    groups.append((start, prev - start + 1))
    return groups


def _likely_all_rows_changed(
    original: list[list[Any]],
    final: list[list[Any]],
) -> bool:
    """
    サンプリングで「ほぼ全行が変化」と推定できるか。
    真のとき全行書込 fast path を使う（偽陽性は余分な書込のみで安全）。
    """
    n = min(len(original), len(final))
    if n <= 0:
        return False
    if n <= 256:
        return all(_row_changed(original[i], final[i]) for i in range(n))
    step = max(1, n // 64)
    for i in range(0, n, step):
        if not _row_changed(original[i], final[i]):
            return False
    if not _row_changed(original[-1], final[-1]):
        return False
    return True


def rows_with_any_change(
    original: list[list[Any]],
    final: list[list[Any]],
) -> list[int]:
    """行ごとにセル値が変わった行の 0 始まりインデックス一覧。"""
    n = min(len(original), len(final))
    changed: list[int] = []
    for i in range(n):
        if _row_changed(original[i], final[i]):
            changed.append(i)
    return changed


def count_rows_to_write(
    original: list[list[Any]],
    final: list[list[Any]],
    *,
    mostly_changed: bool = False,
) -> int:
    n = min(len(original), len(final))
    if n <= 0:
        return 0
    if mostly_changed or _likely_all_rows_changed(original, final):
        return n
    return len(rows_with_any_change(original, final))


def write_changed_slices(
    sheet_pointer: Any,
    start_y: int,
    start_x: int,
    original: list[list[Any]],
    final: list[list[Any]],
    *,
    progress_cb: Optional[Callable[[int], None]] = None,
    text_mode: bool = False,
    mostly_changed: bool = False,
) -> int:
    """変更行のみ write_chunk で書込む。ほぼ全行変化時は一括書込 fast path。"""
    from core import core_xlc

    n = min(len(original), len(final))
    if n <= 0:
        return 0

    if mostly_changed or _likely_all_rows_changed(original, final):
        core_xlc.write_chunk(
            sheet_pointer,
            start_y,
            start_x,
            final,
            progress_cb=None,
            text_mode=text_mode,
        )
        if progress_cb is not None:
            try:
                progress_cb(n)
            except Exception:
                pass
        return n

    groups = _group_contiguous_row_indices(rows_with_any_change(original, final))
    if not groups:
        return 0
    written = 0
    for row_off, row_count in groups:
        chunk = final[row_off : row_off + row_count]
        core_xlc.write_chunk(
            sheet_pointer,
            start_y + row_off,
            start_x,
            chunk,
            progress_cb=None,
            text_mode=text_mode,
        )
        written += row_count
        if progress_cb is not None:
            try:
                progress_cb(written)
            except Exception:
                pass
    return written
