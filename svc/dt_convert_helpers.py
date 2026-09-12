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
    if not isinstance(ser_dt, pd.Series):
        ser_dt = pd.Series(ser_dt, index=ser_col.index)
    # .where 結合や COM 混在で object 化することがある → .dt 前に必ず datetime 化
    if not pd.api.types.is_datetime64_any_dtype(ser_dt):
        ser_dt = pd.to_datetime(ser_dt, errors="coerce")
    mask_ok = ser_dt.notna()
    success = int(mask_ok.sum())
    if success == 0 or not pd.api.types.is_datetime64_any_dtype(ser_dt):
        return ser_col.map(shape_fallback), success, non_empty
    ser_fmt = ser_dt.dt.strftime(fmt)
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


def elapsed_ms(since: float) -> int:
    """perf_counter 起点からの経過ミリ秒（負にならない）。"""
    import time

    return max(0, int((time.perf_counter() - since) * 1000))


def status_bar_save(book: Any) -> str:
    """現在の Excel ステータスバー文言を退避する。"""
    try:
        return str(book.app.api.StatusBar or "")
    except Exception:
        return ""


def status_bar_set(book: Any, msg: str) -> None:
    """Excel のステータスバーに指定メッセージを表示する。"""
    try:
        book.app.api.DisplayStatusBar = True
        book.app.api.StatusBar = str(msg)
    except Exception:
        pass


def status_bar_restore(book: Any, saved: str) -> None:
    """ステータスバーを status_bar_save で退避した文言に戻す。"""
    try:
        book.app.api.StatusBar = saved
    except Exception:
        pass


def restore_excel_screen_updating(ptr_a: Any) -> None:
    try:
        ptr_a.api.ScreenUpdating = True
    except Exception:
        pass


def get_window_rect(hwnd: int) -> tuple[int, int, int, int] | None:
    """Win32 GetWindowRect → (left, top, right, bottom)。非 NT / 失敗時 None。"""
    import os

    if not int(hwnd or 0) or os.name != "nt":
        return None
    try:
        import ctypes
        from ctypes import wintypes

        r = wintypes.RECT()
        if ctypes.windll.user32.GetWindowRect(int(hwnd), ctypes.byref(r)):
            return (int(r.left), int(r.top), int(r.right), int(r.bottom))
    except Exception:
        pass
    return None


def normalize_2d(raw: Any, yn: int, xn: int) -> list[list[Any]]:
    """xlwings Range.value を yn×xn の 2 次元リストに正規化する。"""
    if yn <= 0 or xn <= 0:
        return []
    if yn == 1 and xn == 1:
        return [[raw]]
    if yn == 1:
        row = raw if isinstance(raw, list) else [raw]
        return [row[:xn] + [None] * max(0, xn - len(row))]
    out: list[list[Any]] = []
    if not isinstance(raw, list):
        return [[None] * xn for _ in range(yn)]
    for r in range(yn):
        row = raw[r] if r < len(raw) else None
        if row is None:
            out.append([None] * xn)
        elif isinstance(row, list):
            out.append((row + [None] * xn)[:xn])
        else:
            out.append([row] + [None] * (xn - 1) if xn > 1 else [row])
    return out


def normalize_date_text(v: Any) -> Any:
    """日付文字列の軽い正規化（to_datetime 向け）。変換不能判定は呼び出し側。"""
    import re
    import unicodedata

    if v is None:
        return v
    try:
        if pd.isna(v):
            return v
    except Exception:
        pass

    s = str(v)
    if not s:
        return s

    s = unicodedata.normalize("NFKC", s)
    s = s.replace("\u3000", " ").strip()
    s = re.sub(r"\s+", " ", s)
    s = re.sub(r"[（(][^）)]*[）)]", "", s)
    s = re.sub(r"\s*年\s*", "/", s)
    s = re.sub(r"\s*月\s*", "/", s)
    s = re.sub(r"\s*日\s*", "", s)
    s = s.replace(".", "/").replace("-", "/")
    s = re.sub(r"/{2,}", "/", s)
    return s.strip()


def parse_datetime_with_normalized_fallback(ser_col: pd.Series) -> pd.Series:
    """to_datetime を優先し、失敗分のみ normalize_date_text で再判定する。"""
    if not isinstance(ser_col, pd.Series):
        ser_col = pd.Series(ser_col)
    ser_dt = pd.to_datetime(ser_col, errors="coerce")
    if not isinstance(ser_dt, pd.Series):
        ser_dt = pd.Series(ser_dt, index=ser_col.index)
    mask_failed = ser_dt.isna()
    if bool(mask_failed.any()):
        ser_norm = ser_col.map(normalize_date_text)
        ser_dt_norm = pd.to_datetime(ser_norm, errors="coerce")
        if not isinstance(ser_dt_norm, pd.Series):
            ser_dt_norm = pd.Series(ser_dt_norm, index=ser_col.index)
        # Series.where は dtype が object に落ち .dt が使えなくなることがある
        merged = ser_dt.to_numpy(dtype=object, copy=True)
        failed = mask_failed.to_numpy()
        merged[failed] = ser_dt_norm.to_numpy(dtype=object)[failed]
        ser_dt = pd.to_datetime(pd.Series(merged, index=ser_col.index), errors="coerce")
    if not pd.api.types.is_datetime64_any_dtype(ser_dt):
        ser_dt = pd.to_datetime(ser_dt, errors="coerce")
    return ser_dt
