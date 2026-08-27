# -*- coding: utf-8 -*-
"""
Python: 3.12
Module: core/core_xlc.py
Created: 2026-01-xx
Updated: 2026-06-06
Version: 2.5.12
Purpose:
  Excel COM 操作の薄いヘルパ（UI非依存 / core から ui import 禁止）。
  - シートカスタムプロパティ（GUID 等）の読み書き
  - Workbook Names への通知書き込み（VBAポーリング連携）
  - 大量書き込み（チャンク）など、Excelが固まりやすい箇所の安全策

Design:
  - core は ui/svc に依存しない（規約）
  - 失敗時に例外を投げない（Excelロック解除漏れの方が致命）

History (latest 3):
  - 2.5.12 (2026-06-14): get_excel_context_from_hwnd: sheet_id 指定時は全ブック走査（アクティブブック誤結合を防止）。
  - 2.5.11 (2026-06-06) restore_screen_updating ヘルパを追加（restore_on_exit=False 利用後の復帰用）。
  - 2.5.10 (2026-06-06) suspend_sheet_updates: restore_on_exit=False で ScreenUpdating 復帰を呼び出し側に委譲。
"""

from __future__ import annotations

import base64
import numbers
import re
import secrets
import time
from contextlib import contextmanager
from typing import Any, Callable, List, Optional

import pythoncom
from core.core_log import get_logger
from core import core_cst as cst


# 変数: バージョン情報
__version__ = "2.5.11"

EXCEL_SHEET_NAME_MAX_LEN: int = 31

logger = get_logger(__name__)


def sanitize_excel_sheet_name(name: str, *, fallback: str = "Sheet1") -> str:
    """Excel シート名に使えない文字を除去し、最大 31 文字に切り詰める。"""
    t = re.sub(r"[\[\]:*?/\\]", "", (name or "").strip())
    t = t[:EXCEL_SHEET_NAME_MAX_LEN]
    fb = (fallback or "Sheet1").strip()[:EXCEL_SHEET_NAME_MAX_LEN] or "Sheet1"
    return t if t else fb


def unique_excel_sheet_name_in_names(
    existing_names: set[str],
    base: str,
    *,
    fallback: str = "Sheet1",
) -> str:
    """既存シート名と衝突しないよう、31 文字以内でユニーク名を返す。"""
    b = sanitize_excel_sheet_name(base, fallback=fallback)
    if b not in existing_names:
        return b
    for i in range(2, 10000):
        suf = f"_{i}"
        cand = (b[: max(1, EXCEL_SHEET_NAME_MAX_LEN - len(suf))] + suf)[
            :EXCEL_SHEET_NAME_MAX_LEN
        ]
        if cand not in existing_names:
            return cand
    tail = "_99"
    return (b[: max(1, EXCEL_SHEET_NAME_MAX_LEN - len(tail))] + tail)[
        :EXCEL_SHEET_NAME_MAX_LEN
    ]


def excel_sheet_name_for_split_part(base: str, part_idx: int) -> str:
    """分割読込用シート名（例: base-2）。接尾辞分を確保して base を切り詰める。"""
    suffix = f"-{max(1, int(part_idx))}"
    stem = sanitize_excel_sheet_name(base, fallback="CSV")[
        : max(1, EXCEL_SHEET_NAME_MAX_LEN - len(suffix))
    ]
    return (stem + suffix)[:EXCEL_SHEET_NAME_MAX_LEN]


def com_excel_scalar_int(val: Any, default: int = 0) -> int:
    """
    Excel COM が返すスカラー（int / float または .Value 付きオブジェクト）を int にする。
    Range.Row / Column / Count がすでに int のとき、.Value 前提の getattr だと 0 になるため分岐する。
    """
    if val is None:
        return default
    if isinstance(val, numbers.Integral):
        return int(val)
    if isinstance(val, numbers.Real) and not isinstance(val, bool):
        return int(float(val))
    try:
        inner = getattr(val, "Value", val)
        if inner is None:
            return default
        if isinstance(inner, numbers.Integral):
            return int(inner)
        if isinstance(inner, numbers.Real) and not isinstance(inner, bool):
            return int(float(inner))
        return int(inner)
    except (TypeError, ValueError, AttributeError):
        return default


# ==============================================================================
# GUID helpers
# ==============================================================================
def create_guid_b64() -> str:
    """URL-safe な Base64 GUID を生成する。"""
    raw = secrets.token_bytes(16)
    return base64.urlsafe_b64encode(raw).decode("ascii").rstrip("=")


# ==============================================================================
# Workbook / Sheet helpers
# ==============================================================================
def get_sheet_prop(sheet_pointer: Any, key_name_string: str) -> str:
    """シートの CustomProperties から値を取得する（無ければ空）。"""
    try:
        api_sheet = sheet_pointer.api
        for prop in api_sheet.CustomProperties:
            if prop.Name == key_name_string:
                return str(prop.Value)
    except Exception:
        pass
    return ""


def set_sheet_prop(sheet_pointer: Any, key_name_string: str, value_string: str) -> bool:
    """シートの CustomProperties に値を設定する（既存は置換）。"""
    try:
        api_sheet = sheet_pointer.api
        for prop in api_sheet.CustomProperties:
            if prop.Name == key_name_string:
                prop.Delete()
        api_sheet.CustomProperties.Add(Name=key_name_string, Value=str(value_string))
        return True
    except Exception:
        return False


def find_sheet_by_guid(workbook_pointer: Any, target_guid_string: str) -> Optional[Any]:
    """ブック内の全シートを走査し GUID を持つシートを返す（無ければ None）。

    Notes:
        - VBA待機中の Excel へ Python が再 COM 侵入すると、特定環境で OLE/三角待ちを誘発しうる。
        - 再構築フェーズの切り分けでは、全シート走査（COM列挙）を避けるため、
          `cst.SVC_SKIP_FIND_SHEET_BY_GUID=True` の場合は **即 None** を返す。
    """
    if getattr(cst, "SVC_SKIP_FIND_SHEET_BY_GUID", False):
        try:
            logger.info("[COM_SKIP] find_sheet_by_guid skipped by cst flag")
        except Exception:
            pass
        return None

    if not target_guid_string or workbook_pointer is None:
        return None

    t0 = time.perf_counter()
 
    try:
        for sh in workbook_pointer.sheets:
            found = get_sheet_prop(sh, "HC_GUID_B64")
            if found == target_guid_string:
                return sh

    except Exception as ex:
        # ここで落とすと Excelロック解除漏れの方が致命。ログだけ残して None で戻す。
        try:
            dt_ms = int((time.perf_counter() - t0) * 1000)
            logger.warning("[COM_NG] find_sheet_by_guid failed dt_ms=%s ex=%r", dt_ms, ex)
        except Exception:
            pass
    return None


def find_book_and_sheet_by_guid_in_app(
    app: Any,
    target_guid_string: str,
) -> tuple[Any, Any] | None:
    """App 内の全ブックを走査し、GUID を持つシートとそのブックを返す。

    マルチブック／アクティブブック切替後も、起動時 sheet_id で正しいブックに結び付ける。
    """
    guid = str(target_guid_string or "").strip()
    if not guid:
        return None
    try:
        for book in app.books:
            sheet = find_sheet_by_guid(book, guid)
            if sheet is not None:
                return book, sheet
    except Exception as ex:
        try:
            logger.warning(
                "[XLC_CTX] find_book_and_sheet_by_guid_in_app ex=%r guid=%r",
                ex,
                guid,
            )
        except Exception:
            pass
    return None


# ==============================================================================
# Excel コンテキスト取得（HWND から app / book / sheet）
# ==============================================================================
def excel_try_set_main_commandbars_enabled(xw_app: Any, enabled: bool) -> None:
    """リボン／主要 CommandBar の有効・無効（ベストエフォート）。

    メインのデータ集約 UI 表示中にリボン操作を抑止し、ワークシート上の操作は継続しやすくする。

    ExecuteExcel4Macro の SHOW.TOOLBAR は Fluent UI のリボンを消せるが、タイトルバーまわりの違和感が大きいため使わない。
    Office 2007 以降、Fluent リボンの操作感は CommandBars だけでは抑止しきれないことがある。
    """
    try:
        api = getattr(xw_app, "api", None)
        if api is None:
            logger.info(
                "[XLC_CMD_BAR] excel_try_set_main_commandbars_enabled skipped: no api on app enabled=%s",
                enabled,
            )
            return
        cbs = api.CommandBars
        names = (
            "Ribbon",
            "Worksheet Menu Bar",
            "Chart Menu Bar",
            "Standard",
            "Formatting",
        )
        parts: list[str] = []
        item = getattr(cbs, "Item", None)
        if item is None or not callable(item):
            logger.info(
                "[XLC_CMD_BAR] excel_try_set_main_commandbars_enabled skipped: CommandBars.Item missing enabled=%s",
                enabled,
            )
            return
        for name in names:
            try:
                bar = item(name)
                setattr(bar, "Enabled", bool(enabled))
                parts.append("%s=ok" % name)
            except Exception as ex:
                parts.append("%s=fail:%s" % (name, ex))
        logger.info(
            "[XLC_CMD_BAR] enabled=%s per_bar=%s",
            enabled,
            "; ".join(parts),
        )
    except Exception as ex:
        logger.warning(
            "[XLC_CMD_BAR] excel_try_set_main_commandbars_enabled outer fail enabled=%s ex=%r",
            enabled,
            ex,
        )


def get_excel_context_from_hwnd(hwnd: int, sheet_id: str = "") -> Optional[tuple]:
    """HWND から xlwings の app, book, sheet を取得する。

    Returns:
        (app, book, sheet, hwnd) のタプル。失敗時は None。
    """
    ph = int(hwnd or 0)
    if ph == 0:
        logger.info("[XLC_CTX] get_excel_context_from_hwnd fail: hwnd_zero")
        return None
    try:
        import xlwings as xw  # noqa: PLC0415
        from xlwings._xlwindows import App as WinApp  # noqa: PLC0415

        app = xw.App(impl=WinApp(xl=ph))
        sheet_id_s = str(sheet_id or "").strip()
        if sheet_id_s:
            hit = find_book_and_sheet_by_guid_in_app(app, sheet_id_s)
            if hit is None:
                logger.info(
                    "[XLC_CTX] get_excel_context_from_hwnd fail: sheet_guid_not_found hwnd=%s sheet_id=%r",
                    ph,
                    sheet_id_s,
                )
                return None
            book, sheet = hit
            logger.info(
                "[XLC_CTX] get_excel_context_from_hwnd ok hwnd=%s sheet_id=%r book=%s via=guid_scan",
                ph,
                sheet_id_s,
                getattr(book, "name", "?"),
            )
            return (app, book, sheet, ph)

        book = app.books.active
        if not book:
            logger.info(
                "[XLC_CTX] get_excel_context_from_hwnd fail: no_active_workbook hwnd=%s",
                ph,
            )
            return None
        sheet = book.sheets.active
        if sheet is None:
            logger.info(
                "[XLC_CTX] get_excel_context_from_hwnd fail: no_sheet hwnd=%s sheet_id=%r",
                ph,
                sheet_id_s,
            )
            return None
        logger.info(
            "[XLC_CTX] get_excel_context_from_hwnd ok hwnd=%s sheet_id=%r book=%s",
            ph,
            sheet_id_s,
            getattr(book, "name", "?"),
        )
        return (app, book, sheet, ph)
    except Exception as ex:
        try:
            logger.warning(
                "[XLC_CTX] get_excel_context_from_hwnd exception hwnd=%s ex=%r",
                hwnd,
                ex,
            )
        except Exception:
            pass
        return None


def set_workbook_name(book: Any, name: str, refers_to: str) -> None:
    """Workbook.Names に値（RefersTo）を設定する（存在すれば置換）。

    Args:
        book: xlwings Book
        name: Name（例: "HC_NOTIFY_RETV"）
        refers_to: RefersTo 文字列（例: '="READY_UI|ok"'）

    Notes:
        - VBA 側は ThisWorkbook.Names(name).RefersTo を監視する想定。
        - Excel の Names は “式” 扱いなので、文字列は必ず ="..." 形式にする。
        - COM 例外は端末差で頻発しうるため、ここは落とさない。
    """
    try:
        api = book.api
        try:
            api.Names(name).Delete()
        except Exception:
            pass
        api.Names.Add(Name=name, RefersTo=refers_to)
    except Exception:
        pass


# ==============================================================================
# 一括書込み時の表示停止・有効領域外クリア
# ==============================================================================
def _get_app_api(sheet_or_book: Any) -> Any:
    """シートまたはブックから Excel Application (COM API) を取得する。"""
    book = getattr(sheet_or_book, "book", None) or sheet_or_book
    app = getattr(book, "app", None) if book else None
    return getattr(app, "api", None) if app else None


def restore_screen_updating(sheet_or_book: Any) -> None:
    """suspend_sheet_updates(restore_on_exit=False) 利用後に ScreenUpdating を復帰する。"""
    api = _get_app_api(sheet_or_book)
    if api is not None:
        try:
            api.ScreenUpdating = True
        except Exception:
            pass


@contextmanager
def suspend_sheet_updates(sheet_or_book: Any, *, restore_on_exit: bool = True) -> Any:
    """シート更新表示を停止して処理を行い、終了時に再開するコンテキストマネージャ。
    一括書込みの前に with で囲むと高速化に有効。
    restore_on_exit=False のとき ScreenUpdating は with 終了時に戻さない（呼び出し側で復帰）。
    """
    api = _get_app_api(sheet_or_book)
    try:
        if api is not None:
            api.ScreenUpdating = False
        yield
    finally:
        if restore_on_exit and api is not None:
            try:
                api.ScreenUpdating = True
            except Exception:
                pass


def clear_used_range_overflow(
    sheet_pointer: Any, data_rows: int, data_cols: int
) -> None:
    """シートの有効データ範囲 (data_rows, data_cols) を超える UsedRange の余白をクリアする。
    書込み後に有効領域が拡大しないように呼び出す。
    """
    if data_rows <= 0 or data_cols <= 0:
        return
    try:
        ur = getattr(sheet_pointer, "used_range", None)
        if ur is None:
            return
        uapi = getattr(ur, "api", None)
        if uapi is None:
            return
        rows_o = getattr(uapi, "Rows", None)
        cols_o = getattr(uapi, "Columns", None)
        old_rows = com_excel_scalar_int(getattr(rows_o, "Count", None), 0) if rows_o else 0
        old_cols = com_excel_scalar_int(getattr(cols_o, "Count", None), 0) if cols_o else 0
        if old_rows <= data_rows and old_cols <= data_cols:
            return
        # UsedRange を縮小させるため Clear（内容+書式）を優先。Ctrl+End がデータ領域外へ飛ぶ対策。
        def _clear(r: Any) -> None:
            try:
                clear_all = getattr(r, "clear", None) or getattr(r, "Clear", None)
                if callable(clear_all):
                    clear_all()
                    return
            except Exception:
                pass
            try:
                cf = getattr(r, "clear_contents", None) or getattr(r, "ClearContents", None)
                if callable(cf):
                    cf()
                else:
                    r.value = None
            except Exception:
                try:
                    r.value = None
                except Exception:
                    pass
        if old_cols > data_cols:
            rng = sheet_pointer.range(
                (1, data_cols + 1), (max(old_rows, data_rows), old_cols)
            )
            _clear(rng)
        if old_rows > data_rows:
            rng = sheet_pointer.range((data_rows + 1, 1), (old_rows, old_cols))
            _clear(rng)
    except Exception:
        pass


def clear_used_range_overflow_at(
    sheet_pointer: Any,
    top_row: int,
    top_col: int,
    data_row_count: int,
    data_col_count: int,
) -> None:
    """(top_row, top_col) を左上とする data_row_count 行 × data_col_count 列の矩形をデータ領域とみなし、
    その右隣（同じ行帯）および下側（anchor 列から UsedRange 右端まで）の余白をクリアする。
    """
    if data_row_count <= 0 or data_col_count <= 0 or top_row < 1 or top_col < 1:
        return
    bottom_data = top_row + data_row_count - 1
    right_data = top_col + data_col_count - 1
    try:
        ur = getattr(sheet_pointer, "used_range", None)
        if ur is None:
            return
        uapi = getattr(ur, "api", None)
        if uapi is None:
            return
        first_row = com_excel_scalar_int(getattr(uapi, "Row", None), 0)
        first_col = com_excel_scalar_int(getattr(uapi, "Column", None), 0)
        rows_o = getattr(uapi, "Rows", None)
        cols_o = getattr(uapi, "Columns", None)
        old_rows = com_excel_scalar_int(getattr(rows_o, "Count", None), 0) if rows_o else 0
        old_cols = com_excel_scalar_int(getattr(cols_o, "Count", None), 0) if cols_o else 0
        if old_rows <= 0 or old_cols <= 0:
            return
        last_row = first_row + old_rows - 1
        last_col = first_col + old_cols - 1

        def _clear(r: Any) -> None:
            try:
                clear_all = getattr(r, "clear", None) or getattr(r, "Clear", None)
                if callable(clear_all):
                    clear_all()
                    return
            except Exception:
                pass
            try:
                cf = getattr(r, "clear_contents", None) or getattr(r, "ClearContents", None)
                if callable(cf):
                    cf()
                else:
                    r.value = None
            except Exception:
                try:
                    r.value = None
                except Exception:
                    pass

        if last_col > right_data:
            rng = sheet_pointer.range(
                (top_row, right_data + 1), (bottom_data, last_col)
            )
            _clear(rng)
        if last_row > bottom_data:
            rng = sheet_pointer.range(
                (bottom_data + 1, top_col), (last_row, last_col)
            )
            _clear(rng)
    except Exception:
        pass


def clear_sheet_used_range(sheet_pointer: Any) -> None:
    """シートの UsedRange を Clear（内容・書式）する。クリア書込みの事前処理向け。"""
    try:
        ur = getattr(sheet_pointer, "used_range", None)
        if ur is None:
            return
        clear_all = getattr(ur, "clear", None) or getattr(ur, "Clear", None)
        if callable(clear_all):
            clear_all()
            return
        api = getattr(ur, "api", None)
        if api is not None:
            clr = getattr(api, "Clear", None)
            if callable(clr):
                clr()
    except Exception:
        pass


# ==============================================================================
# Performance helpers
# ==============================================================================
def yield_to_excel() -> None:
    """COM メッセージを捌き、Excel の無応答を避ける（ベストエフォート）。"""
    try:
        for _ in range(10):
            pythoncom.PumpWaitingMessages()
            time.sleep(0.05)
            pythoncom.PumpWaitingMessages()
    except Exception:
        pass


def write_chunk(
    sheet_pointer: Any,
    start_y: int,
    start_x: int,
    data_list: List[List[Any]],
    progress_ui: Any = None,
    *,
    chunk_rows: Optional[int] = None,
    progress_notify_rows: Optional[int] = None,
    progress_cb: Optional[Callable[[int], None]] = None,
    text_mode: bool = False,
) -> None:
    """Excel へデータを書き込む（チャンク分割）。

    互換性:
        旧来の write_chunk(sheet, y, x, data, progress_ui) 呼び出しはそのまま動作する。
        chunk_rows / progress_cb は進捗を滑らかにしたい場合のみ指定する。
        progress_notify_rows < chunk_rows のとき COM 分割を細かくし progress_cb の頻度を上げる。
        text_mode=True のとき書込前に @ 書式を設定し、各セル値を ' 付き文字列に変換する（CSV 読込等）。
    """
    total_rows = len(data_list)
    if total_rows == 0:
        return
    total_cols = len(data_list[0])

    step_unit = int(chunk_rows) if chunk_rows and int(chunk_rows) > 0 else 50000
    if progress_cb is not None and progress_notify_rows and int(progress_notify_rows) > 0:
        notify = max(500, int(progress_notify_rows))
        if notify < step_unit:
            step_unit = notify

    for i_offset in range(0, total_rows, step_unit):
        slice_h = min(step_unit, total_rows - i_offset)
        chunk = data_list[i_offset : i_offset + slice_h]
        if text_mode:
            from core.core_excel_text import matrix_as_excel_forced_text

            chunk = matrix_as_excel_forced_text(chunk)
        rng = sheet_pointer.range((start_y + i_offset, start_x)).resize(
            slice_h, total_cols
        )
        if text_mode:
            try:
                rng.number_format = "@"
            except Exception:
                pass
        rng.value = chunk

        if progress_cb is not None:
            try:
                progress_cb(i_offset + slice_h)
            except Exception:
                pass


def _autofit_range_best_effort(rng: Any) -> bool:
    """Range の列に AutoFit を試みる。成功時 True。"""
    try:
        cols = getattr(rng, "columns", None)
        if cols is not None:
            autofit_fn = getattr(cols, "autofit", None) or getattr(cols, "AutoFit", None)
            if callable(autofit_fn):
                autofit_fn()
                return True
        api = getattr(rng, "api", None)
        if api is not None:
            api_cols = getattr(api, "Columns", None)
            if api_cols is not None:
                api_autofit = getattr(api_cols, "AutoFit", None)
                if callable(api_autofit):
                    api_autofit()
                    return True
    except Exception:
        pass
    return False


def autofit_sheet_columns(
    sheet_pointer: Any,
    *,
    min_row: int = 1,
    min_col: int = 1,
    max_row: int,
    max_col: int,
    sheet_name_for_visible_fallback: str = "",
) -> None:
    """
    シート上の矩形 (min_row..max_row, min_col..max_col) の列幅をオートフィットする。
    行数が core_cst.AUTOFIT_MAX_ROWS を超えるときは 1 行目のみにフォールバックする（csv_ld と同趣旨）。
    """
    if max_row < min_row or max_col < min_col:
        return
    try:
        max_rows = int(getattr(cst, "AUTOFIT_MAX_ROWS", 100000) or 100000)
    except Exception:
        max_rows = 100000
    try:
        if max_row - min_row + 1 <= max_rows:
            rng = sheet_pointer.range((min_row, min_col), (max_row, max_col))
            if _autofit_range_best_effort(rng):
                return
        try:
            ur = getattr(sheet_pointer, "used_range", None)
            if ur is not None and _autofit_range_best_effort(ur):
                return
        except Exception:
            pass
        if max_row - min_row + 1 > max_rows:
            try:
                sheet_pointer.activate()
            except Exception:
                pass
            try:
                book = getattr(sheet_pointer, "book", None)
                app = getattr(book, "app", None) if book else None
                api = getattr(app, "api", None) if app else None
                aw = getattr(api, "ActiveWindow", None) if api else None
                if aw is not None:
                    vis = getattr(aw, "VisibleRange", None)
                    if vis is not None:
                        parent = getattr(vis, "Parent", None)
                        pname = getattr(parent, "Name", "") if parent else ""
                        if (
                            not sheet_name_for_visible_fallback
                            or pname == sheet_name_for_visible_fallback
                        ):
                            cols = getattr(vis, "Columns", None)
                            if cols is not None:
                                af = getattr(cols, "AutoFit", None) or getattr(
                                    cols, "Autofit", None
                                )
                                if callable(af):
                                    af()
                                    return
            except Exception:
                pass
        try:
            hdr = sheet_pointer.range((min_row, min_col), (min_row, max_col))
            _autofit_range_best_effort(hdr)
        except Exception:
            pass
    except Exception:
        try:
            logger.debug("[XLC] autofit_sheet_columns skipped", exc_info=True)
        except Exception:
            pass
