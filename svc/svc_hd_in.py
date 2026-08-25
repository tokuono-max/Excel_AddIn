# -*- coding: utf-8 -*-
"""
Python: 3.10+
Module: svc/svc_hd_in
Created: 2026-03-05
Updated: 2026-08-20
Version: 1.3.0
Purpose:
  出荷履歴用定型ヘッダをシート先頭へ挿入する（新方式: svc_server + book 渡し）。
  項目名は JSON（config/hd_in.json、任意で {app}/出荷履歴項目.json）から読む。
  UI なし・core_xlc / core_stat / core_w32 使用。

History (latest 3):
  - 1.3.0 (2026-08-20) 項目名を JSON から読む。py 直書きの LABELS は廃止。
  - 1.2.0 (2026-05-01) 破壊的処理直前に Undo スナップショットを保存（元に戻す対応）。
  - 1.1.0 (2026-04-06) HC_LOG_PERF: [HD_IN_PERF]。診断: [HD_IN_TRACE]。
"""
from __future__ import annotations

import json
import os
import sys
import time
from pathlib import Path
from typing import Any, Optional

_path_here = os.path.abspath(os.path.dirname(__file__))
_path_root = os.path.dirname(_path_here)
if _path_root not in sys.path:
    sys.path.insert(0, _path_root)

from core.core_log import get_diag_logger, get_logger, get_perf_logger

logger = get_logger(__name__)
_hd_in_diag = get_diag_logger("hc_csv_tool.diag.hd_in")
_perf = get_perf_logger("svc.svc_hd_in.perf")
__version__ = "1.3.0"


def _elapsed_ms(since: float) -> int:
    return max(0, int((time.perf_counter() - since) * 1000))


def _hd_in_trace(phase: str, t0: float, **kv: object) -> None:
    try:
        if kv:
            _hd_in_diag.info(
                "[HD_IN_TRACE] phase=%s cumulative_ms=%d %s",
                phase,
                _elapsed_ms(t0),
                " ".join("%s=%s" % (k, v) for k, v in kv.items()),
            )
        else:
            _hd_in_diag.info(
                "[HD_IN_TRACE] phase=%s cumulative_ms=%d", phase, _elapsed_ms(t0)
            )
    except Exception:
        pass


def _perf_hd_in(phase: str, t0: float, **kv: object) -> None:
    try:
        if kv:
            _perf.info(
                "[HD_IN_PERF] phase=%s cumulative_ms=%d %s",
                phase,
                _elapsed_ms(t0),
                " ".join("%s=%s" % (k, v) for k, v in kv.items()),
            )
        else:
            _perf.info("[HD_IN_PERF] phase=%s cumulative_ms=%d", phase, _elapsed_ms(t0))
    except Exception:
        pass


try:
    from core import core_xlc as xlc
    from core import core_stat
    from core import core_w32 as w32
except ImportError:
    xlc = None  # type: ignore[assignment]
    core_stat = None  # type: ignore[assignment]
    w32 = None  # type: ignore[assignment]

_DEFAULT_CONFIG_FILE = "hd_in.json"
_DEFAULT_OVERRIDE_FILE = "出荷履歴項目.json"


class HdInConfigError(Exception):
    """出荷履歴項目の JSON が無い・壊れている・LABELS が空。"""


def _override_basename(name: object) -> str | None:
    """OVERRIDE_FILE はベース名のみ。パスや '..' は無効。"""
    if not isinstance(name, str):
        return None
    raw = name.strip()
    if not raw:
        return None
    fn = raw.replace("\\", "/").split("/")[-1]
    if not fn or fn != raw or ".." in fn:
        return None
    return fn


def _labels_from_obj(obj: Any) -> list[str]:
    if not isinstance(obj, dict):
        return []
    raw = obj.get("LABELS")
    if not isinstance(raw, list):
        return []
    out: list[str] = []
    for x in raw:
        s = str(x).strip() if x is not None else ""
        if s:
            out.append(s)
    return out


def _read_json_obj(path: Path) -> dict[str, Any]:
    try:
        text = path.read_text(encoding="utf-8-sig")
    except OSError as e:
        raise HdInConfigError(
            f"ERROR: 出荷履歴項目の設定ファイルを読めません。({path})"
        ) from e
    try:
        data = json.loads(text)
    except json.JSONDecodeError as e:
        raise HdInConfigError(
            f"ERROR: 出荷履歴項目の設定ファイルの形式が正しくありません。({path})"
        ) from e
    if not isinstance(data, dict):
        raise HdInConfigError(
            f"ERROR: 出荷履歴項目の設定ファイルの形式が正しくありません。({path})"
        )
    return data


def load_hd_in_labels(
    *,
    default_path: Path | None = None,
    override_dir: Path | None = None,
) -> list[str]:
    """
    出荷履歴ヘッダの LABELS を返す。

    使える LABELS を次の順で探す（片方だけ壊れていても、もう片方が使えれば進む）。

    1. ``{override_dir}/{OVERRIDE_FILE}`` に非空の LABELS があればそれを使う。
    2. 無ければ（ファイル無し・壊れている・LABELS 空）``default_path``
       （既定: ``config/hd_in.json``）の非空 LABELS。
    3. どちらも無い／壊れている／空なら ``HdInConfigError``（挿入しない）。

    ``OVERRIDE_FILE`` は既定 JSON の値（ファイル名のみ）。無ければ ``出荷履歴項目.json``。
    """
    if default_path is None:
        from core.core_cst import resolve_config_file_path

        default_path = resolve_config_file_path(_DEFAULT_CONFIG_FILE)
    if override_dir is None:
        from core import runtime_layout

        override_dir = runtime_layout.runtime_project_root(__file__)

    default_obj: dict[str, Any] | None = None
    default_err: HdInConfigError | None = None
    if default_path.is_file():
        try:
            default_obj = _read_json_obj(default_path)
        except HdInConfigError as e:
            default_err = e

    override_name = _DEFAULT_OVERRIDE_FILE
    if default_obj is not None:
        bn = _override_basename(default_obj.get("OVERRIDE_FILE"))
        if bn:
            override_name = bn

    op = override_dir / override_name
    if op.is_file():
        try:
            labels = _labels_from_obj(_read_json_obj(op))
            if labels:
                logger.info(
                    "[HD_IN] labels source=override path=%s n=%d", op, len(labels)
                )
                return labels
            logger.warning("[HD_IN] override LABELS empty path=%s; fallback default", op)
        except HdInConfigError as e:
            logger.warning(
                "[HD_IN] override unreadable path=%s err=%s; fallback default", op, e
            )

    if default_obj is not None:
        labels = _labels_from_obj(default_obj)
        if labels:
            logger.info(
                "[HD_IN] labels source=default path=%s n=%d", default_path, len(labels)
            )
            return labels
        raise HdInConfigError(
            f"ERROR: 出荷履歴項目の LABELS が空です。({default_path})"
        )
    if default_err is not None:
        raise default_err
    raise HdInConfigError(
        f"ERROR: 出荷履歴項目の設定ファイルが見つかりません。({default_path})"
    )


def _get_sheet(book: Any, sheet_id: str) -> Any:
    """sheet_id またはアクティブシートを返す。"""
    if sheet_id and xlc:
        sh = xlc.find_sheet_by_guid(book, sheet_id)
        if sh is not None:
            return sh
    try:
        return book.sheets.active
    except Exception:
        return None


def insert_header(
    book: Any,
    sheet_id: str = "",
    target_hwnd: Optional[int] = None,
    excel_hwnd: Optional[int] = None,
    **kwargs: Any,
) -> None:
    """
    出荷履歴用の定型ヘッダをシート先頭へ挿入する。
    svc_server から book, sheet_id, target_hwnd で呼ばれる。
    """
    t_flow = time.perf_counter()
    _perf_hd_in("enter", t_flow)
    _hd_in_trace("enter", t_flow)

    hwnd = int(target_hwnd or excel_hwnd or 0)
    if book is None:
        logger.warning("[HD_IN] 対象ブックなし")
        _perf_hd_in("abort_no_book", t_flow)
        _hd_in_trace("abort_no_book", t_flow)
        return
    ptr_s = _get_sheet(book, sheet_id)
    if ptr_s is None:
        logger.warning("[HD_IN] 対象シートなし sheet_id=%s", sheet_id)
        _perf_hd_in("abort_no_sheet", t_flow, sheet_id=sheet_id or "")
        _hd_in_trace("abort_no_sheet", t_flow, sheet_id=sheet_id or "")
        if core_stat:
            try:
                core_stat.set_status_info(
                    book.sheets.active, "ERROR: シートを特定できませんでした。"
                )
            except Exception:
                pass
        return

    ptr_a = getattr(book, "app", None)
    if ptr_a is None:
        logger.warning("[HD_IN] book.app なし")
        _perf_hd_in("abort_no_app", t_flow)
        _hd_in_trace("abort_no_app", t_flow)
        return

    logger.info("[HD_IN] 開始 sheet_id=%s", sheet_id or "")
    _perf_hd_in("after_resolve", t_flow, sheet_id=sheet_id or "")
    _hd_in_trace("after_resolve", t_flow, sheet_id=sheet_id or "")

    try:
        labels = load_hd_in_labels()
    except HdInConfigError as e:
        err_msg = str(e)
        logger.warning("[HD_IN] %s", err_msg)
        _perf_hd_in("abort_no_labels", t_flow)
        _hd_in_trace("abort_no_labels", t_flow)
        if core_stat:
            try:
                core_stat.set_status_info(ptr_s, err_msg)
            except Exception:
                pass
        return

    # 共通仕様: 破壊的処理の直前で Undo 用スナップショットを保存（元に戻すで復元可能にする）
    try:
        from svc.svc_undo import save_undo_snapshot

        save_undo_snapshot(book, sheet_id=sheet_id, target_hwnd=hwnd, excel_hwnd=hwnd)
        _perf_hd_in("after_undo_snapshot", t_flow)
        _hd_in_trace("after_undo_snapshot", t_flow)
    except Exception as e:
        logger.warning("[HD_IN] save_undo_snapshot failed (undo unavailable): %s", e)
        _perf_hd_in("after_undo_snapshot_failed", t_flow)
        _hd_in_trace("after_undo_snapshot_failed", t_flow)

    try:
        api = getattr(ptr_a, "api", None) or ptr_a
        api.Interactive = False
    except Exception:
        pass

    try:
        # 1行目のフォント退避
        back_font_name = None
        back_font_size = None
        try:
            r1 = ptr_s.range("1:1")
            api_r1 = getattr(r1, "api", None)
            if api_r1:
                back_font_name = getattr(api_r1.Font, "Name", None)
                back_font_size = getattr(api_r1.Font, "Size", None)
        except Exception:
            pass

        # 1行目に挿入
        ptr_target = ptr_s.range("1:1")
        api_target = getattr(ptr_target, "api", None)
        if api_target:
            api_target.Insert()

        # ヘッダ書き込み（列数は LABELS の件数に追随）
        ptr_s.range((1, 1)).value = [labels]

        # スタイル復元
        try:
            new_row_api = getattr(ptr_s.range("1:1"), "api", None)
            if new_row_api:
                if back_font_name is not None:
                    new_row_api.Font.Name = back_font_name
                if back_font_size is not None:
                    new_row_api.Font.Size = back_font_size
                new_row_api.Font.Bold = False
        except Exception:
            pass

        # 列幅オートフィット
        try:
            ur = getattr(ptr_s, "used_range", None)
            if ur is not None:
                cols = getattr(ur, "columns", None)
                if cols is not None:
                    af = getattr(cols, "autofit", None) or getattr(
                        cols, "AutoFit", None
                    )
                    if callable(af):
                        af()
        except Exception:
            pass

        msg = "出荷履歴用の定型ヘッダ項目をシート先頭へ物理挿入しました。"
        if core_stat:
            core_stat.set_status_info(ptr_s, msg)
        logger.info("[HD_IN] %s", msg)
        _perf_hd_in("after_insert_ok", t_flow)
        _hd_in_trace("after_insert_ok", t_flow)

    except Exception as ex_in:
        err_msg = f"ERROR: ヘッダ挿入不全 Detail: {ex_in}"
        if core_stat:
            try:
                core_stat.set_status_info(ptr_s, err_msg)
            except Exception:
                pass
        logger.exception("[HD_IN] 致命的エラー: %s", err_msg)
        _perf_hd_in("exception", t_flow)
        _hd_in_trace("exception", t_flow)

    finally:
        try:
            api = getattr(ptr_a, "api", None) or ptr_a
            api.Interactive = True
        except Exception:
            pass
        if hwnd and w32:
            try:
                w32.bring_to_front(hwnd)
            except Exception:
                pass
        _perf_hd_in("flow_end", t_flow)
        _hd_in_trace("flow_end", t_flow)


# hc_main が insert_shuka_header で呼ぶ場合の互換エイリアス（直接呼び出し時用）
insert_shuka_header = insert_header
