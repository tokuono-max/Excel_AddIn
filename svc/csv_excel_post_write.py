# -*- coding: utf-8 -*-
"""CSV読込・CSV結合: Excel展開後の AutoFit / AutoFilter（ui_*.json の AUTOFIT_MAX_ROWS / AUTOFILTER）。

AUTOFILTER 適用成功時は 1 行目（ヘッダ）固定も連動して実施する。
"""
from __future__ import annotations

from typing import Any

from core.core_log import get_logger

logger = get_logger(__name__)

CSV_LD_FEATURE_KEY = "csv_ld"
CSV_MG_FEATURE_KEY = "csv_mg"


def parse_autofit_max_rows(cfg: dict[str, Any] | None) -> int:
    """AUTOFIT_MAX_ROWS を解釈する。0 または省略は「行数によるスキップなし」。"""
    if not isinstance(cfg, dict):
        return 0
    try:
        return max(0, int(cfg.get("AUTOFIT_MAX_ROWS") or 0))
    except (TypeError, ValueError):
        return 0


def parse_autofilter(cfg: dict[str, Any] | None) -> bool:
    """AUTOFILTER を解釈する。省略時は false。"""
    if not isinstance(cfg, dict):
        return False
    return bool(cfg.get("AUTOFILTER", False))


def load_csv_ui_excel_opts(feature_key: str) -> tuple[int, bool]:
    """config/ui_<feature>.json から (AUTOFIT_MAX_ROWS, AUTOFILTER) を読む。"""
    from core import core_cst as cst  # noqa: WPS433

    cfg = cst.get_ui_config_from_file_required(feature_key)
    return parse_autofit_max_rows(cfg), parse_autofilter(cfg)


def should_apply_csv_autofit(output_rows: int, autofit_max_rows: int) -> bool:
    """出力行数が limit 未満なら AutoFit。limit<=0 なら常に実施（行数>0 のとき）。"""
    rows = max(0, int(output_rows))
    if rows <= 0:
        return False
    limit = int(autofit_max_rows or 0)
    if limit <= 0:
        return True
    return rows < limit


def apply_csv_autofit_sheet(
    sheet: Any,
    *,
    min_row: int = 1,
    max_row: int,
    max_col: int,
    min_col: int = 1,
) -> None:
    """指定矩形の列幅をオートフィット（best-effort）。"""
    if max_row < min_row or max_col < min_col:
        return
    try:
        from core import core_xlc as xlc  # noqa: WPS433

        xlc.autofit_sheet_columns(
            sheet,
            min_row=min_row,
            min_col=min_col,
            max_row=max_row,
            max_col=max_col,
        )
    except Exception:
        pass


def _sheet_name_best_effort(sheet: Any) -> str:
    try:
        return str(getattr(sheet, "name", "") or "") or "-"
    except Exception:
        return "-"


def _log_csv_post_write_freeze(
    sheet: Any,
    *,
    rows: int,
    cols: int,
    ok_fr: bool,
    reason: str = "",
) -> None:
    """freeze 結果をログ出力（COM 例外残留で logger が落ちないよう PyErr_Clear）。"""
    from svc.svc_data_agg_write import _clear_native_exception_state  # noqa: WPS433

    _clear_native_exception_state()
    sheet_name = _sheet_name_best_effort(sheet)
    if ok_fr:
        logger.info(
            "[CSV_POST_WRITE] ヘッダ行固定 sheet=%s rows=%s cols=%s",
            sheet_name,
            rows,
            cols,
        )
    elif reason:
        logger.warning(
            "[CSV_POST_WRITE] ヘッダ行固定未適用 sheet=%s rows=%s cols=%s reason=%s",
            sheet_name,
            rows,
            cols,
            reason,
        )
    else:
        logger.warning(
            "[CSV_POST_WRITE] ヘッダ行固定未適用 sheet=%s rows=%s cols=%s",
            sheet_name,
            rows,
            cols,
        )
    _clear_native_exception_state()


def apply_csv_autofilter_ld(
    sheet: Any,
    *,
    last_row: int,
    max_col: int,
) -> bool:
    """CSV読込: 1行目ヘッダにオートフィルタ。成功時は 1 行目固定も連動実施。"""
    rows = max(0, int(last_row))
    cols = max(0, int(max_col))
    if rows <= 0 or cols <= 0:
        return False
    try:
        from svc.svc_data_agg_write import (  # noqa: WPS433
            _clear_native_exception_state,
            apply_autofilter_to_block,
            freeze_sheet_below_header_row,
        )

        _clear_native_exception_state()
        ok = apply_autofilter_to_block(
            sheet,
            top_row=1,
            left_col=1,
            n_rows=rows,
            n_cols=cols,
        )
        _clear_native_exception_state()
        if ok:
            ok_fr = False
            freeze_reason = ""
            try:
                ok_fr = bool(
                    freeze_sheet_below_header_row(sheet, 1, left_col=1)
                )
            except BaseException as ex:
                _clear_native_exception_state()
                ok_fr = False
                freeze_reason = repr(ex)
            _log_csv_post_write_freeze(
                sheet,
                rows=rows,
                cols=cols,
                ok_fr=ok_fr,
                reason=freeze_reason,
            )
        return ok
    except Exception as ex:
        from svc.svc_data_agg_write import _clear_native_exception_state  # noqa: WPS433

        _clear_native_exception_state()
        logger.warning(
            "[CSV_POST_WRITE] autofilter/freeze 例外 sheet=%s rows=%s cols=%s: %s",
            _sheet_name_best_effort(sheet),
            rows,
            cols,
            ex,
        )
        _clear_native_exception_state()
        return False


def should_apply_csv_mg_autofilter(
    *,
    enabled: bool,
    start_row: int,
    mode: str,
) -> bool:
    """結合: 空シート先頭書込・mode_append のみオートフィルタ対象。"""
    if not enabled:
        return False
    if int(start_row) != 1:
        return False
    return str(mode or "").strip() == "mode_append"


def apply_csv_autofilter_mg(
    sheet: Any,
    *,
    last_row: int,
    max_col: int,
    enabled: bool,
    start_row: int,
    mode: str,
) -> bool:
    """ファイル結合: 条件を満たすとき 1 行目ヘッダにオートフィルタ（成功時は 1 行目固定も連動）。"""
    if not should_apply_csv_mg_autofilter(
        enabled=enabled, start_row=start_row, mode=mode
    ):
        return False
    return apply_csv_autofilter_ld(sheet, last_row=last_row, max_col=max_col)


def csv_post_write_step_phase_label(
    step: str,
    *,
    phase_prefix: str = "3/4",
    sheet_part: str = "",
) -> str:
    """進捗工程3用ラベル。step は autofit_run / autofit_skip / autofilter_run / autofilter_skip。"""
    labels = {
        "autofit_run": "AutoFit 実行中",
        "autofit_skip": "AutoFit 省略",
        "autofilter_run": "AutoFilter 実行中",
        "autofilter_skip": "AutoFilter 省略",
    }
    base = labels.get(str(step or "").strip(), str(step or "").strip() or "仕上げ中")
    if sheet_part:
        return f"{phase_prefix} {base} ({sheet_part})"
    return f"{phase_prefix} {base}"


def post_write_csv_ld_sheet(
    sheet: Any,
    *,
    last_row: int,
    max_col: int,
    autofit_max_rows: int,
    autofilter: bool,
) -> None:
    """CSV読込: 1シート分の展開後処理（AutoFit → AutoFilter）。"""
    rows = max(0, int(last_row))
    cols = max(0, int(max_col))
    if cols <= 0 or rows <= 0:
        return
    if should_apply_csv_autofit(rows, autofit_max_rows):
        apply_csv_autofit_sheet(sheet, min_row=1, max_row=rows, max_col=cols)
    if autofilter:
        apply_csv_autofilter_ld(sheet, last_row=rows, max_col=cols)
