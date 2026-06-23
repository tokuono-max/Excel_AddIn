# -*- coding: utf-8 -*-
"""データ集約: 縦反復抽出の上限解決と打ち切り検知。"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any, Optional

from core import core_env
from core.core_log import get_data_agg_diag_logger, get_logger

_logger = get_logger(__name__)
_agg_diag = get_data_agg_diag_logger()

_extract_truncation_buffer: list["ExtractTruncationRecord"] = []


class DataAggExtractTruncated(Exception):
    """読取上限に達し、未読データが残っている（方針 abort）。"""

    def __init__(self, records: list["ExtractTruncationRecord"]) -> None:
        self.records = list(records)
        super().__init__(self._format_message())

    def _format_message(self) -> str:
        if not self.records:
            return "extract truncated"
        r0 = self.records[0]
        extra = len(self.records) - 1
        msg = (
            "%s の「%s」: 読取上限 %s 件に達しました（未読データの可能性）"
            % (Path(r0.file_path).name, r0.item_label, r0.limit)
        )
        if extra > 0:
            msg += "（他 %s 件）" % extra
        return msg


@dataclass(frozen=True)
class ExtractTruncationRecord:
    file_path: str
    item_label: str
    limit: int
    read_count: int
    source_index: int = 0


def clear_extract_truncation_records() -> None:
    _extract_truncation_buffer.clear()


def take_extract_truncation_records() -> list[ExtractTruncationRecord]:
    out = list(_extract_truncation_buffer)
    _extract_truncation_buffer.clear()
    return out


def skip_extract_truncation_peek(
    *,
    repeat_max: Optional[int],
    repeat_until_empty: bool,
) -> bool:
    """
    明示 repeat_max=1（空白まででない）のとき peek / 打ち切り記録を省略する。

    ODN-375 品名_ユニット / MAC RMT 等: 1 ファイル 1 行が意図。2 行目非空は正常。
    """
    if repeat_until_empty:
        return False
    if repeat_max is None:
        return False
    return int(repeat_max) == 1


def resolve_extract_repeat_limit(
    *,
    repeat_max: Optional[int],
    repeat_until_empty: bool,
    max_primary_rows: Optional[int] = None,
) -> int:
    abs_max = core_env.data_agg_extract_absolute_max()
    if repeat_max is not None and int(repeat_max) > 0:
        limit = min(int(repeat_max), abs_max)
    elif repeat_until_empty:
        limit = abs_max
    else:
        limit = core_env.data_agg_extract_default_max()
    if max_primary_rows is not None and int(max_primary_rows) > 0:
        limit = min(limit, int(max_primary_rows))
    return max(1, limit)


def cell_value_nonempty(val: Any) -> bool:
    if val is None:
        return False
    if isinstance(val, str):
        return str(val).strip() != ""
    return True


def record_extract_truncation_if_needed(
    vals: list[Any],
    *,
    limit: int,
    peek_next: Any,
    file_path: str | Path,
    item_label: str,
    source_index: int = 0,
) -> None:
    if len(vals) < limit:
        return
    if not cell_value_nonempty(peek_next):
        return
    rec = ExtractTruncationRecord(
        file_path=str(file_path),
        item_label=str(item_label or "").strip() or "-",
        limit=int(limit),
        read_count=len(vals),
        source_index=int(source_index),
    )
    _extract_truncation_buffer.append(rec)
    try:
        _agg_diag.warning(
            "[DATA_AGG_WARN] extract_truncated file=%s item=%s limit=%s read=%s src_idx=%s",
            Path(rec.file_path).name,
            rec.item_label,
            rec.limit,
            rec.read_count,
            rec.source_index,
        )
    except Exception:
        pass


def enforce_extract_truncation_policy(
    records: list[ExtractTruncationRecord],
    *,
    scenario_id: str = "",
    probe_caller: str = "",
    preview_master_mode: bool = False,
) -> None:
    if not records:
        return
    for rec in records:
        try:
            _logger.warning(
                "[DATA_AGG] extract_truncated scenario=%s caller=%s file=%s item=%s limit=%s",
                scenario_id or "-",
                probe_caller or "-",
                Path(rec.file_path).name,
                rec.item_label,
                rec.limit,
            )
        except Exception:
            pass
    policy = core_env.data_agg_extract_trunc_policy(
        probe_caller=probe_caller or None,
        preview_master=preview_master_mode,
    )
    if policy == "warn":
        return
    raise DataAggExtractTruncated(records)
