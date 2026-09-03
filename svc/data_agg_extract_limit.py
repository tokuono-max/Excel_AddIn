# -*- coding: utf-8 -*-
"""データ集約: 縦反復抽出の上限解決と打ち切り検知。"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any, Optional, Sequence

from core import core_env
from core.core_log import get_data_agg_diag_logger, get_logger

_logger = get_logger(__name__)
_agg_diag = get_data_agg_diag_logger()

_extract_truncation_buffer: list["ExtractTruncationRecord"] = []


class DataAggExtractTruncated(Exception):
    """読取上限に達し、未読データが残っている（方針 abort）。"""

    def __init__(self, records: list["ExtractTruncationRecord"]) -> None:
        self.records = list(records)
        super().__init__(format_extract_truncation_user_message(self.records))


@dataclass(frozen=True)
class ExtractTruncationRecord:
    file_path: str
    item_label: str
    limit: int
    read_count: int
    source_index: int = 0


def format_extract_truncation_user_message(
    records: Sequence[ExtractTruncationRecord] | list[ExtractTruncationRecord],
) -> str:
    """ユーザー向けの読取上限警告文（要点を短く）。"""
    recs = list(records or [])
    lines = [
        "主キーの取得件数が設定されている上限に達しました。 その先にも",
        "データがありそうです。",
        "シナリオの「取得件数」を確認し、必要に応じて修正してください。",
    ]
    if not recs:
        return "\n".join(lines)
    show = recs[:5]
    for r0 in show:
        lines.append("・ファイル: %s" % Path(r0.file_path).name)
        lines.append("・主キー項目: %s" % (r0.item_label or "-"))
        lines.append("・上限: %s 件（取込済み %s 件）" % (r0.limit, r0.read_count))
    extra = len(recs) - len(show)
    if extra > 0:
        lines.append("（他 %s 件の項目／ファイルでも同様）" % extra)
    return "\n".join(lines)


def extract_truncation_records_to_dicts(
    records: Sequence[ExtractTruncationRecord] | list[ExtractTruncationRecord],
) -> list[dict[str, Any]]:
    out: list[dict[str, Any]] = []
    for r0 in records or []:
        out.append(
            {
                "file": Path(r0.file_path).name,
                "file_path": str(r0.file_path),
                "item": str(r0.item_label or "").strip() or "-",
                "limit": int(r0.limit),
                "read": int(r0.read_count),
                "source_index": int(r0.source_index),
            }
        )
    return out


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
    repeat_until_last: bool = False,
) -> bool:
    """
    明示 repeat_max=1（空白まで／終端でない）のとき peek / 打ち切り記録を省略する。

    ODN-375 品名_ユニット / MAC RMT 等: 1 ファイル 1 行が意図。2 行目非空は正常。
    """
    if repeat_until_empty or repeat_until_last:
        return False
    if repeat_max is None:
        return False
    return int(repeat_max) == 1


def resolve_extract_repeat_limit(
    *,
    repeat_max: Optional[int],
    repeat_until_empty: bool,
    max_primary_rows: Optional[int] = None,
    repeat_until_last: bool = False,
) -> int:
    abs_max = core_env.data_agg_extract_absolute_max()
    if repeat_max is not None and int(repeat_max) > 0:
        limit = min(int(repeat_max), abs_max)
    elif repeat_until_empty or repeat_until_last:
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
    if preview_master_mode:
        by_item: dict[str, list[ExtractTruncationRecord]] = {}
        for rec in records:
            by_item.setdefault(rec.item_label, []).append(rec)
        for item_label, recs in by_item.items():
            try:
                _logger.warning(
                    "[DATA_AGG] extract_truncated scenario=%s caller=%s item=%s files=%s limit=%s",
                    scenario_id or "-",
                    probe_caller or "-",
                    item_label,
                    len(recs),
                    recs[0].limit,
                )
            except Exception:
                pass
    else:
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


def is_extract_truncated_batch_notify(payload: dict[str, Any] | None) -> bool:
    """一括完了通知が読取上限打ち切り（継続選択対象）かどうか。"""
    if not isinstance(payload, dict):
        return False
    if payload.get("ok", True):
        return False
    err = str(payload.get("error") or "").strip().lower()
    phase = str(payload.get("abort_phase") or "").strip().lower()
    return err == "extract_truncated" or phase == "extract_truncated"
