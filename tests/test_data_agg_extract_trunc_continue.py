# -*- coding: utf-8 -*-
"""読取上限打ち切り時の継続選択判定テスト。"""
from __future__ import annotations

import sys
from pathlib import Path

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from svc.data_agg_extract_limit import is_extract_truncated_batch_notify  # noqa: E402


def test_is_extract_truncated_by_error() -> None:
    assert is_extract_truncated_batch_notify(
        {"ok": False, "error": "extract_truncated", "message": "上限"}
    )


def test_is_extract_truncated_by_abort_phase() -> None:
    assert is_extract_truncated_batch_notify(
        {"ok": False, "abort_phase": "extract_truncated", "message": "上限"}
    )


def test_is_extract_truncated_ignores_ok() -> None:
    assert not is_extract_truncated_batch_notify(
        {"ok": True, "error": "extract_truncated"}
    )


def test_is_extract_truncated_ignores_other_errors() -> None:
    assert not is_extract_truncated_batch_notify(
        {"ok": False, "error": "cancelled", "abort_phase": "compute"}
    )


def test_is_extract_truncated_none_or_empty() -> None:
    assert not is_extract_truncated_batch_notify(None)
    assert not is_extract_truncated_batch_notify({})
