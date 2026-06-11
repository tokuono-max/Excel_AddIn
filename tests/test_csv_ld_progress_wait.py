# -*- coding: utf-8 -*-
"""進捗 DONE 後待機ヘルパのユニットテスト。"""
from __future__ import annotations

from unittest.mock import patch

from core import core_progress_wait as pw


def test_wait_after_progress_done_yields_until_deadline() -> None:
    calls: list[float] = []

    def _fake_yield() -> None:
        calls.append(1.0)

    with patch.object(pw, "xlc") as m_xlc, patch.object(pw.time, "monotonic", side_effect=[0.0, 0.0, 0.5, 1.1]):
        m_xlc.yield_to_excel = _fake_yield
        with patch.object(pw.time, "sleep"):
            pw.wait_after_progress_done(min_sec=1.0)
    assert len(calls) >= 1
