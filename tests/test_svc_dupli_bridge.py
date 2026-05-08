# -*- coding: utf-8 -*-
"""svc_dupli bridge アドレス解析とログ用 extra の単体テスト。"""
from __future__ import annotations

from unittest.mock import MagicMock, patch

import pytest


def test_bridge_sheet_and_local_column() -> None:
    from svc.svc_dupli import _bridge_sheet_and_local_from_external

    sn, loc = _bridge_sheet_and_local_from_external("[Book1.xlsm]SMRT_01!$C:$C")
    assert sn == "SMRT_01"
    assert loc == "$C:$C"


def test_bridge_sheet_and_local_row() -> None:
    from svc.svc_dupli import _bridge_sheet_and_local_from_external

    sn, loc = _bridge_sheet_and_local_from_external("[b]S!$1:$1048576")
    assert sn == "S"
    assert loc == "$1:$1048576"


def test_full_sheet_from_bridge_count_large() -> None:
    from svc.svc_dupli import _full_sheet_from_bridge_count_large

    ok, why = _full_sheet_from_bridge_count_large(100, 100)
    assert ok is True
    assert "bridge_count_large_ok" in why

    ok2, _ = _full_sheet_from_bridge_count_large(50, 100)
    assert ok2 is False

    ok3, _ = _full_sheet_from_bridge_count_large(100, 0)
    assert ok3 is False

    ok4, _ = _full_sheet_from_bridge_count_large(None, 100)
    assert ok4 is False


def test_corner_full_sheet_flags_prefers_bridge_counts(monkeypatch: pytest.MonkeyPatch) -> None:
    from svc import svc_dupli

    calls: list[tuple] = []

    def fake_com(*_a: object, **_k: object) -> tuple[bool, str]:
        calls.append((True,))
        return False, "com_should_not_run"

    monkeypatch.setattr(svc_dupli, "_selection_full_sheet_flags", fake_com)
    ok, why = svc_dupli._corner_full_sheet_flags(
        MagicMock(), MagicMock(), 10, 10
    )
    assert ok is True
    assert "bridge_count_large_ok" in why
    assert len(calls) == 0


def test_bridge_sheet_and_local_no_bang() -> None:
    from svc.svc_dupli import _bridge_sheet_and_local_from_external

    sn, loc = _bridge_sheet_and_local_from_external("A1")
    assert sn is None
    assert loc == "A1"


def test_dupli_intersection_log_extra_bridge_has_no_selection_diag_keys() -> None:
    from svc.svc_dupli import _dupli_intersection_log_extra

    d = _dupli_intersection_log_extra(MagicMock(), MagicMock(), "bridge", [(1, 1, 2, 3)])
    assert d.get("rects_source") == "bridge"
    assert "full_sheet" not in d
    assert "full_sheet_why" not in d


@patch("svc.svc_dupli._dupli_selection_diag_snapshot")
def test_dupli_sel_log_merge_pops_full_sheet_keys(mock_snap: MagicMock) -> None:
    """COM 系 rects_source で snapshot に full_sheet が入っても、明示キーとマージで重複しない。"""
    from svc.svc_dupli import _dupli_intersection_log_extra

    mock_snap.return_value = {
        "tag": "selection_diag",
        "full_sheet": True,
        "full_sheet_why": "from_snapshot",
    }
    extra = dict(_dupli_intersection_log_extra(MagicMock(), MagicMock(), "com_fallback", [(1, 1, 1, 1)]))
    extra.pop("full_sheet", None)
    extra.pop("full_sheet_why", None)
    merged = {**extra, "full_sheet": False, "full_sheet_why": "from_flow"}
    assert merged["full_sheet"] is False
    assert merged["full_sheet_why"] == "from_flow"


@patch("svc.svc_dupli._dupli_selection_diag_snapshot")
def test_dupli_empty_mode_b_path_pops_only_full_sheet_why(mock_snap: MagicMock) -> None:
    mock_snap.return_value = {
        "tag": "selection_diag",
        "full_sheet": True,
        "full_sheet_why": "snap",
    }
    from svc.svc_dupli import _dupli_intersection_log_extra

    extra = dict(_dupli_intersection_log_extra(MagicMock(), MagicMock(), "com_fallback", None))
    extra.pop("full_sheet_why", None)
    merged = {**extra, "full_sheet_why": "flow"}
    assert merged["full_sheet"] is True
    assert merged["full_sheet_why"] == "flow"
