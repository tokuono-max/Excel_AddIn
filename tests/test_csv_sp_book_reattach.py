# -*- coding: utf-8 -*-
"""CSV 分割: 保存直前の Book/Sheet 再取得。"""
from __future__ import annotations

import sys
from pathlib import Path
from typing import Any

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from svc import svc_csv_sp as sp_mod  # noqa: E402


class _FakeSheet:
    def __init__(self) -> None:
        self.used_range = type(
            "UR",
            (),
            {
                "rows": type("R", (), {"count": 3})(),
                "columns": type("C", (), {"count": 2})(),
            },
        )()


def test_sp_refresh_sheet_dims_reads_header(monkeypatch) -> None:
    sheet = _FakeSheet()
    monkeypatch.setattr(sp_mod, "_sp_read_header_row", lambda _p, _n, **_: ["H1", "H2"])
    out = sp_mod._sp_refresh_sheet_dims(sheet, use_display_text=True)
    assert out is not None
    ncols, headers, nrows = out
    assert ncols == 2
    assert headers == ["H1", "H2"]
    assert nrows == 3


def test_resolve_book_and_sheet_accepts_ptr_s_kwarg(monkeypatch) -> None:
    """csv_sp が ptr_s= を渡しても TypeError にならない（csv_ld 経由）。"""
    import svc.svc_server as svc_server
    from svc import svc_csv_ld as ld_mod

    book = object()
    sheet = object()
    monkeypatch.setattr(
        svc_server,
        "resolve_fresh_book_after_ui_wait",
        lambda b, **_k: b,
    )
    monkeypatch.setattr(ld_mod.xlc, "find_sheet_by_guid", lambda _b, _sid: sheet)
    out_book, out_sh = ld_mod._resolve_book_and_sheet(
        book,
        "g1",
        100,
        attach_keys=(100, "", ""),
        ptr_s=object(),
    )
    assert out_book is book
    assert out_sh is sheet


def test_sp_refresh_sheet_dims_none_when_no_used_range() -> None:
    assert sp_mod._sp_refresh_sheet_dims(object(), use_display_text=False) is None
