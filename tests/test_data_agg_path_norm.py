# -*- coding: utf-8 -*-
from __future__ import annotations

import sys
from pathlib import Path

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from svc.data_agg_path_norm import (  # noqa: E402
    normalize_source_path,
    normalize_source_path_literal,
)
from svc.svc_data_agg import (  # noqa: E402
    _batch_paths_rank_index,
    _row_file_path_matches_host,
)


def test_normalize_source_path_literal_skips_resolve(monkeypatch) -> None:
    calls: list[str] = []

    def _boom(self: Path) -> Path:
        calls.append(str(self))
        raise OSError("resolve blocked")

    monkeypatch.setattr(Path, "resolve", _boom, raising=False)
    unc = r"\\server\share\folder\book.xlsx"
    out = normalize_source_path_literal(unc)
    assert calls == []
    assert out == "//server/share/folder/book.xlsx"
    assert normalize_source_path(unc, resolve=False) == out


def test_normalize_source_path_resolve_local(tmp_path: Path) -> None:
    f = tmp_path / "A.xlsx"
    f.write_bytes(b"x")
    lit = normalize_source_path_literal(f)
    full = normalize_source_path(f)
    assert lit
    assert full
    assert lit == full.casefold() if __import__("os").name == "nt" else lit == full


def test_row_file_path_matches_via_norm_path() -> None:
    display = r"\\server\share\book.xlsx"
    norm = normalize_source_path_literal(display)
    row = {"__file_path": display, "__norm_path": norm}
    assert _row_file_path_matches_host(row, display)


def test_batch_paths_rank_uses_literal_norm() -> None:
    p = r"\\server\share\a.xlsx"
    rank = _batch_paths_rank_index([p])
    assert rank[p] == 0
    assert rank[normalize_source_path_literal(p)] == 0
