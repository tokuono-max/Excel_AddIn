# -*- coding: utf-8 -*-
"""一覧組立: チャンク一括変換の parity・上限・性能。"""
from __future__ import annotations

import sys
import time
from pathlib import Path

import pytest

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from svc.svc_data_agg import (  # noqa: E402
    _append_merged_rows_to_table_chunked,
    _batch_sparse_merged_row_noise,
    _merged_dict_rows_to_table_rows,
)


def _legacy_append(
    table_rows: list[list],
    merged_rows: list[dict],
    headers: list[str],
    *,
    max_table_rows: int | None,
    apply_sparse: bool,
) -> None:
    """旧行ループ相当（進捗なし）。"""
    rows_for_table: list[dict] = []
    for r in merged_rows:
        if max_table_rows is not None and max_table_rows > 0 and (
            len(table_rows) + len(rows_for_table) >= max_table_rows
        ):
            break
        if not isinstance(r, dict):
            continue
        if apply_sparse and _batch_sparse_merged_row_noise(r, headers):
            continue
        rows_for_table.append(r)
    if rows_for_table:
        table_rows.extend(_merged_dict_rows_to_table_rows(rows_for_table, headers))


def _sample_merged(n: int, headers: list[str]) -> list[dict]:
    out: list[dict] = []
    for i in range(n):
        row = {h: "" for h in headers}
        row[headers[0]] = "pk-%s" % i
        row[headers[1]] = "val-%s" % i
        out.append(row)
    return out


def test_chunked_matches_legacy_loop() -> None:
    headers = ["key", "val", "note"]
    merged = _sample_merged(120, headers)
    merged[3] = {h: "" for h in headers}
    merged[7] = {"key": "noise", "val": "", "note": "only2"}

    legacy_out: list[list] = []
    _legacy_append(legacy_out, merged, headers, max_table_rows=None, apply_sparse=True)

    chunked_out: list[list] = []
    _append_merged_rows_to_table_chunked(
        chunked_out,
        merged,
        headers,
        max_table_rows=None,
        row_skip=lambda r: _batch_sparse_merged_row_noise(r, headers),
        chunk_size=37,
    )
    assert chunked_out == legacy_out


def test_chunked_respects_max_table_rows() -> None:
    headers = ["a", "b"]
    merged = _sample_merged(50, headers)
    legacy_out: list[list] = []
    _legacy_append(legacy_out, merged, headers, max_table_rows=12, apply_sparse=False)

    chunked_out: list[list] = []
    hit = _append_merged_rows_to_table_chunked(
        chunked_out,
        merged,
        headers,
        max_table_rows=12,
        chunk_size=5,
    )
    assert hit is True
    assert chunked_out == legacy_out
    assert len(chunked_out) == 12


def test_chunked_iteration_contexts_parity() -> None:
    headers = ["k", "v"]
    merged = _sample_merged(8, headers)
    ctx_legacy: list[dict] = []
    rows_legacy: list[list] = []
    rows_for_table: list[dict] = []
    for gi, r in enumerate(merged):
        if not isinstance(r, dict):
            continue
        rows_for_table.append(r)
        ctx_legacy.append({"file_path": "f.csv", "iter_index": gi, "k": r.get("k")})
    rows_legacy.extend(_merged_dict_rows_to_table_rows(rows_for_table, headers))

    ctx_chunk: list[dict] = []
    rows_chunk: list[list] = []
    _append_merged_rows_to_table_chunked(
        rows_chunk,
        merged,
        headers,
        iteration_contexts_out=ctx_chunk,
        iteration_context_for_row=lambda r, gi: {
            "file_path": "f.csv",
            "iter_index": gi,
            "k": r.get("k"),
        },
        chunk_size=3,
    )
    assert rows_chunk == rows_legacy
    assert ctx_chunk == ctx_legacy


def test_chunked_large_under_time_budget() -> None:
    headers = ["c%d" % i for i in range(22)]
    merged = _sample_merged(8000, headers)
    out: list[list] = []
    t0 = time.perf_counter()
    _append_merged_rows_to_table_chunked(
        out,
        merged,
        headers,
        chunk_size=2000,
    )
    elapsed = time.perf_counter() - t0
    assert len(out) == 8000
    assert elapsed < 5.0
