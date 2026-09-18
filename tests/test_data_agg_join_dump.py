# -*- coding: utf-8 -*-
"""#20: 結合ダンプヘルパ（フラグ OFF 時は早期 return）。"""
from __future__ import annotations

from svc.data_agg_join_dump import (
    join_dump_col_filter_accepts,
    join_dump_ctx_prefix,
    join_dump_post_merge_file,
    join_dump_pv,
)


def test_join_dump_pv_truncates() -> None:
    assert join_dump_pv("abc") == "abc"
    assert join_dump_pv("x" * 200, max_len=10).endswith("…")
    assert len(join_dump_pv("x" * 200, max_len=10)) == 10


def test_join_dump_col_filter(monkeypatch) -> None:
    monkeypatch.delenv("DATA_AGG_JOIN_DUMP_COL", raising=False)
    monkeypatch.delenv("HC_DIAG_DATA_AGG_JOIN_COL", raising=False)
    assert join_dump_col_filter_accepts("任意")
    monkeypatch.setenv("DATA_AGG_JOIN_DUMP_COL", "PT")
    assert join_dump_col_filter_accepts("PT番号")
    assert not join_dump_col_filter_accepts("機器番号")


def test_join_dump_ctx_prefix() -> None:
    s = join_dump_ctx_prefix(
        {"scenario_id": "s1", "file_path": r"C:\a\b.xlsx", "caller": "batch"}
    )
    assert "scenario=s1" in s
    assert "file=b.xlsx" in s
    assert "caller=batch" in s


def test_join_dump_post_merge_noop_when_disabled(monkeypatch) -> None:
    monkeypatch.delenv("HC_DIAG_DATA_AGG_JOIN", raising=False)
    monkeypatch.delenv("DATA_AGG_JOIN_DUMP", raising=False)
    # 例外なく即 return
    join_dump_post_merge_file(
        [{"PT番号": "1"}],
        ["PT番号"],
        [
            {
                "sources": [
                    {
                        "type": "cell",
                        "ui_scenario_source_v1": {
                            "join_defs": [{"item": "MAC", "cell": "A1"}]
                        },
                    }
                ]
            }
        ],
        file_path=r"C:\x.xlsx",
        scenario_id="t",
        caller="test",
        preview_master=False,
    )
