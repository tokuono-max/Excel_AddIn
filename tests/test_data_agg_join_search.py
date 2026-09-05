# -*- coding: utf-8 -*-
from __future__ import annotations

import sys
from pathlib import Path
from typing import Any

import pytest

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from svc.svc_data_agg import (  # noqa: E402
    _apply_join_key_search_write,
    _batch_file_extract_and_merge,
    _build_cross_file_join_search_plan,
    _build_join_search_index,
    _cross_join_should_use_emit_driven,
    _join_search_pool_scope,
    _join_search_rows_for_slice_with_host_supplement,
    _row_satisfies_join_and,
)
from core.core_join_compare import join_compare_display_key  # noqa: E402


def test_row_satisfies_join_and() -> None:
    jds = [{"item": "A"}, {"item": "B"}]
    row = {"A": "1", "B": " 2 "}
    assert _row_satisfies_join_and(row, jds, {"A": "1", "B": "2"})
    assert not _row_satisfies_join_and(row, jds, {"A": "x", "B": "2"})


def test_row_satisfies_join_and_apostrophe_mismatch() -> None:
    """Excel 文字列固定 ' と本体のみが一致する。"""
    jds = [{"item": "出荷日"}]
    row = {"出荷日": "'20220527"}
    assert _row_satisfies_join_and(row, jds, {"出荷日": "20220527"})
    assert _row_satisfies_join_and(row, jds, {"出荷日": "'20220527"})
    assert not _row_satisfies_join_and(row, jds, {"出荷日": "20220601"})


def test_join_defs_checks_trim_enables_match() -> None:
    """join_defs の加工（トリム）が抽出後処理に効き、照合に使われる。"""
    from svc.data_agg_value_post import postprocess_link_rule_value  # noqa: WPS433

    jd = {"item": "K", "checks": ["トリム"]}
    extracted = postprocess_link_rule_value("  x  ", jd)
    row = {"K": "'x"}
    assert _row_satisfies_join_and(row, [jd], {"K": extracted})


def test_join_defs_without_checks_unchanged_match() -> None:
    """checks / value_shape_script 無し join_defs は従来どおり（後方互換）。"""
    jds = [{"item": "K", "cell": "A1"}]
    row = {"K": "'02301"}
    assert _row_satisfies_join_and(row, jds, {"K": "02301"})
    assert _row_satisfies_join_and(row, jds, {"K": "'02301"})


def test_validate_scenario_join_defs_value_shape_script() -> None:
    from svc.svc_data_agg_scenario import validate_scenario  # noqa: WPS433

    ok = {
        "items": [
            {
                "id": "item_0",
                "name": "出荷番号",
                "sources": [
                    {
                        "type": "cell",
                        "cell_ref": "E2",
                        "row_offset": 1,
                        "repeat_until_empty": True,
                        "ui_scenario_source_v1": {
                            "file_pattern": "x",
                            "join_defs": [
                                {
                                    "cell": "A2",
                                    "row": 0,
                                    "col": 0,
                                    "item": "機器番号",
                                    "value_shape_script": "wide",
                                }
                            ],
                        },
                    }
                ],
            }
        ],
    }
    bad = {
        "items": [
            {
                "id": "item_0",
                "name": "出荷番号",
                "sources": [
                    {
                        "type": "cell",
                        "cell_ref": "E2",
                        "row_offset": 1,
                        "repeat_until_empty": True,
                        "ui_scenario_source_v1": {
                            "file_pattern": "x",
                            "join_defs": [
                                {
                                    "cell": "A2",
                                    "row": 0,
                                    "col": 0,
                                    "item": "機器番号",
                                    "value_shape_script": "bad_cmd(",
                                }
                            ],
                        },
                    }
                ],
            }
        ],
    }
    assert not validate_scenario(ok)
    errs = validate_scenario(bad)
    assert any("join_defs" in e and "value_shape_script" in e for e in errs)


def test_batch_extract_and_merge_reads_full_rows_for_master_preview_join_file(
    monkeypatch, tmp_path: Path
) -> None:
    p = tmp_path / "join_target.xlsx"
    p.write_text("", encoding="utf-8")
    item = {
        "id": "i1",
        "name": "HostCol",
        "sources": [
            {
                "type": "cell",
                "sheet_name": "Sheet",
                "cell_ref": "A1",
                "ui_scenario_source_v1": {"file_pattern": "join_target"},
            }
        ],
    }
    seen: list[int | None] = []

    def _fake_extract_item_bundle(*args, **kwargs):
        seen.append(kwargs.get("max_primary_rows"))
        return {"primary_values": ["HOSTVAL"]}

    from svc import svc_data_agg_extract as extract_mod  # noqa: WPS433

    monkeypatch.setattr(extract_mod, "extract_item_bundle", _fake_extract_item_bundle)

    res = _batch_file_extract_and_merge(
        p,
        items=[item],
        headers=["HostCol"],
        header_set={"HostCol"},
        column_modes=["fill_in"],
        linked_targets=set(),
        join_targets=set(),
        path_col="",
        master_preview_cap_idx=None,
        master_preview_extract_allow=None,
        master_preview_join_full_read_patterns=("join_target",),
        preview_master_mode=True,
        use_join_search_merge=False,
        max_primary_rows=500,
    )
    from svc.data_agg_master_preview_perf import master_preview_scan_row_cap  # noqa: WPS433

    assert seen == [master_preview_scan_row_cap()]
    assert join_compare_display_key((res.bundles[0].get("primary_values") or [])[0]) == "HOSTVAL"


def test_apply_join_key_search_1_primary_n_join_slices_union() -> None:
    """1主値・結合値Nスライス: 各スライスで一致行を集め、和集合に同一主値を書く。"""
    pool = [
        {"KeyCol": "a", "Out": None, "__file_path": "f", "__iter_index": 0},
        {"KeyCol": "b", "Out": None, "__file_path": "f", "__iter_index": 0},
    ]
    item = {
        "name": "Out",
        "sources": [
            {
                "type": "cell",
                "ui_scenario_source_v1": {
                    "join_defs": [{"item": "KeyCol", "cell": "X1"}],
                },
            }
        ],
    }
    bundle = {
        "primary_values": ["MAIN"],
        "join_values": {"KeyCol": ["a", "b"]},
    }
    _apply_join_key_search_write(pool, item, "Out", bundle, "overwrite")
    assert pool[0]["Out"] == "MAIN"
    assert pool[1]["Out"] == "MAIN"


def test_apply_join_key_search_ignore_n_to_1() -> None:
    """主複数・結合1スライスは無視。"""
    pool = [{"KeyCol": "a", "Out": None}]
    item = {
        "name": "Out",
        "sources": [
            {
                "type": "cell",
                "ui_scenario_source_v1": {
                    "join_defs": [{"item": "KeyCol", "cell": "X1"}],
                },
            }
        ],
    }
    bundle = {
        "primary_values": ["P1", "P2"],
        "join_values": {"KeyCol": ["a"]},
    }
    _apply_join_key_search_write(pool, item, "Out", bundle, "overwrite")
    assert pool[0]["Out"] is None


def test_apply_join_key_no_match_leaves_out_empty() -> None:
    """結合スライスに一致する行が無いとき、Out 列へ主値を入れない。"""
    pool = [
        {"KeyCol": "x", "Out": None, "__file_path": "f", "__iter_index": 0},
        {"KeyCol": "y", "Out": None, "__file_path": "f", "__iter_index": 0},
    ]
    item = {
        "name": "Out",
        "sources": [
            {
                "type": "cell",
                "ui_scenario_source_v1": {
                    "join_defs": [{"item": "KeyCol", "cell": "X1"}],
                },
            }
        ],
    }
    bundle = {"primary_values": ["MAIN"], "join_values": {"KeyCol": ["a"]}}
    _apply_join_key_search_write(pool, item, "Out", bundle, "overwrite")
    assert pool[0]["Out"] is None
    assert pool[1]["Out"] is None


def test_row_satisfies_join_float_matches_string_int() -> None:
    jds = [{"item": "K"}]
    row = {"K": 1.0}
    assert _row_satisfies_join_and(row, jds, {"K": "1"})


def test_apply_join_key_paired_min_length() -> None:
    pool = [
        {"KeyCol": "a", "Out": None, "__iter_index": 0},
        {"KeyCol": "b", "Out": None, "__iter_index": 1},
    ]
    item = {
        "name": "Out",
        "sources": [
            {
                "type": "cell",
                "ui_scenario_source_v1": {
                    "join_defs": [{"item": "KeyCol", "cell": "X1"}],
                },
            }
        ],
    }
    bundle = {
        "primary_values": ["P0", "P1"],
        "join_values": {"KeyCol": ["a", "b"]},
    }
    _apply_join_key_search_write(pool, item, "Out", bundle, "overwrite")
    assert pool[0]["Out"] == "P0"
    assert pool[1]["Out"] == "P1"


def test_join_search_pool_scope_cross_file_limits_to_side_and_host() -> None:
    """cross_file 時は比較列 side の pattern 行と当該ホストファイル行のみ。"""
    pool = [
        {"MAC": "a", "__file_path": r"C:\光特性履歴_1.xlsx"},
        {"MAC": "b", "__file_path": r"C:\光特性履歴_2.xlsx"},
        {"MAC": "c", "__file_path": r"C:\紐づけ履歴_1.xlsx"},
        {"MAC": "d", "__file_path": r"C:\紐づけ履歴_2.xlsx"},
        {"MAC": "e", "__file_path": r"C:\other.xlsx"},
    ]
    host_item = {
        "name": "PT番号",
        "sources": [
            {
                "type": "cell",
                "ui_scenario_source_v1": {
                    "file_pattern": "紐づけ",
                    "join_defs": [{"item": "MACアドレス", "cell": "P5"}],
                },
            }
        ],
    }
    items = [
        {
            "name": "機器番号",
            "sources": [
                {
                    "type": "cell",
                    "ui_scenario_source_v1": {
                        "file_pattern": "光特性",
                        "link_defs": [{"item": "MACアドレス", "cell": "M7"}],
                    },
                }
            ],
        },
        host_item,
        {"name": "MACアドレス", "sources": []},
    ]
    headers = ["機器番号", "PT番号", "MACアドレス"]
    scoped = _join_search_pool_scope(
        pool,
        r"C:\紐づけ履歴_1.xlsx",
        True,
        host_item=host_item,
        items=items,
        headers=headers,
    )
    paths = {r["__file_path"] for r in scoped}
    assert paths == {
        r"C:\光特性履歴_1.xlsx",
        r"C:\光特性履歴_2.xlsx",
        r"C:\紐づけ履歴_1.xlsx",
    }


def test_join_search_pool_scope_same_file_unchanged() -> None:
    pool = [
        {"K": "a", "__file_path": "f1"},
        {"K": "b", "__file_path": "f2"},
    ]
    scoped = _join_search_pool_scope(pool, "f1", False)
    assert len(scoped) == 1
    assert scoped[0]["__file_path"] == "f1"


def test_join_search_host_supplement_finds_host_only_match() -> None:
    """side 索引 + host 線形スキャンで host 行の一致も拾う。"""
    side_rows = [
        {"MAC": "a", "Out": None, "__file_path": r"C:\光特性_1.xlsx"},
    ]
    host_rows = [
        {"MAC": "b", "Out": None, "__file_path": r"C:\紐づけ_1.xlsx"},
    ]
    join_defs = [{"item": "MAC", "cell": "P1"}]
    side_index = (
        ["MAC"],
        {("a",): [side_rows[0]]},
    )
    jv = {"MAC": ["b"]}
    matched = _join_search_rows_for_slice_with_host_supplement(
        side_index, host_rows, join_defs, jv, ["MAC"], 0
    )
    assert len(matched) == 1
    assert matched[0]["__file_path"] == r"C:\紐づけ_1.xlsx"


def test_cross_file_join_plan_builds_emit_row_index_once() -> None:
    """横断結合索引は file_pattern 側ではなく Excel 出力対象行のみ（紐づけ専用行は除外）。"""
    global_pool = [
        {"MAC": "x", "PT": None, "__file_path": r"C:\光特性_a.xlsx"},
        {"MAC": "y", "PT": None, "__file_path": r"C:\光特性_b.xlsx"},
        {"MAC": "z", "PT": None, "__file_path": r"C:\紐づけ_a.xlsx"},
    ]
    host_item = {
        "name": "PT",
        "sources": [
            {
                "type": "cell",
                "ui_scenario_source_v1": {
                    "file_pattern": "紐づけ",
                    "join_defs": [{"item": "MAC", "cell": "P1"}],
                },
            }
        ],
    }
    items = [
        {
            "name": "Dev",
            "sources": [
                {
                    "type": "cell",
                    "ui_scenario_source_v1": {
                        "file_pattern": "光特性",
                        "link_defs": [{"item": "MAC", "cell": "M1"}],
                    },
                }
            ],
        },
        host_item,
    ]
    plan = _build_cross_file_join_search_plan(global_pool, host_item, items, ["Dev", "PT"])
    assert len(plan.side_rows) == 2
    assert all("光特性" in str(r.get("__file_path") or "") for r in plan.side_rows)
    assert plan.side_index[0] == ["MAC"]
    assert ("x",) in plan.side_index[1]


def test_batch_extract_skips_items_without_matching_file_pattern(tmp_path: Path) -> None:
    """file_pattern が当該入力ファイルに合わない項目は extract をスキップする。"""
    from openpyxl import Workbook

    from svc.svc_data_agg import _batch_file_extract_and_merge

    p = tmp_path / "only_host.xlsx"
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws["A1"] = "HOSTVAL"
    wb.save(p)

    item_match = {
        "id": "i1",
        "name": "HostCol",
        "sources": [
            {
                "type": "cell",
                "sheet_name": "Sheet",
                "cell_ref": "A1",
                "ui_scenario_source_v1": {"file_pattern": "only_host"},
            }
        ],
    }
    item_skip = {
        "id": "i2",
        "name": "SideCol",
        "sources": [
            {
                "type": "cell",
                "sheet_name": "Sheet",
                "cell_ref": "A1",
                "ui_scenario_source_v1": {"file_pattern": "other_file"},
            }
        ],
    }
    headers = ["HostCol", "SideCol"]
    res = _batch_file_extract_and_merge(
        p,
        items=[item_match, item_skip],
        headers=headers,
        header_set=set(headers),
        column_modes=["fill_in", "fill_in"],
        linked_targets=set(),
        join_targets=set(),
        path_col="",
        master_preview_cap_idx=None,
        master_preview_extract_allow=None,
        preview_master_mode=False,
        use_join_search_merge=True,
        max_primary_rows=None,
    )
    assert join_compare_display_key((res.bundles[0].get("primary_values") or [])[0]) == "HOSTVAL"
    assert res.bundles[1] == {}


def test_cross_join_emit_driven_gate_auto_threshold(monkeypatch: Any) -> None:
    idx = _build_join_search_index(
        [{"K": "a", "Out": None}],
        [{"item": "K"}],
    )
    monkeypatch.setenv("HC_DATA_AGG_JOIN_EMIT_DRIVEN", "auto")
    monkeypatch.setenv("HC_DATA_AGG_JOIN_EMIT_DRIVEN_MIN_SLICES", "2000")
    assert not _cross_join_should_use_emit_driven(
        cross_file=True,
        stacked_join=False,
        n_join=232,
        search_pool_len=15710,
        join_index=idx,
    )
    # ホスト ≫ emit
    assert _cross_join_should_use_emit_driven(
        cross_file=True,
        stacked_join=False,
        n_join=47076,
        search_pool_len=15710,
        join_index=idx,
    )
    # ホスト ＜ emit でも閾値以上なら発火（PT / ダミーQR 年次）
    assert _cross_join_should_use_emit_driven(
        cross_file=True,
        stacked_join=False,
        n_join=4958,
        search_pool_len=15710,
        join_index=idx,
    )
    assert not _cross_join_should_use_emit_driven(
        cross_file=True,
        stacked_join=True,
        n_join=47076,
        search_pool_len=15710,
        join_index=idx,
    )
    monkeypatch.setenv("HC_DATA_AGG_JOIN_EMIT_DRIVEN", "off")
    assert not _cross_join_should_use_emit_driven(
        cross_file=True,
        stacked_join=False,
        n_join=47076,
        search_pool_len=15710,
        join_index=idx,
    )


def test_join_search_rows_for_slice_indexed_no_list_copy() -> None:
    """索引ヒットは共有参照を返し、毎スライスの list() コピーをしない。"""
    from svc.svc_data_agg import (
        _join_cell_compare_norm,
        _join_search_rows_for_slice_indexed,
    )

    rows = [{"K": "a", "Out": None}, {"K": "a", "Out": None}]
    idx_cols, idx_map = _build_join_search_index(rows, [{"item": "K"}])
    key = (_join_cell_compare_norm("a"),)
    stored = idx_map[key]
    hit = _join_search_rows_for_slice_indexed(idx_cols, idx_map, {"K": ["a"]}, 0)
    assert hit is stored
    assert _join_search_rows_for_slice_indexed(idx_cols, idx_map, {"K": ["missing"]}, 0) == []


def _clone_join_pool(pool: list[dict[str, Any]]) -> list[dict[str, Any]]:
    return [dict(r) for r in pool]


def test_emit_driven_cross_join_matches_classic_paired_fill_in(
    monkeypatch: Any,
) -> None:
    """同一キー複数ホスト寄与でも fill_in の最終状態が classic と一致する。"""
    item = {
        "name": "出荷番号",
        "sources": [
            {
                "type": "cell",
                "ui_scenario_source_v1": {
                    "file_pattern": "出荷指示書",
                    "join_defs": [
                        {"item": "出荷年月日", "cell": "A1"},
                        {"item": "型名", "cell": "G1"},
                    ],
                    "link_defs": [
                        {"item": "エンドユーザ", "cell": "C1", "mode": "セル座標"},
                    ],
                },
            }
        ],
    }
    # emit（比較側）
    emit = [
        {
            "出荷年月日": "20220101",
            "型名": "T1",
            "出荷番号": None,
            "エンドユーザ": None,
            "__file_path": r"C:\光.xlsx",
            "__iter_index": 0,
        },
        {
            "出荷年月日": "20220101",
            "型名": "T1",
            "出荷番号": None,
            "エンドユーザ": None,
            "__file_path": r"C:\光.xlsx",
            "__iter_index": 1,
        },
        {
            "出荷年月日": "20220102",
            "型名": "T2",
            "出荷番号": "KEEP",
            "エンドユーザ": "OLD",
            "__file_path": r"C:\光.xlsx",
            "__iter_index": 2,
        },
    ]
    host_rows = [
        {
            "出荷年月日": "20220101",
            "型名": "T1",
            "出荷番号": "H0",
            "__file_path": r"C:\指示.xlsx",
            "__iter_index": 0,
        },
        {
            "出荷年月日": "20220101",
            "型名": "T1",
            "出荷番号": "H1",
            "__file_path": r"C:\指示.xlsx",
            "__iter_index": 1,
        },
    ]
    bundle = {
        "primary_values": ["S0", "S1", "S2"],
        "join_values": {
            "出荷年月日": ["20220101", "20220101", "20220102"],
            "型名": ["T1", "T1", "T2"],
        },
        "link_values": {"エンドユーザ": ["U0", "U1", "U2"]},
    }
    join_defs = [{"item": "出荷年月日"}, {"item": "型名"}]
    header_set = {"出荷番号", "出荷年月日", "型名", "エンドユーザ"}

    def _run(path_mode: str) -> list[dict[str, Any]]:
        monkeypatch.setenv("HC_DATA_AGG_JOIN_EMIT_DRIVEN", path_mode)
        pool = _clone_join_pool(emit + host_rows)
        emit_part = pool[: len(emit)]
        host_part = pool[len(emit) :]
        join_index = _build_join_search_index(emit_part, join_defs)
        _apply_join_key_search_write(
            pool,
            item,
            "出荷番号",
            bundle,
            "fill_in",
            search_pool=emit_part,
            cross_file=True,
            join_index=join_index,
            join_host_rows=host_part,
            header_set=header_set,
        )
        return pool

    classic = _run("off")
    driven = _run("force")
    assert len(classic) == len(driven)
    for a, b in zip(classic, driven):
        assert a.get("出荷番号") == b.get("出荷番号")
        assert a.get("エンドユーザ") == b.get("エンドユーザ")
    # fill_in: 先勝ち → T1 行は S0 / U0、KEEP は維持
    assert classic[0]["出荷番号"] == "S0"
    assert classic[0]["エンドユーザ"] == "U0"
    assert classic[2]["出荷番号"] == "KEEP"
    assert classic[2]["エンドユーザ"] == "OLD"


def test_emit_driven_cross_join_matches_classic_1prim_overwrite(
    monkeypatch: Any,
) -> None:
    item = {
        "name": "Out",
        "sources": [
            {
                "type": "cell",
                "ui_scenario_source_v1": {
                    "join_defs": [{"item": "KeyCol", "cell": "X1"}],
                    "link_defs": [
                        {"item": "L", "cell": "Y1", "mode": "セル座標"},
                    ],
                },
            }
        ],
    }
    emit = [
        {"KeyCol": "a", "Out": None, "L": None, "__file_path": r"C:\e.xlsx"},
        {"KeyCol": "b", "Out": None, "L": None, "__file_path": r"C:\e.xlsx"},
    ]
    bundle = {
        "primary_values": ["MAIN"],
        "join_values": {"KeyCol": ["a", "b", "a"]},
        "link_values": {"L": ["L0", "L1", "L2"]},
    }
    def _run(mode: str) -> list[dict[str, Any]]:
        monkeypatch.setenv("HC_DATA_AGG_JOIN_EMIT_DRIVEN", mode)
        pool = _clone_join_pool(emit)
        join_index = _build_join_search_index(pool, [{"item": "KeyCol"}])
        _apply_join_key_search_write(
            pool,
            item,
            "Out",
            bundle,
            "overwrite",
            search_pool=pool,
            cross_file=True,
            join_index=join_index,
            join_host_rows=[],
            header_set={"Out", "KeyCol", "L"},
        )
        return pool

    classic = _run("off")
    driven = _run("force")
    assert classic[0]["Out"] == driven[0]["Out"] == "MAIN"
    assert classic[1]["Out"] == driven[1]["Out"] == "MAIN"
    # overwrite: 同一キー a は L0→L2 で最終 L2
    assert classic[0]["L"] == driven[0]["L"] == "L2"
    assert classic[1]["L"] == driven[1]["L"] == "L1"


def test_emit_driven_cancel_during_by_key_build(monkeypatch: Any) -> None:
    """emit_driven の by_key 構築中も poll され、途中中止できる。"""
    from svc.data_agg_cancel import DataAggCancelled, batch_cancel_scope

    monkeypatch.setenv("HC_DATA_AGG_JOIN_EMIT_DRIVEN", "force")
    n = 40
    emit = [{"K": "k%s" % i, "Out": None, "__file_path": r"C:\e.xlsx"} for i in range(n)]
    item = {
        "name": "Out",
        "sources": [
            {
                "type": "cell",
                "ui_scenario_source_v1": {
                    "file_pattern": "host",
                    "join_defs": [{"item": "K", "cell": "A1"}],
                },
            }
        ],
    }
    bundle = {
        "primary_values": ["MAIN"],
        "join_values": {"K": ["k%s" % i for i in range(n)]},
    }
    polls = {"n": 0}

    def _chk(*, force: bool = False) -> None:
        polls["n"] += 1
        if polls["n"] >= 8:
            raise DataAggCancelled()

    join_index = _build_join_search_index(emit, [{"item": "K"}])
    with batch_cancel_scope(_chk):
        with pytest.raises(DataAggCancelled):
            _apply_join_key_search_write(
                emit,
                item,
                "Out",
                bundle,
                "fill_in",
                search_pool=emit,
                cross_file=True,
                join_index=join_index,
                join_host_rows=[],
            )
    assert polls["n"] >= 8
    # 途中中止のため全行が埋まっている必要はない
    assert sum(1 for r in emit if r.get("Out") == "MAIN") < n
