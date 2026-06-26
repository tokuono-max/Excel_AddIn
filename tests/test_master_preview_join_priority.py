# -*- coding: utf-8 -*-
from __future__ import annotations

import sys
from pathlib import Path

import pytest
from openpyxl import Workbook
from core.core_join_compare import join_compare_display_key

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from svc.svc_data_agg import (  # noqa: E402
    _master_preview_join_item_effective,
    _master_preview_per_file_pool_cap,
    compute_batch_table_rows,
    filter_file_paths_for_master_preview,
    reorder_paths_for_master_preview_join_priority,
)
from svc.data_agg_master_preview import table_rows_to_join_search_seed_pool  # noqa: E402
from tests.test_data_agg_batch_stability import (  # noqa: E402
    _active_worksheet,
    _cross_join_mini_scenario,
)


def test_reorder_paths_interleaves_side_and_host(tmp_path: Path) -> None:
    """PT番号ホスト: side（光特性）と host（紐づけ）をラウンドロビンで交互に並べる。"""
    data, paths = _cross_join_mini_scenario(tmp_path)
    anchor, join_f = paths[0], paths[1]
    shuffled = [join_f, join_f, anchor, join_f]
    headers = [str(it.get("name") or "") for it in data["items"]]
    dd: dict = {"mi_idx": 2}
    ordered = reorder_paths_for_master_preview_join_priority(
        shuffled, data["items"], headers, dd
    )
    assert ordered[0] == anchor
    assert ordered[1] == join_f
    assert dd.get("master_preview_priority_files")
    assert dd.get("master_preview_join_side_patterns")
    assert dd.get("master_preview_join_host_patterns")


def test_master_preview_per_file_pool_cap_is_read_limit() -> None:
    data, _paths = _cross_join_mini_scenario(Path("."))
    host = data["items"][2]
    headers = [str(it.get("name") or "") for it in data["items"]]
    cap = _master_preview_per_file_pool_cap(200, host, data["items"], headers)
    assert cap == 200


def test_master_preview_join_priority_fills_link_columns(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """横断 join プレビューで PT番号・製番が結合結果に載る。"""
    data, paths = _cross_join_mini_scenario(tmp_path)
    anchor, join_f = paths[0], paths[1]
    data["__debug_diag"] = {
        "enabled": True,
        "source": "ui_data_agg_debug.master_preview",
        "mi_idx": 2,
        "join_search_skip_seed": True,
    }
    paths_ordered = [join_f, anchor]
    monkeypatch.setenv("DATA_AGG_FILE_PARALLEL_WORKERS", "0")
    monkeypatch.setenv("DATA_AGG_MASTER_PARALLEL_EXTRACT", "0")
    headers, rows, _ev, _je = compute_batch_table_rows(
        data,
        paths_ordered,
        max_primary_rows=4,
        max_table_rows=4,
        probe_caller="test_join_priority",
    )
    idx_pt = headers.index("PT番号")
    idx_seq = headers.index("製番")
    pt_vals = [r[idx_pt] for r in rows if r[idx_pt]]
    seq_vals = [r[idx_seq] for r in rows if r[idx_seq]]
    assert pt_vals, "PT番号が空（join プレビューが効いていない）"
    assert seq_vals, "製番が空（link allowlist 経由の取得が効いていない）"


def test_master_preview_per_file_cap_allows_second_file_type(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """1 ファイルが cap を独占しない（side と host の両方がプールに入る）。"""
    data, paths = _cross_join_mini_scenario(tmp_path)
    anchor = Path(paths[0])
    wb = Workbook()
    ws = _active_worksheet(wb)
    ws.title = "ﾃﾞｰﾀ"
    for i in range(8):
        ws.cell(row=7 + i, column=3, value="DEV%s" % i)
        ws.cell(row=7 + i, column=13, value="MAC-A")
    wb.save(anchor)
    for src in data["items"][0]["sources"]:
        src["repeat_max"] = 8
    data["__debug_diag"] = {
        "enabled": True,
        "source": "ui_data_agg_debug.master_preview",
        "mi_idx": 2,
        "join_search_skip_seed": True,
    }
    monkeypatch.setenv("DATA_AGG_FILE_PARALLEL_WORKERS", "0")
    monkeypatch.setenv("DATA_AGG_MASTER_PARALLEL_EXTRACT", "0")
    _h, rows, _ev, _je = compute_batch_table_rows(
        data,
        paths,
        max_primary_rows=4,
        max_table_rows=4,
        probe_caller="test_per_file_cap",
    )
    headers = list(_h)
    dd = data.get("__debug_diag") or {}
    idx_pt = headers.index("PT番号")
    idx_seq = headers.index("製番")
    pt_vals = [r[idx_pt] for r in rows if r[idx_pt]]
    seq_vals = [r[idx_seq] for r in rows if r[idx_seq]]
    assert pt_vals, "PT番号が空（side と host の両方がプールに入っていない）"
    assert seq_vals, "製番が空"
    assert len(rows) <= 4
    assert len(rows) >= 1
    assert int(dd.get("master_preview_pool_row_cap") or 0) == 8


def test_master_preview_interleave_reads_host_after_side_cap(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """光特性が複数あっても host（紐づけ）がプールに入り join が走る。"""
    data, paths = _cross_join_mini_scenario(tmp_path)
    anchor, join_f = paths[0], paths[1]
    anchor2 = tmp_path / "光特性履歴_test2.xlsx"
    anchor3 = tmp_path / "光特性履歴_test3.xlsx"
    for extra in (anchor2, anchor3):
        extra.write_bytes(Path(anchor).read_bytes())
    wb = Workbook()
    ws = _active_worksheet(wb)
    ws.title = "ﾃﾞｰﾀ"
    for i in range(50):
        ws.cell(row=7 + i, column=3, value="DEV%s" % i)
        ws.cell(row=7 + i, column=13, value="MAC-A")
    wb.save(anchor)
    for src in data["items"][0]["sources"]:
        src["repeat_max"] = 50
    data["__debug_diag"] = {
        "enabled": True,
        "source": "ui_data_agg_debug.master_preview",
        "mi_idx": 2,
        "join_search_skip_seed": True,
    }
    paths_many_side = [str(anchor), str(anchor2), str(anchor3), str(join_f)]
    monkeypatch.setenv("DATA_AGG_FILE_PARALLEL_WORKERS", "0")
    monkeypatch.setenv("DATA_AGG_MASTER_PARALLEL_EXTRACT", "0")
    headers, rows, _ev, _je = compute_batch_table_rows(
        data,
        paths_many_side,
        max_primary_rows=200,
        max_table_rows=100,
        probe_caller="test_interleave_host",
    )
    dd = data.get("__debug_diag") or {}
    idx_pt = headers.index("PT番号")
    idx_seq = headers.index("製番")
    pt_vals = [r[idx_pt] for r in rows if r[idx_pt]]
    seq_vals = [r[idx_seq] for r in rows if r[idx_seq]]
    assert pt_vals, "PT番号が空（host がプールに入っていない）"
    assert seq_vals, "製番が空（link が効いていない）"
    assert len(rows) >= 1
    pool_cap = int(dd.get("master_preview_pool_row_cap") or 0)
    assert pool_cap == 800


def test_master_preview_anchor_rows_keep_prior_row_identity(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    data, paths = _cross_join_mini_scenario(tmp_path)
    headers = [str(it.get("name") or "") for it in data["items"]]
    seed_rows = [["DEV1", "MAC-A", None, None]]
    seed_pool = table_rows_to_join_search_seed_pool(headers, seed_rows)
    data["__debug_diag"] = {
        "enabled": True,
        "source": "ui_data_agg_debug.master_preview",
        "mi_idx": 2,
        "join_search_skip_seed": False,
        "join_search_seed_pool": seed_pool,
        "preview_anchor_row_keys": [[seed_pool[0]["__file_path"], seed_pool[0]["__iter_index"]]],
    }
    monkeypatch.setenv("DATA_AGG_FILE_PARALLEL_WORKERS", "0")
    monkeypatch.setenv("DATA_AGG_MASTER_PARALLEL_EXTRACT", "0")
    out_headers, out_rows, _ev, _je = compute_batch_table_rows(
        data,
        paths,
        max_primary_rows=10,
        max_table_rows=10,
        probe_caller="test_anchor_rows",
    )
    idx_dev = out_headers.index("機器番号")
    idx_pt = out_headers.index("PT番号")
    idx_seq = out_headers.index("製番")
    assert len(out_rows) == 1
    assert out_rows[0][idx_dev] == "DEV1"
    assert join_compare_display_key(out_rows[0][idx_pt]) == "PT-001"
    assert join_compare_display_key(out_rows[0][idx_seq]) == "SEQ-1"


def test_reorder_paths_uses_topology_when_prior_sources_cleared(
    tmp_path: Path,
) -> None:
    """carry-forward で前項目 sources が空でも topology で横断 join 判定できる。"""
    import copy

    data, paths = _cross_join_mini_scenario(tmp_path)
    stepped_items = copy.deepcopy(data["items"])
    for j, it in enumerate(stepped_items):
        if j < 2:
            it["sources"] = []
    headers = [str(it.get("name") or "") for it in data["items"]]
    dd: dict = {
        "mi_idx": 2,
        "preview_join_topology_items": copy.deepcopy(data["items"]),
    }
    ordered = reorder_paths_for_master_preview_join_priority(
        paths, stepped_items, headers, dd
    )
    assert ordered[0] == paths[0]
    assert dd.get("master_preview_join_side_patterns")
    assert dd.get("master_preview_join_host_patterns")


def test_carry_forward_topology_enables_cross_join_pt(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """mpv UI 相当: carry-forward + topology で PT が table_rows に載る。"""
    import copy

    from svc.data_agg_master_preview import (  # noqa: WPS433
        MASTER_PREVIEW_DIAG_SOURCE,
        scenario_for_stepped_preview,
        table_rows_to_join_search_seed_pool,
    )
    from svc.svc_data_agg import master_preview_extract_item_allowlist  # noqa: WPS433

    data, paths = _cross_join_mini_scenario(tmp_path)
    headers = [str(it.get("name") or "") for it in data["items"]]
    prior_rows = [["DEV1", "MAC-A", None, None]]
    scen = scenario_for_stepped_preview(
        data,
        mi_idx=2,
        master_step_idx=1,
        active_slot_indices=[0],
        carry_forward_completed_items=True,
    )
    dd = scen["__debug_diag"]
    dd["preview_join_topology_items"] = copy.deepcopy(data["items"])
    dd["source"] = MASTER_PREVIEW_DIAG_SOURCE
    dd["master_preview_join_read_full_files"] = True
    allow_ix = master_preview_extract_item_allowlist(data, mi_idx=2)
    if allow_ix is not None:
        dd["preview_extract_item_allowlist"] = list(allow_ix)
    dd["join_search_seed_pool"] = table_rows_to_join_search_seed_pool(
        headers, prior_rows
    )
    monkeypatch.setenv("DATA_AGG_FILE_PARALLEL_WORKERS", "0")
    monkeypatch.setenv("DATA_AGG_MASTER_PARALLEL_EXTRACT", "0")
    out_headers, out_rows, _ev, _je = compute_batch_table_rows(
        scen,
        paths,
        max_primary_rows=10,
        max_table_rows=10,
        probe_caller="test_carry_forward_topology",
    )
    idx_pt = out_headers.index("PT番号")
    pt_vals = [r[idx_pt] for r in out_rows if r[idx_pt] not in (None, "")]
    assert pt_vals, "carry-forward + topology でも PT が空"
    assert join_compare_display_key(pt_vals[0]) == "PT-001"


def test_record_join_host_patterns_for_non_cross_chain() -> None:
    from svc.svc_data_agg import _master_preview_record_join_host_patterns_only

    host = {
        "sources": [
            {
                "type": "cell",
                "ui_scenario_source_v1": {"file_pattern": "梱包"},
            }
        ]
    }
    dd: dict = {}
    _master_preview_record_join_host_patterns_only(dd, host=host, topo=[host])
    assert dd.get("master_preview_join_host_patterns") == ["梱包"]


def test_seed_pool_prefers_anchor_file_path_over_synthetic() -> None:
    rows = table_rows_to_join_search_seed_pool(
        ["機器番号", "MACアドレス"],
        [["DEV1", "MAC-A"]],
        anchor_file_path=r"C:\data\光特性履歴.xlsx",
    )
    assert rows[0]["__file_path"] == r"C:\data\光特性履歴.xlsx"
    assert not str(rows[0]["__file_path"]).startswith("mpv_table_seed://")


def test_cross_detection_uses_topology_not_empty_stepped_sources(
    tmp_path: Path,
) -> None:
    import copy

    from svc.svc_data_agg import (  # noqa: WPS433
        _join_host_needs_cross_file_pool,
    )

    data, _paths = _cross_join_mini_scenario(tmp_path)
    headers = [str(it.get("name") or "") for it in data["items"]]
    stepped = copy.deepcopy(data["items"])
    for it in stepped:
        it["sources"] = []
    host_topo = data["items"][2]
    host_stepped = stepped[2]
    assert not _join_host_needs_cross_file_pool(host_stepped, stepped, headers)
    assert _join_host_needs_cross_file_pool(host_topo, data["items"], headers)


def test_carry_forward_step0_topology_cross_join_pt(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """step0（n_pick=0）でも topology により横断 join が成立する。"""
    import copy

    from svc.data_agg_master_preview import (  # noqa: WPS433
        MASTER_PREVIEW_DIAG_SOURCE,
        scenario_for_stepped_preview,
        table_rows_to_join_search_seed_pool,
    )
    from svc.svc_data_agg import master_preview_extract_item_allowlist  # noqa: WPS433

    data, paths = _cross_join_mini_scenario(tmp_path)
    headers = [str(it.get("name") or "") for it in data["items"]]
    anchor, join_f = paths[0], paths[1]
    prior_rows = [["DEV1", "MAC-A", None, None]]
    scen = scenario_for_stepped_preview(
        data,
        mi_idx=2,
        master_step_idx=0,
        active_slot_indices=[0],
        carry_forward_completed_items=True,
    )
    dd = scen["__debug_diag"]
    dd["preview_join_topology_items"] = copy.deepcopy(data["items"])
    dd["source"] = MASTER_PREVIEW_DIAG_SOURCE
    dd["master_preview_join_read_full_files"] = True
    allow_ix = master_preview_extract_item_allowlist(data, mi_idx=2)
    if allow_ix is not None:
        dd["preview_extract_item_allowlist"] = list(allow_ix)
    dd["join_search_seed_pool"] = table_rows_to_join_search_seed_pool(
        headers,
        prior_rows,
        anchor_file_path=anchor,
    )
    monkeypatch.setenv("DATA_AGG_FILE_PARALLEL_WORKERS", "0")
    monkeypatch.setenv("DATA_AGG_MASTER_PARALLEL_EXTRACT", "0")
    out_headers, out_rows, _ev, _je = compute_batch_table_rows(
        scen,
        [join_f],
        max_primary_rows=10,
        max_table_rows=10,
        probe_caller="test_step0_topology",
    )
    idx_pt = out_headers.index("PT番号")
    pt_vals = [r[idx_pt] for r in out_rows if r[idx_pt] not in (None, "")]
    assert pt_vals, "step0 + topology でも PT が空"
    assert join_compare_display_key(pt_vals[0]) == "PT-001"


def test_master_preview_join_item_effective_restores_defs_at_step0(
    tmp_path: Path,
) -> None:
    """step0 で stepped sources が空でも join_defs を topology から復元する。"""
    import copy

    from svc.data_agg_master_preview import (  # noqa: WPS433
        scenario_for_stepped_preview,
    )
    from svc.svc_data_agg import (  # noqa: WPS433
        _item_join_defs_list,
        _master_preview_join_item_effective,
    )

    data, _paths = _cross_join_mini_scenario(tmp_path)
    scen = scenario_for_stepped_preview(
        data,
        mi_idx=2,
        master_step_idx=0,
        active_slot_indices=[0],
        carry_forward_completed_items=True,
    )
    items = list(scen.get("items") or [])
    dd = scen.get("__debug_diag") or {}
    dd["preview_join_topology_items"] = copy.deepcopy(data["items"])
    stepped_host = items[2]
    assert not _item_join_defs_list(stepped_host)
    eff = _master_preview_join_item_effective(
        items,
        2,
        dd,
        preview_master_mode=True,
    )
    assert _item_join_defs_list(eff)
    assert _item_join_defs_list(eff) == _item_join_defs_list(data["items"][2])


def test_reorder_paths_when_stepped_host_has_no_join_defs(
    tmp_path: Path,
) -> None:
    """step0 の結合ホスト（sources 空）でも topology で path 優先が効く。"""
    import copy

    from svc.data_agg_master_preview import scenario_for_stepped_preview  # noqa: WPS433

    data, paths = _cross_join_mini_scenario(tmp_path)
    scen = scenario_for_stepped_preview(
        data,
        mi_idx=2,
        master_step_idx=0,
        active_slot_indices=[0],
        carry_forward_completed_items=True,
    )
    items = list(scen.get("items") or [])
    headers = [str(it.get("name") or "") for it in data["items"]]
    dd = scen.get("__debug_diag") or {}
    dd["mi_idx"] = 2
    dd["preview_join_topology_items"] = copy.deepcopy(data["items"])
    ordered = reorder_paths_for_master_preview_join_priority(
        [paths[1], paths[0]], items, headers, dd
    )
    assert ordered[0] == paths[0]
    assert dd.get("master_preview_join_host_patterns")


def test_filter_master_preview_topology_fallback_when_stepped_empty(
    tmp_path: Path,
) -> None:
    """carry-forward で stepped sources が全空でも topology で path を復元する。"""
    import copy

    from svc.data_agg_master_preview import MASTER_PREVIEW_DIAG_SOURCE  # noqa: WPS433

    data, paths = _cross_join_mini_scenario(tmp_path)
    stepped = copy.deepcopy(data["items"])
    for it in stepped:
        it["sources"] = []
    files = [str(paths[0]), str(paths[1]), r"C:\data\other.xlsx"]
    dd = {
        "source": MASTER_PREVIEW_DIAG_SOURCE,
        "mi_idx": 2,
        "preview_join_topology_items": copy.deepcopy(data["items"]),
        "preview_extract_item_allowlist": [0, 1, 2],
    }
    out = filter_file_paths_for_master_preview(files, stepped, dd)
    assert str(paths[0]) in out
    assert str(paths[1]) in out
    assert r"C:\data\other.xlsx" not in out


def test_filter_master_preview_cross_join_prefers_topology_over_single_stepped(
    tmp_path: Path,
) -> None:
    """横断 join: stepped がホスト 1 パターンだけでも topology で side を含める。"""
    import copy

    from svc.data_agg_master_preview import MASTER_PREVIEW_DIAG_SOURCE  # noqa: WPS433

    data, paths = _cross_join_mini_scenario(tmp_path)
    stepped = copy.deepcopy(data["items"])
    for j, it in enumerate(stepped):
        if j < 2:
            it["sources"] = []
    files = [str(paths[0]), str(paths[1])]
    dd = {
        "source": MASTER_PREVIEW_DIAG_SOURCE,
        "mi_idx": 2,
        "preview_join_topology_items": copy.deepcopy(data["items"]),
        "preview_extract_item_allowlist": [0, 1, 2],
    }
    out = filter_file_paths_for_master_preview(files, stepped, dd)
    assert len(out) == 2
    assert set(out) == {str(paths[0]), str(paths[1])}


def test_master_preview_pool_cap_applies_with_join_full_read(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """join full read 有効時もプール行 cap が効く（巨大 pool で join 列が空になるのを防ぐ）。"""
    import copy

    from svc.data_agg_master_preview import (  # noqa: WPS433
        MASTER_PREVIEW_DIAG_SOURCE,
        scenario_for_stepped_preview,
        table_rows_to_join_search_seed_pool,
    )

    data, paths = _cross_join_mini_scenario(tmp_path)
    headers = [str(it.get("name") or "") for it in data["items"]]
    scen = scenario_for_stepped_preview(
        data,
        mi_idx=2,
        master_step_idx=1,
        active_slot_indices=[0],
        carry_forward_completed_items=True,
    )
    dd = scen["__debug_diag"]
    dd["source"] = MASTER_PREVIEW_DIAG_SOURCE
    dd["preview_join_topology_items"] = copy.deepcopy(data["items"])
    dd["master_preview_join_read_full_files"] = True
    dd["join_search_seed_pool"] = table_rows_to_join_search_seed_pool(
        headers, [["DEV1", "MAC-A", None, None]]
    )
    monkeypatch.setenv("DATA_AGG_FILE_PARALLEL_WORKERS", "0")
    monkeypatch.setenv("DATA_AGG_MASTER_PARALLEL_EXTRACT", "0")
    _h, _rows, _ev, _je = compute_batch_table_rows(
        scen,
        paths,
        max_primary_rows=4,
        max_table_rows=4,
        probe_caller="test_pool_cap_join_full_read",
    )
    cap = int(dd.get("master_preview_pool_row_cap") or 0)
    assert cap > 0
    pool_rows = int(dd.get("master_preview_pool_rows") or 0)
    assert pool_rows <= cap


def test_filter_master_preview_stacked_join_host_only(
    tmp_path: Path,
) -> None:
    """積み上げ join: path filter は当ステップのホストファイルのみ。"""
    import copy

    from svc.data_agg_master_preview import MASTER_PREVIEW_DIAG_SOURCE  # noqa: WPS433

    data, paths = _cross_join_mini_scenario(tmp_path)
    stepped = copy.deepcopy(data["items"])
    for j, it in enumerate(stepped):
        if j < 2:
            it["sources"] = []
    files = [str(paths[0]), str(paths[1])]
    dd = {
        "source": MASTER_PREVIEW_DIAG_SOURCE,
        "mi_idx": 2,
        "preview_join_topology_items": copy.deepcopy(data["items"]),
        "master_preview_stacked_join": True,
        "join_search_seed_from_table_rows": True,
        "join_search_seed_pool": [{"機器番号": "DEV1"}],
    }
    out = filter_file_paths_for_master_preview(files, stepped, dd)
    assert out == [str(paths[1])]


def test_stacked_join_cross_join_pt_host_only(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """積み上げ join: 前項目表示行 + ホストのみ読込で PT・製番が埋まる。"""
    import copy

    from svc.data_agg_master_preview import (  # noqa: WPS433
        MASTER_PREVIEW_DIAG_SOURCE,
        scenario_for_stepped_preview,
        table_rows_to_join_search_seed_pool,
    )

    data, paths = _cross_join_mini_scenario(tmp_path)
    anchor, join_f = paths[0], paths[1]
    headers = [str(it.get("name") or "") for it in data["items"]]
    prior_rows = [["DEV1", "MAC-A", None, None]]
    scen = scenario_for_stepped_preview(
        data,
        mi_idx=2,
        master_step_idx=1,
        active_slot_indices=[0],
        carry_forward_completed_items=True,
    )
    dd = scen["__debug_diag"]
    dd["source"] = MASTER_PREVIEW_DIAG_SOURCE
    dd["mi_idx"] = 2
    dd["preview_join_topology_items"] = copy.deepcopy(data["items"])
    seed_pool = table_rows_to_join_search_seed_pool(headers, prior_rows)
    dd["join_search_seed_pool"] = seed_pool
    dd["join_search_seed_from_table_rows"] = True
    dd["master_preview_stacked_join"] = True
    dd["preview_extract_item_allowlist"] = [2]
    dd["master_preview_join_read_full_files"] = False
    dd["preview_anchor_row_keys"] = [
        [str(seed_pool[0]["__file_path"]), int(seed_pool[0]["__iter_index"])]
    ]
    monkeypatch.setenv("DATA_AGG_FILE_PARALLEL_WORKERS", "0")
    monkeypatch.setenv("DATA_AGG_MASTER_PARALLEL_EXTRACT", "0")
    out_headers, out_rows, _ev, _je = compute_batch_table_rows(
        scen,
        [anchor, join_f],
        max_primary_rows=10,
        max_table_rows=10,
        probe_caller="test_stacked_join_pt",
    )
    assert int(dd.get("master_preview_stats_files_read") or 0) <= 1
    idx_pt = out_headers.index("PT番号")
    idx_seq = out_headers.index("製番")
    pt_vals = [r[idx_pt] for r in out_rows if r[idx_pt] not in (None, "")]
    seq_vals = [r[idx_seq] for r in out_rows if r[idx_seq] not in (None, "")]
    assert pt_vals, "積み上げ join で PT が空"
    assert seq_vals, "積み上げ join で製番が空"
    assert join_compare_display_key(pt_vals[0]) == "PT-001"
