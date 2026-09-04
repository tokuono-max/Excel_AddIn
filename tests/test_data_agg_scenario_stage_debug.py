# -*- coding: utf-8 -*-
"""シナリオデバッグの UNC ステージ読取マップ／共有セッション。"""
from __future__ import annotations

from pathlib import Path

from svc.svc_data_agg_debug_run import (
    ScenarioDebugStageSession,
    _scenario_staged_read_map,
    _sheet_column_preview,
)


def test_scenario_staged_read_map_local_passthrough(tmp_path: Path, monkeypatch) -> None:
    local = tmp_path / "a.xlsx"
    local.write_bytes(b"data")
    monkeypatch.setenv("DATA_AGG_NETWORK_STAGE", "1")
    with _scenario_staged_read_map([str(local)], scan_root=str(tmp_path)) as m:
        assert m == {}


def test_scenario_staged_read_map_network_copies(tmp_path: Path, monkeypatch) -> None:
    src = tmp_path / "book.xls"
    src.write_bytes(b"hello-xls")
    unc = str(src.resolve())
    monkeypatch.setenv("DATA_AGG_NETWORK_STAGE", "1")
    monkeypatch.setattr(
        "svc.data_agg_path_network.path_is_network",
        lambda _p: True,
    )
    monkeypatch.setattr(
        "svc.data_agg_network_stage.path_is_network",
        lambda _p: True,
    )
    with _scenario_staged_read_map([unc], scan_root=str(tmp_path)) as m:
        assert unc in m
        assert m[unc] != unc
        assert Path(m[unc]).is_file()
        assert Path(m[unc]).read_bytes() == b"hello-xls"


def test_scenario_stage_session_reuse_same_paths(tmp_path: Path, monkeypatch) -> None:
    """シート名→主キーで同一パスなら再コピーせず io を再利用する。"""
    src = tmp_path / "book.xlsx"
    src.write_bytes(b"PK-fake")
    unc = str(src.resolve())
    monkeypatch.setenv("DATA_AGG_NETWORK_STAGE", "1")
    monkeypatch.setattr(
        "svc.data_agg_path_network.path_is_network",
        lambda _p: True,
    )
    monkeypatch.setattr(
        "svc.data_agg_network_stage.path_is_network",
        lambda _p: True,
    )
    builds: list[list[str]] = []

    import svc.data_agg_network_stage as stage_mod

    real_build = stage_mod.build_network_stage_batch

    def _counting_build(paths, **kwargs):
        builds.append([str(p) for p in paths])
        return real_build(paths, **kwargs)

    monkeypatch.setattr(
        "svc.data_agg_network_stage.build_network_stage_batch",
        _counting_build,
    )

    sess = ScenarioDebugStageSession()
    m1 = sess.ensure([unc], scan_root=str(tmp_path))
    m2 = sess.ensure([unc], scan_root=str(tmp_path))
    assert len(builds) == 1
    assert m1[unc] == m2[unc]
    assert Path(m1[unc]).is_file()
    # prefetch 相当: カバー済みの部分集合でも再コピーしない
    m3 = sess.ensure([unc], scan_root=str(tmp_path))
    assert len(builds) == 1
    assert m3[unc] == m1[unc]
    sess.clear()


def test_scenario_stage_session_superset_rebuilds(tmp_path: Path, monkeypatch) -> None:
    """カバー外パスを足すと再構築する（prefetch で全scanを渡すとこうなる）。"""
    a = tmp_path / "a.xlsx"
    b = tmp_path / "b.xlsx"
    a.write_bytes(b"aaa")
    b.write_bytes(b"bbb")
    ua, ub = str(a.resolve()), str(b.resolve())
    monkeypatch.setenv("DATA_AGG_NETWORK_STAGE", "1")
    monkeypatch.setattr(
        "svc.data_agg_path_network.path_is_network",
        lambda _p: True,
    )
    monkeypatch.setattr(
        "svc.data_agg_network_stage.path_is_network",
        lambda _p: True,
    )
    builds: list[int] = []
    import svc.data_agg_network_stage as stage_mod

    real_build = stage_mod.build_network_stage_batch

    def _counting_build(paths, **kwargs):
        builds.append(len(list(paths)))
        return real_build(paths, **kwargs)

    monkeypatch.setattr(
        "svc.data_agg_network_stage.build_network_stage_batch",
        _counting_build,
    )
    sess = ScenarioDebugStageSession()
    sess.ensure([ua], scan_root=str(tmp_path))
    sess.ensure([ua, ub], scan_root=str(tmp_path))
    assert builds == [1, 2]
    sess.clear()


def test_scenario_stage_session_union_keeps_prior(tmp_path: Path, monkeypatch) -> None:
    """不足パス追加時は和集合で再構築し、既存パスもカバーする。"""
    a = tmp_path / "a.xlsx"
    b = tmp_path / "b.xlsx"
    a.write_bytes(b"aaa")
    b.write_bytes(b"bbb")
    ua, ub = str(a.resolve()), str(b.resolve())
    monkeypatch.setenv("DATA_AGG_NETWORK_STAGE", "1")
    monkeypatch.setattr(
        "svc.data_agg_path_network.path_is_network",
        lambda _p: True,
    )
    monkeypatch.setattr(
        "svc.data_agg_network_stage.path_is_network",
        lambda _p: True,
    )
    sess = ScenarioDebugStageSession()
    m1 = sess.ensure([ua], scan_root=str(tmp_path))
    m2 = sess.ensure([ua, ub], scan_root=str(tmp_path))
    assert ua in m2 and ub in m2
    assert Path(m2[ua]).is_file() and Path(m2[ub]).is_file()
    # 再構築後も ua は読める（旧 batch cleanup 済みでも新 batch に含まれる）
    assert m1[ua] != m2[ua] or Path(m2[ua]).read_bytes() == b"aaa"
    sess.clear()


def test_sheet_column_preview_uses_read_map_io(tmp_path: Path, monkeypatch) -> None:
    """read_map があるとき表示パスではなく io 側を開く。"""
    import openpyxl

    display = tmp_path / "display.xlsx"
    staged = tmp_path / "staged.xlsx"
    wb = openpyxl.Workbook()
    wb.active.title = "OnlyOnStaged"
    wb.save(staged)
    wb.close()
    # display は空に近い別ファイル（シート名が違う）
    wb2 = openpyxl.Workbook()
    wb2.active.title = "WrongSheet"
    wb2.save(display)
    wb2.close()

    disp_s = str(display)
    preview = _sheet_column_preview(
        {
            "sheet_name": "OnlyOnStaged",
            "ui_scenario_source_v1": {"sheet_rule": "完全一致"},
        },
        [disp_s],
        10,
        read_map={disp_s: str(staged)},
    )
    assert preview == ["OnlyOnStaged"]
