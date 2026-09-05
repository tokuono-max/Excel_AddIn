# -*- coding: utf-8 -*-
from __future__ import annotations

import time
from pathlib import Path
from unittest.mock import MagicMock

import pytest

from svc.data_agg_network_stage import (
    build_network_stage_batch,
    cleanup_all_network_stage_dirs,
    register_stage_dir,
    unregister_stage_dir,
)


def test_on_file_staged_callback_local_passthrough(tmp_path: Path) -> None:
    f = tmp_path / "a.xlsx"
    f.write_bytes(b"x")
    staged: list[tuple[int, str, str]] = []

    batch = build_network_stage_batch(
        [f],
        enabled=True,
        on_file_staged=lambda i, d, io: staged.append((i, d, io)),
    )
    assert len(staged) == 1
    assert staged[0][0] == 0
    assert batch.stage_dir is None


def test_cleanup_all_registered_dirs(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(
        "svc.data_agg_network_stage._stage_temp_base",
        lambda: tmp_path / "data_agg_stage",
    )
    base = tmp_path / "data_agg_stage"
    base.mkdir(parents=True)
    sd = base / "abc123"
    sd.mkdir()
    (sd / "f.xlsx").write_bytes(b"1")
    register_stage_dir(sd)
    assert cleanup_all_network_stage_dirs(prune_orphans=False) >= 1
    assert not sd.is_dir()
    unregister_stage_dir(sd)


def test_cleanup_prunes_orphan_with_contents(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """非登録の残骸は中途半端な中身ごと削除する。"""
    monkeypatch.setattr(
        "svc.data_agg_network_stage._stage_temp_base",
        lambda: tmp_path / "data_agg_stage",
    )
    base = tmp_path / "data_agg_stage"
    base.mkdir(parents=True)
    orphan = base / "orphan_deadbeef"
    orphan.mkdir()
    (orphan / "partial.xlsx").write_bytes(b"half")
    (orphan / "x.part").write_bytes(b"part")
    # mtime を古くする（60 秒閾値を超える）
    old = time.time() - 120
    import os

    os.utime(orphan, (old, old))
    n = cleanup_all_network_stage_dirs(prune_orphans=True)
    assert n >= 1
    assert not orphan.is_dir()


def test_pipeline_invokes_extract_on_staged(monkeypatch, tmp_path: Path) -> None:
    from svc.data_agg_stage_pipeline import run_stage_extract_pipeline

    src = tmp_path / "src.xlsx"
    src.write_bytes(b"hi")
    unc = str(src.resolve())
    monkeypatch.setattr(
        "svc.data_agg_network_stage.path_is_network",
        lambda _p: True,
    )
    extract_calls: list[tuple[int, str, str]] = []

    def _extract(fi: int, io: str, disp: str) -> MagicMock:
        extract_calls.append((fi, io, disp))
        return MagicMock()

    batch, results = run_stage_extract_pipeline(
        [unc],
        scan_root=tmp_path,
        cancel_check=None,
        progress_callback=None,
        extract_work=_extract,
        target_workers=2,
        cold_workers=1,
        ramp_files=1,
    )
    try:
        assert len(extract_calls) == 1
        assert extract_calls[0][0] == 1
        assert Path(extract_calls[0][1]).is_file()
        assert extract_calls[0][2] == unc
        assert 1 in results
    finally:
        batch.cleanup()


def test_pipeline_ramp_cold_to_target_no_semaphore_error(
    monkeypatch, tmp_path: Path
) -> None:
    """cold→target ランプで BoundedSemaphore 由来の ValueError が出ないこと。"""
    from svc.data_agg_stage_pipeline import run_stage_extract_pipeline

    files = []
    for i in range(24):
        p = tmp_path / ("f%s.xlsx" % i)
        p.write_bytes(b"x")
        files.append(str(p.resolve()))

    monkeypatch.setattr(
        "svc.data_agg_network_stage.path_is_network",
        lambda _p: True,
    )

    def _extract(fi: int, io: str, disp: str) -> MagicMock:
        return MagicMock(name="res_%s" % fi)

    batch, results = run_stage_extract_pipeline(
        files,
        scan_root=tmp_path,
        cancel_check=None,
        progress_callback=None,
        extract_work=_extract,
        target_workers=4,
        cold_workers=2,
        ramp_files=16,
    )
    try:
        assert len(results) == 24
    finally:
        batch.cleanup()


def test_pipeline_cancel_stops_queueing(monkeypatch, tmp_path: Path) -> None:
    """キャンセル後は extract 投入を止め、未開始 future を打ち切る。"""
    from svc.data_agg_cancel import DataAggCancelled
    from svc.data_agg_stage_pipeline import run_stage_extract_pipeline

    files = []
    for i in range(12):
        p = tmp_path / ("f%s.xlsx" % i)
        p.write_bytes(b"x")
        files.append(str(p.resolve()))

    monkeypatch.setattr(
        "svc.data_agg_network_stage.path_is_network",
        lambda _p: True,
    )

    extract_calls: list[int] = []
    cancel_after = {"n": 3}

    def _extract(fi: int, io: str, disp: str) -> MagicMock:
        extract_calls.append(fi)
        time.sleep(0.05)
        return MagicMock(name="res_%s" % fi)

    polls = {"n": 0}

    def _chk(*, force: bool = False) -> None:
        polls["n"] += 1
        if polls["n"] >= cancel_after["n"]:
            raise DataAggCancelled()

    with pytest.raises(DataAggCancelled):
        run_stage_extract_pipeline(
            files,
            scan_root=tmp_path,
            cancel_check=_chk,
            progress_callback=None,
            extract_work=_extract,
            target_workers=2,
            cold_workers=1,
            ramp_files=1,
        )
    assert len(extract_calls) < 12
