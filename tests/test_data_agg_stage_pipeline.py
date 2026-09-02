# -*- coding: utf-8 -*-
from __future__ import annotations

from pathlib import Path
from unittest.mock import MagicMock

import pytest

from svc.data_agg_network_stage import (
    NetworkStageBatch,
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
