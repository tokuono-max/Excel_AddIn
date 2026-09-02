# -*- coding: utf-8 -*-
from __future__ import annotations

from pathlib import Path

import pytest

from svc.data_agg_network_stage import build_network_stage_batch
from svc.data_agg_path_network import path_class, path_is_network


def test_path_class_unc() -> None:
    assert path_class(r"\\server\share\a.xlsx") == "unc"


def test_path_is_network_local(tmp_path: Path) -> None:
    p = tmp_path / "a.xlsx"
    p.write_bytes(b"x")
    assert not path_is_network(p)


def test_stage_passthrough_local(tmp_path: Path) -> None:
    f = tmp_path / "a.xlsx"
    f.write_bytes(b"data")
    batch = build_network_stage_batch([f], enabled=True)
    assert batch.io_paths == [str(f.resolve())]
    assert batch.display_paths == [str(f.resolve())]
    assert batch.stage_dir is None
    assert batch.staged_files == 0


def test_stage_copies_unc_path(monkeypatch, tmp_path: Path) -> None:
    src = tmp_path / "src.xlsx"
    src.write_bytes(b"hello")
    unc = str(src.resolve())

    monkeypatch.setattr(
        "svc.data_agg_network_stage.path_is_network",
        lambda _p: True,
    )
    batch = build_network_stage_batch([unc], enabled=True)
    assert batch.staged_files == 1
    assert batch.stage_dir is not None
    staged = Path(batch.io_paths[0])
    assert staged.is_file()
    assert staged.read_bytes() == b"hello"
    assert batch.display_paths == [unc]
    stage_dir = batch.stage_dir
    batch.cleanup()
    assert stage_dir is not None
    assert not stage_dir.is_dir()


def test_stage_parallel_copy(monkeypatch, tmp_path: Path) -> None:
    files = []
    for i in range(4):
        p = tmp_path / ("f%s.xlsx" % i)
        p.write_bytes(("x%s" % i).encode())
        files.append(str(p.resolve()))

    monkeypatch.setattr(
        "svc.data_agg_network_stage.path_is_network",
        lambda _p: True,
    )
    progress: list[tuple[int, int, str]] = []

    def _prog(done: int, total: int, name: str) -> None:
        progress.append((done, total, name))

    batch = build_network_stage_batch(
        files,
        enabled=True,
        copy_workers=2,
        progress_callback=_prog,
    )
    assert batch.staged_files == 4
    assert len(progress) == 4
    batch.cleanup()


def test_stage_cancel_during_copy(monkeypatch, tmp_path: Path) -> None:
    from svc.data_agg_cancel import DataAggCancelled
    from svc.data_agg_network_stage import _stage_temp_base

    p = tmp_path / "one.xlsx"
    p.write_bytes(b"z")
    unc = str(p.resolve())
    monkeypatch.setattr(
        "svc.data_agg_network_stage.path_is_network",
        lambda _p: True,
    )
    calls = [0]

    def _cancel(**_kw: object) -> None:
        calls[0] += 1
        if calls[0] >= 1:
            raise DataAggCancelled()

    base = _stage_temp_base()
    before = set(base.iterdir()) if base.is_dir() else set()
    with pytest.raises(DataAggCancelled):
        build_network_stage_batch([unc], enabled=True, cancel_check=_cancel)
    after = set(base.iterdir()) if base.is_dir() else set()
    assert after == before


def test_stage_removes_empty_dirs(monkeypatch, tmp_path: Path) -> None:
    src = tmp_path / "nested" / "deep" / "f.xlsx"
    src.parent.mkdir(parents=True)
    src.write_bytes(b"x")
    unc = str(src.resolve())
    monkeypatch.setattr(
        "svc.data_agg_network_stage.path_is_network",
        lambda _p: True,
    )
    batch = build_network_stage_batch([unc], scan_root=tmp_path, enabled=True)
    staged = Path(batch.io_paths[0])
    assert staged.is_file()
    # コピー後に空の中間ディレクトリが残らないこと
    for d in batch.stage_dir.rglob("*"):  # type: ignore[union-attr]
        if d.is_dir():
            assert any(d.iterdir())
    batch.cleanup()
