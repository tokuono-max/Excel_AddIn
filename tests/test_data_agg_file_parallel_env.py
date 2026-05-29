# -*- coding: utf-8 -*-
from __future__ import annotations

import os

from core import core_env


def test_file_parallel_workers_auto(monkeypatch) -> None:
    monkeypatch.delenv("DATA_AGG_FILE_PARALLEL_WORKERS", raising=False)
    assert core_env.data_agg_file_parallel_workers(n_files=1) == 0
    w = core_env.data_agg_file_parallel_workers(n_files=19)
    assert 1 <= w <= 4


def test_file_parallel_workers_zero(monkeypatch) -> None:
    monkeypatch.setenv("DATA_AGG_FILE_PARALLEL_WORKERS", "0")
    assert core_env.data_agg_file_parallel_workers(n_files=19) == 0


def test_file_parallel_workers_explicit(monkeypatch) -> None:
    monkeypatch.setenv("DATA_AGG_FILE_PARALLEL_WORKERS", "2")
    assert core_env.data_agg_file_parallel_workers(n_files=19) == 2
