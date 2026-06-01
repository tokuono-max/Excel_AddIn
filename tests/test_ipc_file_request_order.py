# -*- coding: utf-8 -*-
from __future__ import annotations

import sys
import time
from pathlib import Path

import pytest

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from ui_qt import ipc_file  # noqa: E402


def test_safe_mtime_for_sort_handles_oserror(monkeypatch: pytest.MonkeyPatch) -> None:
    class _BrokenPath:
        def stat(self) -> object:
            raise OSError("stat failed")

    v = ipc_file._safe_mtime_for_sort(_BrokenPath())  # pyright: ignore[reportArgumentType]
    assert v == float("inf")


def test_pop_next_request_survives_stat_race(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    req_ok = tmp_path / "req_ok.pkl"
    req_bad = tmp_path / "req_bad.pkl"
    req_ok.write_bytes(b"ok")
    req_bad.write_bytes(b"bad")

    monkeypatch.setattr(ipc_file, "get_request_dir", lambda: tmp_path)

    orig = Path.stat

    def _stat_with_race(self: Path) -> object:
        if self.name == "req_bad.pkl":
            raise OSError("simulated race")
        return orig(self)

    monkeypatch.setattr(Path, "stat", _stat_with_race, raising=True)
    claimed = ipc_file.pop_next_request()
    assert claimed is not None
    assert claimed.name == "req_ok.work.pkl"


def test_read_pickle_retries_permission_error(monkeypatch: pytest.MonkeyPatch) -> None:
    payload = {"k": "v"}
    raw = ipc_file.pickle.dumps(payload, protocol=ipc_file.pickle.HIGHEST_PROTOCOL)
    calls = {"n": 0}

    class _PathLike:
        def read_bytes(self) -> bytes:
            calls["n"] += 1
            if calls["n"] <= 3:
                raise PermissionError("locked")
            return raw

    monkeypatch.setattr(ipc_file.time, "sleep", lambda _: None)
    got = ipc_file.read_pickle(_PathLike())  # pyright: ignore[reportArgumentType]
    assert got == payload
    assert calls["n"] == 4


def test_read_pickle_raises_after_retry_exhausted(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    class _PathLike:
        def read_bytes(self) -> bytes:
            raise PermissionError("always locked")

    monkeypatch.setattr(ipc_file.time, "sleep", lambda _: None)
    with pytest.raises(PermissionError):
        ipc_file.read_pickle(_PathLike())  # pyright: ignore[reportArgumentType]


def test_cleanup_failed_requests_removes_only_expired(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    req_dir = tmp_path / "requests"
    failed_dir = req_dir / "_failed"
    failed_dir.mkdir(parents=True, exist_ok=True)
    old_p = failed_dir / "old.bad.pkl"
    new_p = failed_dir / "new.bad.pkl"
    old_p.write_bytes(b"old")
    new_p.write_bytes(b"new")
    now = time.time()
    old_ts = now - 120
    new_ts = now - 5
    old_p.touch()
    new_p.touch()
    old_p_stat_ns = int(old_ts * 1_000_000_000)
    new_p_stat_ns = int(new_ts * 1_000_000_000)
    old_p.chmod(0o666)
    new_p.chmod(0o666)
    old_p.unlink(missing_ok=False)
    old_p.write_bytes(b"old")
    new_p.unlink(missing_ok=False)
    new_p.write_bytes(b"new")
    import os

    os.utime(old_p, ns=(old_p_stat_ns, old_p_stat_ns))
    os.utime(new_p, ns=(new_p_stat_ns, new_p_stat_ns))

    monkeypatch.setattr(ipc_file, "get_request_dir", lambda: req_dir)
    n = ipc_file.cleanup_failed_requests(ttl_sec=60, max_remove=10)
    assert n == 1
    assert not old_p.exists()
    assert new_p.exists()


def test_cleanup_failed_requests_respects_max_remove(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    req_dir = tmp_path / "requests"
    failed_dir = req_dir / "_failed"
    failed_dir.mkdir(parents=True, exist_ok=True)
    paths = []
    base = int(time.time()) - 300
    import os

    for i in range(3):
        p = failed_dir / f"f{i}.bad.pkl"
        p.write_bytes(b"x")
        ts_ns = int((base + i) * 1_000_000_000)
        os.utime(p, ns=(ts_ns, ts_ns))
        paths.append(p)

    monkeypatch.setattr(ipc_file, "get_request_dir", lambda: req_dir)
    n = ipc_file.cleanup_failed_requests(ttl_sec=60, max_remove=2)
    assert n == 2
    remain = [p for p in paths if p.exists()]
    assert len(remain) == 1


def test_cap_failed_requests_keeps_newest(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    req_dir = tmp_path / "requests"
    failed_dir = req_dir / "_failed"
    failed_dir.mkdir(parents=True, exist_ok=True)
    import os

    paths = []
    base = int(time.time()) - 300
    for i in range(4):
        p = failed_dir / f"f{i}.bad.pkl"
        p.write_bytes(b"x")
        ts_ns = int((base + i) * 1_000_000_000)
        os.utime(p, ns=(ts_ns, ts_ns))
        paths.append(p)
    monkeypatch.setattr(ipc_file, "get_request_dir", lambda: req_dir)
    n = ipc_file.cap_failed_requests(max_keep=2)
    assert n == 2
    remain = sorted([p.name for p in failed_dir.glob("*.bad.pkl")])
    assert remain == ["f2.bad.pkl", "f3.bad.pkl"]
