# -*- coding: utf-8 -*-
"""#18: 依存レイヤの宣言整合（lock と意図の乖離防止）。"""
from __future__ import annotations

from pathlib import Path

_ROOT = Path(__file__).resolve().parents[1]


def _pkgs(path: Path) -> set[str]:
    names: set[str] = set()
    for raw in path.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or line.startswith("-r "):
            continue
        name = line.split("==", 1)[0].split(">=", 1)[0].strip()
        if name:
            names.add(name.lower().replace("_", "-"))
    return names


def test_requirements_layer_files_exist() -> None:
    for name in (
        "requirements-runtime.txt",
        "requirements-optional.txt",
        "requirements-dev.txt",
        "requirements-build.txt",
        "requirements.lock.txt",
    ):
        assert (_ROOT / name).is_file(), name


def test_lock_includes_all_layers() -> None:
    text = (_ROOT / "requirements.lock.txt").read_text(encoding="utf-8")
    for layer in (
        "requirements-runtime.txt",
        "requirements-optional.txt",
        "requirements-dev.txt",
        "requirements-build.txt",
    ):
        assert f"-r {layer}" in text


def test_runtime_excludes_dev_and_dead_build_tools() -> None:
    runtime = _pkgs(_ROOT / "requirements-runtime.txt")
    forbidden = {
        "debugpy",
        "ruff",
        "pytest",
        "pyinstaller",
        "pyinstaller-hooks-contrib",
        "altgraph",
        "pefile",
        "pywin32-ctypes",
        "nuitka",
        "polars",
    }
    assert not (runtime & forbidden)


def test_optional_keeps_polars() -> None:
    optional = _pkgs(_ROOT / "requirements-optional.txt")
    assert "polars" in optional


def test_runtime_includes_bsdiff4_and_core_stack() -> None:
    runtime = _pkgs(_ROOT / "requirements-runtime.txt")
    for name in ("bsdiff4", "pandas", "numpy", "pyside6", "xlwings", "filelock"):
        assert name in runtime, name


def test_dev_has_debug_tools_not_in_runtime() -> None:
    dev = _pkgs(_ROOT / "requirements-dev.txt")
    runtime = _pkgs(_ROOT / "requirements-runtime.txt")
    assert "debugpy" in dev and "ruff" in dev
    assert not (dev & runtime)
