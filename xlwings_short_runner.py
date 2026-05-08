# -*- coding: utf-8 -*-
"""Short-lived xlwings entry: runs a UTF-8 script file (replaces python -c for Nuitka EXE).

Build with Nuitka alongside deployment; xlwings.conf INTERPRETER / INTERPRETER_WIN should point to this executable when USE_PACKAGED_RUNPYTHON is True.
"""
from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path


def main() -> int:
    """短寿命ランナー: 引数の UTF-8 スクリプトを exec する。配布時は ``app\\bin`` に同梱された DLL を EXE 横で解決。"""
    if os.name == "nt":
        from core.shared_dll_bootstrap import ensure_shared_dll_search_path_next_to_executable

        ensure_shared_dll_search_path_next_to_executable()

    parser = argparse.ArgumentParser()
    parser.add_argument("--script-file", required=True, help="UTF-8 Python snippet path")
    args, _unknown = parser.parse_known_args()

    script_path = Path(args.script_file)
    if not script_path.is_file():
        sys.stderr.write(f"xlwings_short_runner: missing script file: {script_path}\n")
        return 2

    src = script_path.read_text(encoding="utf-8-sig")

    root = (os.environ.get("HC_INSTALL_ROOT") or "").strip()
    if root:
        if root not in sys.path:
            sys.path.insert(0, root)
        try:
            os.chdir(root)
        except OSError:
            pass

    g: dict = {"__name__": "__main__", "__file__": str(script_path)}
    exec(compile(src, str(script_path), "exec"), g)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
