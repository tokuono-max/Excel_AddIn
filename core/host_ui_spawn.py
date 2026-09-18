# -*- coding: utf-8 -*-
"""UI サーバの起動。更新画面もここを呼ぶ。spawn コマンドと Nuitka は変えない。"""
from __future__ import annotations

import os
import subprocess
import sys
import time
from pathlib import Path

from core import runtime_layout
from core.core_log import get_logger
from core.host_svc_spawn import (
    _clear_shutdown_flags,
    _is_expected_venv_interpreter,
    _is_project_venv_interpreter,
    _project_pythonw,
    project_root,
)
from ui_qt import ipc_file

logger = get_logger(__name__)

_UI_MUTEX_NAME = "Global\\HC_QT_UI_SERVER"
_SYNCHRONIZE = 0x00100000
_MUTEX_WAIT_SEC = 8.0
_MUTEX_GRACE_SEC = 0.5
_STARTING_FLAG = "ui_server_starting.flag"
_STARTING_FLAG_TTL_SEC = 5.0


def ui_server_script() -> Path:
    return (project_root() / "ui_qt" / "ui_server.py").resolve()


def _ui_mutex_exists() -> bool:
    if os.name != "nt":
        return False
    try:
        import ctypes
        from ctypes import wintypes

        kernel32 = ctypes.windll.kernel32  # type: ignore[attr-defined]
        handle = kernel32.OpenMutexW(_SYNCHRONIZE, False, wintypes.LPCWSTR(_UI_MUTEX_NAME))
        if handle:
            kernel32.CloseHandle(handle)
            return True
        return False
    except Exception:
        return False


def is_ui_server_running() -> bool:
    return _ui_mutex_exists()


def _wait_until_running(poll_sec: float) -> None:
    t0 = time.time()
    while time.time() - t0 < _MUTEX_WAIT_SEC:
        if is_ui_server_running():
            return
        time.sleep(poll_sec)
    time.sleep(_MUTEX_GRACE_SEC)
    if is_ui_server_running():
        logger.info("[QT_UI_SERVER] mutex observed after grace wait")
        return
    logger.warning("[QT_UI_SERVER] spawn requested but mutex not observed yet")


def spawn_ui_server() -> None:
    """Qt UI サーバを起動する（未起動想定）。"""
    root = project_root()
    packaged = runtime_layout.use_packaged_server_commands()
    server_py: Path | None = None
    ui_exe: Path | None = None

    if packaged:
        ui_exe = runtime_layout.packaged_app_exe("hc_ui_server.exe")
        if ui_exe is None:
            logger.warning(
                "[QT_UI_SERVER] packaged hc_ui_server.exe not found under HC_INSTALL_ROOT/app/bin",
            )
            return
        proj = runtime_layout.runtime_project_root(str(Path(__file__).resolve()))
    else:
        server_py = ui_server_script()
        if not server_py.exists():
            raise FileNotFoundError(str(server_py))
        proj = server_py.parent.parent

    if not _is_project_venv_interpreter(root):
        logger.warning(
            "[HOST] skip spawn: interpreter is not project venv: %s",
            sys.executable,
        )
        return
    if not _is_expected_venv_interpreter(root):
        logger.warning("skip spawn: unexpected interpreter: %s", sys.executable)
        return

    import sys

    ipc_root = str(ipc_file.get_ipc_root())
    logs_dir = Path(ipc_root) / "logs"
    logs_dir.mkdir(parents=True, exist_ok=True)
    boot_log = logs_dir / ("ui_server_boot_%s.log" % int(time.time() * 1000))

    env = os.environ.copy()
    env["HC_IPC_ROOT"] = ipc_root
    env["HC_QT_IPC_DIR"] = ipc_root
    env["HC_PROJECT_ROOT"] = str(proj)
    env["PYTHONPATH"] = str(proj) + os.pathsep + env.get("PYTHONPATH", "")
    env["HC_UI_PARENT_PID"] = str(os.getpid())
    try:
        env["HC_EXCEL_PID"] = str(os.getppid())
    except Exception:
        pass

    if packaged:
        ir = runtime_layout.install_root()
        if ir is not None:
            env = runtime_layout.env_with_packaged_dll_search_path(env, ir)
        if ui_exe is None:
            return
        cmd = [str(ui_exe)]
        spawn_label = ui_exe
        ui_cwd = str(ui_exe.resolve().parent)
    else:
        exe = _project_pythonw(proj)
        cmd = [exe, "-u", str(server_py)]
        spawn_label = server_py
        ui_cwd = str(proj)

    with boot_log.open("w", encoding="utf-8") as f:
        f.write("[BOOT] cmd=%s\n" % cmd)
        f.write("[BOOT] cwd=%s\n" % ui_cwd)
        f.write("[BOOT] HC_PROJECT_ROOT=%s\n" % proj)
        f.write("[BOOT] HC_QT_IPC_DIR=%s\n" % ipc_root)

    popen_kw: dict = {
        "cwd": ui_cwd,
        "env": env,
        "stdout": boot_log.open("a", encoding="utf-8"),
        "stderr": boot_log.open("a", encoding="utf-8"),
    }
    if os.name == "nt":
        popen_kw["creationflags"] = 0x08000000
    subprocess.Popen(cmd, **popen_kw)  # noqa: S603,S607
    logger.info(
        "[QT_UI_SERVER] spawned: %s IPC=%s (boot_log=%s)", spawn_label, ipc_root, boot_log
    )


def ensure_ui_server() -> None:
    """起動済みなら何もしない。未起動なら UI サーバだけ起動する。"""
    _clear_shutdown_flags("ensure_ui_server")
    if is_ui_server_running():
        logger.info("[QT_UI_SERVER] already running (mutex exists)")
        return

    ipc_root = Path(str(ipc_file.get_ipc_root()))
    flag = ipc_root / "control" / _STARTING_FLAG
    flag.parent.mkdir(parents=True, exist_ok=True)
    try:
        if flag.exists() and (time.time() - flag.stat().st_mtime) < _STARTING_FLAG_TTL_SEC:
            logger.info("[QT_UI_SERVER] startup in progress (flag exists); skip spawn")
        else:
            flag.write_text(str(int(time.time() * 1000)), encoding="utf-8")
            spawn_ui_server()
        _wait_until_running(0.02)
    finally:
        try:
            if flag.exists():
                flag.unlink()
        except Exception:
            pass
