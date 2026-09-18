# -*- coding: utf-8 -*-
"""リボンからの svc_server 起動。UI / bridge / 更新画面の spawn はここへ移さない。"""
from __future__ import annotations

import os
import subprocess
import sys
import time
from pathlib import Path

from core.core_log import get_logger
from core import runtime_layout
from ui_qt import ipc_file

logger = get_logger(__name__)

_SVC_MUTEX_NAME = "Global\\HC_SVC_SERVER"
_SYNCHRONIZE = 0x00100000
_MUTEX_WAIT_SEC = 8.0
_MUTEX_GRACE_SEC = 0.5
_STARTING_FLAG = "svc_server_starting.flag"
_STARTING_FLAG_TTL_SEC = 5.0


def project_root() -> Path:
    """core/ の親。svc_host.py の parent.parent と同じ。"""
    return Path(__file__).resolve().parent.parent


def svc_server_script() -> Path:
    return (project_root() / "svc" / "svc_server.py").resolve()


def _is_expected_venv_interpreter(root: Path) -> bool:
    if runtime_layout.packaged_spawn_requested():
        return True
    try:
        expected_dir = (root / ".venv" / "Scripts").resolve()
        if not expected_dir.exists():
            return True
        exe = Path(sys.executable).resolve()
        return exe.parent == expected_dir
    except Exception:
        return True


def _is_project_venv_interpreter(root: Path) -> bool:
    if getattr(sys, "frozen", False):
        return True
    if runtime_layout.packaged_spawn_requested():
        return True
    if os.name != "nt":
        return True
    venv_scripts = (root / ".venv" / "Scripts").resolve()

    def _r(p: str) -> Path:
        try:
            return Path(p).resolve()
        except Exception:
            return Path(p)

    exe = _r(getattr(sys, "executable", "") or "")
    argv0 = _r((sys.argv[0] if sys.argv else "") or "")
    if not venv_scripts.exists():
        return True
    exe_ok = exe.parent == venv_scripts and exe.name.lower() in {
        "python.exe",
        "pythonw.exe",
    }
    argv0_name = argv0.name.lower()
    if argv0_name in {"python.exe", "pythonw.exe", "-c"}:
        argv_ok = True
    else:
        try:
            argv_ok = argv0.suffix.lower() == ".py" and root.resolve() in argv0.parents
        except Exception:
            argv_ok = False
    return exe_ok and argv_ok


def _project_pythonw(root: Path) -> str:
    try:
        venv_pythonw = (root / ".venv" / "Scripts" / "pythonw.exe").resolve()
        if venv_pythonw.exists():
            return str(venv_pythonw)
        venv_python = (root / ".venv" / "Scripts" / "python.exe").resolve()
        if venv_python.exists():
            return str(venv_python)
    except Exception:
        pass
    try:
        pythonw = Path(sys.executable).with_name("pythonw.exe")
        if pythonw.exists():
            return str(pythonw)
    except Exception:
        pass
    return sys.executable


def _svc_mutex_exists() -> bool:
    if os.name != "nt":
        return False
    try:
        import ctypes
        from ctypes import wintypes

        kernel32 = ctypes.windll.kernel32  # type: ignore[attr-defined]
        handle = kernel32.OpenMutexW(_SYNCHRONIZE, False, wintypes.LPCWSTR(_SVC_MUTEX_NAME))
        if handle:
            kernel32.CloseHandle(handle)
            return True
        return False
    except Exception:
        return False


def is_svc_server_running() -> bool:
    return _svc_mutex_exists()


def _clear_shutdown_flags(reason: str) -> None:
    from core.core_host_ipc import host_control_dir

    d = host_control_dir()
    removed: list[str] = []
    for name in ("shutdown.flag", "svc_shutdown.flag"):
        p = d / name
        try:
            if p.exists():
                p.unlink()
                removed.append(name)
        except OSError:
            pass
    if removed:
        logger.info("[CONTROL] cleared flags=%s reason=%s", ",".join(removed), reason)


def _wait_until_running(poll_sec: float) -> None:
    t0 = time.time()
    while time.time() - t0 < _MUTEX_WAIT_SEC:
        if is_svc_server_running():
            return
        time.sleep(poll_sec)
    time.sleep(_MUTEX_GRACE_SEC)
    if is_svc_server_running():
        logger.info("[SVC_SERVER] mutex observed after grace wait")
        return
    logger.warning("[SVC_SERVER] spawn requested but mutex not observed yet")


def spawn_svc_server() -> None:
    """svc_server を起動する（未起動想定）。UI / bridge は起動しない。"""
    try:
        bn = os.path.basename(sys.argv[0]).lower()
        if bn in ("svc_server.py", "svc_server.exe", "hc_svc_server.exe"):
            logger.info("[HOST] skip spawn: already inside svc_server process")
            return
    except Exception:
        pass

    root = project_root()
    packaged = runtime_layout.use_packaged_server_commands()
    server_py: Path | None = None
    svc_exe: Path | None = None

    if packaged:
        svc_exe = runtime_layout.packaged_app_exe("hc_svc_server.exe")
        if svc_exe is None:
            logger.warning(
                "[SVC_SERVER] packaged hc_svc_server.exe not found under HC_INSTALL_ROOT/app/bin",
            )
            return
        proj = runtime_layout.runtime_project_root(str(Path(__file__).resolve()))
    else:
        server_py = svc_server_script()
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

    ipc_root = str(ipc_file.get_ipc_root())
    logs_dir = Path(ipc_root) / "logs"
    logs_dir.mkdir(parents=True, exist_ok=True)
    boot_log = logs_dir / ("svc_server_boot_%s.log" % int(time.time() * 1000))

    env = os.environ.copy()
    env["HC_IPC_ROOT"] = ipc_root
    env["HC_QT_IPC_DIR"] = ipc_root
    env["HC_PROJECT_ROOT"] = str(proj)
    env["PYTHONPATH"] = str(proj) + os.pathsep + env.get("PYTHONPATH", "")

    if packaged:
        ir = runtime_layout.install_root()
        if ir is not None:
            env = runtime_layout.env_with_packaged_dll_search_path(env, ir)
        cmd = [str(svc_exe)]
        spawn_label = svc_exe
    else:
        exe = _project_pythonw(proj)
        cmd = [exe, "-u", str(server_py)]
        spawn_label = server_py

    with boot_log.open("w", encoding="utf-8") as f:
        f.write("[BOOT] cmd=%s\n" % cmd)
        f.write("[BOOT] cwd=%s\n" % proj)
        f.write("[BOOT] HC_QT_IPC_DIR=%s\n" % ipc_root)

    popen_kw: dict = {
        "cwd": str(proj),
        "env": env,
        "stdout": boot_log.open("a", encoding="utf-8"),
        "stderr": boot_log.open("a", encoding="utf-8"),
    }
    if os.name == "nt":
        popen_kw["creationflags"] = 0x08000000
    subprocess.Popen(cmd, **popen_kw)  # noqa: S603,S607
    logger.info(
        "[SVC_SERVER] spawned: %s IPC=%s (boot_log=%s)", spawn_label, ipc_root, boot_log
    )


def ensure_svc_server() -> None:
    """起動済みなら何もしない。未起動なら svc_server だけ起動する。"""
    _clear_shutdown_flags("ensure_svc_server")
    if is_svc_server_running():
        logger.info("[SVC_SERVER] already running (mutex exists)")
        return

    ipc_root = Path(str(ipc_file.get_ipc_root()))
    flag = ipc_root / "control" / _STARTING_FLAG
    flag.parent.mkdir(parents=True, exist_ok=True)
    try:
        if flag.exists() and (time.time() - flag.stat().st_mtime) < _STARTING_FLAG_TTL_SEC:
            logger.info("[SVC_SERVER] startup in progress (flag exists); skip spawn")
        else:
            flag.write_text(str(int(time.time() * 1000)), encoding="utf-8")
            spawn_svc_server()
        _wait_until_running(0.05)
    finally:
        try:
            if flag.exists():
                flag.unlink()
        except OSError:
            pass
