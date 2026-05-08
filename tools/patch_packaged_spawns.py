# -*- coding: utf-8 -*-
"""Replace spawn_svc_server and spawn_bridge in svc_host.py (ASCII-only script).

正本は ``svc/svc_host.py``。子プロセス env（packaged 時の PATH 補強など）を変えたら、
埋め込みブロックを手で同期するか、このスクリプトを更新してから実行する。
（``spawn_ui_server`` はここでは置換しない）
"""
from __future__ import annotations

from pathlib import Path

SVC_NEW = r'''def spawn_svc_server() -> None:
    """svc_server を起動する（未起動想定）。"""

    # --- v0.4.16 in-process spawn guard ---
    # Prevent svc_server from spawning itself
    try:
        bn = os.path.basename(sys.argv[0]).lower()
        if bn in ("svc_server.py", "svc_server.exe", "hc_svc_server.exe"):
            logger.info("[HOST] skip spawn: already inside svc_server process")
            return
    except Exception:
        pass
    # --------------------------------------
    dev_root = Path(__file__).resolve().parent.parent
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
        project_root = runtime_layout.runtime_project_root(str(Path(__file__).resolve()))
    else:
        server_py = _resolve_svc_server_path()
        if not server_py.exists():
            raise FileNotFoundError(str(server_py))
        project_root = server_py.parent.parent

    if not _is_project_venv_interpreter(dev_root):
        logger.warning(
            "[HOST] skip spawn: interpreter is not project venv: %s",
            sys.executable,
        )
        return
    # guard: system python からの誤起動（二重起動）を抑止
    if not _is_expected_venv_interpreter(dev_root):
        try:
            from core import core_log

            core_log.get_logger(__name__).warning(
                "skip spawn: unexpected interpreter: %s", sys.executable
            )
        except Exception:
            pass
        return

    ipc_root = str(ipc_file.get_ipc_root())
    logs_dir = Path(ipc_root) / "logs"
    logs_dir.mkdir(parents=True, exist_ok=True)
    boot_log = logs_dir / ("svc_server_boot_%s.log" % int(time.time() * 1000))

    env = os.environ.copy()
    env["HC_IPC_ROOT"] = ipc_root
    env["HC_QT_IPC_DIR"] = ipc_root
    env["HC_PROJECT_ROOT"] = str(project_root)
    env["PYTHONPATH"] = str(project_root) + os.pathsep + env.get("PYTHONPATH", "")

    if packaged:
        ir = runtime_layout.install_root()
        if ir is not None:
            env = runtime_layout.env_with_packaged_dll_search_path(env, ir)

    if packaged:
        cmd = [str(svc_exe)]
        spawn_label = svc_exe
    else:
        exe = _project_pythonw(project_root)
        cmd = [exe, "-u", str(server_py)]
        spawn_label = server_py

    with boot_log.open("w", encoding="utf-8") as f:
        f.write("[BOOT] cmd=%s\n" % cmd)
        f.write("[BOOT] cwd=%s\n" % project_root)
        f.write("[BOOT] HC_QT_IPC_DIR=%s\n" % ipc_root)

    popen_kw = {
        "cwd": str(project_root),
        "env": env,
        "stdout": boot_log.open("a", encoding="utf-8"),
        "stderr": boot_log.open("a", encoding="utf-8"),
    }
    if os.name == "nt":
        popen_kw["creationflags"] = 0x08000000  # CREATE_NO_WINDOW
    subprocess.Popen(cmd, **popen_kw)  # noqa: S603,S607
    logger.info(
        "[SVC_SERVER] spawned: %s IPC=%s (boot_log=%s)", spawn_label, ipc_root, boot_log
    )


'''

BRIDGE_NEW = r'''def spawn_bridge() -> None:
    """Bridge resident process (hc_main). load_csv etc. without RunPython."""
    try:
        bn = os.path.basename(sys.argv[0]).lower()
        if bn == "hc_main.py" or bn == "hc_main.exe":
            logger.info("[HOST] skip spawn: already inside bridge process")
            return
    except Exception:
        pass
    dev_root = Path(__file__).resolve().parent.parent
    packaged = runtime_layout.use_packaged_server_commands()
    bridge_py: Path | None = None
    bridge_exe: Path | None = None

    if packaged:
        bridge_exe = runtime_layout.packaged_app_exe("hc_main.exe")
        if bridge_exe is None:
            logger.warning(
                "%s packaged hc_main.exe not found under HC_INSTALL_ROOT/app/bin",
                LOG_MAIN_PREFIX,
            )
            return
        project_root = runtime_layout.runtime_project_root(str(Path(__file__).resolve()))
    else:
        bridge_py = _resolve_bridge_path()
        if not bridge_py.exists():
            logger.warning("%s bridge script not found: %s", LOG_MAIN_PREFIX, bridge_py)
            return
        project_root = _bridge_script_project_root(bridge_py)

    if not _is_project_venv_interpreter(dev_root):
        logger.warning("[HOST] skip bridge spawn: interpreter is not project venv: %s", sys.executable)
        return
    if not _is_expected_venv_interpreter(dev_root):
        return
    ipc_root = str(ipc_file.get_ipc_root())
    logs_dir = Path(ipc_root) / "logs"
    logs_dir.mkdir(parents=True, exist_ok=True)
    boot_log = logs_dir / ("hc_main_boot_%s.log" % int(time.time() * 1000))
    env = os.environ.copy()
    env["HC_IPC_ROOT"] = ipc_root
    env["HC_QT_IPC_DIR"] = ipc_root
    env["HC_PROJECT_ROOT"] = str(project_root)
    env["PYTHONPATH"] = str(project_root) + os.pathsep + env.get("PYTHONPATH", "")

    if packaged:
        ir = runtime_layout.install_root()
        if ir is not None:
            env = runtime_layout.env_with_packaged_dll_search_path(env, ir)

    if packaged:
        cmd = [str(bridge_exe)]
        spawn_label = bridge_exe
    else:
        exe = _project_pythonw(project_root)
        cmd = [exe, "-u", str(bridge_py)]
        spawn_label = bridge_py
    with boot_log.open("w", encoding="utf-8") as f:
        f.write("[BOOT] cmd=%s\n" % cmd)
        f.write("[BOOT] cwd=%s\n" % project_root)
    popen_kw = {
        "cwd": str(project_root),
        "env": env,
        "stdout": boot_log.open("a", encoding="utf-8"),
        "stderr": boot_log.open("a", encoding="utf-8"),
    }
    if os.name == "nt":
        popen_kw["creationflags"] = 0x08000000
    subprocess.Popen(cmd, **popen_kw)  # noqa: S603,S607
    logger.info("%s spawned: %s IPC=%s", LOG_MAIN_PREFIX, spawn_label, ipc_root)


'''


def _replace_between(text: str, start: str, end: str, new_block: str) -> str:
    i = text.find(start)
    if i < 0:
        raise SystemExit(f"start not found: {start!r}")
    j = text.find(end, i + len(start))
    if j < 0:
        raise SystemExit(f"end not found after {start!r}")
    return text[:i] + new_block + text[j:]


def main() -> None:
    root = Path(__file__).resolve().parent.parent
    p = root / "svc" / "svc_host.py"
    s = p.read_text(encoding="utf-8")
    s = _replace_between(s, "def spawn_svc_server() -> None:", "def _resolve_bridge_path() -> Path:", SVC_NEW)
    s = _replace_between(s, "def spawn_bridge() -> None:", "def ensure_bridge() -> None:", BRIDGE_NEW)
    p.write_text(s, encoding="utf-8", newline="\n")
    print("patched", p)


if __name__ == "__main__":
    main()
