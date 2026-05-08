# -*- coding: utf-8 -*-
"""One-shot patch for spawn_ui_server in svc_host.py."""
from __future__ import annotations

from pathlib import Path


def main() -> None:
    root = Path(__file__).resolve().parent.parent
    p = root / "svc" / "svc_host.py"
    s = p.read_text(encoding="utf-8")
    old1 = (
        "    if not _is_expected_venv_interpreter(project_root):\n"
        "        try:\n"
        "            from core import core_log\n\n"
        "            core_log.get_logger(__name__).warning(\n"
        '                "skip spawn: unexpected interpreter: %s", sys.executable\n'
        "            )\n"
        "        except Exception:\n"
        "            pass\n"
        "        return\n\n"
        "    ipc_root = str(ipc_file.get_ipc_root())\n"
        '    logs_dir = Path(ipc_root) / "logs"\n'
        "    logs_dir.mkdir(parents=True, exist_ok=True)\n"
        '    boot_log = logs_dir / f"ui_server_boot_{int(time.time() * 1000)}.log"'
    )
    new1 = old1.replace("project_root", "dev_root", 1)
    if old1 not in s:
        raise SystemExit("block1 (ui expected venv) not found")
    s = s.replace(old1, new1, 1)

    marker = '[QT_UI_SERVER] spawned:'
    idx = s.find(marker)
    if idx < 0:
        raise SystemExit("QT_UI_SERVER log not found")
    # find start of exe = _project_pythonw before this logger line
    chunk_start = s.rfind("    exe = _project_pythonw(project_root)", 0, idx)
    if chunk_start < 0:
        raise SystemExit("exe = _project_pythonw block not found for ui")
    chunk_end = s.find("\n\n\ndef ensure_ui_server", chunk_start)
    if chunk_end < 0:
        raise SystemExit("ensure_ui_server boundary not found")
    new_chunk = """    if packaged:
        cmd = [str(ui_exe)]
        spawn_label = ui_exe
    else:
        exe = _project_pythonw(project_root)
        cmd = [exe, "-u", str(server_py)]
        spawn_label = server_py

    with boot_log.open("w", encoding="utf-8") as f:
        f.write(f"[BOOT] cmd={cmd}\n")
        f.write(f"[BOOT] cwd={project_root}\n")
        f.write(f"[BOOT] HC_QT_IPC_DIR={ipc_root}\n")

    popen_kw: dict = {
        "cwd": str(project_root),
        "env": env,
        "stdout": boot_log.open("a", encoding="utf-8"),
        "stderr": boot_log.open("a", encoding="utf-8"),
    }
    if os.name == "nt":
        popen_kw["creationflags"] = 0x08000000  # CREATE_NO_WINDOW
    subprocess.Popen(cmd, **popen_kw)  # noqa: S603,S607
    logger.info(
        "[QT_UI_SERVER] spawned: %s IPC=%s (boot_log=%s)", spawn_label, ipc_root, boot_log
    )
"""
    s = s[:chunk_start] + new_chunk + s[chunk_end:]
    p.write_text(s, encoding="utf-8", newline="\n")
    print("patched", p)


if __name__ == "__main__":
    main()
