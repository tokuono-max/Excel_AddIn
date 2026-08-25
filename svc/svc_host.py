# -*- coding: utf-8 -*-
"""
Python: 3.12
Module: svc/svc_host.py
Created: 2026-02-11
Updated: 2026-06-26
Version: 0.4.42
Purpose:
  UI Host（common foundation）。
  - Qt UI Server 起動・生存判定・終了要求を 1か所に集約する。
  - 判定は Windows named mutex(OpenMutex) により「存在確認のみ」（所有しない）。
  - Excel終了時は request_shutdown_all() / shutdown_all_with_force_kill() を呼ぶ想定。
  - ブリッジ常駐（ensure_bridge）で load_csv を RunPython なしで受け付け、待ち時間短縮。

History (latest 3):
  - 0.4.42 (2026-06-26): 常駐ホスト生存時のリボン fast path（spawn/prewarm/register_book COM 省略）。
  - 0.4.41 (2026-06-16): restart_svc_server ログを recovery restart に変更（救済用である旨を明示）。
  - 0.4.40 (2026-06-14): B+ — 常駐 svc_server を維持。事前 COM 再起動を廃止（汚染時のみ recycle）。
"""

  # bootstrap: allow direct script execution
# ==============================================================================
import os as _os
import sys as _sys


def _bootstrap_sys_path() -> None:
    here = _os.path.abspath(_os.path.dirname(__file__))  # .../svc
    root = _os.path.normpath(_os.path.join(here, ".."))  # project root
    if root not in _sys.path:
        _sys.path.insert(0, root)


_bootstrap_sys_path()
# ruff: noqa: E402
# ==============================================================================
__version__ = "0.4.42"
import os
import shlex
import subprocess
import ctypes
from ctypes import wintypes
import sys
import time
import tempfile
from pathlib import Path
from typing import Callable

from core import runtime_layout


def _is_expected_venv_interpreter(project_root: Path) -> bool:
    """このプロジェクトの .venv 配下の Python で実行されているか判定する。

    目的:
      - system python からの誤起動（=二重起動）を抑止する。
      - venv が存在しない開発環境では判定をスキップ（許容）する。
    """
    if runtime_layout.packaged_spawn_requested():
        return True
    try:
        expected_dir = (project_root / ".venv" / "Scripts").resolve()
        if not expected_dir.exists():
            return True  # venv 構成でない環境は許容
        exe = Path(sys.executable).resolve()
        return exe.parent == expected_dir
    except Exception:
        return True


from core.core_env import LOG_MAIN_PREFIX
from core.core_log import get_logger, get_perf_logger
from ui_qt import ipc_file

logger = get_logger(__name__)

# spawn 後、子プロセスが named mutex を取るまでの待ち（秒）。初回 import が遅い環境向けに余裕を持たせる。
# Nuitka の hc_ui_server.exe 初回起動は PySide DLL 読み込みで 5s を超えることがある
_MUTEX_WAIT_SEC: float = 8.0
# 上記ループ終了直後の追加待ち＋再判定（誤 WARNING 低減）
_MUTEX_GRACE_SEC: float = 0.5


def _wait_until_running(
    is_running: Callable[[], bool],
    log_label: str,
    *,
    max_wait_sec: float = _MUTEX_WAIT_SEC,
    poll_sec: float = 0.02,
) -> None:
    """子プロセス起動後、mutex 相当の生存判定が真になるまで待つ。"""
    t0 = time.time()
    while time.time() - t0 < max_wait_sec:
        if is_running():
            return
        time.sleep(poll_sec)
    time.sleep(_MUTEX_GRACE_SEC)
    if is_running():
        logger.info("%s mutex observed after grace wait", log_label)
        return
    logger.warning("%s spawn requested but mutex not observed yet", log_label)


def _log_startup_ui_gate_skip(perf_prefix: str, hwnd: int, reason: str) -> None:
    try:
        from core.packaged_update import _append_update_log, _install_root

        lr = _install_root()
        if lr and lr.is_dir():
            _append_update_log(
                lr,
                "startup: skip_duplicate_update_ui prefix={p} reason={r} hwnd={h}".format(
                    p=perf_prefix,
                    r=reason or "-",
                    h=hwnd,
                ),
            )
    except Exception:
        pass


def _excel_startup_svc_ui_bridge_register(target_hwnd: int, perf_prefix: str) -> None:
    """ensure_svc / ensure_ui / ensure_bridge / register_book と perf ログ（プレフィックスのみ異なる）。"""
    from core.excel_session import register_book as _register_book
    from core.packaged_update import maybe_apply_pending_bootstrap_update
    from core.startup_session_gate import excel_startup_ui_gate

    plog = get_perf_logger(f"{__name__}.excel_startup")
    t0 = time.perf_counter()
    hwnd = int(target_hwnd or 0)
    plog.info("%s phase=enter cumulative_ms=0 hwnd=%s", perf_prefix, hwnd)
    with excel_startup_ui_gate(hwnd, perf_prefix) as ui_gate:
        if ui_gate.skip_update_ui:
            plog.info(
                "%s phase=skip_duplicate_startup_ui reason=%s hwnd=%s",
                perf_prefix,
                ui_gate.reason or "-",
                hwnd,
            )
            _log_startup_ui_gate_skip(perf_prefix, hwnd, ui_gate.reason)
        else:
            maybe_apply_pending_bootstrap_update(owner_hwnd=hwnd, sheet_id="_")
            plog.info(
                "%s phase=after_apply_pending_bootstrap cumulative_ms=%d",
                perf_prefix,
                int((time.perf_counter() - t0) * 1000),
            )
        ensure_svc_ui_bridge_parallel()
        plog.info(
            "%s phase=after_ensure_svc_ui_bridge cumulative_ms=%d",
            perf_prefix,
            int((time.perf_counter() - t0) * 1000),
        )
        _register_book(target_hwnd=target_hwnd)
        if hwnd > 0:
            from core.excel_book_register_gate import mark_excel_book_registered

            mark_excel_book_registered(hwnd)
        plog.info(
            "%s phase=after_register_book cumulative_ms=%d",
            perf_prefix,
            int((time.perf_counter() - t0) * 1000),
        )
        if ui_gate.skip_update_ui:
            plog.info(
                "%s phase=skip_duplicate_version_check reason=%s hwnd=%s",
                perf_prefix,
                ui_gate.reason or "-",
                hwnd,
            )
        else:
            try:
                from core.packaged_update import maybe_check_updates_on_startup

                maybe_check_updates_on_startup(owner_hwnd=hwnd, sheet_id="_")
            except Exception as e:
                logger.warning("%s packaged update check skipped: %s", perf_prefix, e)


def excel_startup_workbook_open_full(target_hwnd: int) -> None:
    """Workbook_Open 用: 1 回の RunPython で svc / ui / bridge / register_book まで完了する。"""
    _excel_startup_svc_ui_bridge_register(target_hwnd, "startup_full")


def excel_startup_workbook_open_warmup() -> None:
    """後方互換: svc/ui のみ。新規は excel_startup_workbook_open_full を使用。"""
    plog = get_perf_logger(f"{__name__}.excel_startup")
    t0 = time.perf_counter()
    plog.info("warmup phase=enter cumulative_ms=0")
    ensure_svc_server()
    plog.info(
        "warmup phase=after_ensure_svc_server cumulative_ms=%d",
        int((time.perf_counter() - t0) * 1000),
    )
    ensure_ui_server()
    plog.info(
        "warmup phase=after_ensure_ui_server cumulative_ms=%d",
        int((time.perf_counter() - t0) * 1000),
    )


def excel_startup_after_excel_idle(target_hwnd: int) -> None:
    """VBA InitPythonServer（フォールバック）・Manual_Init 用。full と同一処理、perf ラベルのみ init_bridge。"""
    _excel_startup_svc_ui_bridge_register(target_hwnd, "init_bridge")


def _is_project_venv_interpreter(project_root: Path) -> bool:
    r"""True if current interpreter is the project's .venv Python (Windows).

    Notes
    - When Python is invoked as `python -c ...`, `sys.argv[0]` may be `-c` (or a
      pseudo-path ending with `\-c`). We must treat that as valid.
    - When Python is invoked to run a script (e.g. `...pythonw.exe svc_server.py`),
      `sys.argv[0]` is the script path. That must be accepted if the script is
      under the project root.
    - We still require `sys.executable` to be the project venv interpreter under
      `.venv\Scripts` to avoid spawning servers from a base/system interpreter.
    """
    if getattr(sys, "frozen", False):
        return True
    if runtime_layout.packaged_spawn_requested():
        return True
    if os.name != "nt":
        return True

    venv_scripts = (project_root / ".venv" / "Scripts").resolve()

    def _r(p: str) -> Path:
        try:
            return Path(p).resolve()
        except Exception:
            return Path(p)

    exe = _r(getattr(sys, "executable", "") or "")
    argv0 = _r((sys.argv[0] if sys.argv else "") or "")
    base = _r(getattr(sys, "_base_executable", "") or "")
    launcher = os.environ.get("__PYVENV_LAUNCHER__", "")

    try:
        logger.info(
            "[HOST_ENV] pid=%s ppid=%s exe=%s argv0=%s base=%s launcher=%s",
            os.getpid(),
            os.getppid(),
            exe,
            argv0,
            base,
            launcher,
        )
    except Exception:
        pass

    if not venv_scripts.exists():
        return True

    exe_ok = exe.parent == venv_scripts and exe.name.lower() in {
        "python.exe",
        "pythonw.exe",
    }

    argv0_name = argv0.name.lower()

    # 1) direct interpreter invocation or -c
    if argv0_name in {"python.exe", "pythonw.exe", "-c"}:
        argv_ok = True
    else:
        # 2) script path under project root (e.g., svc_server.py / ui_server.py)
        try:
            argv_ok = (
                argv0.suffix.lower() == ".py"
                and project_root.resolve() in argv0.parents
            )
        except Exception:
            argv_ok = False

    return exe_ok and argv_ok


# module load log (version)
try:
    logger.info(
        "[MODULE_LOAD] %s version=%s pid=%s file=%s",
        __name__,
        __version__,
        os.getpid(),
        __file__,
    )
except Exception:
    pass

# Permanent import diagnostics
try:
    logger.info(
        "[MODULE] name=%s version=%s file=%s",
        __name__,
        __version__,
        Path(__file__).resolve(),
    )
except Exception:
    pass

import threading

# Excel 終了などで svc 側の処理を止める共通フラグ
_STOP_EVENT = threading.Event()


# ------------------------------------------------------------------------------
# IPC root (best-effort)
# ------------------------------------------------------------------------------
try:
    ipc_root = Path(ipc_file.get_ipc_root())
except Exception:  # pragma: no cover
    ipc_root = Path(tempfile.gettempdir()) / "csv_tool"


def _control_dir() -> Path:
    d = ipc_root / "control"
    d.mkdir(parents=True, exist_ok=True)
    return d


def clear_shutdown_flags(reason: str = "") -> None:
    """残留している shutdown フラグを削除する（次回起動の誤判定を防ぐ）。"""
    d = _control_dir()
    removed = []
    for name in ("shutdown.flag", "svc_shutdown.flag"):
        p = d / name
        try:
            if p.exists():
                p.unlink()
                removed.append(name)
        except Exception:
            pass
    if removed:
        logger.info("[CONTROL] cleared flags=%s reason=%s", ",".join(removed), reason)


_MUTEX_NAME = "Global\\HC_QT_UI_SERVER"
_SVC_MUTEX_NAME = "Global\\HC_SVC_SERVER"
# 常駐 hc_main（新）。移行期間は旧名も is_* で検知する。
_MAIN_RUNNER_MUTEX_NAME = "Global\\HC_MAIN_RUNNER"
_LEGACY_BRIDGE_RUNNER_MUTEX_NAME = "Global\\HC_BRIDGE_RUNNER"

_SYNCHRONIZE = 0x00100000
_kernel32 = ctypes.windll.kernel32  # type: ignore[attr-defined]


def _is_mutex_exists(name: str) -> bool:
    try:
        handle = _kernel32.OpenMutexW(_SYNCHRONIZE, False, wintypes.LPCWSTR(name))
        if handle:
            _kernel32.CloseHandle(handle)
            return True
        return False
    except Exception:
        return False


def is_ui_server_running() -> bool:
    return _is_mutex_exists(_MUTEX_NAME)


def _project_pythonw(project_root: Path) -> str:
    """Return the project venv pythonw.exe if present; otherwise fallback to python.exe / current interpreter.

    Rationale:
      Prefer pythonw.exe to avoid showing a console window when launching UI/svc servers.
      Even if the entry interpreter is a system/base python, we must start servers with
      the *project* venv to avoid double servers (system+venv).
    """
    try:
        venv_pythonw = (project_root / ".venv" / "Scripts" / "pythonw.exe").resolve()
        if venv_pythonw.exists():
            return str(venv_pythonw)
        venv_python = (project_root / ".venv" / "Scripts" / "python.exe").resolve()
        if venv_python.exists():
            return str(venv_python)
    except Exception:
        pass
    # final fallback: sibling pythonw of current interpreter, else sys.executable
    try:
        pythonw = Path(sys.executable).with_name("pythonw.exe")
        if pythonw.exists():
            return str(pythonw)
    except Exception:
        pass
    return sys.executable


def _resolve_ui_server_path() -> Path:
    # .../svc/svc_host.py -> project root -> ui_qt/ui_server.py
    here = Path(__file__).resolve().parent
    root = here.parent
    return (root / "ui_qt" / "ui_server.py").resolve()


def spawn_ui_server() -> None:
    """Qt UI サーバを起動する（未起動想定）。"""
    dev_root = Path(__file__).resolve().parent.parent
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
        project_root = runtime_layout.runtime_project_root(str(Path(__file__).resolve()))
    else:
        server_py = _resolve_ui_server_path()
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
    boot_log = logs_dir / f"ui_server_boot_{int(time.time() * 1000)}.log"

    env = os.environ.copy()
    env["HC_IPC_ROOT"] = ipc_root
    env["HC_QT_IPC_DIR"] = ipc_root
    env["HC_PROJECT_ROOT"] = str(project_root)
    env["PYTHONPATH"] = str(project_root) + os.pathsep + env.get("PYTHONPATH", "")

    env["HC_UI_PARENT_PID"] = str(os.getpid())
    try:
        env["HC_EXCEL_PID"] = str(os.getppid())
    except Exception:
        pass

    if packaged:
        ir = runtime_layout.install_root()
        if ir is not None:
            env = runtime_layout.env_with_packaged_dll_search_path(env, ir)

    if packaged:
        if ui_exe is None:
            return
        cmd = [str(ui_exe)]
        spawn_label = ui_exe
        # EXE と同じフォルダを cwd に（Nuitka+PySide6 の相対パス・DLL 解決を安定化。HC_PROJECT_ROOT は引き続きインストールルート）
        ui_cwd = str(ui_exe.resolve().parent)
    else:
        exe = _project_pythonw(project_root)
        cmd = [exe, "-u", str(server_py)]
        spawn_label = server_py
        ui_cwd = str(project_root)

    with boot_log.open("w", encoding="utf-8") as f:
        f.write(f"[BOOT] cmd={cmd}\n")
        f.write(f"[BOOT] cwd={ui_cwd}\n")
        f.write(f"[BOOT] HC_PROJECT_ROOT={project_root}\n")
        f.write(f"[BOOT] HC_QT_IPC_DIR={ipc_root}\n")

    popen_kw: dict = {
        "cwd": ui_cwd,
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



def ensure_ui_server() -> None:
    """起動済みなら何もしない。未起動なら spawn する（別プロセス常駐）。"""
    clear_shutdown_flags("ensure_ui_server")
    if is_ui_server_running():
        logger.info("[QT_UI_SERVER] already running (mutex exists)")
        return

    # ------------------------------------------------------------------
    # double-spawn guard
    #  - RunPython が短時間に複数回呼ばれると、mutex が張られる前に
    #    spawn が重複し、結果として ui_server が複数常駐することがある。
    #  - そのため IPC ルート配下の flag で「起動中」を共有する。
    # ------------------------------------------------------------------
    ipc_root = Path(str(ipc_file.get_ipc_root()))
    flag = ipc_root / "control" / "ui_server_starting.flag"
    flag.parent.mkdir(parents=True, exist_ok=True)

    try:
        if flag.exists() and (time.time() - flag.stat().st_mtime) < 5.0:
            logger.info("[QT_UI_SERVER] startup in progress (flag exists); skip spawn")
        else:
            flag.write_text(str(int(time.time() * 1000)), encoding="utf-8")
            spawn_ui_server()

        # wait for mutex (初回起動が遅い環境向けに猶予付き)
        _wait_until_running(is_ui_server_running, "[QT_UI_SERVER]", poll_sec=0.02)
    finally:
        try:
            if flag.exists():
                flag.unlink()
        except Exception:
            pass


def _resolve_svc_server_path() -> Path:
    # .../svc/svc_host.py -> same folder -> svc_server.py
    here = Path(__file__).resolve().parent
    return (here / "svc_server.py").resolve()


def is_svc_server_running() -> bool:
    return _is_mutex_exists(_SVC_MUTEX_NAME)


def is_main_runner_running() -> bool:
    """常駐 hc_main が生存しているか（`HC_MAIN_RUNNER` または移行用旧名 `HC_BRIDGE_RUNNER`）。"""
    return _is_mutex_exists(_MAIN_RUNNER_MUTEX_NAME) or _is_mutex_exists(
        _LEGACY_BRIDGE_RUNNER_MUTEX_NAME
    )


def is_bridge_running() -> bool:
    """`is_main_runner_running` と同一（旧 API 名の互換）。"""
    return is_main_runner_running()


def all_python_hosts_running() -> bool:
    """svc / ui / bridge がすべて mutex 生存しているか。"""
    return (
        is_svc_server_running()
        and is_ui_server_running()
        and is_bridge_running()
    )


def spawn_svc_server() -> None:
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


def _resolve_bridge_path() -> Path:
    """常駐ブリッジ: プロジェクトルート直下の hc_main.py（必須。無い場合は spawn 側で警告して中止）。"""
    root = Path(__file__).resolve().parent.parent
    return (root / "hc_main.py").resolve()


def _bridge_script_project_root(bridge_py: Path) -> Path:
    """hc_main.py はルート直下配置を前提とする。"""
    return bridge_py.parent.resolve()


def spawn_bridge() -> None:
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


def ensure_bridge() -> None:
    """ブリッジが起動済みなら何もしない。未起動なら spawn する。"""
    clear_shutdown_flags("ensure_bridge")
    if is_bridge_running():
        logger.info("%s already running (mutex exists)", LOG_MAIN_PREFIX)
        return
    ipc_root = Path(str(ipc_file.get_ipc_root()))
    flag = ipc_root / "control" / "bridge_starting.flag"
    flag.parent.mkdir(parents=True, exist_ok=True)
    try:
        if flag.exists() and (time.time() - flag.stat().st_mtime) < 5.0:
            logger.info("%s startup in progress (flag exists); skip spawn", LOG_MAIN_PREFIX)
        else:
            flag.write_text(str(int(time.time() * 1000)), encoding="utf-8")
            spawn_bridge()
        _wait_until_running(is_bridge_running, LOG_MAIN_PREFIX, poll_sec=0.02)
    finally:
        try:
            if flag.exists():
                flag.unlink()
        except Exception:
            pass


_STARTUP_SPAWN_FLAG_TTL_SEC = 5.0
_SVC_STARTING_FLAG = "svc_server_starting.flag"
_UI_STARTING_FLAG = "ui_server_starting.flag"
_BRIDGE_STARTING_FLAG = "bridge_starting.flag"


def _control_flag_path(name: str) -> Path:
    ipc_root = Path(str(ipc_file.get_ipc_root()))
    return ipc_root / "control" / name


def _clear_control_flags(*names: str) -> None:
    for name in names:
        try:
            p = _control_flag_path(name)
            if p.exists():
                p.unlink()
        except Exception:
            pass


def _spawn_server_if_needed(
    is_running: Callable[[], bool],
    spawn_fn: Callable[[], None],
    flag_name: str,
    log_label: str,
) -> None:
    """未起動なら spawn のみ（mutex 待ちは呼び出し側で一括）。"""
    if is_running():
        logger.info("%s already running (mutex exists)", log_label)
        return
    flag = _control_flag_path(flag_name)
    flag.parent.mkdir(parents=True, exist_ok=True)
    if flag.exists() and (time.time() - flag.stat().st_mtime) < _STARTUP_SPAWN_FLAG_TTL_SEC:
        logger.info("%s startup in progress (flag exists); skip spawn", log_label)
        return
    flag.write_text(str(int(time.time() * 1000)), encoding="utf-8")
    spawn_fn()


def _wait_until_all_running(
    checks: list[tuple[Callable[[], bool], str]],
    *,
    max_wait_sec: float = _MUTEX_WAIT_SEC,
    poll_sec: float = 0.02,
) -> None:
    """複数常駐プロセスの mutex を一括待ち（並列 spawn 後）。"""
    t0 = time.time()
    while time.time() - t0 < max_wait_sec:
        if all(fn() for fn, _ in checks):
            return
        time.sleep(poll_sec)
    time.sleep(_MUTEX_GRACE_SEC)
    if all(fn() for fn, _ in checks):
        for _fn, label in checks:
            logger.info("%s mutex observed after grace wait", label)
        return
    missing = [label for fn, label in checks if not fn()]
    logger.warning(
        "[HOST_STARTUP] parallel spawn mutex not observed yet: %s",
        ",".join(missing),
    )


def ensure_python_hosts_ready(target_hwnd: int | None = None) -> None:
    """起動時・リボン操作時: svc/ui/bridge が死んでいれば起動。生存中は何もしない。

    マルチ Excel / 新規ブックでは Workbook_Open の register_book 前にリボンが押されることがあるため、
    target_hwnd を渡してブック登録も行う。

    常駐ホストがすべて生存かつ HWND 登録済み IPC がある場合は spawn / prewarm / register_book COM を省略する。
    B+: 常駐 svc_server は HWND キャッシュでマルチ Excel を扱い、COM 汚染時のみ recycle する。
    """
    hwnd = int(target_hwnd or 0)
    plog = get_perf_logger(f"{__name__}.ensure_hosts")
    t0 = time.perf_counter()

    if all_python_hosts_running():
        from core.excel_book_register_gate import (
            mark_excel_book_registered,
            should_skip_register_book_com,
        )

        if hwnd > 0 and should_skip_register_book_com(hwnd):
            clear_shutdown_flags("ensure_python_hosts_ready_fast")
            plog.info(
                "ribbon_fast phase=skip_hosts_and_register cumulative_ms=%d hwnd=%s",
                int((time.perf_counter() - t0) * 1000),
                hwnd,
            )
            return
        clear_shutdown_flags("ensure_python_hosts_ready")
        if hwnd > 0:
            from core.excel_session import register_book

            register_book(target_hwnd=hwnd)
            mark_excel_book_registered(hwnd)
            plog.info(
                "ribbon_fast phase=register_only cumulative_ms=%d hwnd=%s",
                int((time.perf_counter() - t0) * 1000),
                hwnd,
            )
        else:
            plog.info(
                "ribbon_fast phase=skip_hosts cumulative_ms=%d",
                int((time.perf_counter() - t0) * 1000),
            )
        return

    if hwnd > 0:
        from core.excel_com_session import prepare_com_session_before_request

        prepare_com_session_before_request(hwnd)
    ensure_svc_ui_bridge_parallel()
    if hwnd > 0:
        from core.excel_book_register_gate import mark_excel_book_registered
        from core.excel_session import register_book

        register_book(target_hwnd=hwnd)
        mark_excel_book_registered(hwnd)


def ensure_svc_ui_bridge_parallel() -> None:
    """svc / ui / bridge を並列 spawn し、xlwings prewarm と並行して mutex を待つ。"""
    clear_shutdown_flags("ensure_svc_ui_bridge_parallel")
    if all_python_hosts_running():
        logger.info(
            "[HOST] all python hosts already running — skip parallel spawn/prewarm"
        )
        return

    from core.ribbon_invoke import start_xlwings_import_prewarm

    start_xlwings_import_prewarm()
    try:
        _spawn_server_if_needed(
            is_svc_server_running,
            spawn_svc_server,
            _SVC_STARTING_FLAG,
            "[SVC_SERVER]",
        )
        _spawn_server_if_needed(
            is_ui_server_running,
            spawn_ui_server,
            _UI_STARTING_FLAG,
            "[QT_UI_SERVER]",
        )
        _spawn_server_if_needed(
            is_bridge_running,
            spawn_bridge,
            _BRIDGE_STARTING_FLAG,
            LOG_MAIN_PREFIX,
        )
        _wait_until_all_running(
            [
                (is_svc_server_running, "[SVC_SERVER]"),
                (is_ui_server_running, "[QT_UI_SERVER]"),
                (is_bridge_running, LOG_MAIN_PREFIX),
            ],
            poll_sec=0.02,
        )
    finally:
        _clear_control_flags(
            _SVC_STARTING_FLAG,
            _UI_STARTING_FLAG,
            _BRIDGE_STARTING_FLAG,
        )


def ensure_svc_server() -> None:
    """起動済みなら何もしない。未起動なら spawn する（別プロセス常駐）。"""
    clear_shutdown_flags("ensure_svc_server")
    if is_svc_server_running():
        logger.info("[SVC_SERVER] already running (mutex exists)")
        return

    ipc_root = Path(str(ipc_file.get_ipc_root()))
    flag = ipc_root / "control" / "svc_server_starting.flag"
    flag.parent.mkdir(parents=True, exist_ok=True)

    try:
        if flag.exists() and (time.time() - flag.stat().st_mtime) < 5.0:
            logger.info("[SVC_SERVER] startup in progress (flag exists); skip spawn")
        else:
            flag.write_text(str(int(time.time() * 1000)), encoding="utf-8")
            spawn_svc_server()

        _wait_until_running(is_svc_server_running, "[SVC_SERVER]", poll_sec=0.05)
    finally:
        try:
            if flag.exists():
                flag.unlink()
        except Exception:
            pass


def is_shutdown_requested() -> bool:
    """svc 側で shutdown が要求されているか。"""
    return _STOP_EVENT.is_set()


def request_svc_shutdown() -> None:
    """svc 側の処理停止を要求する（UIサーバとは別）。"""
    _STOP_EVENT.set()


def request_ui_server_shutdown() -> None:
    """ui_server に終了要求を出す（次の idle で quit）。"""
    try:
        ipc_file.write_shutdown_flag()
        logger.info("[QT_UI_SERVER] shutdown requested (flag written)")
    except Exception as ex:
        logger.exception("[QT_UI_SERVER] shutdown request failed: %s", ex)


def _write_svc_shutdown_flag() -> None:
    d = _control_dir()
    p = d / "svc_shutdown.flag"
    try:
        p.write_text("shutdown", encoding="utf-8")
    except Exception:
        try:
            p.open("a").close()
        except Exception:
            pass


_SVC_LAST_COM_HWND_FILE = "svc_last_com_hwnd.txt"
_SVC_COM_RECYCLE_WAIT_SEC = 3.0


def read_last_svc_com_hwnd() -> int:
    """svc_server が最後に COM 接続した Excel HWND（IPC 永続化）。"""
    try:
        p = _control_dir() / _SVC_LAST_COM_HWND_FILE
        if not p.exists():
            return 0
        return int((p.read_text(encoding="utf-8") or "0").strip() or "0")
    except Exception:
        return 0


def write_last_svc_com_hwnd(hwnd: int) -> None:
    """svc_server の COM 接続先 HWND を記録する。"""
    ph = int(hwnd or 0)
    p = _control_dir() / _SVC_LAST_COM_HWND_FILE
    p.parent.mkdir(parents=True, exist_ok=True)
    try:
        if ph <= 0:
            p.unlink(missing_ok=True)
        else:
            p.write_text(str(ph), encoding="utf-8")
    except Exception:
        pass


def _list_svc_server_pids(project_root: Path | None = None) -> list[int]:
    """稼働中の svc_server プロセス PID を列挙する。"""
    if os.name != "nt":
        return []
    root = project_root or Path(__file__).resolve().parent.parent
    root_key = str(root.resolve()).lower()
    pids: list[int] = []
    try:
        cmd = (
            "Get-CimInstance Win32_Process | "
            "Where-Object { $_.Name -eq 'hc_svc_server.exe' -or "
            "($_.CommandLine -match 'svc_server\\.py') } | "
            "ForEach-Object { $_.ProcessId.ToString() + [char]9 + ($_.CommandLine ?? '') }"
        )
        cp = subprocess.run(
            ["powershell", "-NoProfile", "-Command", cmd],
            capture_output=True,
            text=True,
            check=False,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        for line in (cp.stdout or "").splitlines():
            low = line.lower()
            if "svc_server" not in low and "hc_svc_server" not in low:
                continue
            if root_key not in low and "hc_svc_server.exe" not in low:
                continue
            try:
                pid = int(line.split("\t", 1)[0].strip())
            except Exception:
                continue
            if pid > 0 and pid not in pids:
                pids.append(pid)
    except Exception:
        pass
    return pids


def restart_svc_server(*, reason: str = "") -> None:
    """救済用: svc_server を明示的に終了し新インスタンスを起動する（ui/bridge は維持）。

    通常のリボン操作では呼ばない（B+ 常駐）。COM 汚染時は svc_server 自プロセスの
    com_recycle が先に動く。本関数は手動復旧・将来の明示再起動用に残す。
    """
    logger.info("[SVC_SERVER] recovery restart begin reason=%s", reason or "-")
    if is_svc_server_running():
        _write_svc_shutdown_flag()
        t0 = time.time()
        while is_svc_server_running() and (time.time() - t0) < _SVC_COM_RECYCLE_WAIT_SEC:
            time.sleep(0.05)
        if is_svc_server_running():
            root = Path(__file__).resolve().parent.parent
            for pid in _list_svc_server_pids(root):
                if pid != os.getpid():
                    _safe_kill_pid_windows(pid)
            time.sleep(0.15)
    clear_shutdown_flags("restart_svc_server")
    _clear_control_flags(_SVC_STARTING_FLAG)
    if not is_svc_server_running():
        spawn_svc_server()
        _wait_until_running(is_svc_server_running, "[SVC_SERVER]", poll_sec=0.05)
    logger.info("[SVC_SERVER] recovery restart done reason=%s", reason or "-")


def restart_svc_server_for_com_if_needed(target_hwnd: int) -> bool:
    """COM 操作前の svc_server 再起動判定（B+: 常駐維持のため事前再起動は行わない）。

    COM 汚染時の recycle は svc_server 側で com_recycle を予約する。
    本関数は API 互換のため残し、常に False を返す。
    """
    _ = int(target_hwnd or 0)
    return False


def request_shutdown_all() -> None:
    """Excel終了時の共通終了処理。"""
    request_svc_shutdown()
    request_ui_server_shutdown()
    try:
        _write_svc_shutdown_flag()
        logger.info("[SVC_SERVER] shutdown requested (flag written)")
    except Exception:
        pass


_HC_PACKAGED_EXE_MARKERS = (
    "hc_main.exe",
    "hc_svc_server.exe",
    "hc_ui_server.exe",
)


def _safe_kill_pid_windows(pid: int) -> bool:
    """指定 PID を taskkill で終了する（ベストエフォート）。"""
    if os.name != "nt" or int(pid or 0) <= 0:
        return False
    try:
        cp = subprocess.run(
            ["taskkill", "/F", "/T", "/PID", str(int(pid))],
            capture_output=True,
            text=True,
            check=False,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        return cp.returncode in (0, 128)
    except Exception:
        return False


def _list_hc_python_target_pids(project_root: Path) -> list[int]:
    """現在のプロジェクト配下の hc_main / svc_server / ui_server の PID を列挙する。"""
    if os.name != "nt":
        return []
    try:
        script_targets = (
            str((project_root / "hc_main.py").resolve()).lower(),
            str((project_root / "svc" / "svc_server.py").resolve()).lower(),
            str((project_root / "ui_qt" / "ui_server.py").resolve()).lower(),
        )
        rel_markers = (
            "hc_main.py",
            "svc\\svc_server.py",
            "svc/svc_server.py",
            "ui_qt\\ui_server.py",
            "ui_qt/ui_server.py",
        )
        root_key = str(project_root.resolve()).lower()
    except Exception:
        return []
    cmd = (
        "Get-CimInstance Win32_Process | "
        "Where-Object { $_.Name -match '^(python|pythonw)\\.exe$' -or "
        "$_.Name -match '^hc_(main|svc_server|ui_server)\\.exe$' } | "
        "ForEach-Object { $_.ProcessId.ToString() + [char]9 + ($_.CommandLine ?? '') }"
    )
    try:
        cp = subprocess.run(
            ["powershell", "-NoProfile", "-Command", cmd],
            capture_output=True,
            text=True,
            check=False,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except Exception:
        return []
    if cp.returncode != 0:
        return []
    pids: list[int] = []
    for line in (cp.stdout or "").splitlines():
        low = line.lower()
        matched = any(t in low for t in script_targets)
        if not matched:
            matched = any(m in low for m in _HC_PACKAGED_EXE_MARKERS)
        if not matched and root_key in low:
            matched = any(m in low for m in rel_markers)
        if not matched:
            continue
        try:
            if "\t" in line:
                pid_token = line.split("\t", 1)[0]
                pid = int(pid_token.strip())
            else:
                pid = int(shlex.split(line.strip())[0])
        except Exception:
            parts = line.strip().split()
            if not parts:
                continue
            try:
                pid = int(parts[0])
            except Exception:
                continue
        if pid > 0 and pid not in pids:
            pids.append(pid)
    return pids


def excel_shutdown_workbook_close(
    target_hwnd: int,
    sheet_id: str = "",
    reason: str = "excel_shutdown",
) -> None:
    """Workbook_BeforeClose 用: 1 回の RunPython で restore / shutdown / registry clear まで完了する。"""
    from core.excel_host_restore import restore_excel_host_ui_state
    from core.excel_session import clear_internal_registry

    plog = get_perf_logger(f"{__name__}.excel_shutdown")
    t0 = time.perf_counter()
    plog.info("shutdown phase=enter cumulative_ms=0 hwnd=%s", int(target_hwnd or 0))
    restore_excel_host_ui_state(int(target_hwnd or 0), str(sheet_id or ""))
    plog.info(
        "shutdown phase=after_restore cumulative_ms=%d",
        int((time.perf_counter() - t0) * 1000),
    )
    shutdown_all_with_force_kill(reason)
    plog.info(
        "shutdown phase=after_shutdown cumulative_ms=%d",
        int((time.perf_counter() - t0) * 1000),
    )
    clear_internal_registry()
    plog.info(
        "shutdown phase=done cumulative_ms=%d",
        int((time.perf_counter() - t0) * 1000),
    )


_SHUTDOWN_GRACE_MAX_SEC: float = 1.2
_SHUTDOWN_POLL_INTERVAL_SEC: float = 0.15


def _shutdown_grace_wait_for_targets(
    project_root: Path,
    self_pid: int,
    *,
    max_wait_sec: float = _SHUTDOWN_GRACE_MAX_SEC,
    poll_interval_sec: float = _SHUTDOWN_POLL_INTERVAL_SEC,
) -> list[int]:
    """shutdown フラグ書込後、常駐プロセスが自然終了するまで短い間隔で待つ（最大 max_wait_sec）。"""
    deadline = time.monotonic() + max(0.0, float(max_wait_sec))
    interval = max(0.05, float(poll_interval_sec))
    target_pids = [
        p for p in _list_hc_python_target_pids(project_root) if p != self_pid
    ]
    while target_pids and time.monotonic() < deadline:
        time.sleep(min(interval, max(0.0, deadline - time.monotonic())))
        target_pids = [
            p for p in _list_hc_python_target_pids(project_root) if p != self_pid
        ]
    return target_pids


def shutdown_all_with_force_kill(reason: str = "excel_shutdown") -> None:
    """通常 shutdown 要求 + 残留 Python 常駐プロセスのフェイルセーフ終了。"""
    request_shutdown_all()
    if os.name != "nt":
        return
    project_root = Path(__file__).resolve().parent.parent
    self_pid = os.getpid()
    target_pids = _shutdown_grace_wait_for_targets(project_root, self_pid)
    killed = 0
    for pid in target_pids:
        if _safe_kill_pid_windows(pid):
            killed += 1
    logger.info(
        "[HOST_SHUTDOWN] reason=%s force_kill_target=%s force_kill_done=%s self_pid=%s",
        reason,
        len(target_pids),
        killed,
        self_pid,
    )

# Release: hc_host_v0.4.16
