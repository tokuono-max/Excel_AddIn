# -*- coding: utf-8 -*-
"""
Python: 3.12
Module: ui_qt/ipc_file.py
Created: 2026-02-11
Updated: 2026-06-06
Version: 1.2.13
Purpose:
  Qt UI サーバ (ui_qt) と svc 層の間で、ファイル(Pickle)で通信する最小 IPC。
  - req_*.pkl : svc -> ui_server
  - res_*.pkl : ui_server -> svc
  - ready_*.pkl : ui_server -> svc（初回描画完了などの早期通知）

History (latest 3):
  - 1.2.13 (2026-06-06) waitform_ready_signal_path / write_waitform_ready_signal（VBA WaitForm 合図ファイル）。
  - 1.2.12 (2026-04-06) IPC ルート解決を core.core_env.ipc_dir_raw() に統一（HC_IPC_ROOT / HC_QT_IPC_DIR）。
  - 1.2.1 (2026-02-11) TEMP配下固定(%TEMP%\\csv_tool)とログパスを安定化。構文不備を修正。
"""


import hashlib
import os
import pickle
import tempfile
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Optional, Tuple

import ctypes
from ctypes import wintypes

from core import core_env

# NOTE:
#  - UIサーバは別プロセスのため、TEMP が別になる事故を避ける。
#  - IPC ルートは core.core_env.ipc_dir_raw()（HC_IPC_ROOT / HC_QT_IPC_DIR）で解決。
#  - 未設定の場合も、%TEMP%\\csv_tool に固定して双方で一致させる。

# Windows named mutex（多重起動防止）
_MUTEX_NAME = "Global\\HC_QT_UI_SERVER"
_ERROR_ALREADY_EXISTS = 183

# DIAG: which file is actually imported (no console needed)
try:
    __version__ = "1.2.13"
    VERSION = __version__
    from pathlib import Path
    import os
    from core.core_log import append_text_with_cap
    _ipc_dir = core_env.ipc_dir_raw()
    if _ipc_dir:
        Path(_ipc_dir).mkdir(parents=True, exist_ok=True)
        append_text_with_cap(
            Path(_ipc_dir) / "ipc_import.log",
            f"ipc_file imported from: {Path(__file__).resolve()}  version={globals().get('VERSION','?')}\n",
        )
except Exception:
    pass



def get_ipc_root() -> Path:
    """IPC ルートディレクトリを返す（必ず存在させる）。"""
    forced = core_env.ipc_dir_raw()
    if forced:
        d = Path(forced)
    else:
        d = Path(tempfile.gettempdir()) / "csv_tool"

    d.mkdir(parents=True, exist_ok=True)
    return d


def waitform_ready_signal_path(parent_hwnd: int) -> Path:
    """VBA WaitForm 解除合図ファイル（Application.hwnd / parent_hwnd と同一キー）。"""
    hwnd = int(parent_hwnd or 0)
    temp = os.environ.get("TEMP") or os.environ.get("TMP") or tempfile.gettempdir()
    return Path(temp) / "csv_tool" / "waitform" / f"{hwnd}.ready"


def write_waitform_ready_signal(parent_hwnd: int) -> None:
    """ui_server が UI 表示直前に書く。VBA は DoEvents でこのファイルを待つ。"""
    hwnd = int(parent_hwnd or 0)
    if hwnd <= 0:
        return
    path = waitform_ready_signal_path(hwnd)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("READY_UI\n", encoding="utf-8")


# 基準フォルダ: ファイル読込/保存で直近使ったフォルダを共有（CSV読込・CSV保存で共通）
_LAST_FOLDER_FILENAME = "last_folder.txt"


def get_last_folder() -> str:
    """基準フォルダ（直近のファイル読込/保存で使ったフォルダ）を返す。無ければ空文字。"""
    try:
        p = get_ipc_root() / _LAST_FOLDER_FILENAME
        if not p.exists() or p.stat().st_size <= 0:
            return ""
        raw = p.read_text(encoding="utf-8-sig").replace("\r\n", "\n").replace("\r", "\n")
        line = ""
        for part in raw.split("\n"):
            t = part.strip()
            if t:
                line = t
                break
        if not line:
            return ""
        if (line.startswith('"') and line.endswith('"')) or (
            line.startswith("'") and line.endswith("'")
        ):
            line = line[1:-1].strip()
        line = os.path.expandvars(os.path.expanduser(line.strip()))
        if not line:
            return ""
        cand = Path(line)
        if cand.is_dir():
            return str(cand.resolve())
    except Exception:
        pass
    return ""


def set_last_folder(dir_path: str) -> None:
    """基準フォルダを保存する。dir_path はディレクトリの絶対パス（空なら書かない）。"""
    dir_path = (dir_path or "").strip()
    if not dir_path:
        return
    try:
        p = Path(dir_path)
        if not p.is_dir():
            p = p.parent
        if not p.is_dir():
            return
        out = get_ipc_root() / _LAST_FOLDER_FILENAME
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(str(p.resolve()), encoding="utf-8")
    except Exception:
        pass


def get_request_dir() -> Path:
    """req_*.pkl を置くディレクトリを返す。"""
    d = get_ipc_root() / "requests"
    d.mkdir(parents=True, exist_ok=True)
    return d


def get_server_log_path() -> Path:
    """UIサーバのログファイル（%TEMP% 配下）を返す。"""
    return get_ipc_root() / "ui_server.log"



# -----------------------------------------------------------------------------
# Control (shutdown etc.)
# -----------------------------------------------------------------------------

def get_control_dir() -> Path:
    """control 用ディレクトリを返す（shutdown 等）。"""
    d = get_ipc_root() / "control"
    d.mkdir(parents=True, exist_ok=True)
    return d


def get_shutdown_flag_path() -> Path:
    """ui_server 停止要求フラグのパスを返す。"""
    return get_control_dir() / "shutdown.flag"


def write_shutdown_flag() -> None:
    """ui_server 停止要求フラグを作成する（best-effort）。"""
    p = get_shutdown_flag_path()
    p.parent.mkdir(parents=True, exist_ok=True)
    try:
        p.write_text("shutdown", encoding="utf-8")
    except Exception:
        try:
            p.open("a").close()
        except Exception:
            pass


def clear_shutdown_flag() -> None:
    """ui_server 停止要求フラグを削除する（best-effort）。"""
    p = get_shutdown_flag_path()
    try:
        p.unlink(missing_ok=True)
    except Exception:
        pass

def _atomic_write_bytes(path: Path, data: bytes) -> None:
    """中途半端なファイルを残さないために atomic write する。

    注意:
        - 監視側が「存在した瞬間」に読みに来るケースがあるため、
          0byte ファイルが見えないようにすることが重要。
        - 同一ディレクトリ内の一時ファイルに書いてから os.replace する。
          （Windows でも同一ボリューム内なら原子的に置換される）
    """
    path.parent.mkdir(parents=True, exist_ok=True)

    # 同一ディレクトリにユニークな tmp を作る（衝突防止）
    fd, tmp_name = tempfile.mkstemp(
        prefix=path.name + ".",
        suffix=".tmp",
        dir=str(path.parent),
    )
    try:
        with os.fdopen(fd, "wb") as f:
            f.write(data)
            f.flush()
            os.fsync(f.fileno())
        os.replace(tmp_name, path)
    finally:
        # replace 前に落ちた場合の掃除
        try:
            if os.path.exists(tmp_name):
                os.remove(tmp_name)
        except OSError:
            pass


def write_pickle(path: Path, data: Any) -> None:
    """pickle を atomic に書き込む。"""
    payload = pickle.dumps(data, protocol=pickle.HIGHEST_PROTOCOL)
    _atomic_write_bytes(path, payload)


def batch_done_notify_path(sheet_id: str) -> Path:
    """一括実行完了を親 Qt ダイアログがポーリングするための pickle パス（sheet_id で安定ハッシュ）。"""
    h = hashlib.sha256(str(sheet_id or "").encode("utf-8", errors="replace")).hexdigest()[:24]
    d = get_ipc_root() / "result"
    d.mkdir(parents=True, exist_ok=True)
    return d / ("data_agg_batch_done_%s.pkl" % h)


def write_batch_done_notify(
    sheet_id: str,
    title: str,
    message: str,
    *,
    ok: bool,
    run_id: str = "",
    error: str = "",
    abort_phase: str = "",
) -> None:
    """別プロセス一括実行の完了を、メイン UI と同一 IPC ルートへ書き出す。"""
    write_pickle(
        batch_done_notify_path(sheet_id),
        {
            "title": str(title or "データ集約"),
            "message": str(message or ""),
            "ok": bool(ok),
            "run_id": str(run_id or ""),
            "error": str(error or ""),
            "abort_phase": str(abort_phase or ""),
            "ts": time.time(),
        },
    )


def try_read_batch_done_notify(sheet_id: str) -> dict[str, Any] | None:
    p = batch_done_notify_path(sheet_id)
    if not p.is_file():
        return None
    try:
        raw = read_pickle(p)
        return raw if isinstance(raw, dict) else None
    except Exception:
        return None


def delete_batch_done_notify(sheet_id: str) -> None:
    try:
        batch_done_notify_path(sheet_id).unlink(missing_ok=True)
    except OSError:
        pass


def read_pickle(path: Path) -> Any:
    """pickle を読み込む（Windows の瞬間的な PermissionError を短いリトライで吸収）。"""
    last_exc: Exception | None = None
    max_attempts = 24
    for i in range(max_attempts):
        try:
            return pickle.loads(path.read_bytes())
        except PermissionError as e:
            last_exc = e
            time.sleep(min(0.06, 0.004 * (i + 1)))
        except EOFError as e:
            last_exc = e
            time.sleep(min(0.06, 0.004 * (i + 1)))
    if last_exc is not None:
        raise last_exc
    return pickle.loads(path.read_bytes())


@dataclass(frozen=True)
class UiRequest:
    """ui_server へ渡す要求。"""

    parent_hwnd: int
    result_path: str
    ready_path: str = ""
    sheet_id: str = ""
    # NOTE: ログパスを要求側で指定したい場合に使用（未指定ならサーバ既定）
    log_path: str = ""
    action: str = "csv_mg"

    # ui_server側で import するモジュール（未指定なら既定）
    module: str = ""
    # 追加の要求ペイロード（後方互換のため、未指定でもOK）
    req_dict: dict[str, Any] | None = None

    def to_dict(self) -> dict[str, Any]:
        return {
            "parent_hwnd": int(self.parent_hwnd),
            "result_path": self.result_path,
            "ready_path": self.ready_path,
            "sheet_id": self.sheet_id,
            "log_path": self.log_path,
            "action": self.action,
            "module": self.module,
            "req_dict": self.req_dict,
        }

    @staticmethod
    def from_dict(d: dict[str, Any]) -> "UiRequest":
        return UiRequest(
            parent_hwnd=int(d.get("parent_hwnd", 0) or 0),
            result_path=str(d.get("result_path", "")),
            ready_path=str(d.get("ready_path", "")),
            sheet_id=str(d.get("sheet_id", "")),
            log_path=str(d.get("log_path", "")),
            action=str(d.get("action", "csv_mg")),
            module=str(d.get("module", "") or ""),
            req_dict=d.get("req_dict") if isinstance(d.get("req_dict"), dict) else None,
        )


def submit_request(req: UiRequest) -> Path:
    """req_*.pkl を生成して投入する。

    注意:
      - req ファイルは ui_server が最初に「claim（リネーム）」してから処理する。
      - 書き込み後に存在/サイズを確認し、「存在しないのに submitted」と誤認しない。
    """
    req_dir = get_request_dir()
    ts_ms = int(time.time() * 1000)
    path = req_dir / f"req_{ts_ms}_{os.getpid()}.pkl"
    write_pickle(path, req.to_dict())

    # 書き込み直後の防衛（ごく稀な競合/失敗時に「存在しないのに submitted」を防ぐ）
    try:
        if not path.exists() or path.stat().st_size <= 0:
            raise FileNotFoundError(str(path))
    except OSError as e:
        raise RuntimeError(f"request file not created: {path}") from e

    return path


def _claim_request_file(path: Path) -> Optional[Path]:
    """req を処理中として claim する（原子的にリネーム）。

    - 成功: .work.pkl にリネームした Path を返す
    - 失敗/競合: None
    """
    work = path.with_suffix(".work.pkl")
    try:
        return path.replace(work)
    except FileNotFoundError:
        return None
    except PermissionError:
        return None
    except OSError:
        return None


def _safe_mtime_for_sort(path: Path) -> float:
    """mtime 取得失敗時は末尾扱いにする（列挙中の競合で全体が落ちないように）。"""
    try:
        return path.stat().st_mtime
    except OSError:
        return float("inf")


def pop_next_request() -> Optional[Path]:
    """次に処理すべき req を1つ取り出す（無ければ None）。

    取り出しは「claim（リネーム）」で行い、二重処理やレースでの FileNotFound を抑止する。
    """
    req_dir = get_request_dir()
    items = sorted(req_dir.glob("req_*.pkl"), key=_safe_mtime_for_sort)
    for p in items:
        claimed = _claim_request_file(p)
        if claimed is not None:
            return claimed
    return None


def cleanup_failed_requests(*, ttl_sec: int = 24 * 60 * 60, max_remove: int = 20) -> int:
    """隔離済み req（requests/_failed/*.bad.pkl）を TTL ベースで掃除する。"""
    req_dir = get_request_dir()
    failed_dir = req_dir / "_failed"
    if not failed_dir.is_dir():
        return 0
    ttl = max(0, int(ttl_sec))
    limit = max(1, int(max_remove))
    cutoff = time.time() - ttl
    removed = 0
    for p in sorted(failed_dir.glob("*.bad.pkl"), key=_safe_mtime_for_sort):
        if removed >= limit:
            break
        try:
            if p.stat().st_mtime > cutoff:
                continue
            p.unlink(missing_ok=True)
            removed += 1
        except OSError:
            continue
    return removed


def cap_failed_requests(*, max_keep: int = 200) -> int:
    """隔離済み req の件数上限を維持し、超過分（古い順）を削除する。"""
    req_dir = get_request_dir()
    failed_dir = req_dir / "_failed"
    if not failed_dir.is_dir():
        return 0
    keep = max(0, int(max_keep))
    try:
        files = sorted(failed_dir.glob("*.bad.pkl"), key=_safe_mtime_for_sort)
    except OSError:
        return 0
    excess = max(0, len(files) - keep)
    if excess <= 0:
        return 0
    removed = 0
    for p in files[:excess]:
        try:
            p.unlink(missing_ok=True)
            removed += 1
        except OSError:
            continue
    return removed


def create_single_instance_mutex(name: str | None = None) -> Tuple[int, bool]:
    """UIサーバ多重起動防止 mutex を作成する。

    Returns:
        (handle, already_running)
    """
    kernel32 = ctypes.windll.kernel32
    kernel32.CreateMutexW.argtypes = [wintypes.LPVOID, wintypes.BOOL, wintypes.LPCWSTR]
    kernel32.CreateMutexW.restype = wintypes.HANDLE

    mutex_name = name or _MUTEX_NAME
    h = kernel32.CreateMutexW(None, True, mutex_name)
    already = kernel32.GetLastError() == _ERROR_ALREADY_EXISTS
    return int(h), bool(already)


def release_mutex(handle: int) -> None:
    """mutex を解放する（失敗しても無視）。"""
    try:
        ctypes.windll.kernel32.ReleaseMutex(wintypes.HANDLE(handle))
        ctypes.windll.kernel32.CloseHandle(wintypes.HANDLE(handle))
    except Exception:
        pass
