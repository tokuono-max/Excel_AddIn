# -*- coding: utf-8 -*-
"""
Python: 3.12+
Module: hc_main (project root)
Created: 2026-03-11 (logic from svc/bridge_runner)
Updated: 2026-04-11
Version: 0.2.4

Code map（本リビジョン時点の目安）:
  - 総行数: 約 500 行（import・bootstrap・`if __name__` 含む）
  - トップレベル関数（`def`）: 18 本（公開エントリは `main`、他はブリッジ／IPC 内部処理）

Purpose:
  常駐ブリッジ: bridge_requests の JSON を svc_requests に転送し svc_server が処理する。
  フェーズ D によりプロジェクトルートの本ファイルがエントリ（xlwings の sys.path 基点とも一致）。
  xlwings 短寿命の invoke は core.ribbon_invoke（core.excel_session 経由）。

History (latest 3):
  - 0.2.4 (2026-04-11) check_duplicates: bridge の `selection_count_large` / `sheet_cells_count_large` を kwargs に転送。
  - 0.2.3 (2026-04-11) bridge JSON の `selection_areas` を検証のうえ `svc_requests` の kwargs に転送（重複チェック用）。
  - 0.2.2 (2026-04-11) Mutex `HC_MAIN_RUNNER` ＋ 旧 `HC_BRIDGE_RUNNER` を両方取得（旧 svc_host との生存検知互換）。
  - 0.2.1 (2026-04-11) フェーズ E: ログ `[MAIN]`、`HC_MAIN_*` env（`HC_BRIDGE_*` フォールバック）。
  - 0.2.0 (2026-04-11) svc/bridge_runner からルート hc_main.py へ移設。bootstrap は「本ファイルの親＝プロジェクトルート」。
  - 0.1.8 (2026-04-11) core.ribbon_public_to_svc に action 対応を集約。
  - 0.1.7 (2026-04-10) リボン action 全件転送、JSON payload 引き継ぎ。
"""
from __future__ import annotations

import json
import os
import sys
import time
import tempfile
import pickle
from pathlib import Path

# 旧 bootstrap（_here を sys.path 先頭へ）の記録: フェーズ D 以降は下記 BASE_DIR ブロックに統一済み。
# 本ファイルはプロジェクトルート直下に置く（親ディレクトリはルートではない想定）。

# ===== パス制御（ここを置き換え）=====
# Nuitka は PyInstaller と異なり常に sys.frozen を立てない。__compiled__ を併用する。
if getattr(sys, "frozen", False) or globals().get("__compiled__") is not None:
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.abspath(os.path.dirname(__file__))

if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

os.chdir(BASE_DIR)

# 配布では Nuitka 成果物を ``app\bin`` に集約。DLL は EXE 横で解決されることが多い（必要時は shared_dll_bootstrap）。
if os.name == "nt":
    from core.shared_dll_bootstrap import ensure_shared_dll_search_path_for_layout

    ensure_shared_dll_search_path_for_layout(Path(BASE_DIR))
# ===== ここまで =====

__version__ = "0.2.4"

# bridge JSON 読み取り失敗などをファイルパス単位でカウントし、連続ポーリックを抑制する。
_fail_polls: dict[str, int] = {}

from core import core_env
from core.ribbon_public_to_svc import RIBBON_PUBLIC_TO_SVC_ACTION
from core.core_log import get_logger

logger = get_logger(__name__)
_TAG = core_env.LOG_MAIN_PREFIX


def _bridge_fail_key(p: Path) -> str:
    """同一ファイルを一意に指すキー（正規化パス文字列）。"""
    return str(p.resolve())


def _bump_bridge_fail(p: Path) -> int:
    """読み取り／転送失敗回数を 1 増やし、現在の累計を返す。"""
    k = _bridge_fail_key(p)
    _fail_polls[k] = _fail_polls.get(k, 0) + 1
    return _fail_polls[k]


def _clear_bridge_fail(p: Path) -> None:
    """ファイル処理が成功したあと、当該パスの失敗カウンタを捨てる。"""
    _fail_polls.pop(_bridge_fail_key(p), None)


def _min_file_age_sec() -> float:
    """bridge JSON を触るまで待つ最短経過秒（書き込み完了待ち）。"""
    return core_env.hc_main_min_file_age_sec()


def _max_bad_polls() -> int:
    """デコード等の失敗を何回まで再試行するかの上限。"""
    return core_env.hc_main_bad_file_max_polls()


def _load_bridge_json_dict(raw: bytes) -> dict:
    """utf-8（BOM 可）優先、旧 VBA ANSI 出力は cp932。失敗時は UnicodeDecodeError / JSONDecodeError / TypeError。"""
    last: BaseException | None = None
    for enc in ("utf-8-sig", "utf-8", "cp932"):
        try:
            text = raw.decode(enc)
            obj = json.loads(text)
            if isinstance(obj, dict):
                return obj
            raise TypeError(f"expected JSON object, got {type(obj).__name__}")
        except (UnicodeDecodeError, json.JSONDecodeError) as e:
            last = e
            continue
        except TypeError:
            raise
    if last:
        raise last
    raise ValueError("empty bridge request file")


def _get_ipc_root() -> Path:
    """bridge_requests / svc_requests など IPC サブフォルダの親。環境変数または tempfile 配下。"""
    forced = core_env.ipc_dir_raw()
    if forced:
        d = Path(forced)
    else:
        d = Path(tempfile.gettempdir()) / "csv_tool"
    d.mkdir(parents=True, exist_ok=True)
    return d


def _bridge_requests_dir() -> Path:
    """VBA / xlwings 側が書き込む JSON 要求の置き場。"""
    d = _get_ipc_root() / "bridge_requests"
    d.mkdir(parents=True, exist_ok=True)
    return d


def _svc_requests_dir() -> Path:
    """本プロセスが pickle 要求を書き込み、svc_server が読む置き場。"""
    d = _get_ipc_root() / "svc_requests"
    d.mkdir(parents=True, exist_ok=True)
    return d


def _atomic_write_pickle(path: Path, obj: object) -> None:
    """一時ファイルへ書き込み後 `os.replace` で置換し、読み手に半端なファイルを見せない。"""
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp = tempfile.mkstemp(
        prefix=path.name + ".", suffix=".tmp", dir=str(path.parent)
    )
    try:
        with os.fdopen(fd, "wb") as f:
            f.write(pickle.dumps(obj, protocol=pickle.HIGHEST_PROTOCOL))
            f.flush()
            os.fsync(f.fileno())
        os.replace(tmp, path)
    finally:
        try:
            if os.path.exists(tmp):
                os.remove(tmp)
        except OSError:
            pass


def _submit_svc_request(
    action: str,
    excel_hwnd: int,
    book_fullname: str,
    book_name: str,
    sheet_id: str,
    *,
    extra_kwargs: dict | None = None,
) -> None:
    """ribbon 由来の 1 件を svc 向け pickle 1 ファイルとしてキューイングする。"""
    req_dir = _svc_requests_dir()
    ts_ms = int(time.time() * 1000)
    req_path = req_dir / f"svc_req_{ts_ms}_{os.getpid()}.pkl"
    kwargs: dict = {
        "excel_hwnd": int(excel_hwnd or 0),
        "book_fullname": str(book_fullname or ""),
        "book_name": str(book_name or ""),
        "sheet_id": str(sheet_id or ""),
    }
    if extra_kwargs:
        for k, v in extra_kwargs.items():
            ks = str(k or "")
            if ks in ("excel_hwnd", "book_fullname", "book_name", "sheet_id"):
                continue
            kwargs[ks] = v
    req = {
        "action": action,
        "args": [],
        "kwargs": kwargs,
    }
    _atomic_write_pickle(req_path, req)
    # 連続キューでファイル名のタイムスタンプが衝突しないよう、次要求まで短い間隔を空ける。
    time.sleep(0.5)


def _normalize_bridge_selection_areas(raw: object) -> list[str] | None:
    """bridge JSON の selection_areas を list[str] に正規化。無効なら None。"""
    if not isinstance(raw, list):
        return None
    out: list[str] = []
    for x in raw:
        if isinstance(x, str):
            s = x.strip()
            if s:
                out.append(s)
    return out if out else None


def _normalize_bridge_count_large_value(raw: object) -> int | None:
    """bridge JSON の CountLarge 系数値を非負 int に正規化。無効なら None。"""
    if raw is None:
        return None
    if isinstance(raw, bool):
        return None
    try:
        if isinstance(raw, float):
            if raw != raw or raw < 0:
                return None
            return int(raw)
        n = int(raw)
        if n < 0:
            return None
        return n
    except (TypeError, ValueError, OverflowError):
        return None


def _process_bridge_request(data: dict) -> bool:
    """bridge の 1 レコードを解釈し、対応する svc アクションへ転送する。処理したら True。"""
    action = (data.get("action") or "").strip()
    if not action:
        return False
    svc_action = RIBBON_PUBLIC_TO_SVC_ACTION.get(action)
    if not svc_action:
        return False
    hwnd = int(data.get("hwnd", 0) or 0)
    sheet_id = str(data.get("sheet_id") or "").strip()
    book_fullname = str(data.get("book_fullname") or "").strip()
    book_name = str(data.get("book_name") or "").strip()
    extra: dict = {}
    if "payload" in data:
        extra["payload"] = data.get("payload")
    if "selection_areas" in data:
        sa = _normalize_bridge_selection_areas(data.get("selection_areas"))
        if sa is not None:
            extra["selection_areas"] = sa
        else:
            try:
                logger.warning(
                    "%s ignoring invalid or empty selection_areas type=%r",
                    _TAG,
                    type(data.get("selection_areas")).__name__,
                )
            except Exception:
                pass
    if action == "check_duplicates":
        scl = _normalize_bridge_count_large_value(data.get("selection_count_large"))
        shcl = _normalize_bridge_count_large_value(data.get("sheet_cells_count_large"))
        if scl is not None:
            extra["selection_count_large"] = scl
        if shcl is not None:
            extra["sheet_cells_count_large"] = shcl
    _submit_svc_request(
        svc_action, hwnd, book_fullname, book_name, sheet_id, extra_kwargs=extra or None
    )
    try:
        logger.info(
            "%s forwarded public_action=%s svc_action=%s sheet_id=%s hwnd=%s",
            _TAG,
            action,
            svc_action,
            sheet_id,
            hwnd,
        )
    except Exception:
        pass
    return True


def _poll_once() -> bool:
    """bridge_requests 内の古い順に 1 ファイルだけ処理する。処理して削除できれば True。"""
    req_dir = _bridge_requests_dir()
    min_age = _min_file_age_sec()
    max_polls = _max_bad_polls()
    try:
        files = sorted(req_dir.glob("*.json"), key=lambda p: p.stat().st_mtime)
    except Exception:
        return False
    for p in files:
        try:
            # 書き込み途中の JSON を掴まないよう、更新時刻がしきい値より古いものだけ対象にする。
            try:
                st = p.stat()
            except OSError:
                continue
            if time.time() - st.st_mtime < min_age:
                continue
            raw_bytes = p.read_bytes()
            try:
                data = _load_bridge_json_dict(raw_bytes)
            except TypeError as e:
                try:
                    logger.warning(
                        "%s reject file=%s: %s: %s",
                        _TAG,
                        p,
                        type(e).__name__,
                        e,
                    )
                except Exception:
                    pass
                try:
                    p.unlink(missing_ok=True)
                except OSError:
                    pass
                _clear_bridge_fail(p)
                continue
            except (UnicodeDecodeError, json.JSONDecodeError, ValueError) as e:
                n = _bump_bridge_fail(p)
                try:
                    logger.debug(
                        "%s decode wait file=%s n=%s %s: %s",
                        _TAG,
                        p,
                        n,
                        type(e).__name__,
                        e,
                    )
                except Exception:
                    pass
                if n > max_polls:
                    try:
                        logger.warning(
                            "%s giving up file=%s after %s polls: %s: %s",
                            _TAG,
                            p,
                            n,
                            type(e).__name__,
                            e,
                        )
                    except Exception:
                        pass
                    try:
                        p.unlink(missing_ok=True)
                    except OSError:
                        pass
                    _clear_bridge_fail(p)
                continue

            _clear_bridge_fail(p)
            try:
                ok = _process_bridge_request(data)
            except Exception as e:
                n = _bump_bridge_fail(p)
                try:
                    logger.debug(
                        "%s forward retry file=%s n=%s %s: %s",
                        _TAG,
                        p,
                        n,
                        type(e).__name__,
                        e,
                    )
                except Exception:
                    pass
                if n > max_polls:
                    try:
                        logger.warning(
                            "%s giving up file=%s after %s forward errors: %s: %s",
                            _TAG,
                            p,
                            n,
                            type(e).__name__,
                            e,
                        )
                    except Exception:
                        pass
                    try:
                        p.unlink(missing_ok=True)
                    except OSError:
                        pass
                    _clear_bridge_fail(p)
                continue

            if not ok:
                try:
                    logger.warning(
                        "%s unsupported file=%s action=%r",
                        _TAG,
                        p,
                        (data.get("action") or "")[:80],
                    )
                except Exception:
                    pass
                try:
                    p.unlink(missing_ok=True)
                except OSError:
                    pass
                _clear_bridge_fail(p)
                continue

            try:
                p.unlink(missing_ok=True)
            except OSError:
                pass
            _clear_bridge_fail(p)
            return True
        except Exception as e:
            try:
                logger.warning(
                    "%s poll skip file=%s: %s: %s",
                    _TAG,
                    p,
                    type(e).__name__,
                    e,
                )
            except Exception:
                pass
            try:
                p.unlink(missing_ok=True)
            except OSError:
                pass
            _clear_bridge_fail(p)
    return False


def _shutdown_requested() -> bool:
    """control/shutdown.flag があればメインループを終了する。"""
    try:
        flag = _get_ipc_root() / "control" / "shutdown.flag"
        return flag.exists()
    except Exception:
        return False


def _acquire_mutex() -> bool:
    """Windows: 単一インスタンス。`HC_MAIN_RUNNER` と旧名 `HC_BRIDGE_RUNNER` の両方を取得する。

    - 他プロセスがいずれかを保持していれば失敗（二重起動防止）。
    - 旧 svc_host は `HC_BRIDGE_RUNNER` のみ参照するため、新プロセスも旧 Mutex を取り、互換を保つ。
    """
    if os.name != "nt":
        return True
    try:
        import ctypes
        from ctypes import wintypes

        k = ctypes.windll.kernel32
        SYNCHRONIZE = 0x00100000
        ERROR_ALREADY_EXISTS = 183
        legacy = "Global\\HC_BRIDGE_RUNNER"
        main_name = "Global\\HC_MAIN_RUNNER"
        k.OpenMutexW.restype = wintypes.HANDLE
        k.OpenMutexW.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.LPCWSTR]
        h_probe = k.OpenMutexW(SYNCHRONIZE, False, wintypes.LPCWSTR(legacy))
        if h_probe:
            k.CloseHandle(h_probe)
            return False
        k.CreateMutexW.restype = wintypes.HANDLE
        k.GetLastError.restype = wintypes.DWORD
        k.SetLastError(0)
        h_main = k.CreateMutexW(None, False, wintypes.LPCWSTR(main_name))
        if h_main and k.GetLastError() == ERROR_ALREADY_EXISTS:
            k.CloseHandle(h_main)
            return False
        if not h_main:
            return False
        k.SetLastError(0)
        h_legacy = k.CreateMutexW(None, False, wintypes.LPCWSTR(legacy))
        if h_legacy and k.GetLastError() == ERROR_ALREADY_EXISTS:
            k.CloseHandle(h_legacy)
            k.CloseHandle(h_main)
            return False
        return bool(h_legacy)
    except Exception:
        return True


def main() -> int:
    """Mutex 取得後、起動時スイープを行い bridge ポーリングを無限ループする。終了コード 0。"""
    if not _acquire_mutex():
        logger.info("%s already running (mutex), exit", _TAG)
        return 0
    logger.info("%s started pid=%s", _TAG, os.getpid())
    try:
        from core.ipc_cleanup import run_bridge_startup_sweeps

        run_bridge_startup_sweeps(_get_ipc_root())
    except Exception:
        pass
    idle_sec = core_env.hc_main_poll_sec()
    while True:
        if _shutdown_requested():
            logger.info("%s shutdown requested", _TAG)
            break
        _poll_once()
        time.sleep(idle_sec)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
