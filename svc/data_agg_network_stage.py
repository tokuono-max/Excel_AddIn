# -*- coding: utf-8 -*-
"""ネットワーク上の集約ソースをローカル TEMP へステージングしてから読む。"""
from __future__ import annotations

import hashlib
import logging
import os
import shutil
import threading
import time
import uuid
from concurrent.futures import ThreadPoolExecutor, as_completed
from contextlib import contextmanager
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Iterator, Sequence

from core.core_log import get_data_agg_diag_logger, get_logger
from svc.data_agg_path_network import path_is_network
from svc.data_agg_path_norm import normalize_source_path, normalize_source_path_literal

logger = get_logger(__name__)
_diag = get_data_agg_diag_logger()

StageProgressCallback = Callable[[int, int, str], None]
FileStagedCallback = Callable[[int, str, str], None]
CancelCheck = Callable[..., None]

_registry_lock = threading.Lock()
_registered_stage_dirs: set[str] = set()


def register_stage_dir(path: Path) -> None:
    with _registry_lock:
        _registered_stage_dirs.add(str(path.resolve()))


def unregister_stage_dir(path: Path | None) -> None:
    if path is None:
        return
    try:
        key = str(path.resolve())
    except OSError:
        key = str(path)
    with _registry_lock:
        _registered_stage_dirs.discard(key)


# ロック解放待ちの短いリトライ（秒）。最終失敗時は登録を残し次回 cleanup_all で再試行する。
_REMOVE_RETRY_DELAYS_SEC: tuple[float, ...] = (0.05, 0.2, 0.5, 1.0)


def _force_remove_stage_dir(path: Path, *, retries: Sequence[float] | None = None) -> bool:
    """ステージ dir を削除する。成功時のみ True（dir が無い場合も True）。"""
    if not path.exists():
        return True
    if not path.is_dir():
        return False
    delays = list(_REMOVE_RETRY_DELAYS_SEC if retries is None else retries)
    attempts = 1 + len(delays)
    last_err: BaseException | None = None
    for i in range(attempts):
        try:
            _remove_orphan_part_files(path)
            _remove_empty_dirs(path, remove_root=True)
            if path.is_dir():
                shutil.rmtree(path, ignore_errors=True)
            if not path.exists():
                return True
        except OSError as e:
            last_err = e
        if i < len(delays):
            try:
                time.sleep(float(delays[i]))
            except Exception:
                pass
    try:
        logger.warning(
            "[DATA_AGG_STAGE] force remove incomplete dir=%s exists=%s err=%s",
            path,
            path.exists(),
            last_err,
        )
    except Exception:
        pass
    return not path.exists()


def cleanup_all_network_stage_dirs(*, prune_orphans: bool = True) -> int:
    """登録済みステージ dir を削除。マスタ／シナリオデバッグ終了・キャンセル後の掃除用。

    削除に失敗した dir はレジストリに残し、次回呼び出しで再試行する。
    prune_orphans 時は非登録残骸を年齢 0（即時）で掃除する（明示クリーンアップ経路）。
    """
    with _registry_lock:
        dirs = [Path(p) for p in list(_registered_stage_dirs)]
    removed = 0
    for sd in dirs:
        if _force_remove_stage_dir(sd):
            removed += 1
            unregister_stage_dir(sd)
        else:
            try:
                logger.warning(
                    "[DATA_AGG_STAGE] cleanup_all kept registered (still on disk) dir=%s",
                    sd,
                )
            except Exception:
                pass
    if prune_orphans:
        # 明示掃除経路では年齢ゲートなし（部分削除残骸を即拾う）
        removed += _prune_stale_stage_dirs(max_age_sec=0)
    return removed


def wipe_data_agg_stage_base(*, include_registered: bool = True) -> int:
    """
    ``data_agg_stage`` 配下の UUID dir をすべて削除試行する（デバッグダイアログ閉鎖用）。

    include_registered=True のときは登録中も含めて消す（閉鎖時は自プロセスの
    ステージを残さない）。失敗した登録 dir はレジストリに残す。
    """
    base = _stage_temp_base()
    if not base.is_dir():
        return 0
    removed = 0
    with _registry_lock:
        active = set(_registered_stage_dirs)
    try:
        children = list(base.iterdir())
    except OSError:
        return 0
    for child in children:
        if not child.is_dir():
            continue
        try:
            key = str(child.resolve())
        except OSError:
            key = str(child)
        if (not include_registered) and key in active:
            continue
        if _force_remove_stage_dir(child):
            removed += 1
            unregister_stage_dir(child)
            try:
                logger.info("[DATA_AGG_STAGE] wipe dir=%s", child)
            except Exception:
                pass
        else:
            try:
                logger.warning(
                    "[DATA_AGG_STAGE] wipe incomplete dir=%s registered=%s",
                    child,
                    key in active,
                )
            except Exception:
                pass
    try:
        _remove_empty_dirs(base, remove_root=False)
    except Exception:
        pass
    return removed


@dataclass
class NetworkStageBatch:
    """io_paths で読取、display_paths で __file_path 等の表示用パスを保持する。"""

    display_paths: list[str]
    io_paths: list[str]
    stage_dir: Path | None = None
    copy_ms: int = 0
    staged_files: int = 0
    _norm_to_display: dict[str, str] = field(default_factory=dict, repr=False)

    def display_path_for_io(self, io_path: str | Path) -> str:
        key = normalize_source_path(io_path)
        return self._norm_to_display.get(key, str(io_path))

    def cleanup(self) -> bool:
        """ステージ dir を削除する。成功（または元々無し）なら True。

        削除失敗時はレジストリに残し stage_dir も保持して再試行可能にする。
        """
        sd = self.stage_dir
        if sd is None:
            return True
        if not sd.exists():
            unregister_stage_dir(sd)
            self.stage_dir = None
            return True
        ok = _force_remove_stage_dir(sd)
        if ok:
            try:
                logger.info("[DATA_AGG_STAGE] cleanup dir=%s", sd)
            except Exception:
                pass
            unregister_stage_dir(sd)
            self.stage_dir = None
            return True
        try:
            logger.warning(
                "[DATA_AGG_STAGE] cleanup incomplete dir=%s (kept registered for retry)",
                sd,
            )
        except Exception:
            pass
        return False


def _stage_temp_base() -> Path:
    return Path(os.environ.get("TEMP", "C:\\Temp")) / "csv_tool" / "data_agg_stage"


_DEFAULT_ORPHAN_MAX_AGE_SEC = 3600


def _prune_stale_stage_dirs(*, max_age_sec: int = _DEFAULT_ORPHAN_MAX_AGE_SEC) -> int:
    """
    非登録の古いステージ dir を中身ごと削除（.part・途中コピー含む）。
    実行中 dir はレジストリで保護する。戻り値は削除した件数。
    """
    base = _stage_temp_base()
    if not base.is_dir():
        return 0
    with _registry_lock:
        active = set(_registered_stage_dirs)
    cutoff = time.time() - max(0, int(max_age_sec))
    removed = 0
    try:
        for child in list(base.iterdir()):
            if not child.is_dir():
                continue
            try:
                key = str(child.resolve())
            except OSError:
                key = str(child)
            if key in active:
                continue
            try:
                if child.stat().st_mtime >= cutoff and max_age_sec > 0:
                    continue
                if _force_remove_stage_dir(child):
                    removed += 1
                    try:
                        logger.info("[DATA_AGG_STAGE] prune orphan dir=%s", child)
                    except Exception:
                        pass
            except OSError:
                pass
        _remove_empty_dirs(base, remove_root=False)
    except OSError:
        pass
    return removed


def _remove_empty_dirs(root: Path, *, remove_root: bool) -> None:
    """root 配下の空ディレクトリを下から削除する。"""
    if not root.is_dir():
        return
    try:
        for child in list(root.iterdir()):
            if child.is_dir():
                _remove_empty_dirs(child, remove_root=True)
    except OSError:
        return
    if not remove_root:
        return
    try:
        if not any(root.iterdir()):
            root.rmdir()
    except OSError:
        pass


def _remove_orphan_part_files(root: Path) -> None:
    """コピー中断で残った .part を削除する。"""
    if not root.is_dir():
        return
    try:
        for p in root.rglob("*.part"):
            try:
                if p.is_file():
                    p.unlink(missing_ok=True)
            except OSError:
                pass
    except OSError:
        pass


def _stage_rel_name(src: Path, *, scan_root: Path | None) -> str:
    if scan_root is not None:
        try:
            src_lit = normalize_source_path_literal(src)
            root_lit = normalize_source_path_literal(scan_root)
            rel = Path(src_lit).relative_to(Path(root_lit))
            return rel.as_posix()
        except (OSError, ValueError):
            pass
    digest = hashlib.sha1(normalize_source_path_literal(src).encode("utf-8")).hexdigest()[:12]
    return f"{digest}_{src.name}"


def _copy_file_atomic(src: Path, dst: Path) -> None:
    dst.parent.mkdir(parents=True, exist_ok=True)
    tmp = dst.with_suffix(dst.suffix + ".part")
    try:
        if tmp.is_file():
            tmp.unlink(missing_ok=True)
    except OSError:
        pass
    shutil.copy2(src, tmp)
    os.replace(tmp, dst)


def _poll_cancel(cancel_check: CancelCheck | None) -> None:
    if cancel_check is not None:
        cancel_check(force=True)


def _copy_one_network_file(
    disp: str,
    stage_dir: Path,
    scan_root: Path | None,
) -> tuple[str, str, str]:
    """Returns (display_path, io_path, norm_key)."""
    src = Path(disp)
    rel = _stage_rel_name(src, scan_root=scan_root)
    dst = stage_dir / rel
    _copy_file_atomic(src, dst)
    io_s = str(dst)
    return disp, io_s, normalize_source_path(io_s)


def build_network_stage_batch(
    file_paths: Sequence[str | Path],
    *,
    scan_root: str | Path | None = None,
    enabled: bool = True,
    cancel_check: CancelCheck | None = None,
    progress_callback: StageProgressCallback | None = None,
    copy_workers: int | None = None,
    on_file_staged: FileStagedCallback | None = None,
) -> NetworkStageBatch:
    """
    ネットワークパスを TEMP へコピー。ローカルのみなら io=display のまま返す。
    enabled=False のときもパススルー。
    on_file_staged: 各ファイルの io パス確定直後 (index, display, io) を通知（パイプライン用）。
    """
    display = [str(p) for p in file_paths]

    def _notify_staged(idx: int, disp: str, io_s: str) -> None:
        if on_file_staged is not None:
            on_file_staged(idx, disp, io_s)

    if not enabled or not display:
        batch = NetworkStageBatch(display_paths=display, io_paths=list(display))
        if on_file_staged is not None:
            for i, (disp, io_s) in enumerate(zip(display, batch.io_paths)):
                _notify_staged(i, disp, io_s)
        return batch

    need_stage = any(path_is_network(p) for p in display)
    if not need_stage:
        batch = NetworkStageBatch(display_paths=display, io_paths=list(display))
        if on_file_staged is not None:
            for i, (disp, io_s) in enumerate(zip(display, batch.io_paths)):
                _notify_staged(i, disp, io_s)
        return batch

    _prune_stale_stage_dirs()
    root = Path(scan_root) if scan_root else None
    stage_dir = _stage_temp_base() / uuid.uuid4().hex
    stage_dir.mkdir(parents=True, exist_ok=True)
    register_stage_dir(stage_dir)
    n_total = len(display)
    staged = 0
    copy_ms = 0
    workers = 1

    try:
        io_out: list[str | None] = [None] * len(display)
        norm_map: dict[str, str] = {}
        t0 = time.perf_counter()

        network_jobs: list[tuple[int, str]] = []
        for i, disp in enumerate(display):
            if not path_is_network(disp):
                io_out[i] = disp
                norm_map[normalize_source_path_literal(disp)] = disp
                _notify_staged(i, disp, io_out[i])
                continue
            network_jobs.append((i, disp))

        if copy_workers is None:
            from core import core_env

            copy_workers = core_env.data_agg_network_stage_copy_workers(
                n_files=max(1, len(network_jobs))
            )
        workers = max(1, int(copy_workers or 1))
        done_n = 0

        def _report_progress(disp: str) -> None:
            nonlocal done_n
            done_n += 1
            if progress_callback is not None:
                try:
                    progress_callback(done_n, len(network_jobs), Path(disp).name)
                except Exception:
                    pass

        def _apply_copy_result(
            idx: int, disp: str, io_s: str, norm_key: str, *, ok: bool
        ) -> None:
            nonlocal staged
            if ok:
                io_out[idx] = io_s
                norm_map[norm_key] = disp
                staged += 1
            else:
                io_out[idx] = disp
                norm_map[normalize_source_path_literal(disp)] = disp
            _notify_staged(idx, disp, str(io_out[idx]))

        if network_jobs:
            _poll_cancel(cancel_check)
            if workers <= 1 or len(network_jobs) == 1:
                for idx, disp in network_jobs:
                    _poll_cancel(cancel_check)
                    try:
                        d, io_s, nk = _copy_one_network_file(disp, stage_dir, root)
                        _apply_copy_result(idx, d, io_s, nk, ok=True)
                    except OSError as e:
                        logger.warning(
                            "[DATA_AGG_STAGE] copy failed src=%s err=%s; fallback to direct read",
                            disp,
                            e,
                        )
                        _apply_copy_result(idx, disp, disp, "", ok=False)
                    _report_progress(disp)
            else:
                with ThreadPoolExecutor(max_workers=workers) as pool:
                    fut_map = {
                        pool.submit(_copy_one_network_file, disp, stage_dir, root): (
                            idx,
                            disp,
                        )
                        for idx, disp in network_jobs
                    }
                    for fut in as_completed(fut_map):
                        _poll_cancel(cancel_check)
                        idx, disp = fut_map[fut]
                        try:
                            d, io_s, nk = fut.result()
                            _apply_copy_result(idx, d, io_s, nk, ok=True)
                        except OSError as e:
                            logger.warning(
                                "[DATA_AGG_STAGE] copy failed src=%s err=%s; fallback to direct read",
                                disp,
                                e,
                            )
                            _apply_copy_result(idx, disp, disp, "", ok=False)
                        except Exception as e:
                            logger.warning(
                                "[DATA_AGG_STAGE] copy error src=%s err=%s; fallback to direct read",
                                disp,
                                e,
                            )
                            _apply_copy_result(idx, disp, disp, "", ok=False)
                        _report_progress(disp)

        _remove_empty_dirs(stage_dir, remove_root=False)
        _remove_orphan_part_files(stage_dir)

        io_final = [str(p) if p is not None else display[i] for i, p in enumerate(io_out)]
        copy_ms = int((time.perf_counter() - t0) * 1000)
        batch = NetworkStageBatch(
            display_paths=display,
            io_paths=io_final,
            stage_dir=stage_dir,
            copy_ms=copy_ms,
            staged_files=staged,
            _norm_to_display=norm_map,
        )
    except BaseException:
        _force_remove_stage_dir(stage_dir)
        unregister_stage_dir(stage_dir)
        raise
    try:
        logger.info(
            "[DATA_AGG_STAGE] prepared dir=%s staged=%s total=%s copy_ms=%s workers=%s",
            stage_dir,
            staged,
            n_total,
            copy_ms,
            workers,
        )
        _diag.info(
            "[DATA_AGG_IO_PROFILE] stage_copy ms=%s staged_files=%s total_files=%s dir=%s workers=%s",
            copy_ms,
            staged,
            n_total,
            stage_dir,
            workers,
        )
    except Exception:
        pass
    return batch


@contextmanager
def network_stage_batch(
    file_paths: Sequence[str | Path],
    *,
    scan_root: str | Path | None = None,
    enabled: bool = True,
    cancel_check: CancelCheck | None = None,
    progress_callback: StageProgressCallback | None = None,
    copy_workers: int | None = None,
) -> Iterator[NetworkStageBatch]:
    batch = build_network_stage_batch(
        file_paths,
        scan_root=scan_root,
        enabled=enabled,
        cancel_check=cancel_check,
        progress_callback=progress_callback,
        copy_workers=copy_workers,
    )
    try:
        yield batch
    finally:
        batch.cleanup()
