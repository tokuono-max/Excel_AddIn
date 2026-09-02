# -*- coding: utf-8 -*-
"""
TEMP: ネットワーク参照 vs ローカル比較用 I/O 計測。
本改善完了後に本ファイルと、各所の「io_profile_temp」参照を削除する。

有効化: DATA_AGG_IO_PROFILE=1（別名 HC_DIAG_DATA_AGG_IO_PROFILE=1）
出力先: %TEMP%\\csv_tool\\hc_csv_diag.log（HC_LOG_DIAG=1 等で診断ログ有効時）
タグ: [DATA_AGG_IO_PROFILE]
"""
from __future__ import annotations

import threading
from dataclasses import dataclass
from pathlib import Path

from svc.data_agg_path_network import path_class as _path_class_impl


def enabled() -> bool:
    from core import core_env

    return core_env.data_agg_io_profile_enabled()


def _path_key(file_path: str | Path) -> str:
    try:
        return str(Path(file_path).resolve())
    except OSError:
        return str(file_path)


def path_class(file_path: str | Path) -> str:
    return _path_class_impl(file_path)


@dataclass
class _FileIoStats:
    cache_open_ms: int = 0
    cache_open_count: int = 0
    sheet_name_open_ms: int = 0
    sheet_name_open_count: int = 0
    hidden_open_ms: int = 0
    hidden_open_count: int = 0
    materialize_ms: int = 0
    materialize_count: int = 0

    def total_open_count(self) -> int:
        return (
            self.cache_open_count
            + self.sheet_name_open_count
            + self.hidden_open_count
        )


_tls = threading.local()
_global_lock = threading.Lock()
_global_file_stats: dict[str, _FileIoStats] = {}
_batch_lock = threading.Lock()
_batch_totals = _FileIoStats()
_stage_copy_ms: int | None = None
_scan_ms: int | None = None
_scan_files: int = 0
_scan_path: str = ""


def _file_map() -> dict[str, _FileIoStats]:
    m = getattr(_tls, "by_path", None)
    if m is None:
        m = {}
        _tls.by_path = m
    return m


def _stats_for(path_key: str) -> _FileIoStats:
    m = _file_map()
    st = m.get(path_key)
    if st is None:
        st = _FileIoStats()
        m[path_key] = st
    return st


def record_cache_open(file_path: str | Path, seconds: float) -> None:
    if not enabled() or seconds <= 0:
        return
    key = _path_key(file_path)
    st = _stats_for(key)
    st.cache_open_ms += int(seconds * 1000.0 + 0.5)
    st.cache_open_count += 1


def record_sheet_name_open(file_path: str | Path, seconds: float) -> None:
    if not enabled() or seconds <= 0:
        return
    key = _path_key(file_path)
    st = _stats_for(key)
    st.sheet_name_open_ms += int(seconds * 1000.0 + 0.5)
    st.sheet_name_open_count += 1


def record_hidden_open(file_path: str | Path, seconds: float) -> None:
    if not enabled() or seconds <= 0:
        return
    key = _path_key(file_path)
    st = _stats_for(key)
    st.hidden_open_ms += int(seconds * 1000.0 + 0.5)
    st.hidden_open_count += 1


def _merge_stats(dst: _FileIoStats, src: _FileIoStats) -> None:
    dst.cache_open_ms += src.cache_open_ms
    dst.cache_open_count += src.cache_open_count
    dst.sheet_name_open_ms += src.sheet_name_open_ms
    dst.sheet_name_open_count += src.sheet_name_open_count
    dst.hidden_open_ms += src.hidden_open_ms
    dst.hidden_open_count += src.hidden_open_count
    dst.materialize_ms += src.materialize_ms
    dst.materialize_count += src.materialize_count


def flush_file_stats(file_path: str | Path) -> None:
    """並列ワーカー終了時: スレッドローカル計測をグローバルへ集約。"""
    if not enabled():
        return
    key = _path_key(file_path)
    m = getattr(_tls, "by_path", None)
    if not isinstance(m, dict):
        return
    st = m.pop(key, None)
    if st is None:
        return
    with _global_lock:
        g = _global_file_stats.get(key)
        if g is None:
            _global_file_stats[key] = st
        else:
            _merge_stats(g, st)


def record_stage_copy_ms(ms: int) -> None:
    if not enabled():
        return
    global _stage_copy_ms
    _stage_copy_ms = int(ms)


def record_materialize(file_path: str | Path, seconds: float) -> None:
    if not enabled() or seconds <= 0:
        return
    key = _path_key(file_path)
    st = _stats_for(key)
    st.materialize_ms += int(seconds * 1000.0 + 0.5)
    st.materialize_count += 1


def consume_file_stats(file_path: str | Path) -> _FileIoStats | None:
    if not enabled():
        return None
    key = _path_key(file_path)
    st = _FileIoStats()
    m = getattr(_tls, "by_path", None)
    if isinstance(m, dict):
        tls_st = m.pop(key, None)
        if tls_st is not None:
            _merge_stats(st, tls_st)
    with _global_lock:
        g = _global_file_stats.pop(key, None)
        if g is not None:
            _merge_stats(st, g)
    with _batch_lock:
        _merge_stats(_batch_totals, st)
    return st


def record_scan(*, ms: int, n_files: int, start_path: str | Path) -> None:
    if not enabled():
        return
    global _scan_ms, _scan_files, _scan_path
    _scan_ms = int(ms)
    _scan_files = int(n_files)
    _scan_path = str(start_path)
    try:
        from core.core_log import get_data_agg_diag_logger

        get_data_agg_diag_logger().info(
            "[DATA_AGG_IO_PROFILE] scan ms=%s files=%s path_class=%s path=%s",
            _scan_ms,
            _scan_files,
            path_class(start_path),
            start_path,
        )
    except Exception:
        pass


def emit_per_file(
    *,
    scenario_id: str,
    caller: str,
    file_index: int,
    n_files: int,
    file_path: str | Path,
    profile_io_path: str | Path | None = None,
    wall_total_ms: int | None = None,
    pf_open_ms: int | None = None,
    pf_read_extract_ms: int | None = None,
) -> None:
    if not enabled():
        return
    io_path = profile_io_path if profile_io_path is not None else file_path
    st = consume_file_stats(io_path)
    if st is None:
        return
    try:
        from core.core_log import get_data_agg_diag_logger

        get_data_agg_diag_logger().info(
            "[DATA_AGG_IO_PROFILE] per_file scenario=%s caller=%s i=%s/%s file=%s "
            "path_class=%s cache_open_ms=%s cache_open_n=%s sheet_name_open_ms=%s "
            "sheet_name_open_n=%s hidden_open_ms=%s hidden_open_n=%s materialize_ms=%s "
            "materialize_n=%s total_open_n=%s legacy_open_ms=%s read_extract_ms=%s "
            "wall_total_ms=%s",
            scenario_id,
            caller or "-",
            file_index,
            n_files,
            Path(str(file_path)).name,
            path_class(file_path),
            st.cache_open_ms,
            st.cache_open_count,
            st.sheet_name_open_ms,
            st.sheet_name_open_count,
            st.hidden_open_ms,
            st.hidden_open_count,
            st.materialize_ms,
            st.materialize_count,
            st.total_open_count(),
            pf_open_ms if pf_open_ms is not None else "-",
            pf_read_extract_ms if pf_read_extract_ms is not None else "-",
            wall_total_ms if wall_total_ms is not None else "-",
        )
    except Exception:
        pass


def emit_batch_summary(
    *,
    scenario_id: str,
    caller: str,
    n_files: int,
    n_items: int,
    parallel_workers: int,
    compute_total_ms: int,
) -> None:
    if not enabled():
        return
    with _batch_lock:
        bt = _FileIoStats(
            cache_open_ms=_batch_totals.cache_open_ms,
            cache_open_count=_batch_totals.cache_open_count,
            sheet_name_open_ms=_batch_totals.sheet_name_open_ms,
            sheet_name_open_count=_batch_totals.sheet_name_open_count,
            hidden_open_ms=_batch_totals.hidden_open_ms,
            hidden_open_count=_batch_totals.hidden_open_count,
            materialize_ms=_batch_totals.materialize_ms,
            materialize_count=_batch_totals.materialize_count,
        )
        scan_ms = _scan_ms
        scan_files = _scan_files
        scan_path = _scan_path
        stage_ms = _stage_copy_ms
        _reset_batch_state_unlocked()
    try:
        from core.core_log import get_data_agg_diag_logger

        get_data_agg_diag_logger().info(
            "[DATA_AGG_IO_PROFILE] batch_summary scenario=%s caller=%s files=%s items=%s "
            "workers=%s scan_ms=%s scan_files=%s scan_path=%s stage_copy_ms=%s "
            "compute_total_ms=%s sum_cache_open_ms=%s sum_cache_open_n=%s "
            "sum_sheet_name_open_ms=%s sum_sheet_name_open_n=%s sum_hidden_open_ms=%s "
            "sum_hidden_open_n=%s sum_materialize_ms=%s sum_materialize_n=%s "
            "sum_total_open_n=%s",
            scenario_id,
            caller or "-",
            n_files,
            n_items,
            parallel_workers,
            scan_ms if scan_ms is not None else "-",
            scan_files,
            scan_path or "-",
            stage_ms if stage_ms is not None else "-",
            compute_total_ms,
            bt.cache_open_ms,
            bt.cache_open_count,
            bt.sheet_name_open_ms,
            bt.sheet_name_open_count,
            bt.hidden_open_ms,
            bt.hidden_open_count,
            bt.materialize_ms,
            bt.materialize_count,
            bt.total_open_count(),
        )
    except Exception:
        pass


def _reset_batch_state_unlocked() -> None:
    global _scan_ms, _scan_files, _scan_path, _stage_copy_ms
    _batch_totals.cache_open_ms = 0
    _batch_totals.cache_open_count = 0
    _batch_totals.sheet_name_open_ms = 0
    _batch_totals.sheet_name_open_count = 0
    _batch_totals.hidden_open_ms = 0
    _batch_totals.hidden_open_count = 0
    _batch_totals.materialize_ms = 0
    _batch_totals.materialize_count = 0
    _scan_ms = None
    _scan_files = 0
    _scan_path = ""
    _stage_copy_ms = None
    _global_file_stats.clear()


def reset_batch_state() -> None:
    """テスト用。"""
    if not enabled():
        return
    with _batch_lock:
        _reset_batch_state_unlocked()
    with _global_lock:
        _global_file_stats.clear()
    m = getattr(_tls, "by_path", None)
    if isinstance(m, dict):
        m.clear()
