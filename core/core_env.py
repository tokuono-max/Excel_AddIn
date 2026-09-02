# -*- coding: utf-8 -*-
"""
Module: core_env
Updated: 2026-04-12
Purpose:
    CSV Tool 用環境変数の正規名とレガシー別名の解決を 1 か所に集約する。
    ドキュメント: docs/environment_variables.md
"""

from __future__ import annotations

import os
from typing import Optional

# 子プロセス（UI サーバ等）へ Excel ウィンドウを伝えるための環境変数名（値の設定は set_excel_hwnd_for_spawn）。
ENV_EXCEL_HWND = "HC_EXCEL_HWND"


def truthy(val: Optional[str], *, empty_means_false: bool = True) -> bool:
    """1 / true / yes / on / y（大小無視）を真とする。"""
    if val is None:
        return False
    s = str(val).strip()
    if empty_means_false and s == "":
        return False
    return s.lower() in ("1", "true", "yes", "on", "y")


def get(name: str, default: Optional[str] = None) -> Optional[str]:
    v = os.environ.get(name)
    if v is not None and str(v).strip() != "":
        return v
    return default


def get_first(*names: str, default: Optional[str] = None) -> Optional[str]:
    """複数名のうち最初に非空の値を返す。"""
    for n in names:
        v = os.environ.get(n)
        if v is not None and str(v).strip() != "":
            return v
    return default


def ipc_dir_raw() -> str:
    """IPC ルートの上書きパス。未設定なら空文字。"""
    return (get_first("HC_IPC_ROOT", "HC_QT_IPC_DIR") or "").strip()


def set_excel_hwnd_for_spawn(hwnd: int) -> None:
    """UI サーバ等の子プロセスが参照する Excel HWND を `ENV_EXCEL_HWND` に書き込む。"""
    os.environ[ENV_EXCEL_HWND] = str(int(hwnd))


def progress_window_startup_wait_sec() -> float:
    """進捗ウィンドウが立ち上がるまで待つ秒数（CSV 読込 `svc_csv_ld` 等）。

    環境変数 `HC_PROGRESS_WINDOW_STARTUP_WAIT_SEC`。未設定時は 1.0。解釈失敗時も 1.0。
    値は 0.0〜10.0 にクランプする。
    """
    raw = get("HC_PROGRESS_WINDOW_STARTUP_WAIT_SEC")
    if raw is None:
        return 1.0
    try:
        return max(0.0, min(10.0, float(str(raw).strip())))
    except ValueError:
        return 1.0


def log_diag_enabled() -> bool:
    """診断ログファイル・詳細トレースのマスタ（HC_LOG_DIAG または従来 HC_DEBUG）。"""
    if truthy(os.environ.get("HC_LOG_DIAG"), empty_means_false=False):
        return True
    if truthy(os.environ.get("HC_DEBUG"), empty_means_false=False):
        return True
    return False


def log_perf_enabled() -> bool:
    """計測専用ログ（hc_csv_perf.log）。"""
    return truthy(os.environ.get("HC_LOG_PERF"), empty_means_false=False)


def csv_sp_conflict_hwnd_diag_enabled() -> bool:
    """ui_server: 分割・同名確認の HWND 診断（HC_CSV_SP_CONFLICT_HWND_DIAG）。単独で hc_csv_diag.log を有効化する。"""
    return truthy(os.environ.get("HC_CSV_SP_CONFLICT_HWND_DIAG"), empty_means_false=False)


def ui_native_file_diag_enabled() -> bool:
    """ネイティブファイル選択の FG/Z 順調査（HC_UI_NATIVE_FILE_DIAG）。単独で hc_csv_diag.log を有効化する。"""
    return truthy(os.environ.get("HC_UI_NATIVE_FILE_DIAG"), empty_means_false=False)


def ui_fg_diag_enabled() -> bool:
    """ui_server: Excel 前面／Z 順の調査ログ（HC_UI_FG_DIAG）。単独で hc_csv_diag.log を有効化する。"""
    return truthy(os.environ.get("HC_UI_FG_DIAG"), empty_means_false=False)


def ui_excel_lock_diag_enabled() -> bool:
    """enable_excel_window のロック／解除ログ（HC_UI_EXCEL_LOCK_DIAG）。症状 A（操作不能）切り分け。"""
    return truthy(os.environ.get("HC_UI_EXCEL_LOCK_DIAG"), empty_means_false=False)


def ui_window_caption_diag_enabled() -> bool:
    """ui_server: タイトルバー最小化/最大化・GWL_STYLE の調査（HC_UI_WINDOW_CAPTION_DIAG）。単独で hc_csv_diag.log を有効化する。"""
    return truthy(os.environ.get("HC_UI_WINDOW_CAPTION_DIAG"), empty_means_false=False)


def ui_keep_console_enabled() -> bool:
    """ui_server: True のとき起動時 FreeConsole をスキップ（HC_UI_KEEP_CONSOLE。CMD から起動して標準出力を見るデバッグ用）。"""
    return truthy(os.environ.get("HC_UI_KEEP_CONSOLE"), empty_means_false=False)


def data_agg_name_path_diag_enabled() -> bool:
    return truthy(os.environ.get("HC_DIAG_DATA_AGG_NAMES"), empty_means_false=False) or (
        os.environ.get("DATA_AGG_NAME_PATH_DIAG", "").strip() == "1"
    )


def data_agg_name_path_max_rows() -> int:
    raw = get_first(
        "HC_DIAG_DATA_AGG_NAMES_MAX_ROWS",
        "DATA_AGG_NAME_PATH_DIAG_MAX_ROWS",
        default="8",
    )
    try:
        v = int(str(raw or "8").strip())
        return max(1, min(v, 50))
    except ValueError:
        return 8


def data_agg_name_path_col_filter() -> str:
    return str(
        get_first("HC_DIAG_DATA_AGG_NAMES_COL", "DATA_AGG_NAME_PATH_DIAG_COL") or ""
    ).strip()


def data_agg_batch_timing_enabled() -> bool:
    return truthy(os.environ.get("HC_DIAG_DATA_AGG_BATCH_TIMING"), empty_means_false=False) or (
        os.environ.get("DATA_AGG_COMPUTE_BATCH_TIMING", "").strip() == "1"
    )


def data_agg_file_timing_enabled() -> bool:
    return truthy(os.environ.get("HC_DIAG_DATA_AGG_FILE_TIMING"), empty_means_false=False) or (
        os.environ.get("DATA_AGG_PER_FILE_TIMING", "").strip() == "1"
    )


def data_agg_io_profile_enabled() -> bool:
    """TEMP: ネットワーク I/O 比較計測（data_agg_io_profile_temp）。本改善後に削除予定。"""
    return truthy(
        os.environ.get("HC_DIAG_DATA_AGG_IO_PROFILE"), empty_means_false=False
    ) or (os.environ.get("DATA_AGG_IO_PROFILE", "").strip() == "1")


def data_agg_network_stage_enabled() -> bool:
    """
    ネットワーク上の入力ファイルを TEMP へコピーしてから読む。
    DATA_AGG_NETWORK_STAGE: 0=無効, 1/未設定=有効（ネットワークパスがあるときのみコピー）。
    """
    raw = os.environ.get("DATA_AGG_NETWORK_STAGE", "").strip().lower()
    if raw in ("0", "false", "no", "off"):
        return False
    return True


def data_agg_join_dump_enabled() -> bool:
    """結合キー検索書込みの診断ダンプ（`[DATA_AGG_JOIN_DUMP]`）を出すか。"""
    return truthy(os.environ.get("HC_DIAG_DATA_AGG_JOIN"), empty_means_false=False) or (
        os.environ.get("DATA_AGG_JOIN_DUMP", "").strip() == "1"
    )


def data_agg_join_dump_max_slices() -> int:
    raw = get_first(
        "HC_DIAG_DATA_AGG_JOIN_MAX_SLICES",
        "DATA_AGG_JOIN_DUMP_MAX_SLICES",
        default="8",
    )
    try:
        v = int(str(raw or "8").strip())
        return max(1, min(v, 50))
    except ValueError:
        return 8


def data_agg_join_dump_max_rows() -> int:
    """post_merge スナップショットで先頭から何行分の列値を出すか。"""
    raw = get_first(
        "HC_DIAG_DATA_AGG_JOIN_MAX_ROWS",
        "DATA_AGG_JOIN_DUMP_MAX_ROWS",
        default="12",
    )
    try:
        v = int(str(raw or "12").strip())
        return max(1, min(v, 100))
    except ValueError:
        return 12


def data_agg_join_dump_col_filter() -> str:
    """空でなければ、この文字列が代入先列名に含まれる項目だけ詳細ログする（大小無視）。"""
    return str(
        get_first("HC_DIAG_DATA_AGG_JOIN_COL", "DATA_AGG_JOIN_DUMP_COL") or ""
    ).strip()


def data_agg_diag_file_needed() -> bool:
    """data_agg 専用フラグだけで診断ファイルが必要か（マスタ HC_LOG_DIAG 以外）。"""
    return (
        data_agg_name_path_diag_enabled()
        or data_agg_batch_timing_enabled()
        or data_agg_file_timing_enabled()
        or data_agg_io_profile_enabled()
        or data_agg_join_dump_enabled()
    )


def data_agg_batch_file_path_filter_enabled() -> bool:
    """本番一括: file_pattern 付き cell ソースで走査結果を OR 絞込。DATA_AGG_BATCH_FILE_PATH_FILTER=0 で無効。"""
    raw = os.environ.get("DATA_AGG_BATCH_FILE_PATH_FILTER", "").strip().lower()
    if raw == "0":
        return False
    if raw in ("1", "true", "yes", "on"):
        return True
    return True


def data_agg_network_stage_copy_workers(*, n_files: int) -> int:
    """
    ネットワークステージの並列コピー数。CPU 数を検出して自動設定。
    DATA_AGG_STAGE_COPY_WORKERS: 0=逐次, auto/未設定=min(8, CPU, ファイル数), 正整数=上限。
    """
    if n_files < 1:
        return 1
    raw = os.environ.get("DATA_AGG_STAGE_COPY_WORKERS", "").strip().lower()
    if raw == "0":
        return 1
    cpu = os.cpu_count() or 4
    if not raw or raw == "auto":
        return max(1, min(8, cpu, n_files))
    try:
        w = int(raw)
        return max(1, min(w, n_files))
    except ValueError:
        return max(1, min(8, cpu, n_files))


def data_agg_file_parallel_workers(*, n_files: int) -> int:
    """
    一括集約の入力ファイル並列数。0 で逐次。
    DATA_AGG_FILE_PARALLEL_WORKERS: 0=オフ, auto/未設定= min(8, CPU, ファイル数)（2ファイル以上）, 正整数=上限。
    """
    if n_files < 2:
        return 0
    raw = os.environ.get("DATA_AGG_FILE_PARALLEL_WORKERS", "").strip().lower()
    if raw == "0":
        return 0
    cpu = os.cpu_count() or 4
    if not raw or raw == "auto":
        return max(1, min(8, cpu, n_files))
    try:
        w = int(raw)
        return max(0, min(w, n_files))
    except ValueError:
        return 0


def data_agg_network_stage_pipeline_enabled() -> bool:
    """コピー完了次第の抽出開始（ステージ＋抽出パイプライン）。DATA_AGG_STAGE_PIPELINE=0 で無効。"""
    raw = os.environ.get("DATA_AGG_STAGE_PIPELINE", "").strip().lower()
    if raw in ("0", "false", "no", "off"):
        return False
    return True


def data_agg_parallel_cold_cap() -> int:
    """UNC 起点時の抽出並列上限（ランプ前）。DATA_AGG_PARALLEL_COLD_CAP（既定 2）。"""
    raw = os.environ.get("DATA_AGG_PARALLEL_COLD_CAP", "").strip()
    if not raw:
        return 2
    try:
        return max(1, min(int(raw), 8))
    except ValueError:
        return 2


def data_agg_parallel_ramp_files() -> int:
    """抽出完了がこの件数に達したら並列上限を本番 workers まで解放。DATA_AGG_PARALLEL_RAMP_FILES（既定 16）。"""
    raw = os.environ.get("DATA_AGG_PARALLEL_RAMP_FILES", "").strip()
    if not raw:
        return 16
    try:
        return max(1, int(raw))
    except ValueError:
        return 16


def data_agg_resolve_file_parallel_workers(
    *,
    n_files: int,
    io_paths: Sequence[str | Path] | None = None,
    display_paths: Sequence[str | Path] | None = None,
) -> int:
    """
    目標 workers（auto 等）を返す。UNC 起点のコールド抑制は ramp 側で適用するため、
    ここでは DATA_AGG_FILE_PARALLEL_WORKERS の解決のみ行う。
    """
    return data_agg_file_parallel_workers(n_files=n_files)


def data_agg_parallel_ramp_enabled(
    *,
    n_files: int,
    target_workers: int,
    display_paths: Sequence[str | Path] | None = None,
    io_paths: Sequence[str | Path] | None = None,
) -> bool:
    """UNC 起点かつ target_workers > cold_cap のときランプを有効化。"""
    if target_workers <= 1:
        return False
    try:
        from svc.data_agg_path_network import path_is_network
    except Exception:
        return False
    disp = [str(p) for p in (display_paths or io_paths or []) if str(p).strip()]
    if not disp:
        return False
    if not any(path_is_network(p) for p in disp):
        return False
    return target_workers > data_agg_parallel_cold_cap()


def data_agg_master_progress_prefetch_enabled() -> bool:
    """マスタ進捗の先読みキューを使う（従来 DATA_AGG_MASTER_OFF_PREFETCH=1 と同じ条件）。"""
    if truthy(os.environ.get("HC_DIAG_DATA_AGG_MASTER_PREFETCH"), empty_means_false=False):
        return True
    return os.environ.get("DATA_AGG_MASTER_OFF_PREFETCH", "").strip() == "1"


def data_agg_master_progress_one_shot_enabled() -> bool:
    """マスタ進捗の項目内一括 compute（結合探索なし時）。DATA_AGG_MASTER_ONE_SHOT=0 で無効。"""
    raw = os.environ.get("DATA_AGG_MASTER_ONE_SHOT", "").strip().lower()
    if raw == "0":
        return False
    return True


def data_agg_master_frozen_columns_enabled() -> bool:
    """マスタ進捗: 完了項目列を行キー付き凍結し、次項目以降の compute で再走査しない。DATA_AGG_MASTER_FROZEN_COLUMNS=0 で無効。"""
    raw = os.environ.get("DATA_AGG_MASTER_FROZEN_COLUMNS", "").strip().lower()
    if raw == "0":
        return False
    return True


def data_agg_master_parallel_extract_enabled() -> bool:
    """マスタプレビューでもファイル並列抽出を使う。DATA_AGG_MASTER_PARALLEL_EXTRACT=0 で無効。"""
    raw = os.environ.get("DATA_AGG_MASTER_PARALLEL_EXTRACT", "").strip().lower()
    if raw == "0":
        return False
    return True


def diag_log_file_enabled() -> bool:
    """hc_csv_diag.log へ書き込むか（マスタ診断・data_agg 診断・conflict HWND 診断・UI 前面診断・タイトルバー診断）。"""
    return (
        log_diag_enabled()
        or data_agg_diag_file_needed()
        or csv_sp_conflict_hwnd_diag_enabled()
        or ui_fg_diag_enabled()
        or ui_native_file_diag_enabled()
        or ui_excel_lock_diag_enabled()
        or ui_window_caption_diag_enabled()
    )


def return_early_wait_sec() -> float:
    """HC_RETURN_EARLY=1 時、依頼送出後に return する直前の sleep 秒数。

    既定は 1.0（従来 hc_main の固定値と同一）。短縮は環境変数で行う（例: 0.5）。
    コード既定を 0.5 に変更する方針は取らず、互換のためデフォルトは 1.0 のままとする。
    """
    raw = get("HC_RETURN_EARLY_WAIT_SEC", "1.0")
    try:
        v = float(str(raw or "1.0").strip())
    except ValueError:
        return 1.0
    if v < 0:
        return 0.0
    return min(v, 30.0)


def data_agg_extract_default_max() -> int:
    """縦反復: repeat_max 未設定かつ空白まででもないときの既定上限（互換 9999）。"""
    return 9999


def data_agg_extract_absolute_max() -> int:
    """縦反復: N件・空白までの絶対上限。HC_DATA_AGG_EXTRACT_ABSOLUTE_MAX（既定 999999）。"""
    raw = get_first("HC_DATA_AGG_EXTRACT_ABSOLUTE_MAX", default="999999")
    try:
        v = int(str(raw or "999999").strip())
        return max(1, min(v, 999999))
    except ValueError:
        return 999999


def data_agg_extract_trunc_policy(
    *,
    probe_caller: Optional[str] = None,
    preview_master: bool = False,
) -> str:
    """
    抽出打ち切り検知時の方針。HC_DATA_AGG_EXTRACT_TRUNC_POLICY:
      abort … 結合前に DataAggExtractTruncated（本番一括の既定）
      warn  … ログのみで続行
    """
    raw = (get("HC_DATA_AGG_EXTRACT_TRUNC_POLICY") or "").strip().lower()
    if raw in ("warn", "continue", "log"):
        return "warn"
    if raw in ("abort", "stop", "cancel"):
        return "abort"
    if preview_master or (probe_caller or "") not in ("excel_batch_submit",):
        return "warn"
    return "abort"


# 常駐メイン（ルート hc_main.py）のログプレフィックス（hc_csv.log 等で grep）
LOG_MAIN_PREFIX = "[MAIN]"


def _float_env_dual(primary: str, legacy: str, default: str) -> float:
    raw = get_first(primary, legacy, default=default)
    try:
        return float(str(raw).strip())
    except ValueError:
        return float(default)


def _int_env_dual(primary: str, legacy: str, default: str) -> int:
    raw = get_first(primary, legacy, default=default)
    try:
        return int(str(raw).strip())
    except ValueError:
        return int(default)


def hc_main_poll_sec() -> float:
    """hc_main のポーリング間隔（秒）。`HC_MAIN_POLL_SEC` 優先、無ければ `HC_BRIDGE_POLL_SEC`。"""
    return _float_env_dual("HC_MAIN_POLL_SEC", "HC_BRIDGE_POLL_SEC", "0.05")


def hc_main_min_file_age_sec() -> float:
    """bridge_requests の .json を読む前の最小経過秒数。"""
    return _float_env_dual("HC_MAIN_MIN_FILE_AGE_SEC", "HC_BRIDGE_MIN_FILE_AGE_SEC", "0.05")


def hc_main_bad_file_max_polls() -> int:
    """同一ファイルの解釈・転送の最大ポーリング回数。"""
    return _int_env_dual("HC_MAIN_BAD_FILE_MAX_POLLS", "HC_BRIDGE_BAD_FILE_MAX_POLLS", "100")
