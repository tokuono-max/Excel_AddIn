# -*- coding: utf-8 -*-
"""
Pythonバージョン: 1.2.2
モジュール名: core_log
作成日: 2026-01-29
更新日: 2026-04-06
バージョン: 1.5.0
概要:
    アドイン全体のログ出力を一元管理するコアモジュール。
    - 運用: %TEMP%\\csv_tool\\hc_csv.log（従来どおり get_logger）
    - 診断: hc_csv_diag.log（HC_LOG_DIAG / HC_DEBUG・data_agg 診断・HC_CSV_SP_CONFLICT_HWND_DIAG・HC_UI_FG_DIAG 時）get_diag_logger
    - 計測: hc_csv_perf.log（HC_LOG_PERF 時）get_perf_logger
    [根本解決] LIFO 形式・1MB 制限・filelock によるプロセス間排他。

改訂履歴:
    1.5.0: 2026-04-06 診断・計測ロガー追加。data_agg 診断は hc_csv_diag.log に統合（旧 data_agg_diag.log は廃止）。
    1.4.0: 2026-04-01 get_data_agg_diag_logger 追加。
    1.3.0: 2026-03-10 filelock（hc_csv.log.lock）。
    1.2.0: 2026-01-31 LIFO・1MB 制限強化。
"""

import os
import logging
import threading
from contextlib import nullcontext
from typing import Optional

try:
    from filelock import FileLock
except ImportError:
    FileLock = None  # type: ignore[misc, assignment]

try:
    from core import core_cst as cst
except Exception:  # pragma: no cover
    cst = object()  # type: ignore

try:
    from core import core_env
except Exception:  # pragma: no cover
    core_env = None  # type: ignore

# 診断・計測ルートロガーのハンドラはプロセス内で 1 度だけ付与
_DIAG_ROOT_CONFIGURED: bool = False
_PERF_ROOT_CONFIGURED: bool = False

LOG_BASENAME_OPS: str = "hc_csv.log"
LOG_BASENAME_DIAG: str = "hc_csv_diag.log"
LOG_BASENAME_PERF: str = "hc_csv_perf.log"

# ==============================================================================
# 物理定数定義
# ==============================================================================
# 【目的】ログファイルの最大保持サイズを 1MB (1,048,576 バイト) に物理固定するため。
MAX_LOG_SIZE_BYTES: int = int((getattr(cst, "SYS", {}) or {}).get("LOG_MAX_BYTES", 1048576))


def trim_file_tail_to_max_bytes(
    path: str | os.PathLike[str],
    *,
    max_bytes: Optional[int] = None,
) -> bool:
    """Keep only newest tail bytes when file exceeds max."""
    target = os.fspath(path)
    try:
        cap = int(MAX_LOG_SIZE_BYTES if max_bytes is None else max_bytes)
    except Exception:
        cap = MAX_LOG_SIZE_BYTES
    if cap <= 0:
        cap = MAX_LOG_SIZE_BYTES
    try:
        if not os.path.isfile(target):
            return False
        size = os.path.getsize(target)
        if size <= cap:
            return False
        with open(target, "rb") as f:
            data = f.read()
        if len(data) > cap:
            data = data[-cap:]
        with open(target, "wb") as f:
            f.write(data)
        return True
    except OSError:
        return False


def append_text_with_cap(
    path: str | os.PathLike[str],
    text: str,
    *,
    encoding: str = "utf-8",
    max_bytes: Optional[int] = None,
) -> None:
    """Append text and enforce max file size cap (tail-keep)."""
    target = os.fspath(path)
    os.makedirs(os.path.dirname(target) or ".", exist_ok=True)
    lock_path = target + ".lock"
    file_lock_ctx = FileLock(lock_path) if FileLock else nullcontext()
    with file_lock_ctx:
        with open(target, "ab") as f:
            f.write(str(text).encode(encoding, errors="replace"))
        trim_file_tail_to_max_bytes(target, max_bytes=max_bytes)

# ==============================================================================
# カスタムハンドラ定義セクション
# ==============================================================================

class ReverseFileHandler(logging.Handler):
    """
    クラス名: ReverseFileHandler
    概要:
        ログメッセージをファイルの冒頭に物理挿入する（LIFO形式）カスタムハンドラ。
        標準の追記型ハンドラとは異なり、ファイルを開いた瞬間に最新の挙動を確認可能にする。
    """

    def __init__(self, filename: str, max_bytes: int = MAX_LOG_SIZE_BYTES):
        """
        メソッド名: __init__
        引数:
            filename (str) : 保存先ログファイルの絶対物理パス
            max_bytes (int): 許容する最大ファイルサイズ
        機能概要: ハンドラの初期化と、スレッド排他制御用ロックの生成を行う。
        """
        # 命令分離: 親クラスの初期化。
        super().__init__()

        # 変数: 物理パスの保持。
        self.base_filename_ptr = filename
        # 変数: サイズ上限の保持。
        self.max_size_limit = max_bytes
        # 変数: スレッドセーフ確保のための物理ロック。
        # 【目的】複数の Python スレッドから同時にログ出力が行われた際の競合を防止するため。
        self.io_lock_obj = threading.Lock()

    def emit(self, record: logging.LogRecord) -> None:
        """
        メソッド名: emit
        引数: record (logging.LogRecord) : 出力対象のログレコード
        戻り値: なし
        機能概要: ログメッセージを整形し、既存ファイルの先頭に物理挿入して保存する。
        """
        # 【重要】プロセス間排他: filelock でロックファイルを取得するまで待機。未導入時はスキップ。
        lock_path = self.base_filename_ptr + ".lock"
        file_lock_ctx = FileLock(lock_path) if FileLock else nullcontext()
        with file_lock_ctx:
            # 【重要】スレッド排他制御の開始。
            with self.io_lock_obj:
                try:
                    # 1. ログメッセージの物理整形。
                    str_formatted_msg = self.format(record)
                    str_entry_v = str_formatted_msg + "\n"

                    # 2. 文字コード変換 (UTF-8)。
                    bin_new_data = str_entry_v.encode('utf-8')

                    # 3. 既存データの物理吸引。
                    bin_old_data = b""
                    if os.path.exists(self.base_filename_ptr):
                        with open(self.base_filename_ptr, 'rb') as f_read_ptr:
                            bin_old_data = f_read_ptr.read()

                    # 4. LIFO（最新 + 既存）形式での原子結合。
                    bin_combined_data = bin_new_data + bin_old_data

                    # 5. サイズ上限による原子切り捨て（Truncate）。
                    if len(bin_combined_data) > self.max_size_limit:
                        bin_combined_data = bin_combined_data[:self.max_size_limit]

                    # 6. 物理ファイルへの上書き保存執行。
                    with open(self.base_filename_ptr, 'wb') as f_write_ptr:
                        f_write_ptr.write(bin_combined_data)

                except Exception:
                    self.handleError(record)

# ==============================================================================
# 公開インターフェース（ロガー取得）
# ==============================================================================

def get_logger(module_name: str) -> logging.Logger:
    """
    メソッド名: get_logger
    引数: module_name (str) : 呼び出し元のモジュール識別名
    戻り値: logging.Logger : 設定済みのロガーインスタンス
    機能概要:
        1MB制限および LIFO 形式が設定されたロガーを各モジュールへ提供する。
    """
    # 1. 物理保存先の特定セクション。
    # 変数: OS の一時フォルダパスを取得。
    # 【目的】VBA 側の Environ("TEMP") と物理的な場所を完全に一致させるため。
    str_temp_dir_v = os.environ.get("TEMP", "C:\\Temp")
    # 変数: 物理ログファイル名の構築。
    os.makedirs(os.path.join(str_temp_dir_v, "csv_tool"), exist_ok=True)
    str_log_path_full = os.path.join(str_temp_dir_v, "csv_tool", LOG_BASENAME_OPS)

    # 2. ロガーインスタンスの生成セクション。
    # 変数: 指定された名称でロガーを取得。
    logger_inst = logging.getLogger(module_name)

    # 判定: ハンドラが未登録であるか（二重登録の物理ガード）。
    if not logger_inst.handlers:
        # 命令分離: ログレベルを INFO に原子設定。
        logger_inst.setLevel(logging.INFO)

        # 3. 共通フォーマットの定義。
        # 【目的】[時刻.ミリ秒] [レベル] [モジュール名] メッセージ の形式で物理統一するため。
        # 動作補足: VBA 側のタイムスタンプ形式と可能な限り整合させる。
        str_fmt_expr = '[%(asctime)s.%(msecs)03d] [%(levelname)s] [%(name)s] %(message)s'
        str_date_expr = '%Y-%m-%d %H:%M:%S'

        # 変数: フォーマッタの実体化。
        formatter_obj = logging.Formatter(fmt=str_fmt_expr, datefmt=str_date_expr)

        # 4. LIFO カスタムハンドラの物理追加。
        # 変数: 上記定義の物理ハンドラを実体化。
        handler_ptr = ReverseFileHandler(filename=str_log_path_full)

        # 命令分離: フォーマッタの適用。
        handler_ptr.setFormatter(formatter_obj)

        # 命令分離: ロガーへのハンドラ登録を執行。
        logger_inst.addHandler(handler_ptr)

    # 命令分離: セットアップ済みロガーを返却。
    return logger_inst


def _diag_file_enabled_fallback() -> bool:
    """core_env 未導入時の最小判定。"""
    def _t(k: str) -> bool:
        v = (os.environ.get(k) or "").strip().lower()
        return v in ("1", "true", "yes", "on", "y")

    if _t("HC_LOG_DIAG") or _t("HC_DEBUG"):
        return True
    if os.environ.get("DATA_AGG_NAME_PATH_DIAG", "").strip() == "1":
        return True
    if os.environ.get("DATA_AGG_COMPUTE_BATCH_TIMING", "").strip() == "1":
        return True
    if os.environ.get("DATA_AGG_PER_FILE_TIMING", "").strip() == "1":
        return True
    if _t("HC_CSV_SP_CONFLICT_HWND_DIAG"):
        return True
    if _t("HC_UI_WINDOW_CAPTION_DIAG"):
        return True
    return False


def _diag_file_enabled() -> bool:
    if core_env is not None:
        try:
            return core_env.diag_log_file_enabled()
        except Exception:
            pass
    return _diag_file_enabled_fallback()


def _perf_file_enabled() -> bool:
    if core_env is not None:
        try:
            return core_env.log_perf_enabled()
        except Exception:
            pass
    v = (os.environ.get("HC_LOG_PERF") or "").strip().lower()
    return v in ("1", "true", "yes", "on", "y")


def _configure_diag_root_logger() -> None:
    global _DIAG_ROOT_CONFIGURED
    if _DIAG_ROOT_CONFIGURED:
        return
    _DIAG_ROOT_CONFIGURED = True
    root = logging.getLogger("hc_csv_tool.diag")
    root.handlers.clear()
    root.propagate = False
    str_fmt_expr = '[%(asctime)s.%(msecs)03d] [%(levelname)s] [%(name)s] %(message)s'
    str_date_expr = '%Y-%m-%d %H:%M:%S'
    formatter_obj = logging.Formatter(fmt=str_fmt_expr, datefmt=str_date_expr)
    if not _diag_file_enabled():
        root.addHandler(logging.NullHandler())
        return
    str_temp_dir_v = os.environ.get("TEMP", "C:\\Temp")
    os.makedirs(os.path.join(str_temp_dir_v, "csv_tool"), exist_ok=True)
    path_full = os.path.join(str_temp_dir_v, "csv_tool", LOG_BASENAME_DIAG)
    h = ReverseFileHandler(filename=path_full)
    h.setFormatter(formatter_obj)
    root.addHandler(h)
    root.setLevel(logging.INFO)


def _configure_perf_root_logger() -> None:
    global _PERF_ROOT_CONFIGURED
    if _PERF_ROOT_CONFIGURED:
        return
    _PERF_ROOT_CONFIGURED = True
    root = logging.getLogger("hc_csv_tool.perf")
    root.handlers.clear()
    root.propagate = False
    str_fmt_expr = '[%(asctime)s.%(msecs)03d] [PERF] [%(name)s] %(message)s'
    str_date_expr = '%Y-%m-%d %H:%M:%S'
    formatter_obj = logging.Formatter(fmt=str_fmt_expr, datefmt=str_date_expr)
    if not _perf_file_enabled():
        root.addHandler(logging.NullHandler())
        return
    str_temp_dir_v = os.environ.get("TEMP", "C:\\Temp")
    os.makedirs(os.path.join(str_temp_dir_v, "csv_tool"), exist_ok=True)
    path_full = os.path.join(str_temp_dir_v, "csv_tool", LOG_BASENAME_PERF)
    h = ReverseFileHandler(filename=path_full)
    h.setFormatter(formatter_obj)
    root.addHandler(h)
    root.setLevel(logging.INFO)


def get_diag_logger(name: str = "hc_csv_tool.diag") -> logging.Logger:
    """診断ログ（hc_csv_diag.log）。無効時はルートに NullHandler のみ。"""
    _configure_diag_root_logger()
    return logging.getLogger(name)


def get_perf_logger(name: str = "hc_csv_tool.perf") -> logging.Logger:
    """計測ログ（hc_csv_perf.log）。無効時はルートに NullHandler のみ。"""
    _configure_perf_root_logger()
    return logging.getLogger(name)


def get_data_agg_diag_logger() -> logging.Logger:
    """
    データ集約の診断ログ。%TEMP%\\csv_tool\\hc_csv_diag.log（診断が有効なときのみ出力）。
    旧 data_agg_diag.log は廃止し診断ファイルに統合。
    """
    _configure_diag_root_logger()
    log = logging.getLogger("hc_csv_tool.diag.data_agg")
    log.propagate = True
    return log


# ==============================================================================
# モジュール動作確認セクション (物理デバッグ用)
# ==============================================================================
if __name__ == "__main__":
    # 【目的】単体起動時に LIFO 形式での書き込みが物理成立するか検証するため。
    # 命令分離: テスト用ロガー取得。
    test_log = get_logger("core.core_log_test")
    # 命令分離: 連続出力。
    test_log.info("LIFO Physical Test: Message 1")
    test_log.info("LIFO Physical Test: Message 2")
    # 命令分離: 終了報告。
    print("Log check completed. Please verify 'hc_csv.log' in %TEMP%\\csv_tool\\")

