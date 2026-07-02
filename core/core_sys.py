# -*- coding: utf-8 -*-
"""
Pythonバージョン: 3.10+
モジュール名: core_sys
作成日: 2025-12-05
更新日: 2026-07-02
バージョン: 1.0.3
概要:
    物理ファイルシステム操作、実行パス特定、および物理キャッシュ（Pickle）の管理を担当。
    方程式「hc + sys (System)」に基づき、アドインの動作環境に関する低層ロジックを管理する。
    本バージョンにて、キャッシュ保存先を Windows 標準の一時フォルダ（Tmp）へ物理固定。
    AI による要約・省略を構造的に排除するため、全ステップを原子分解（1命令1行）で記述。
    保守性を確保するため、すべての主要ロジックに詳細な意図（目的）を付与。

改訂履歴:
    1.0.3: CacheManager をキー別ファイル＋zlib 圧縮に変更（Undo 保存・読込の高速化）。旧単一 pkl は読込・削除時のみ互換参照。
    1.0.2: 原子分解レベルの更なる深化と、例外ガードの厳格化。
    1.0.1: キャッシュ保存先を Windows Temp ディレクトリに変更。
    1.0.0: 新規分割作成。パッケージ化構造への対応。
"""

import os
import re
import sys
import pickle
import tempfile
import zlib
from typing import Any, Optional

# ==============================================================================
# 階層化インポート解決セクション (Ruff/PEP8/Pylance対応)
# ==============================================================================
# 【目的】core フォルダ内からルートの他モジュールを確実に参照するため。

# 変数: 現在実行中のファイルの物理絶対パスを取得。
path_script_raw_v = os.path.abspath(__file__)
# 変数: core ディレクトリのパスを特定。
path_core_dir_ptr_v = os.path.dirname(path_script_raw_v)
# 変数: プロジェクトルート（core の一つ上）のパスを算出。
path_project_root_ptr_v = os.path.dirname(path_core_dir_ptr_v)

# 判定コメント: 検索パスの先頭にプロジェクトルートが存在するか確認。
if path_project_root_ptr_v not in sys.path:
    # 命令分離: 検索パスの先頭に物理挿入。
    # 【目的】パッケージとして実行する際や、VSCode での解析エラーを根絶するため。
    sys.path.insert(0, path_project_root_ptr_v)

try:
    # 命令分離: 共通定数モジュールの階層インポート。
    # 【Ruff】E402 (Module level import not at top of file) を抑制。
    from core import core_cst as c  # noqa: E402
except ImportError:
    # 救済コメント: 階層解決に失敗した場合の直接インポート試行。
    import core_cst as c  # type: ignore # noqa: E402


# ==============================================================================
# 物理環境特定メソッド (原子分解版)
# ==============================================================================


def get_app_path() -> str:
    """
    Method Name : get_app_path
    Return      : str : 絶対ディレクトリ物理パス
    概要: 実行環境に応じた資源アクセス基準パスを原子特定。
    """
    # 【目的】実行ファイルの固化（EXE化）状態を物理判定するため。
    # Nuitka: sys.frozen が無い場合でも __main__.__compiled__ がある。
    _main = sys.modules.get("__main__")
    bool_is_frozen_v = getattr(sys, "frozen", False) or (
        _main is not None and getattr(_main, "__compiled__", None) is not None
    )

    if bool_is_frozen_v:
        # 変数: 実行ファイルの物理パス。
        str_exe_path_v = sys.executable
        # 命令分離: フォルダ部分を抽出。
        str_res_path_v = os.path.dirname(str_exe_path_v)
        return str_res_path_v

    # 変数: 自分自身のソースファイルパス。
    str_src_raw_v = __file__
    # 命令分離: 絶対パスへの正規化。
    str_abs_norm_v = os.path.abspath(str_src_raw_v)
    # 命令分離: ディレクトリ部分の原子抽出。
    str_parent_p_v = os.path.dirname(str_abs_norm_v)

    # 変数: 末尾フォルダ名の取得。
    str_tail_nm_v = os.path.basename(str_parent_p_v)

    # 判定コメント: パッケージ構造（core/）に属している場合、ルート階層へ原子調整。
    if str_tail_nm_v == "core":
        # 命令分離: 一階層上の親ディレクトリを取得。
        str_parent_p_v = os.path.dirname(str_parent_p_v)

    return str_parent_p_v


def get_file_size_str(abs_path_val: str) -> str:
    """
    Method Name : get_file_size_str
    Arguments   : abs_path_val (str) : 対象パス
    概要: 物理ファイルのバイトサイズを取得し、適切な単位（KB/MB）へ原子変換。
    """
    # 判定コメント: 物理実在の確認。
    bool_exists_v = os.path.exists(abs_path_val)
    if not bool_exists_v:
        return "0 KB"

    try:
        # 命令分離: OS から物理サイズの取得。
        val_b_raw_v = os.path.getsize(abs_path_val)

        # 判定コメント: 1MB (1,048,576 bytes) 閾値の検証。
        if val_b_raw_v > 1048576:
            # 演算: MB 単位への変換。
            val_mb_f_v = val_b_raw_v / 1048576
            return f"{val_mb_f_v:.2f} MB"

        # 演算: KB 単位への変換（1024 bytes）。
        val_kb_f_v = val_b_raw_v / 1024
        return f"{val_kb_f_v:.1f} KB"
    except Exception:
        # 救済。
        return "Unknown"


# ==============================================================================
# クラス: CacheManager (物理キャッシュ管理・Tmpフォルダ完全対応版)
# ==============================================================================


class CacheManager:
    """
    概要: Pickle 形式を用いた物理キャッシュ（Undo用データ）の読み書きを管理。
    キーごとに %TEMP% 配下の個別ファイルへ保存（zlib 圧縮）。旧単一 pkl は読込互換のみ。
    """

    _COMPRESS_LEVEL = 1

    @staticmethod
    def _get_abs_path_atomic() -> str:
        """
        Method Name : _get_abs_path_atomic
        Return      : str : レガシー単一キャッシュファイルの絶対物理パス
        """
        str_tmp_dir_root = tempfile.gettempdir()
        str_cache_fn = c.CACHE_FILE_NAME
        return os.path.join(str_tmp_dir_root, str_cache_fn)

    @staticmethod
    def _get_entries_dir() -> str:
        """キー別 Undo エントリを格納するディレクトリ。"""
        base = os.path.splitext(os.path.basename(c.CACHE_FILE_NAME))[0] or "header_converter_cache"
        path_dir = os.path.join(tempfile.gettempdir(), f"{base}_undo")
        try:
            os.makedirs(path_dir, exist_ok=True)
        except Exception:
            pass
        return path_dir

    @staticmethod
    def _entry_path(key_str: str) -> str:
        safe = re.sub(r'[\\/:*?"<>|]', "_", str(key_str or "").strip()) or "undo"
        if len(safe) > 180:
            import hashlib

            digest = hashlib.sha256(safe.encode("utf-8", errors="replace")).hexdigest()[:16]
            safe = f"{safe[:120]}_{digest}"
        return os.path.join(CacheManager._get_entries_dir(), f"{safe}.undo.pkl")

    @staticmethod
    def _serialize_payload(data_payload: Any) -> bytes:
        raw = pickle.dumps(data_payload, protocol=pickle.HIGHEST_PROTOCOL)
        return zlib.compress(raw, level=CacheManager._COMPRESS_LEVEL)

    @staticmethod
    def _deserialize_payload(blob: bytes) -> Any:
        try:
            raw = zlib.decompress(blob)
        except zlib.error:
            raw = blob
        return pickle.loads(raw)

    @staticmethod
    def _atomic_write_bytes(path: str, data: bytes) -> None:
        tmp = f"{path}.tmp"
        with open(tmp, "wb") as f_ptr_write:
            f_ptr_write.write(data)
        os.replace(tmp, path)

    @staticmethod
    def _legacy_load(key_str: str) -> Optional[Any]:
        path_p_load = CacheManager._get_abs_path_atomic()
        if not os.path.exists(path_p_load):
            return None
        try:
            with open(path_p_load, "rb") as f_ptr_load:
                dict_master_v = pickle.load(f_ptr_load)
                if key_str in dict_master_v:
                    return dict_master_v.get(key_str)
        except Exception:
            pass
        return None

    @staticmethod
    def _legacy_delete(key_str: str) -> None:
        path_p_del = CacheManager._get_abs_path_atomic()
        if not os.path.exists(path_p_del):
            return
        try:
            with open(path_p_del, "rb") as f_ptr_edit:
                dict_master_edit = pickle.load(f_ptr_edit)
            if key_str not in dict_master_edit:
                return
            del dict_master_edit[key_str]
            with open(path_p_del, "wb") as f_ptr_sync:
                pickle.dump(dict_master_edit, f_ptr_sync)
        except Exception:
            pass

    @staticmethod
    def save(key_str: str, data_payload: Any) -> None:
        """指定キーのペイロードを個別ファイルへ保存（全件 pkl の読込・書込を避ける）。"""
        path_entry = CacheManager._entry_path(key_str)
        try:
            CacheManager._atomic_write_bytes(path_entry, CacheManager._serialize_payload(data_payload))
        except Exception:
            pass

    @staticmethod
    def load(key_str: str) -> Optional[Any]:
        """個別ファイルを優先読込。無ければレガシー単一 pkl から取得。"""
        path_entry = CacheManager._entry_path(key_str)
        if os.path.exists(path_entry):
            try:
                with open(path_entry, "rb") as f_ptr_load:
                    return CacheManager._deserialize_payload(f_ptr_load.read())
            except Exception:
                pass
        return CacheManager._legacy_load(key_str)

    @staticmethod
    def delete(key_str: str) -> None:
        """個別ファイルとレガシー辞書の両方からキーを削除。"""
        path_entry = CacheManager._entry_path(key_str)
        try:
            os.remove(path_entry)
        except FileNotFoundError:
            pass
        except Exception:
            pass
        CacheManager._legacy_delete(key_str)
