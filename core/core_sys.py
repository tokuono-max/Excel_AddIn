# -*- coding: utf-8 -*-
"""
Pythonバージョン: 3.10+
モジュール名: core_sys
作成日: 2025-12-05
更新日: 2026-01-26
バージョン: 1.0.2
概要:
    物理ファイルシステム操作、実行パス特定、および物理キャッシュ（Pickle）の管理を担当。
    方程式「hc + sys (System)」に基づき、アドインの動作環境に関する低層ロジックを管理する。
    本バージョンにて、キャッシュ保存先を Windows 標準の一時フォルダ（Tmp）へ物理固定。
    AI による要約・省略を構造的に排除するため、全ステップを原子分解（1命令1行）で記述。
    保守性を確保するため、すべての主要ロジックに詳細な意図（目的）を付与。

改訂履歴:
    1.0.2: 原子分解レベルの更なる深化と、例外ガードの厳格化。
    1.0.1: キャッシュ保存先を Windows Temp ディレクトリに変更。
    1.0.0: 新規分割作成。パッケージ化構造への対応。
"""

import os
import sys
import pickle
import tempfile
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
    【重要】本バージョンより、Windows の一時フォルダ（Tmp）へ保存先を物理固定。
    """

    @staticmethod
    def _get_abs_path_atomic() -> str:
        """
        Method Name : _get_abs_path_atomic
        Return      : str : キャッシュファイルの絶対物理パス
        概要: OS 標準の一時ディレクトリを原子特定し、保存用パスを構築。
        """
        # 【目的】Windows の Tmp フォルダ（%TEMP%）を物理的に特定するため。
        # 命令分離: tempfile ライブラリの組み込み関数を実行。
        str_tmp_dir_root = tempfile.gettempdir()

        # 変数: 保存ファイル名の取得（core.hc_cst 参照）。
        str_cache_fn = c.CACHE_FILE_NAME

        # 命令分離: 一時フォルダとファイル名の物理的なパス結合。
        str_full_abs_p = os.path.join(str_tmp_dir_root, str_cache_fn)

        # 返却。
        return str_full_abs_p

    @staticmethod
    def save(key_str: str, data_payload: Any) -> None:
        """
        Method Name : save
        Arguments   : key_str (str) : 固体識別キー, data_payload (Any) : 保存データ
        概要: 指定された物理キーでデータを Windows Tmp フォルダへ永続化。
        """
        # 変数: Tmp フォルダ内のパス。
        path_p_save = CacheManager._get_abs_path_atomic()

        # 変数: 全体バッファ辞書の初期化。
        dict_data_buf = {}

        # 判定コメント: 既存キャッシュファイルの物理存在を確認。
        bool_f_exists = os.path.exists(path_p_save)

        if bool_f_exists:
            try:
                # 命令分離: 物理リードオープン。
                with open(path_p_save, "rb") as f_ptr_read:
                    # 命令分離: 全内容をメモリへ吸引。
                    dict_data_buf = pickle.load(f_ptr_read)
            except Exception:
                # 判定コメント: 読込失敗時は空の辞書から再構成。
                dict_data_buf = {}

        # 命令分離: 指定キーに基づいた原子エントリの追加・更新。
        dict_data_buf[key_str] = data_payload

        try:
            # 命令分離: 物理ディスク（Tmp）への書き出しを執行。
            with open(path_p_save, "wb") as f_ptr_write:
                # 命令分離: シリアライズ（Pickle）の実行。
                pickle.dump(dict_data_buf, f_ptr_write)
        except Exception:
            # 沈黙。
            pass

    @staticmethod
    def load(key_str: str) -> Optional[Any]:
        """
        Method Name : load
        Arguments   : key_str (str) : 取得対象キー
        概要: 一時フォルダ内のキャッシュから指定データを原子吸引。
        """
        # 変数: 物理パスの捕捉。
        path_p_load = CacheManager._get_abs_path_atomic()

        # 判定コメント: キャッシュファイル自体の存否確認。
        bool_exists = os.path.exists(path_p_load)
        if not bool_exists:
            return None

        try:
            # 命令分離: 物理オープン（バイナリリード）。
            with open(path_p_load, "rb") as f_ptr_load:
                # 命令分離: デシリアライズ執行。
                dict_master_v = pickle.load(f_ptr_load)

                # 判定コメント: 指定キーが物理的に含まれているか。
                if key_str in dict_master_v:
                    # 変数: データの原子抽出。
                    val_data_res = dict_master_v.get(key_str)
                    return val_data_res
        except Exception:
            # 沈黙。
            pass

        return None

    @staticmethod
    def delete(key_str: str) -> None:
        """
        Method Name : delete
        Arguments   : key_str (str) : 抹消対象キー
        概要: 指定されたキャッシュエントリを原子レベルで抹消。
        """
        # 変数: パス捕捉。
        path_p_del = CacheManager._get_abs_path_atomic()

        # 判定コメント: 存在確認。
        bool_f_exists = os.path.exists(path_p_del)
        if not bool_f_exists:
            return

        try:
            # 命令分離: 全ロードの執行。
            with open(path_p_del, "rb") as f_ptr_edit:
                dict_master_edit = pickle.load(f_ptr_edit)

            # 判定コメント: キーの存在を確認。
            if key_str in dict_master_edit:
                # 命令分離: 辞書からの物理削除。
                del dict_master_edit[key_str]

                # 命令分離: 物理ディスクへの書き戻し同期。
                with open(path_p_del, "wb") as f_ptr_sync:
                    pickle.dump(dict_master_edit, f_ptr_sync)
        except Exception:
            # 沈黙。
            pass
