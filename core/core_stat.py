# -*- coding: utf-8 -*-
"""
Pythonバージョン: 3.12+
モジュール名: core/core_stat
作成日: 2026-02-03
更新日: 2026-02-03
バージョン: 1.0.3
概要:
    Excel シートのカスタムプロパティ（CustomProperties）への物理アクセス（保存・取得）を専門に受け持つサービスモジュール。
    [機能追加] GUID の設定用メソッド (set_guid) を実装し、物理刻印ロジックを共通化。
    [設計] 呼び出し元がキー名称を意識せず、安全に情報の出し入れができるインターフェースを提供。
    [保守性] STATUS/NOTIFY/GUID といったシステム共通キーを一元管理し、不整合を物理的に防止する。

改訂履歴:
    1.0.3: 2026-02-03 [機能追加] セッターメソッド (set_guid) を追加し、識別子の管理を一本化。
    1.0.2: 2026-02-03 [機能追加] ゲッターメソッド（取得処理）を追加。
    1.0.1: 2026-02-03 [不具合修正] Ruff 警告（未使用インポート）の解消。
"""

from typing import Any

# ==============================================================================
# 基盤モジュールの物理吸引セクション
# ==============================================================================
try:
    # 命令分離: 共通 Excel 操作・通信基盤。
    from core import core_xlc as xlc

    # 命令分離: 物理ログ管理基盤。
    from core import core_log
except ImportError:
    # 救済用インポート（パッケージ外環境用）。
    import core_xlc as xlc  # type: ignore
    import core_log  # type: ignore

# 変数: 本モジュール専用の物理ロガーを取得。
# 【目的】プロパティ操作時のシート名、キー、値の状態を物理的に追跡可能にするため。
logger = core_log.get_logger(__name__)

# --- 物理定数（システム共通キーの一元管理） ---
# 【目的】VBA 側の Bridge モジュールおよび Python 側各サービスでキー名称を原子同期させるため。
KEY_STATUS_INFO: str = "HC_STATUS_INFO"
KEY_NOTIFY_RETV: str = "HC_NOTIFY_RETV"
KEY_GUID_B64: str = "HC_GUID_B64"
KEY_BOOK_NAME: str = "HC_BOOK_NAME"


# ==============================================================================
# 公開サービス関数（セッター：保存系）
# ==============================================================================


def set_status_info(sh_target: Any, str_content: str) -> None:
    """
    Method Name : set_status_info
    Arguments   : sh_target, str_content
    Return      : None
    概要: ステータスバー表示用の情報を、標準キー (HC_STATUS_INFO) を用いてシートに保存する。
    """
    # 証跡。
    logger.debug(f"Setting status info to worksheet [{sh_target.name}].")

    # 命令分離: 物理保存の執行。
    set_prop(sh_target, KEY_STATUS_INFO, str_content)


def set_notify_retv(sh_origin: Any, str_content: str) -> None:
    """
    Method Name : set_notify_retv
    Arguments   : sh_origin, str_content
    Return      : None
    概要: 完了通知用の情報を、標準キー (HC_NOTIFY_RETV) を用いてシートに保存する。
    設計: 通常の完了時には呼ばない。特別な時のみ呼び出す（例: エラー通知、VBA 側で MsgBox 表示が必要な場合）。
    VBA の CheckAndNotifyVBA が RunPython 復帰後にこの値を参照し、値があれば MsgBox 表示する。
    """
    # 証跡。
    logger.info(f"Setting notification content to worksheet [{sh_origin.name}].")

    # 命令分離: 物理保存の執行。
    set_prop(sh_origin, KEY_NOTIFY_RETV, str_content)


def set_guid(sh_target: Any, str_guid: str) -> None:
    """
    Method Name : set_guid
    Arguments   :
        sh_target (Any) : 保存対象の Worksheet オブジェクト
        str_guid (str) : 保存する Base64 GUID
    Return      : None
    概要: シートに対して身元識別用の GUID (Base64) を標準キー (HC_GUID_B64) で原子刻印する。
    """
    # 証跡。
    # 【目的】識別情報の刻印タイミングを証跡として残すため。
    logger.info(f"Setting GUID to worksheet [{sh_target.name}].")

    # 命令分離: 物理保存の執行。
    # 【目的】機能モジュール側でキー名を指定せず、本モジュールで一元管理されたキーに保存するため。
    set_prop(sh_target, KEY_GUID_B64, str_guid)


def set_prop(sh_target: Any, str_key: str, str_value: str) -> None:
    """
    Method Name : set_prop
    Arguments   : sh_target, str_key, str_value
    Return      : None
    概要: 汎用的なプロパティ保存メソッド。
    """
    # 判定。
    if sh_target is None:
        # 【目的】無効参照による異常終了を防止するため。
        logger.error(f"Cannot set property [{str_key}]: Worksheet is None.")
        # 命令分離。
        return

    try:
        # 命令分離: 物理刻印。
        # 【目的】xlc モジュールの機能を利用し、シートプロパティへ原子保存するため。
        xlc.set_sheet_prop(sh_target, str_key, str_value)

    except Exception as ex:
        # 異常証跡。
        logger.error(f"Failed to set property [{str_key}] on [{sh_target.name}]: {ex}")


# ==============================================================================
# 公開サービス関数（ゲッター：取得系）
# ==============================================================================


def get_status_info(sh_target: Any) -> str:
    """
    Method Name : get_status_info
    Arguments   : sh_target (Any)
    Return      : str (保存されているステータス文字列)
    概要: ステータスバー表示用の情報を、標準キー (HC_STATUS_INFO) から吸引する。
    """
    # 命令分離: 物理吸引。
    return get_prop(sh_target, KEY_STATUS_INFO)


def get_notify_retv(sh_target: Any) -> str:
    """
    Method Name : get_notify_retv
    Arguments   : sh_target (Any)
    Return      : str (保存されている通知文字列)
    概要: 完了通知用の情報を、標準キー (HC_NOTIFY_RETV) から吸引する。
    """
    # 命令分離: 物理吸引。
    return get_prop(sh_target, KEY_NOTIFY_RETV)


def get_guid(sh_target: Any) -> str:
    """
    Method Name : get_guid
    Arguments   : sh_target (Any)
    Return      : str (Base64 GUID)
    概要: シートに刻印されている GUID を吸引する。
    """
    # 命令分離: 物理吸引。
    return get_prop(sh_target, KEY_GUID_B64)


def get_prop(sh_target: Any, str_key: str) -> str:
    """
    Method Name : get_prop
    Arguments   :
        sh_target (Any) : 取得対象の Worksheet オブジェクト
        str_key (str) : 取得キー名
    Return      : str (取得した値。存在しない場合は空文字)
    概要: 汎用的なプロパティ取得メソッド。
    """
    # 判定。
    if sh_target is None:
        # 【目的】無効参照によるエラーを防止するため。
        return ""

    try:
        # 命令分離: 物理吸引。
        # 【目的】xlc モジュールの基本機能を利用し、シートから値を物理的に取り出すため。
        val_ret = xlc.get_sheet_prop(sh_target, str_key)

        # 戻り値。
        return val_ret

    except Exception as ex:
        # 異常証跡。
        logger.error(
            f"Failed to get property [{str_key}] from [{sh_target.name}]: {ex}"
        )
        # 命令分離。
        return ""


# ---------------------------------------------------------------------------------------------------------------------
# End of core/hc_stat.py
