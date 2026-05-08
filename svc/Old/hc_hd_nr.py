# -*- coding: utf-8 -*-
"""
Pythonバージョン: 3.10+
モジュール名: hc_hd_in
作成日: 2025-12-05
更新日: 2026-01-26
バージョン: 1.0.2
概要:
    出荷履歴データ加工用の定型ヘッダ項目（9項目）を Excel シートの先頭へ物理挿入する。
    方程式「hc + hd (Header) + in (Insert)」に従い、定型ヘッダの物理構築に特化。
    本バージョンにて、未使用インポート（hc_prg, Any）を物理削除し Ruff 警告を根絶。
    挿入先の行が持つ既存フォント設定を物理継承するロジックを実装。
    一切の要約を排し、1 命令 1 行の原子分解記述を徹底。

改訂履歴:
    1.0.2: 未使用インポート（hc_prg, Any）の抹消による Ruff 警告の完全解消。
    1.0.1: 未使用インポート（Any）の抹消試行。
    1.0.0: 新規分割作成。hc_header Ver 1.0.5 より挿入機能を分離。
"""

import os
import sys
from typing import Optional

# ==============================================================================
# 階層化インポート解決セクション (Ruff/Pylance対応)
# ==============================================================================
# 【目的】svc フォルダ内からプロジェクトルートの基盤モジュールを確実に参照するため。

# 変数: 現在実行中のファイルの物理絶対パスを取得。
path_raw_current = os.path.abspath(__file__)
# 変数: 自分自身が属するディレクトリ。
path_svc_dir_ptr = os.path.dirname(path_raw_current)
# 変数: プロジェクトルート（svc の一つ上）のパスを算出。
path_project_root_ptr = os.path.dirname(path_svc_dir_ptr)

# 判定コメント: 検索パスの先頭にプロジェクトルートが存在するか確認。
if path_project_root_ptr not in sys.path:
    # 命令分離: 検索パスの先頭に物理挿入。
    # 【目的】パッケージとして実行する際や、VSCode での解析エラーを根絶するため。
    sys.path.insert(0, path_project_root_ptr)

try:
    # 命令分離: Win32、および Excel 操作の階層インポート。
    # 【目的】未使用の進捗UI（hc_prg）を物理削除し Ruff F401 を解消。
    # 【Ruff】E402 (Module level import not at top of file) を抑制。
    from core import hc_w32 as w32  # noqa: E402
    from core import hc_xlc as xlc  # noqa: E402
except ImportError:
    # 救済コメント: 階層解決に失敗した場合の直接インポート試行。
    import hc_w32 as w32  # type: ignore # noqa: E402
    import hc_xlc as xlc  # type: ignore # noqa: E402


# ==============================================================================
# メメイン処理: insert_header (原子レベル分解版)
# ==============================================================================


def insert_header(target_hwnd: Optional[int] = None) -> None:
    """
    Method Name : insert_header
    Arguments   : target_hwnd (Optional[int]) : 親 Excel ハンドル
    Return      : None
    概要: [VBA呼出] 出荷履歴用の定型ヘッダを物理挿入。
    """
    # 1. 接続環境の原子捕捉。
    # 【目的】Excel の Application, Workbook, Worksheet ポインタを安全に取得するため。
    # 命令分離: Excel コンテキスト取得エンジンの実行。
    tuple_ctx = xlc.get_ctx(hwnd_val=target_hwnd)

    # 変数: ポインタ解体（一行ずつ個別に抽出）。
    ptr_a = tuple_ctx[0]
    ptr_w = tuple_ctx[1]
    ptr_s = tuple_ctx[2]

    # 判定コメント: 接続不全時の例外ガード。
    if ptr_a is None:
        return
    if ptr_w is None:
        return
    if ptr_s is None:
        return

    # 変数: 親 HWND 解決。
    val_xl_h = 0
    # 判定。
    if target_hwnd is not None:
        # 命令分離: 引数のハンドルを優先使用。
        val_xl_h = int(target_hwnd)
    else:
        # 命令分離: アプリポインタから動的に Win32 ID を抽出。
        val_xl_h = xlc.get_h(ptr_a)

    try:
        # ---------------------------------------------------------
        # 2. 定型ラベル配列の原子定義
        # ---------------------------------------------------------
        # 【目的】業務規定に基づいた 9 項目の標準ラベルを物理的に構築。
        list_def_labels = [
            "出荷予定日",
            "伝票番号",
            "顧客コード",
            "顧客名",
            "商品コード",
            "商品名",
            "数量",
            "単位",
            "備考",
        ]

        # ---------------------------------------------------------
        # 3. 物理挿入シーケンス (スタイル継承と展開)
        # ---------------------------------------------------------
        # 【目的】既存行を物理的に押し下げ、1 行目に定型ヘッダを具現化する。

        # 命令分離: アプリケーション API ポインタの捕捉。
        api_app_ptr = ptr_a.api

        # 命令分離: 操作権限の物理ロック（書込中のユーザー介入防止）。
        api_app_ptr.Interactive = False

        # 命令分離: パフォーマンス最適化コンテキスト（画面計算停止）の適用。
        with xlc.Opt(ptr_a):

            # --- スタイル情報の原子退避 ---
            # 【目的】新しく挿入される空行に、元の 1 行目の書式を原子レベルで継承させる。
            back_font_name = None
            back_font_size = None

            try:
                # 変数: 現時点の 1 行目を捕捉。
                ptr_orig_row = ptr_s.range("1:1")
                # 命令分離: フォント名の取得。
                back_font_name = ptr_orig_row.api.Font.Name
                # 命令分離: フォントサイズの取得。
                back_font_size = ptr_orig_row.api.Font.Size
            except Exception:
                # 判定コメント: プロパティ取得失敗時は規定値なし。
                pass

            # --- 物理挿入執行 ---
            # 変数: 基底行ターゲットの捕捉。
            ptr_target_r = ptr_s.range("1:1")
            # 命令分離: 行挿入（Insert）の物理送信。
            # 【目的】シート全体を1行分下方へ物理移動させるため。
            ptr_target_r.api.Insert()

            # --- 内容の物理展開 ---
            # 変数: 二次元リスト形式の書き込みペイロードを構築。
            arr_payload_2d = [list_def_labels]

            # 変数: 書込先起点（新設された 1 行目）の捕捉。
            ptr_write_o = ptr_s.range((1, 1))
            # 命令分離: セル値の原子同期（一括反映）。
            ptr_write_o.value = arr_payload_2d

            # --- スタイルの原子リストア ---
            # 【目的】挿入に伴い自動付与された太字等の意図しない設定を平準化（解除）する。
            try:
                # 変数: 新しく具現化したヘッダ行の API。
                ptr_new_row_api = ptr_s.range("1:1").api

                # 命令分離: フォント名称の物理復元。
                if back_font_name is not None:
                    ptr_new_row_api.Font.Name = back_font_name

                # 命令分離: フォントサイズの物理復元。
                if back_font_size is not None:
                    ptr_new_row_api.Font.Size = back_font_size

                # 命令分離: 物理的な太字解除の執行（正規化前の標準状態へ）。
                ptr_new_row_api.Font.Bold = False
            except Exception:
                pass

            # 変数: 使用済み領域全体の捕捉。
            ptr_u_range = ptr_s.used_range
            # 命令分離: 全列幅の物理オートフィット執行。
            ptr_u_range.columns.autofit()

        # --- 最終物理報告文章 ---
        str_msg_rep = "出荷履歴用の定型ヘッダ項目をシート先頭へ物理挿入しました。"
        # 命令分離: VBA 私書箱（Names）への完了報告保存。
        xlc.save_status(ptr_w, str_msg_rep)

    except Exception as ex_in_fatal:
        # 命令分離: 異常情報の物理報告保存。
        xlc.save_status(ptr_w, f"ERROR: ヘッダ挿入不全 Detail: {ex_in_fatal}")

    finally:
        # --- 物理環境復旧 ---
        try:
            # 変数: API 再接続。
            api_xl_final = ptr_a.api
            # 命令分離: インタラクティブ権限の復旧。
            api_xl_final.Interactive = True
        except Exception:
            # 判定コメント: 断絶時は沈黙。
            pass

        # 命令分離: 親 Excel へのフォーカス還流エンジンの執行。
        # 【重要】本来の Excel を最前面へ物理リフトし、OS フォーカスを安定させる。
        w32.focus_parent(h_target_ptr=val_xl_h)
