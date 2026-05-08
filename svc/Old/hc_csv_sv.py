# -*- coding: utf-8 -*-
"""
Pythonバージョン: 3.12+
モジュール名: hc_csv_sv
作成日: 2025-12-05
更新日: 2026-02-03
バージョン: 1.3.6
概要:
    Excel の指定されたシートの内容を、BOM付き UTF-8 形式の CSV ファイルとして物理保存するサービス。
    [最終仕様] 保存完了時に既存のシートステータス（HC_STATUS_INFO）を一切上書きせず、起点情報を物理保護。
    [設計確立] 保存したファイル情報は HC_NOTIFY_RETV を介してポップアップ通知専用として伝播。
    [品質向上] 1行1命令、詳細な意図コメント（# 【目的】）、詳細なスタックトレースログ記録を完全適用。

改訂履歴:
    1.3.6: 2026-02-03 [設計確立] 起点情報の物理保護を確定。ステータス更新を廃止し直接表示のみ執行。
    1.3.5: 2026-02-03 [仕様調整] 既存ステータス情報の物理維持（上書き廃止）。
    1.3.3: 2026-02-03 [根本解決] ローカルチャンク読込による OLE エラーの物理根絶。
"""

import os
import csv
import pandas as pd
from tkinter import filedialog
from typing import Any, List

# ==============================================================================
# 基盤モジュールの物理吸引セクション
# ==============================================================================
try:
    # 命令分離: Windows API 制御基盤。
    from core import hc_w32 as w32

    # 命令分離: 共通 Excel 操作・通信基盤。
    from core import hc_xlc as xlc

    # 命令分離: ステータス管理共通サービス。
    from core import hc_stat as st

    # 命令分離: GUI 基盤。
    from ui import hc_gui as gui

    # 命令分離: 進捗可視化 UI 基盤。
    from ui import hc_prg as prg

    # 命令分離: 物理ログ管理基盤。
    from core.core_log import get_logger
except ImportError:
    # 救済用インポート（パッケージ外環境用）。
    import hc_w32 as w32  # type: ignore
    import hc_xlc as xlc  # type: ignore
    import hc_stat as st  # type: ignore
    import gui as gui  # type: ignore
    import prg as prg  # type: ignore
    from core.core_log import get_logger  # type: ignore

# 変数: 本モジュール専用の物理ロガーを取得。
# 【目的】保存処理の進行状況および、異常発生時の経緯を精密に監視するため。
logger = get_logger(__name__)

# --- 物理制約定数 ---
# 【役割】1回の通信で吸引する物理的な最大行数。OLE エラーを回避するための境界値。
READ_CHUNK_SIZE: int = 10000


# ==============================================================================
# 公開サービス関数 (VBA から hc_main を経由したリレー先)
# ==============================================================================


def save_csv(book: Any, sheet_id: str = "") -> None:
    """
    Callback Method : save_csv
    Arguments   :
        book (Any) : 操作対象の xw.Book オブジェクト
        sheet_id (str) : 対象シートの不変識別子 (Base64 GUID)
    Return      : None
    概要: [VBAリレー] 指定されたシートの内容を物理保存し、既存の起点ステータス情報を物理維持する。
    """
    # 0. 開始ログ。
    # 【目的】呼び出し元から正しく引数が伝播しているか証跡を記録するため。
    logger.info(f"--- [START] save_csv (v1.3.6) for ID: [{sheet_id}] ---")

    # 1. コンテキスト捕捉フェーズ。
    # 【目的】引数の book オブジェクトから Application ポインタを正確に捕捉するため。
    ptr_a = book.app
    ptr_w = book

    # 変数: 保存対象シート。
    ptr_s = None

    try:
        # 変数: 保存対象シートの物理解決。
        # 【目的】GUID によるシート特定を優先し、該当するシートを確実に捕捉するため。
        if sheet_id != "":
            # 命令分離: 共通基盤経由でシート探索執行。
            ptr_s = xlc.find_sheet_by_guid(ptr_w, sheet_id)
        else:
            # 命令分離: フォールバックとして現在のアクティブシートを採用。
            ptr_s = ptr_w.sheets.active

        # 判定: 捕捉失敗ガード。
        if ptr_s is None:
            # 【目的】対象が見つからない場合に、後続の COM エラーを未然に防止するため。
            logger.error("save_csv: Target sheet retrieval failed.")
            # 命令分離。
            return

        # 変数: 親ウィンドウハンドルの取得。
        # 【目的】座標基点および所有権確立のために利用するため。
        val_xl_h_ptr = xlc.get_h(ptr_a)

        # ---------------------------------------------------------
        # 2. 保存パス選択 (中央配置ハブを利用)
        # ---------------------------------------------------------
        # 【目的】マルチモニタおよび拡大率環境下でも Excel の中央に配置するため。
        hub = gui.get_centered_hub(parent_hwnd=val_xl_h_ptr)

        # 判定。
        if not hub:
            # 【目的】ハブの生成失敗時に処理を中断するため。
            logger.error("save_csv: Centered hub generation failed.")
            # 命令分離。
            return

        # 変数: 既定ファイル名の決定（現在のシート名称を採用）。
        str_default_fn_v = ptr_s.name

        # --- 保存先指定ダイアログの物理起動 ---
        # 【目的】ユーザーに保存パスを入力させ、その物理位置を特定するため。
        str_save_path_v = filedialog.asksaveasfilename(
            parent=hub,
            title="名前を付けてCSVを保存",
            initialfile=str_default_fn_v,
            defaultextension=".csv",
            filetypes=[("CSVファイル", "*.csv"), ("すべてのファイル", "*.*")],
        )

        # 命令分離: ハブ窓の破棄。
        hub.destroy()

        # 判定: キャンセル時は終了。
        if not str_save_path_v:
            # 命令分離。
            logger.info("CSV save operation cancelled by user.")
            # 命令分離。
            return

        # ---------------------------------------------------------
        # 3. 保存執行シーケンス
        # ---------------------------------------------------------
        # 変数: 進捗 UI 初期化。
        p_prog_ui = None

        # 命令分離: 進捗画面の原子具現化。
        p_prog_ui = prg.ProgressWin(
            parent_hwnd=val_xl_h_ptr, title_text="CSV 保存", wb_for_status=ptr_w
        )

        # 変数: 工程総数。
        total_phases = 2

        # --- 工程 1: データ読込 (Excelから分割抽出) ---
        p_prog_ui.set_phase("データ読込中", 1, total_phases)

        # # 【目的】Excel の描画・計算を停止させ、データ吸引速度を物理的に極大化するため。
        xlc.set_performance_mode(ptr_a, True)

        try:
            # 命令分離: ローカルでのチャンク読込執行。
            # 【目的】OLE タイムアウトを回避しながら全セル内容をメモリへ吸引するため。
            list_matrix_2d = _read_matrix_safe(ptr_s, p_prog_ui)
        finally:
            # # 【目的】データ抽出後に Excel の操作性を確実に復元し、不全状態を回避するため。
            xlc.set_performance_mode(ptr_a, False)

        # 判定: データ有無。
        if not list_matrix_2d:
            # 命令分離。
            logger.warning(f"No data found in worksheet [{ptr_s.name}].")
            # 命令分離。
            return

        # --- 工程 2: ファイル保存 (物理ディスクへ書込) ---
        p_prog_ui.set_phase("ファイル保存中", 2, total_phases)
        # 更新提示。
        p_prog_ui.update(0, "CSVファイルを物理生成中...")

        # 変数: 高速エクスポート用の Pandas オブジェクト化。
        df_output_ptr = pd.DataFrame(list_matrix_2d)

        # 命令分離: CSV 物理保存執行。
        # 【目的】指示に基づき UTF-8 (BOM付) で保存し、システム間の物理互換性を確保するため。
        df_output_ptr.to_csv(
            str_save_path_v,
            encoding="utf-8-sig",
            index=False,
            header=False,
            quoting=csv.QUOTE_MINIMAL,
        )

        # 進捗完了。
        p_prog_ui.update(100, "保存完了")

        # ---------------------------------------------------------
        # 4. 情報の後処理（起点情報の物理維持）
        # ---------------------------------------------------------
        # 変数: メトリクス抽出。
        str_fn_saved = os.path.basename(str_save_path_v)
        val_rows_saved = len(list_matrix_2d)
        str_size_saved = _get_formatted_size(str_save_path_v)

        # --- A. ステータス情報の原子維持 ---
        # 【重要：最終仕様】
        # 【目的】既存の HC_STATUS_INFO (読込時の起点情報) を変更せず、そのまま維持するため。
        # 上書き保存を行わないことで、データのトレーサビリティを物理保護する。

        # HC_NOTIFY_RETV は特別な時のみ設定する設計のため、通常の保存完了時は設定しない。

        # --- B. UI 最終反映 ---
        # # 【目的】Excel ステータスバーへ一時的に保存成功の事実を掲示するため。
        ptr_a.api.StatusBar = f"CSV保存終了｜{str_fn_saved} を保存しました。"

        # 正常証跡。
        logger.info(f"Save successful: [{str_fn_saved}]. Origin info preserved.")

    except Exception as ex_fatal_save:
        # 変数: エラーメッセージの構築。
        str_err_msg = f"ERROR: 保存失敗 ｜ {ex_fatal_save}"

        # 異常証跡。
        # 【目的】スタックトレースを含めて物理ログへ詳細を登録し、保守性を極大化するため。
        logger.error(
            f"Fatal error during CSV save sequence: {str_err_msg}", exc_info=True
        )

        # 命令分離: 異常時のみエラー内容をステータスへ原子報告。
        if ptr_s is not None:
            # 命令分離。
            st.set_status_info(ptr_s, str_err_msg)

        # 命令分離: ステータスバーへの物理反映。
        ptr_a.api.StatusBar = f"ERROR: CSV保存不全 ｜ {ex_fatal_save}"

    finally:
        # 資源解放。
        if "p_prog_ui" in locals() and p_prog_ui is not None:
            # 命令分離。
            p_prog_ui.close()

        # 【重要】本来の親 Excel への物理フォーカス還流執行。
        if "val_xl_h_ptr" in locals() and val_xl_h_ptr != 0:
            w32.focus_parent(val_xl_h_ptr)


# ==============================================================================
# 内部ヘルパー関数
# ==============================================================================


def _read_matrix_safe(sheet_ptr: Any, progress_ui: Any) -> List[List[Any]]:
    """
    Method Name : _read_matrix_safe
    概要: 1万行ずつの分割読込を執行し、Excel のフリーズを防止しつつ全データを吸引。
    """
    # 変数: 論理使用領域。
    api_range = sheet_ptr.used_range
    val_row_start = api_range.row
    val_col_start = api_range.column
    val_row_count = api_range.rows.count
    val_col_count = api_range.columns.count

    # 変数: 蓄積コンテナ。
    list_total_matrix = []

    # 命令分離: 分割ループ。
    for i_offset in range(0, val_row_count, READ_CHUNK_SIZE):
        # 変数。
        rows_to_read = min(READ_CHUNK_SIZE, val_row_count - i_offset)
        # 変数: スライス範囲。
        curr_range = sheet_ptr.range(
            (val_row_start + i_offset, val_col_start),
            (
                val_row_start + i_offset + rows_to_read - 1,
                val_col_start + val_col_count - 1,
            ),
        )
        # 命令分離: 物理吸引。
        list_chunk = curr_range.options(ndim=2).value
        # 命令分離: 結合。
        list_total_matrix.extend(list_chunk)

        # 進捗更新。
        val_pct = ((i_offset + rows_to_read) / val_row_count) * 100
        progress_ui.update(
            val_pct, f"データ読込中: {i_offset + rows_to_read:,} / {val_row_count:,} 行"
        )

    # 戻り値。
    return list_total_matrix


def _get_formatted_size(str_path: str) -> str:
    """内部用: 保存されたファイルの物理サイズを取得し、適切な単位で整形する。"""
    try:
        # 物理サイズ吸引。
        sz_val = os.path.getsize(str_path)
        # 判定。
        if sz_val >= 1048576:
            # 命令分離。
            return f"{sz_val / 1048576:.2f} MB"
        else:
            # 命令分離。
            return f"{sz_val / 1024:.1f} KB"
    except Exception:
        # 命令分離。
        return "不明"


# ---------------------------------------------------------------------------------------------------------------------
# End of svc/hc_csv_sv.py
