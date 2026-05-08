# -*- coding: utf-8 -*-
"""
Pythonバージョン: 3.12+
モジュール名: hc_csv_ld
作成日: 2025-12-05
更新日: 2026-02-03
バージョン: 2.2.58
概要:
    外部 CSV ファイルを読み込み、Excel シートへ物理展開するサービスモジュール。
    [構造改善] ステータス・通知・GUID 等のプロパティ操作を共通サービス (hc_stat) へ完全移行。
    [品質向上] 物理制約定数に詳細な役割コメントを付与。1行1命令、詳細な意図コメント、Ruff 準拠を適用。
    [仕様維持] ユーザー指定の「要素名付・｜区切り」の表示フォーマットを厳守し、ロジックの透明性を確保。

改訂履歴:
    2.2.58: 2026-02-03 [保守強化] 物理制約定数の役割コメント拡充。hc_stat / hc_w32 連携の最終原子同期。
    2.2.57: 2026-02-03 [構造改善] プロパティ操作を共通モジュール hc_stat へ外出し。
    2.2.56: 2026-02-03 [仕様変更] ステータスバーおよび通知内容に要素名を付加。
"""

import os
import math
import pandas as pd
from tkinter import filedialog
from typing import Tuple, Any

# ==============================================================================
# 基盤モジュールの物理吸引セクション
# ==============================================================================
try:
    # 命令分離: 共通 Excel 操作・通信基盤。
    from core import hc_xlc as xlc

    # 命令分離: Windows API 制御基盤。
    from core import hc_w32 as w32

    # 命令分離: GUI 基盤。
    from ui_qt import hc_gui as gui

    # 命令分離: 進捗可視化 UI 基盤。
    from ui_qt import hc_prg as prg

    # 命令分離: 物理ログ管理基盤。
    from core.core_log import get_logger

    # 命令分離: 共通ステータス管理サービス。
    from core import hc_stat as st
except ImportError:
    # 救済用インポート（パッケージ外環境用）。
    import hc_xlc as xlc  # type: ignore
    import hc_w32 as w32  # type: ignore
    import gui as gui  # type: ignore
    import prg as prg  # type: ignore
    from core.core_log import get_logger  # type: ignore
    import hc_stat as st  # type: ignore

# 変数: 本モジュール専用の物理ロガーを取得。
# 【目的】読込処理の進行状況および、共通サービスを介したプロパティ保存の証跡を精密に監視するため。
logger = get_logger(__name__)

# --- 物理制約定数 ---
# 【役割】全体の行数に対し、1回のチャンク読込でメモリに展開する比率。負荷分散の指標。
CHUNK_PCT_BASE: float = 0.1
# 【役割】一度に Pandas が読み込む行数の物理的な最大上限値。メモリ溢れを防止するためのガード。
MAX_CHUNK_LIMIT: int = 50000
# 【役割】全体の行数が極端に少ない場合に、チャンクサイズが 0 にならないよう保証する安全下限値。
MIN_CHUNK_LIMIT: int = 100
# 【役割】単一の Excel シートに展開可能な物理的な最大行数。これを超えるとシート分割ロジックが発動する。
MAX_ROWS_PER_SHEET: int = 1000000


# ==============================================================================
# 公開サービス関数 (VBA からの直接リレー先)
# ==============================================================================


def load_csv(book: Any, sheet_id: str = "") -> None:
    """
    Callback Method : load_csv
    Arguments   :
        book (Any) : 操作対象のブックオブジェクト
        sheet_id (str) : VBA 側から伝播されたシート識別子
    Return      : None
    概要: 境界行を厳密に守った分割読込を行い、共通サービス (hc_stat) を介して統計情報を原子保存する。
    """
    # 0. 開始ログ。
    # 【目的】サービスの開始時刻と処理コンテキスト（GUID）を証跡として記録するため。
    logger.info("--- [START] load_csv (v2.2.58) ---")

    # 判定: オブジェクトの物理生存確認。
    if book is None:
        # 【目的】無効な参照によるランタイムエラーを防止するため。
        return

    # 変数: Excel ウィンドウハンドル。
    val_xl_hwnd_ptr = xlc.get_h(book.app)
    # 変数: 起点（ターゲット）シート参照を取得。
    sh_origin_ref = xlc.find_sheet_by_guid(book, sheet_id)

    # 判定。
    if sh_origin_ref is None:
        # 異常証跡。
        logger.error(f"Origin sheet retrieval failed. GUID: {sheet_id}")
        # 【目的】展開先の起点となるシートが特定できない場合に、処理を安全に中断するため。
        return

    # 1. 物理パスの取得（中央配置ハブを利用）。
    str_csv_target_path = _ask_file_path_centered(val_xl_hwnd_ptr)

    # 判定。
    if not str_csv_target_path:
        # 【目的】ユーザーによるキャンセル時は即座に終了し、不要な処理を走らせないため。
        logger.info("Operation cancelled by user.")
        return

    # 2. 事前メトリクスの取得。
    val_total_lines_v = _get_row_count_binary(str_csv_target_path)
    str_readable_size = _get_formatted_size(str_csv_target_path)

    # 解析証跡。
    logger.info(f"Source Metrics: {val_total_lines_v:,} rows / {str_readable_size}")

    # 3. 分割合意確認。
    if val_total_lines_v > MAX_ROWS_PER_SHEET:
        # 変数: 予測枚数の算出。
        val_sheets_total = math.ceil(val_total_lines_v / MAX_ROWS_PER_SHEET)

        # 変数: メッセージ構築。
        str_msg_v = (
            f"総行数 {val_total_lines_v:,} 行を検知しました。\n"
            f"全 {val_sheets_total} 枚のシートに分割して読み込みますか？"
        )

        # ボタン: MB_ICONQUESTION=0x20, MB_YESNO=0x4
        ans_id = gui.show_msg(
            text_body=str_msg_v, icon_flag=0x20, btn_flag=0x4, parent_h=val_xl_hwnd_ptr
        )

        # 判定: 戻り値 IDYES=6 以外は中止。
        if ans_id != 6:
            # 証跡。
            logger.info("Split loading was rejected by user.")
            return

    # 4. 実行環境の物理初期化。
    book.app.api.StatusBar = "CSV読込を開始しています..."
    # 命令分離: 進捗画面の構築。
    p_ui_ptr = prg.ProgressWin(
        parent_hwnd=val_xl_hwnd_ptr, title_text="CSV 読込中", wb_for_status=book
    )

    # 変数: 工程総数。
    total_steps = 3

    try:
        # --- 工程 1: ファイル解析 ---
        p_ui_ptr.set_phase("ファイル解析中", 1, total_steps)
        # 高速化。
        xlc.set_performance_mode(book.app, True)
        # 命令分離: 属性特定。
        _, str_enc_label, str_encoding_name = _detect_encoding_min_unit(
            str_csv_target_path
        )
        # 更新提示。
        p_ui_ptr.update(100, "解析完了")

        # --- 工程 2: Excelへ書き込み ---
        p_ui_ptr.set_phase("Excelへ書き込み中", 2, total_steps)
        _execute_jit_import(
            book=book,
            sh_origin=sh_origin_ref,
            str_path=str_csv_target_path,
            str_enc=str_encoding_name,
            val_total=val_total_lines_v,
            val_size_lbl=str_readable_size,
            str_enc_lbl=str_enc_label,
            p_ui=p_ui_ptr,
        )

        # --- 工程 3: 最終表示調整 ---
        p_ui_ptr.set_phase("表示を最終調整中", 3, total_steps)
        # 最適化解除。
        xlc.set_performance_mode(book.app, False)
        # 物理同期。
        xlc.yield_to_excel()
        # 更新提示。
        p_ui_ptr.update(100, "完了")

        # 正常証跡。
        logger.info(f"Process complete: [{sh_origin_ref.name}]")

    except Exception as ex_fatal:
        # 物理復旧。
        xlc.set_performance_mode(book.app, False)
        # 異常証跡。
        logger.error(f"Fatal error in load_csv: {ex_fatal}")
        # ユーザー通知。
        gui.show_msg(
            text_body=f"ERROR: CSV読込不全 | {ex_fatal}",
            icon_flag=0x10,
            parent_h=val_xl_hwnd_ptr,
        )

    finally:
        # 物理解放。
        if p_ui_ptr is not None:
            # 命令分離。
            p_ui_ptr.close()
        # フォーカス返却。
        # 【目的】ダイアログ終了後、操作権を Excel 本体へ確実に還流させるため。
        w32.focus_parent(val_xl_hwnd_ptr)


# ==============================================================================
# 内部ロジックセクション (Private)
# ==============================================================================


def _execute_jit_import(
    book: Any,
    sh_origin: Any,
    str_path: str,
    str_enc: str,
    val_total: int,
    val_size_lbl: str,
    str_enc_lbl: str,
    p_ui: Any,
) -> None:
    """
    Method Name : _execute_jit_import
    Summary     : 正確な境界行制御を行い、各シートへ統計情報の物理保存（共通サービス経由）を執行する。
    """
    # 変数: 名称特定。
    str_fname = os.path.basename(str_path)
    # 【目的】名称衝突を避けつつ、ハイフン付与前の基底名を決定するため。
    str_base_resolved = _get_unique_base_name(book, os.path.splitext(str_fname)[0])

    # カウンタ初期化。
    val_accum_total = 0
    val_accum_in_sheet = 0
    val_sheets_total = math.ceil(val_total / MAX_ROWS_PER_SHEET)
    # 判定フラグ。
    is_split_mode = val_total > MAX_ROWS_PER_SHEET
    # パート番号。
    curr_part_idx = 0
    # シート参照。
    sh_target = None

    # チャンクサイズ算出。
    calc_chunk_v = int(val_total * CHUNK_PCT_BASE)
    chunk_size_v = min(MAX_CHUNK_LIMIT, max(MIN_CHUNK_LIMIT, calc_chunk_v))

    # Pandas エンジン起動。
    reader = pd.read_csv(str_path, encoding=str_enc, dtype=str, chunksize=chunk_size_v)

    # 巡回開始。
    for df_chunk in reader:
        # 変数。
        val_rows_in_chunk = len(df_chunk)
        processed_chunk_idx = 0

        # --- プレ・スライスループ ---
        while processed_chunk_idx < val_rows_in_chunk:

            # A. シート切り替え・物理初期化判定。
            if sh_target is None or val_accum_in_sheet >= MAX_ROWS_PER_SHEET:

                # 命令分離: 直前シートが存在すれば、その時点の情報コンテキストを原子確定。
                if sh_target is not None:
                    _finalize_sheet_context(
                        book,
                        sh_target,
                        str_fname,
                        val_size_lbl,
                        str_enc_lbl,
                        val_accum_total,
                        val_total,
                        curr_part_idx,
                        val_sheets_total,
                        is_split_mode,
                        val_accum_in_sheet,
                        str_base_resolved,
                        sh_origin,
                    )

                # パート番号更新。
                curr_part_idx += 1

                # 名称生成ロジック。
                if is_split_mode:
                    # 【目的】分割読込時は枝番付きの名称を構成するため。
                    str_target_name = f"{str_base_resolved}-{curr_part_idx}"
                else:
                    # 【目的】単一シート時は基底名をそのまま使用するため。
                    str_target_name = str_base_resolved

                # 判定: 第1パート。
                if curr_part_idx == 1:
                    # 【目的】起点シートが空であれば再利用し、既存データがあれば新規シートを追加するため。
                    if _is_sheet_empty(sh_origin_ref=sh_origin):
                        sh_target = sh_origin
                        xlc.safe_rename_sheet(sh_target, str_target_name)
                    else:
                        sh_target = _add_new_sheet_direct(book, str_target_name)
                else:
                    # 追加。
                    sh_target = _add_new_sheet_direct(book, str_target_name)
                    # 命令分離: 物理描画スタックの同期。
                    xlc.yield_to_excel()

                # シート内累積リセット。
                val_accum_in_sheet = 0
                # フラグセット。
                is_new_sheet_boundary = True
            else:
                is_new_sheet_boundary = False

            # B. 書き込みスライス計算。
            remaining_cap = MAX_ROWS_PER_SHEET - val_accum_in_sheet
            rows_to_write = min(val_rows_in_chunk - processed_chunk_idx, remaining_cap)
            df_slice = df_chunk.iloc[
                processed_chunk_idx : processed_chunk_idx + rows_to_write
            ]

            # C. 身元識別情報の物理刻印（共通サービス hc_stat を利用）。
            # 【目的】VBA 側の RestoreStatBar が行う GUID 存在チェックを通過させるため。
            if st.get_guid(sh_target) == "":
                # 命令分離: 共通窓口経由で GUID を刻印。
                st.set_guid(sh_target, xlc.create_guid_b64())
                # 命令分離: 補助プロパティの原子保存。
                st.set_prop(sh_target, st.KEY_BOOK_NAME, book.name)

            # D. 行列データの物理転送。
            if is_new_sheet_boundary:
                # 【目的】シートの冒頭には CSV ヘッダ（カラム名）を物理付与するため。
                matrix = [df_slice.columns.tolist()] + df_slice.values.tolist()
            else:
                # データのみ。
                matrix = df_slice.values.tolist()

            # 変数: 開始行。
            start_row = val_accum_in_sheet + 1
            # 命令分離: セル書式の文字列化強制。
            sh_target.range(f"A{start_row}").expand().number_format = "@"
            # 物理書込執行。
            xlc.write_chunk(sh_target, start_row, 1, matrix, None)

            # E. 同期処理。
            val_accum_total += rows_to_write
            val_accum_in_sheet += len(matrix)
            processed_chunk_idx += rows_to_write

            # 進捗提示。
            val_curr_pct = (val_accum_total / val_total) * 100
            stat_text = f"{val_accum_total:,} / {val_total:,} 行"
            p_ui.update(val_curr_pct, stat_text)

            # ステータスバー提示。
            if is_split_mode:
                book.app.api.StatusBar = f"Excel書込中... {sh_target.name} [分割：{curr_part_idx}/{val_sheets_total}]"
            else:
                book.app.api.StatusBar = "Excel書込中..."

            # 命令分離。
            xlc.yield_to_excel()

    # --- 最終コンテキストの原子確定 ---
    if sh_target is not None:
        _finalize_sheet_context(
            book,
            sh_target,
            str_fname,
            val_size_lbl,
            str_enc_lbl,
            val_accum_total,
            val_total,
            curr_part_idx,
            val_sheets_total,
            is_split_mode,
            val_accum_in_sheet,
            str_base_resolved,
            sh_origin,
        )


def _finalize_sheet_context(
    book: Any,
    sh: Any,
    fname: str,
    size_lbl: str,
    enc_lbl: str,
    accum_total: int,
    total_file: int,
    p_idx: int,
    p_total: int,
    is_split: bool,
    accum_in_sheet: int,
    base_name: str,
    sh_origin: Any,
) -> None:
    """
    Method Name : _finalize_sheet_context
    Summary     : 共通サービス (hc_stat) を利用し、指示されたラベル付フォーマットで情報を原子保存する。
    """
    # --- 1. 要素名（ラベル）付きパーツの原子構築 ---
    # 【目的】指示に基づき、すべての表示項目に要素名を物理付加するため。
    str_sh_name_v = f"シート名：{sh.name}"
    str_rows_v = f"行数：{accum_in_sheet:,} 行"
    str_fname_v = f"ファイル名：{fname}"
    str_size_v = f"容量：{size_lbl}"
    str_enc_v = f"文字コード：{enc_lbl}"
    str_data_v = f"データ：{max(0, total_file - 1):,} 行"
    # 【仕様】ステータスバー用は「(ヘッダ含む)」。
    str_total_st_v = f"総数(ヘッダ含む)：{total_file:,} 行"

    # --- 2. ステータスバー表示情報の構築（共通サービス同期） ---
    # 【目的】アクティブシート切り替え時に VBA が「要素名：値 ｜ ...」形式で表示できるように整形する。
    if is_split:
        # 分割あり。
        info_status_body = (
            f"{str_sh_name_v} ｜ 分割：{p_idx}/{p_total} ｜ {str_rows_v} ┃ "
            f"{str_fname_v} ｜ {str_size_v} ｜ {str_enc_v} ｜ {str_data_v} ｜ {str_total_st_v}"
        )
    else:
        # 分割なし。
        info_status_body = (
            f"{str_sh_name_v} ｜ {str_rows_v} ┃ "
            f"{str_fname_v} ｜ {str_size_v} ｜ {str_enc_v} ｜ {str_data_v} ｜ {str_total_st_v}"
        )

    # 物理保存：共通サービスの窓口を利用して原子保存。
    # 【目的】保存キー名をロジック側で管理せず、標準の HC_STATUS_INFO へ書き込ませるため。
    st.set_status_info(sh, info_status_body)

    # HC_NOTIFY_RETV は特別な時のみ設定する設計のため、通常の読込完了時は設定しない。

    # --- 3. 読込完了直後のステータスバー最終反映 ---
    # 【目的】指示に基づき「CSV読込終了｜」を接頭辞として付与し即座に表示。
    book.app.api.StatusBar = f"CSV読込終了｜{info_status_body}"


# --- 補助関数群 (インターフェース維持) ---


def _ask_file_path_centered(hwnd_id: int) -> str:
    """内部用: 共通ハブを利用し、Excel 中央へ配置されたファイル選択画面を表示する。"""
    hub = gui.get_centered_hub(parent_hwnd=hwnd_id)
    # 判定。
    if not hub:
        # 命令分離。
        return ""
    # 命令分離。
    path_v = filedialog.askopenfilename(
        parent=hub,
        title="読み込む CSV ファイルを選択してください",
        filetypes=[("CSVファイル", "*.csv"), ("すべてのファイル", "*.*")],
    )
    # 破棄。
    hub.destroy()
    # 戻り値。
    return path_v


def _get_unique_base_name(book_p: Any, str_base: str) -> str:
    """内部用: シート名の衝突を確認し、重複時はアンダースコアを付加した名称を決定する。"""
    list_names = [s.name for s in book_p.sheets]
    # 判定。
    if str_base not in list_names:
        # 命令分離。
        return str_base
    seq = 1
    # 巡回。
    while True:
        target = f"{str_base}_{seq}"
        if target not in list_names:
            # 命令分離。
            return target
        seq += 1


def _add_new_sheet_direct(book_p: Any, target_name: str) -> Any:
    """内部用: 確定した名称でシートを物理追加する。"""
    try:
        # 命令分離。
        return book_p.sheets.add(name=target_name)
    except Exception:
        seq = 1
        # 巡回。
        while True:
            try:
                # 命令分離。
                return book_p.sheets.add(name=f"{target_name}_{seq}")
            except Exception:
                # インクリメント。
                seq += 1


def _is_sheet_empty(sh_origin_ref: Any) -> bool:
    """内部用: 対象シートが完全未使用であるか判定する。"""
    try:
        # 判定。
        return (
            sh_origin_ref.used_range.address == "$A$1"
            and sh_origin_ref.range("A1").value is None
        )
    except Exception:
        # 命令分離。
        return True


def _get_row_count_binary(str_path: str) -> int:
    """内部用: バイナリ走査により行数を高速計数する。"""
    cnt = 0
    # 命令分離。
    with open(str_path, "rb") as f_ptr:
        for _ in f_ptr:
            # インクリメント。
            cnt += 1
    # 戻り値。
    return cnt


def _get_formatted_size(str_path: str) -> str:
    """内部用: 物理サイズを取得し、適切な単位（MB/KB）で整形する。"""
    sz = os.path.getsize(str_path)
    # 判定。
    if sz >= 1048576:
        # 命令分離。
        return f"{sz / 1048576:.2f} MB"
    else:
        # 命令分離。
        return f"{sz / 1024:.1f} KB"


def _detect_encoding_min_unit(str_path: str) -> Tuple[None, str, str]:
    """内部用: 文字コードを物理特定する。"""
    try:
        # 1. BOM 判定。
        with open(str_path, "rb") as f_h:
            if f_h.read(3) == b"\xef\xbb\xbf":
                # 命令分離。
                return None, "UTF-8 (BOM)", "utf-8-sig"
        try:
            # 2. UTF-8 判定。
            with open(str_path, "r", encoding="utf-8") as f_t:
                f_t.read(2048)
            # 命令分離。
            return None, "UTF-8", "utf-8"
        except Exception:
            # 3. 日本語既定判定。
            return None, "Shift-JIS", "cp932"
    except Exception:
        # 4. 不明時。
        return None, "不明", "utf-8"


# ---------------------------------------------------------------------------------------------------------------------
# End of svc/hc_csv_ld.py
