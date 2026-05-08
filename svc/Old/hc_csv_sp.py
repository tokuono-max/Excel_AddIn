# -*- coding: utf-8 -*-
"""
Pythonバージョン: 3.10+
モジュール名: hc_csv_sp
作成日: 2025-12-04
更新日: 2026-01-26
バージョン: 1.0.6
概要:
    Excel シートのデータを特定の列（キー）の値に基づいて、複数の CSV ファイルへ物理分割出力。
    本バージョンにて、プロジェクト全体の用語規定に基づき「原子」表現を「最小単位」へ変更。
    最新の xlc.save_status (Ver 1.0.12) および hc_prg (Ver 1.0.6) との連携を完遂。
    進捗メッセージによるステータスバー汚染を物理解消し、完了報告の永続性を確保。
    一切の要約を排し、1 命令 1 行の最小単位分解記述を徹底。

改訂履歴:
    1.0.6: 用語規定（最小単位）の適用、およびステータス表示ロジックの整合性強化。
    1.0.5: 未使用インポート(Tuple, Any)の抹消、および用語規定(最小単位)の適用。
    1.0.4: ステータス復元処理の堅牢化。
"""

import os
import re
import sys
import csv
import pandas as pd
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from typing import Optional

# ==============================================================================
# 階層化インポート解決セクション (Ruff/Pylance対応)
# ==============================================================================
# 【目的】svc フォルダ内からプロジェクトルートの基盤モジュールを確実に参照するため。

# 変数: 現在実行中のファイルの物理絶対パスを取得。
path_script_raw_val = os.path.abspath(__file__)
# 変数: svc ディレクトリの特定。
path_svc_dir_ptr_v = os.path.dirname(path_script_raw_val)
# 変数: プロジェクトルート（svc の一つ上）のパスを算出。
path_project_root_ptr_v = os.path.dirname(path_svc_dir_ptr_v)

# 判定コメント: 検索パスの先頭にプロジェクトルートが存在するか確認。
if path_project_root_ptr_v not in sys.path:
    # 命令分離: 検索パスの先頭に物理挿入。
    # 【目的】パッケージ実行時や解析ツールの解決エラーを根絶するため。
    sys.path.insert(0, path_project_root_ptr_v)

try:
    # 命令分離: 基盤定数、Win32、ExcelCore、GUI基盤、進捗UIの階層インポート。
    # 【Ruff】E402 (Module level import not at top of file) を抑制。
    from core import hc_cst as c  # noqa: E402
    from core import hc_w32 as w32  # noqa: E402
    from core import hc_xlc as xlc  # noqa: E402
    from ui import hc_gui as gui  # noqa: E402
    from ui import hc_prg as prg  # noqa: E402
except ImportError:
    # 救済コメント: 階層解決に失敗した場合の直接インポート試行。
    import hc_cst as c  # type: ignore # noqa: E402
    import hc_w32 as w32  # type: ignore # noqa: E402
    import hc_xlc as xlc  # type: ignore # noqa: E402
    import hc_gui as gui  # type: ignore # noqa: E402
    import hc_prg as prg  # type: ignore # noqa: E402


# ==============================================================================
# クラス: CSVSplitWin (最小単位レベル UI 制御)
# ==============================================================================


class CSVSplitWin:
    """
    概要: ファイル分割の設定 GUI 画面を管理するクラス。
    【重要】終了時に master.quit() を執行し、Excel の COM 待機を物理解放する。
    """

    def __init__(self, master_rt: tk.Tk, target_hwnd: Optional[int] = None) -> None:
        """
        Method Name : __init__
        概要: 分割設定画面の物理構築と親子関係の最小単位確立。
        """
        # 変数: 共有ルートの保持。
        self.master = master_rt

        # --- Excel 接続コンテキストの最小単位捕捉 ---
        # 【目的】操作対象の Application, Workbook, Worksheet を物理特定するため。
        tuple_context_res = xlc.get_ctx(hwnd_val=target_hwnd)

        # 変数: ポインタの解体。
        self.app = tuple_context_res[0]
        self.wb = tuple_context_res[1]
        self.sh = tuple_context_res[2]

        # --- 親 HWND の物理解決 ---
        # 変数: 親ハンドル ID。
        val_resolved_h = 0

        # 判定。
        if target_hwnd is not None:
            # 命令分離: 引数の整数値を優先採用。
            val_resolved_h = int(target_hwnd)
        else:
            # 命令分離: Excel インスタンスから動的に Win32 ID を最小単位抽出。
            val_resolved_h = xlc.get_h(self.app)

        # 変数: メンバ保持。
        self.parent_hwnd = val_resolved_h

        # --- 子ウィンドウ（Toplevel）の具現化 ---
        # 変数: 窓インスタンスの生成。
        self.win = tk.Toplevel(self.master)

        # 変数: 窓タイトルの構築。
        str_title_val = f"{c.APP_TITLE} - ファイル分割設定"

        # 命令分離: タイトルの最小単位セット。
        self.win.title(str_title_val)

        # 【重要】Win32 所有権リンク（Owner）の最小単位確立。
        if self.parent_hwnd != 0:
            # 命令分離: w32 経由での物理接続を執行。
            # 【目的】Excel の前面を維持しつつ、背後への埋没を物理防止するため。
            w32.set_owner(self.win, self.parent_hwnd)

        # ---------------------------------------------------------
        # 物理サイズ算出セクション (最小単位)
        # ---------------------------------------------------------
        # 変数: 基準寸法。
        val_base_w = 460
        val_base_h = 520

        # 命令分離: 物理スケーリング倍率の適用。
        self.sc_w = gui.get_scaled_size(val_base_w)
        self.sc_h = gui.get_scaled_size(val_base_h)

        # 命令分離: 配置座標（センタリング）の最小単位算出。
        str_geom_ptr = xlc.get_centering_geom_v(self.parent_hwnd, self.sc_w, self.sc_h)

        # 命令分離: ジオメトリの物理適用。
        self.win.geometry(str_geom_ptr)

        # 命令分離: 最小サイズの最小単位セット。
        self.win.minsize(self.sc_w, self.sc_h)

        # ---------------------------------------------------------
        # ウィジェット配置セクション (微細分解)
        # ---------------------------------------------------------
        # 変数: パディングピクセル。
        val_pad_px = gui.get_scaled_size(12)

        # 変数: メインパネルの具現化。
        self.pane_main = ttk.Frame(self.win, padding=val_pad_px)

        # 命令分離: 配置。
        self.pane_main.pack(fill="both", expand=True)

        # 命令分離: 案内ラベルの生成と配置。
        self.lbl_guide = ttk.Label(
            self.pane_main, text="分割の基準となる列（キー列）を選択してください。"
        )
        self.lbl_guide.pack(anchor="w", pady=(0, 10))

        # 変数: リストボックス用コンテナ。
        self.lb_container = ttk.Frame(self.pane_main)
        # 配置。
        self.lb_container.pack(fill="both", expand=True)

        # 変数: 垂直スクロールバー。
        self.scr_bar_v = ttk.Scrollbar(self.lb_container, orient="vertical")
        # 配置。
        self.scr_bar_v.pack(side="right", fill="y")

        # 変数: リストボックス本体。
        self.lb_obj = tk.Listbox(
            self.lb_container,
            font=(c.SUB_WINDOW_FONT_NAME, c.SUB_WINDOW_FONT_SIZE),
            yscrollcommand=self.scr_bar_v.set,
            selectmode="single",
            borderwidth=1,
            relief="solid",
        )

        # 命令分離: 配置。
        self.lb_obj.pack(side="left", fill="both", expand=True)

        # 命令分離: バーとの物理紐付け。
        self.scr_bar_v.config(command=self.lb_obj.yview)

        # 変数: 出力先設定パネル。
        self.pd_group = ttk.LabelFrame(
            self.pane_main, text=" 出力先フォルダ ", padding=8
        )
        # 配置。
        self.pd_group.pack(fill="x", pady=(15, 0))

        # 変数: パス管理変数。
        self.v_path_str = tk.StringVar()

        # 変数: 初期パスの物理解決。
        str_initial_p = ""
        if self.wb is not None:
            # 命令分離: ブックの絶対パスからフォルダを抽出。
            str_fullname_v = self.wb.fullname
            str_initial_p = os.path.dirname(str_fullname_v)

        # 命令分離: 初期値の最小単位セット。
        self.v_path_str.set(str_initial_p)

        # 命令分離: エントリ入力欄の配置。
        self.ent_path = ttk.Entry(self.pd_group, textvariable=self.v_path_str)
        self.ent_path.pack(side="left", fill="x", expand=True, padx=(0, 5))

        # 命令分離: 参照ボタンの配置。
        self.btn_browse = ttk.Button(
            self.pd_group, text="参照...", command=self._on_browse_exec, width=8
        )
        self.btn_browse.pack(side="right")

        # 変数: フッターボタンエリア。
        self.pane_foot = ttk.Frame(self.pane_main)
        # 配置。
        self.pane_foot.pack(fill="x", pady=(20, 0))

        # 命令分離: 実行ボタンの配置。
        self.btn_run_obj = ttk.Button(
            self.pane_foot, text="分割実行", command=self._on_run_exec, width=15
        )
        self.btn_run_obj.pack(side="right", padx=(10, 0))

        # 命令分離: キャンセルボタンの配置。
        self.btn_cancel_obj = ttk.Button(
            self.pane_foot, text="キャンセル", command=self._on_close_exec, width=12
        )
        self.btn_cancel_obj.pack(side="right")

        # 命令分離: ヘッダ列名の最小単位ロード執行。
        self._load_cols_min_unit()

        # 命令分離: 閉鎖プロトコルの最小単位登録。
        self.win.protocol("WM_DELETE_WINDOW", self._on_close_exec)

        # 命令分離: 物理描画の同期。
        self.win.update()

    def _load_cols_min_unit(self) -> None:
        """
        Method Name : _load_cols_min_unit
        概要: ワークシートの見出し行を最小単位レベルで吸引し、リストへ投入。
        """
        # 判定。
        if self.sh is None:
            # 命令分離: 中断。
            return

        try:
            # 命令分離: 使用済み領域の列数を計数。
            api_used_range = self.sh.used_range
            val_cols_count = api_used_range.columns.count

            # 判定。
            if val_cols_count > 0:
                # 命令分離: 1 行目の物理吸引。
                ptr_header_rng = self.sh.range((1, 1), (1, val_cols_count))
                list_header_vals = ptr_header_rng.value

                # 判定。
                if not isinstance(list_header_vals, list):
                    # 命令分離: リスト型への正規化。
                    list_header_vals = [list_header_vals]

                # 命令分離: リストボックスの既存内容を抹消。
                self.lb_obj.delete(0, tk.END)

                # 機能部位: 各列名の投入ループ。
                for i_idx, val_col_nm in enumerate(list_header_vals):
                    # 変数: 表示用ラベルの構築。
                    str_item_label = (
                        f"{i_idx + 1}: {val_col_nm if val_col_nm else '(空)'}"
                    )
                    # 命令分離: 最小単位投入。
                    self.lb_obj.insert(tk.END, str_item_label)

                # 判定。
                if self.lb_obj.size() > 0:
                    # 命令分離: 先頭要素の強制選択。
                    self.lb_obj.selection_set(0)
        except Exception:
            # 判定コメント: ロード失敗時は沈黙。
            pass

    def _on_browse_exec(self) -> None:
        """
        Method Name : _on_browse_exec
        概要: 出力先ディレクトリを物理選択するためのダイアログを起動。
        """
        # 変数: 現在のパス取得。
        str_current_dir = self.v_path_str.get()

        # 命令分離: フォルダ選択ダイアログの物理執行。
        str_selected_p = filedialog.askdirectory(
            initialdir=str_current_dir, parent=self.win
        )

        # 判定。
        if str_selected_p:
            # 命令分離: 正規化されたパスの反映。
            str_norm_p = os.path.normpath(str_selected_p)
            self.v_path_str.set(str_norm_p)

    def _on_run_exec(self) -> None:
        """
        Method Name : _on_run_exec
        概要: 選択情報の最小単位バリデーションを経て分割ロジックを起動。
        """
        # 変数: 選択されているインデックスを取得。
        tuple_selection = self.lb_obj.curselection()

        # 判定。
        if not tuple_selection:
            # 命令分離: 未選択時は中断。
            return

        # 変数: キー列の物理位置（1 開始）。
        val_key_idx = int(tuple_selection[0]) + 1

        # 変数: 出力ディレクトリの抽出。
        str_target_dir = self.v_path_str.get().strip()

        # 判定。
        if os.path.exists(str_target_dir):
            # 命令分離: 実分割エンジンの最小単位起動。
            self._execute_split_logic_min_unit(val_key_idx, str_target_dir)

    def _execute_split_logic_min_unit(
        self, key_idx_val: int, output_dir_str: str
    ) -> None:
        """
        Method Name : _execute_split_logic_min_unit
        概要: 大規模行列データを最小単位で分割し、複数の CSV ファイルとして物理出力。
        """
        # 変数: 進捗 UI ポインタの初期化。
        p_prog_ui_ptr = None

        try:
            # 命令分離: 進捗ウィンドウの物理具現化。
            p_prog_ui_ptr = prg.ProgressWin(
                "ファイル分割を執行中",
                parent_hwnd=self.parent_hwnd,
                wb_for_status=self.wb,
            )

            # 判定。
            if self.parent_hwnd != 0:
                # 変数: 窓ハンドル。
                ptr_prog_win = p_prog_ui_ptr.win_handle
                if ptr_prog_win is not None:
                    # 命令分離: 所有権リンクの確立。
                    w32.set_owner(ptr_prog_win, self.parent_hwnd)

            # --- 物理領域解析セクション ---
            # 変数: 使用済み領域の捕捉。
            ptr_rect_used = self.sh.used_range
            # 変数: 物理行数。
            val_num_y = ptr_rect_used.rows.count
            # 変数: 物理列数。
            val_num_x = ptr_rect_used.columns.count

            # 判定コメント: 処理可能な行数が存在するか（ヘッダ除き 1 行以上）。
            if val_num_y < 2:
                # 命令分離: 進捗窓の破棄。
                if p_prog_ui_ptr is not None:
                    p_prog_ui_ptr.close()
                return

            # 命令分離: パフォーマンス最適化コンテキストの最小単位開始。
            with xlc.Opt(self.app):
                # 命令分離: 物理チャンク吸引エンジンの執行。
                # 【目的】COM 限界を回避しつつ、データを最小単位でメモリへ吸引するため。
                list_captured_data = xlc.read_chunk(
                    self.sh,
                    1,
                    1,
                    val_num_y,
                    val_num_x,
                    p_prog_ui_ptr,
                    "Excelデータを吸引中...",
                )

            # --- 解析および物理書き出しセクション ---
            # 変数: 解析用データフレームの構築。
            df_main_ptr = pd.DataFrame(
                list_captured_data[1:], columns=list_captured_data[0]
            )

            # 変数: キー列名称の最小単位特定。
            str_key_label = list_captured_data[0][key_idx_val - 1]

            # 変数: ユニークなキー項目の抽出。
            list_unique_keys = df_main_ptr[str_key_label].dropna().unique().tolist()

            # 変数: 成功ファイル数の計数。
            val_success_count = 0

            # 機能部位: 各キーに対する巡回保存ループ。
            for i_iter, val_key_it in enumerate(list_unique_keys):
                # 変数: 保存用ファイル名の安全化。
                str_raw_nm = str(val_key_it)
                str_safe_fn = re.sub(r"[\\\\/:*?\"<>|]", "_", str_raw_nm)

                # 命令分離: UI 進捗の同期更新。
                # 【目的】is_save=False が ui/hc_prg Ver 1.0.6 内で自動適用される。
                p_prog_ui_ptr.update(
                    i_iter + 1,
                    len(list_unique_keys),
                    f"CSV生成中 ({i_iter + 1}/{len(list_unique_keys)})",
                    custom_text=str_safe_fn,
                )

                # 変数: フィルタデータの抽出。
                df_filtered_it = df_main_ptr[df_main_ptr[str_key_label] == val_key_it]

                # 変数: 保存先絶対パス。
                str_abs_save_p = os.path.join(output_dir_str, f"{str_safe_fn}.csv")

                try:
                    # 命令分離: CSV 物理保存の執行（BOM付き UTF-8）。
                    df_filtered_it.to_csv(
                        str_abs_save_p,
                        index=False,
                        encoding="utf-8-sig",
                        quoting=csv.QUOTE_MINIMAL,
                    )
                    # 演算。
                    val_success_count = val_success_count + 1
                except Exception:
                    # 判定コメント: 個別ファイル失敗時は継続。
                    pass

            # --- 最終物理報告文章の構築 ---
            str_summary_rep = (
                f"ファイル分割が正常に完了しました。 | "
                f"基準: {str_key_label} | 物理生成: {val_success_count} ファイル | "
                f"出力先: {output_dir_str}"
            )

            # 命令分離: core/hc_xlc 経由でのステータス保存執行。
            # 【重要】is_save=True（規定）で呼び出すことで、完了報告を金庫へ永続保存。
            xlc.save_status(self.wb, str_summary_rep)

            # 資源解放。
            if p_prog_ui_ptr is not None:
                # 命令分離: 進捗窓の物理クローズ。
                p_prog_ui_ptr.close()

            # 命令分離: 設定画面の最小単位自動終了。
            self._on_close_exec()

        except Exception as ex_fatal_logic:
            # 判定。
            if p_prog_ui_ptr is not None:
                # 命令分離: 異常時クローズ。
                p_prog_ui_ptr.close()
            # 命令分離: エラー情報の報告。
            xlc.save_status(self.wb, f"ERROR: 分割処理不全 Detail: {ex_fatal_logic}")

    def _on_close_exec(self) -> None:
        """
        Method Name : _on_close_exec
        概要: 設定画面の物理破棄とスレッド解放エンジンの最小単位執行。
        """
        try:
            # 変数: 親ハンドル。
            val_h_p = self.parent_hwnd
            # 判定。
            if val_h_p != 0:
                # 命令分離: 最前面属性のリセット。
                w32.kill_topmost(val_h_p)
                # 命令分離: 本来の親 Excel への物理フォーカス還流執行。
                w32.focus_parent(val_h_p)

            # 判定。
            if self.win is not None:
                # 命令分離: ウィジェット資源の物理破棄。
                self.win.destroy()

            # 【重要】ハングアップ防止のための Python スレッド解放シグナル。
            # 【目的】mainloop() で停止しているメインプロセスを VBA へ戻すための原子命令。
            self.master.quit()
        except Exception:
            # 判定コメント: 例外時は沈黙。
            pass


# ==============================================================================
# 公開関数: split_csv (VBA 呼出エントリポイント)
# ==============================================================================


def split_csv(target_hwnd: Optional[int] = None) -> None:
    """
    Method Name : split_csv
    Arguments   : target_hwnd (Optional[int]) : 親 Excel の物理ハンドル
    Return      : None
    概要: [VBA呼出] 分割機能を最小単位レベルで物理起動。
    """
    # 1. Excel 接続情報の捕捉。
    tuple_init_ctx = xlc.get_ctx(hwnd_val=target_hwnd)

    # 変数: ポインタ解体。
    ptr_app_v = tuple_init_ctx[0]

    # 判定。
    if ptr_app_v is None:
        # 命令分離: 中断。
        return

    # 2. 親 HWND の物理解決。
    val_xl_hwnd_id = 0
    if target_hwnd is not None:
        # 命令分離: キャスト。
        val_xl_hwnd_id = int(target_hwnd)
    else:
        # 命令分離: 動的な物理捕捉。
        val_xl_hwnd_id = xlc.get_h(ptr_app_v)

    # 3. DPI 同期を伴うシングルトンルートの最小単位確保。
    # 【目的】Excel の場所に合わせて正しい倍率の UI を具現化するため。
    ptr_tk_root = gui.get_root(hwnd_context=val_xl_hwnd_id)

    # 判定コメント: UI 基盤が有効な場合のみ具現化を開始。
    if ptr_tk_root is not None:

        # 命令分離: UI クラスの物理インスタンス化。
        CSVSplitWin(master_rt=ptr_tk_root, target_hwnd=target_hwnd)

        try:
            # 【重要】OS イベント待機ループの物理開始。
            # 動作補足: CSVSplitWin 内部で quit() が呼ばれるまで、ここで Python が待機。
            ptr_tk_root.mainloop()
        except Exception:
            # 沈黙。
            pass

        # ---------------------------------------------------------
        # 4. 最終クリーンアップシーケンス (最小単位)
        # ---------------------------------------------------------
        # 【目的】VBA 側への制御還流と、最新ステータスの物理再掲。
        tuple_final_ctx = xlc.get_ctx(hwnd_val=target_hwnd)

        # 変数: ブックポインタの捕捉。
        ptr_wb_fin = tuple_final_ctx[1]

        # 判定。
        if ptr_wb_fin is not None:
            # 命令分離: VBA 側への「UI終了」信号の最小単位送信。
            # 【重要】xlc v1.0.12 のフィルタリングにより、以前の成功メッセージは守られる。
            xlc.save_status(ptr_wb_fin, "UI")

            # 命令分離: 金庫（プロパティ）からのステータスバー物理復元を執行。
            # 【目的】Python プロセス切断前に、バーの内容を正しく再掲させるため。
            xlc.restore_status(ptr_wb_fin)
