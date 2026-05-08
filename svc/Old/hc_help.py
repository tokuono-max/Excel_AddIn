# -*- coding: utf-8 -*-
"""
Pythonバージョン: 3.10+
モジュール名: hc_help
作成日: 2025-12-03
更新日: 2026-01-26
バージョン: 1.1.4
概要:
    アプリケーションの操作ガイド（Info.txt）を物理的に読み込み、非モーダル窓で表示。
    方程式「hc + help」に従い、マニュアル表示という単一機能に特化。
    本バージョンにて、プロジェクト全体の用語規定に基づき「原子」を「最小単位」へ変更。
    最新の xlc.save_status および restore_status (Ver 1.0.12) に完全適合。
    一切の要約を排し、1 命令 1 行の最小単位分解記述を全域で徹底。

改訂履歴:
    1.1.4: 記述密度の再強化、および最新の xlc v1.0.12 基盤との整合性確保。
    1.1.3: 用語規定（最小単位）の適用、およびステータス復元処理の堅牢化。
    1.1.2: マルチステートメント(E701)の完全解消、原子分解の完遂。
"""

import os
import sys
import tkinter as tk
from tkinter import ttk
from typing import Optional

# ==============================================================================
# 階層化インポート解決セクション (Ruff/Pylance対応)
# ==============================================================================
# 【目的】svc フォルダ内からプロジェクトルートの基盤モジュールを確実に参照するため。

# 変数: 現在実行中のファイルの物理絶対パスを取得。
_path_script_raw_val = os.path.abspath(__file__)
# 変数: 自分自身が属するディレクトリ。
_path_svc_dir_ptr_v = os.path.dirname(_path_script_raw_val)
# 変数: プロジェクトルート（svc の一つ上）のパスを算出。
_path_project_root_ptr_v = os.path.dirname(_path_svc_dir_ptr_v)

# 判定コメント: 検索パスの先頭にプロジェクトルートが存在するか確認。
if _path_project_root_ptr_v not in sys.path:
    # 命令分離: 検索パスの先頭に物理挿入。
    # 【目的】パッケージとして実行する際や、解析ツールの解決エラーを根絶するため。
    sys.path.insert(0, _path_project_root_ptr_v)

try:
    # 命令分離: 基盤定数、Win32、物理システム、ExcelCore、GUI基盤の階層インポート。
    # 【Ruff】E402 (Module level import not at top of file) を抑制。
    from core import hc_cst as c  # noqa: E402
    from core import hc_w32 as w32  # noqa: E402
    from core import hc_sys as hsys  # noqa: E402
    from core import hc_xlc as xlc  # noqa: E402
    from ui import hc_gui as gui  # noqa: E402
except ImportError:
    # 救済コメント: 階層解決に失敗した場合の直接インポート試行。
    import hc_cst as c  # type: ignore # noqa: E402
    import hc_w32 as w32  # type: ignore # noqa: E402
    import hc_sys as hsys  # type: ignore # noqa: E402
    import hc_xlc as xlc  # type: ignore # noqa: E402
    import hc_gui as gui  # type: ignore # noqa: E402


# ==============================================================================
# クラス: HelpWin (非モーダル・最小単位レベル UI 制御)
# ==============================================================================


class HelpWin:
    """
    概要: 操作マニュアル画面を管理するクラス。
    【重要】非モーダル親子関係を死守し、Excel 側の入力を妨げない構造を完遂。
    指示に基づき、全ての構築ステップを最小単位レベルで分解して記述。
    """

    def __init__(self, master_rt: tk.Tk, target_hwnd: Optional[int] = None) -> None:
        """
        Method Name : __init__
        Arguments   : master_rt (tk.Tk) : 親ルート, target_hwnd (Optional[int]) : 親 Excel
        Return      : None
        概要: ヘルプウィンドウの物理構築と親子関係の確立。
        """
        # 変数: Tk ルート（シングルトン）の保持。
        self.master = master_rt

        # --- Excel コンテキストの最小単位捕捉 ---
        # 【目的】操作対象の Application, Workbook を物理特定するため。
        tuple_context_res = xlc.get_ctx(hwnd_val=target_hwnd)

        # 変数: ポインタの解体。
        self.app = tuple_context_res[0]
        self.wb = tuple_context_res[1]

        # --- 子ウィンドウ（Toplevel）の生成 ---
        # 変数: 窓インスタンスの具現化。
        self.win = tk.Toplevel(self.master)

        # 変数: アプリタイトルの取得。
        str_app_nm_v = c.APP_TITLE
        # 変数: 窓見出しの構築。
        str_caption_text = f"{str_app_nm_v} - 操作マニュアル"

        # 命令分離: タイトルの最小単位セット。
        self.win.title(str_caption_text)

        # --- 親子関係の物理解決 ---
        # 変数: 親HWND値。
        val_resolved_parent_h = 0

        # 判定。
        if target_hwnd is not None:
            # 命令分離: 引数の整数値を優先採用。
            val_resolved_parent_h = int(target_hwnd)
        else:
            # 命令分離: アプリポインタから Win32 ID を最小単位抽出。
            val_resolved_parent_h = xlc.get_h(self.app)

        # 変数: メンバ保持。
        self.parent_hwnd = val_resolved_parent_h

        # 【重要】Win32 API による非モーダル所有者（Owner）リンクの最小単位確立。
        if self.parent_hwnd != 0:
            # 命令分離: w32 経由での物理接続を執行。
            # 【目的】Excel の前面を維持しつつ、背後への埋没を物理防止するため。
            w32.set_owner(self.win, self.parent_hwnd)

        # ---------------------------------------------------------
        # ジオメトリ計算セクション (最小単位)
        # ---------------------------------------------------------
        # 変数: 基準寸法。
        val_base_w_px = 580
        val_base_h_px = 420

        # 命令分離: スケーリング後の寸法算出。
        # 【目的】高解像度環境での視認性を最小単位確保するため。
        self.scaled_w = gui.get_scaled_size(val_base_w_px)
        self.scaled_h = gui.get_scaled_size(val_base_h_px)

        # 命令分離: 配置座標（センタリング）の物理算出。
        # 【重要】マルチモニター環境における座標計算エンジンの執行。
        str_geom_val = xlc.get_centering_geom_v(
            self.parent_hwnd, self.scaled_w, self.scaled_h
        )

        # 命令分離: ジオメトリの物理適用。
        self.win.geometry(str_geom_val)

        # 命令分離: リサイズ許可の最小単位セット。
        self.win.resizable(True, True)

        # ---------------------------------------------------------
        # ウィジェット配置セクション (微細分解)
        # ---------------------------------------------------------
        # 変数: 余白パディング。
        val_pad_size_px = gui.get_scaled_size(10)

        # 変数: メインパネルの具現化。
        self.pane_main = ttk.Frame(self.win, padding=val_pad_size_px)

        # 命令分離: 配置。
        self.pane_main.pack(fill="both", expand=True)

        # 変数: テキスト表示コンテナ。
        self.txt_container = ttk.Frame(self.pane_main)
        # 命令分離: 配置。
        self.txt_container.pack(fill="both", expand=True)

        # 変数: 垂直スクロールバー。
        self.scr_bar_v = ttk.Scrollbar(self.txt_container, orient="vertical")
        # 命令分離: 右側配置。
        self.scr_bar_v.pack(side="right", fill="y")

        # 変数: フォント定義。
        str_f_nm = c.SUB_WINDOW_FONT_NAME
        val_f_sz = c.SUB_WINDOW_FONT_SIZE

        # 変数: テキストウィジェットの具現化。
        # 【目的】複数行の操作ガイドを物理的に表示するため。
        self.txt_view_area = tk.Text(
            self.txt_container,
            font=(str_f_nm, val_f_sz),
            yscrollcommand=self.scr_bar_v.set,
            wrap="word",
            padx=10,
            pady=10,
            borderwidth=1,
            relief="solid",
        )

        # 命令分離: 配置。
        self.txt_view_area.pack(side="left", fill="both", expand=True)

        # 命令分離: バーとの物理紐付け。
        self.scr_bar_v.config(command=self.txt_view_area.yview)

        # 変数: フッターパネル。
        self.pane_foot = ttk.Frame(self.pane_main)
        # 命令分離: 下端配置。
        self.pane_foot.pack(fill="x", pady=(10, 0))

        # 変数: 閉じるボタンの具現化。
        self.btn_close_obj = ttk.Button(
            self.pane_foot, text="閉じる", command=self._on_close_exec, width=15
        )
        # 命令分離: 右寄せ配置。
        self.btn_close_obj.pack(side="right")

        # ---------------------------------------------------------
        # 物理データ展開フェーズ
        # ---------------------------------------------------------
        # 【目的】外部ファイルから情報を吸引し画面へ同期。
        # 命令分離: 読込エンジンの最小単位起動。
        self._load_help_content_min_unit()

        # 命令分離: 編集禁止化の執行。
        self.txt_view_area.config(state="disabled")

        # 命令分離: クローズハンドラの最小単位バインド。
        self.win.protocol("WM_DELETE_WINDOW", self._on_close_exec)

        # 命令分離: 描画スタックの最終同期。
        self.win.update()

    def _load_help_content_min_unit(self) -> None:
        """
        Method Name : _load_help_content_min_unit
        概要: 外部ファイルを最小単位レベルで読み込み、テキスト領域へ注入する。
        """
        # 【目的】アプリ実行パスに基づき、Info.txt の場所を物理特定。
        str_app_root = hsys.get_app_path()

        # 変数: ターゲットファイル名。
        str_fn_val = "Info.txt"

        # 命令分離: 物理パスの結合。
        str_abs_p_val = os.path.join(str_app_root, str_fn_val)

        # 変数: 内容バッファ。
        str_content_payload = ""

        # 判定コメント: 物理的な実在確認。
        bool_f_exists = os.path.exists(str_abs_p_val)

        if bool_f_exists:
            try:
                # 【目的】文字化け防止のため UTF-8 で読込。
                # 命令分離: 物理リードオープン。
                with open(str_abs_p_val, "r", encoding="utf-8") as f_ptr:
                    # 命令分離: 全内容の吸引。
                    str_content_payload = f_ptr.read()
            except Exception as ex_read_err:
                # 変数: 救済用エラー文。
                str_content_payload = (
                    f"マニュアルの物理読込に失敗しました。\\n詳細: {ex_read_err}"
                )
        else:
            # 判定コメント: 物理的な欠落時の文章。
            str_content_payload = (
                "ヘルプファイル (Info.txt) が物理配置ディレクトリ内に見つかりません。"
            )

        # --- 表示反映セクション ---
        # 命令分離: 書き換え許可セット。
        self.txt_view_area.config(state="normal")
        # 命令分離: 既存抹消。
        self.txt_view_area.delete("1.0", "end")
        # 命令分離: 物理データ挿入。
        self.txt_view_area.insert("1.0", str_content_payload)

    def _on_close_exec(self) -> None:
        """
        Method Name : _on_close_exec
        概要: 窓の物理破棄と、ハングアップを根絶するためのスレッド解放。
        """
        try:
            # 変数: 親ハンドル。
            val_h_p = self.parent_hwnd
            if val_h_p != 0:
                # 命令分離: 最前面属性の抹消。
                w32.kill_topmost(val_h_p)

                # 命令分離: 本来の Excel へのフォーカス還流エンジンの執行。
                # 【重要】OS レベルで親 Excel を最前面へ物理リフトする。
                w32.focus_parent(val_h_p)
        except Exception:
            # 判定コメント: 還流失敗時は沈黙。
            pass

        try:
            # 判定。
            if self.win is not None:
                # 命令分離: ウィジェット資源の物理破棄。
                self.win.destroy()

            # 【重要】ハングアップ防止のための Python スレッド解放シグナル。
            # 【目的】mainloop() で停止しているメインプロセスを VBA へ戻すための最小単位命令。
            self.master.quit()
        except Exception:
            pass


# ==============================================================================
# 公開関数: show_help (VBA 呼出エントリポイント)
# ==============================================================================


def show_help(target_hwnd: Optional[int] = None) -> None:
    """
    Method Name : show_help
    Arguments   : target_hwnd (Optional[int]) : 呼び出し元 Excel の物理ハンドル
    Return      : None
    概要: [VBA呼出] ヘルプ画面を最小単位レベルで物理起動する。
    """
    # 【目的】シングルトンルート（DPI同期済）を最小単位確保。
    ptr_tk_root = gui.get_root(hwnd_context=target_hwnd)

    # 判定コメント: UI 基盤が有効な場合のみ具現化を執行。
    if ptr_tk_root is not None:

        # 機能部位: UI クラスの物理インスタンス化。
        # 【重要】target_hwnd を渡すことで非モーダル親子関係を完遂させる。
        HelpWin(master_rt=ptr_tk_root, target_hwnd=target_hwnd)

        try:
            # 【重要】OS イベント待機ループの物理開始。
            # 動作補足: HelpWin 内部で quit() が呼ばれるまで Python スレッドが待機。
            ptr_tk_root.mainloop()
        except Exception:
            # 沈黙。
            pass

        # ---------------------------------------------------------
        # 最終クリーンアップセクション
        # ---------------------------------------------------------
        # 【目的】VBA 側への完了通知と、金庫からのステータス復元。
        tuple_fin_ctx = xlc.get_ctx(hwnd_val=target_hwnd)
        ptr_wb_fin = tuple_fin_ctx[1]

        # 判定。
        if ptr_wb_fin is not None:
            # 【重要】"UI" 信号の沈黙送信。
            # 最新の xlc Ver 1.0.12 のフィルタリング機能により、以前の成功メッセージ（金庫）は守られる。
            xlc.save_status(ptr_wb_fin, "UI")

            # 命令分離: 金庫（Property）に保管された最新の「実利メッセージ」をバーへ物理復元。
            # 【目的】ヘルプを閉じた直後に、直前の成功報告（読込完了など）を再掲させるため。
            xlc.restore_status(ptr_wb_fin)


# ==============================================================================
# 直接起動ハンドラ (物理デバッグ用)
# ==============================================================================

if __name__ == "__main__":
    """概要: モジュール単体での物理動作テスト用エントリポイント。"""
    try:
        # 変数: テスト用ダミーハンドル。
        val_dummy_h = 0
        # 命令分離: ヘルプエンジンの物理起動。
        show_help(target_hwnd=val_dummy_h)
    except Exception as ex_fatal_exec:
        # 通知。
        print(f"Direct execution physically failed: {ex_fatal_exec}")
