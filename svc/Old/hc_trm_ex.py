# -*- coding: utf-8 -*-
"""
Pythonバージョン: 3.10+
モジュール名: hc_trm_ex
作成日: 2025-11-28
更新日: 2026-01-26
バージョン: 1.0.2
概要:
    Excel シート全域を物理スキャンし、前後に不要な空白があるセルを特定・修正するサービス。
    本バージョンにて、Ruff が指摘した未使用インポート（re, List, Tuple, hc_cst）を物理抹消。
    全ステップを「1 命令 1 行」に分解し、AI による要約を構造的に排除。
    独立プロセス化により、対話型 UI 表示中も Excel 側のセル編集を物理的に許可。

改訂履歴:
    1.0.2: 未使用インポート（re, List, Tuple, hc_cst）の抹消による Ruff 警告の根絶。
    1.0.1: マルチステートメントの解消、原子分解の徹底。
    1.0.0: 新規分割作成。hc_dataform Ver 1.0.8 より機能を独立。
"""

import os
import sys
import pickle
import tempfile
import subprocess
from typing import Optional, Any, Dict

# ==============================================================================
# 階層化インポート解決セクション (Ruff/Pylance対応)
# ==============================================================================
# 【目的】svc フォルダ内からプロジェクトルートの基盤モジュールを確実に参照するため。

# 変数: 現在実行中のファイルの物理絶対パスを取得。
path_script_raw = os.path.abspath(__file__)
# 変数: 自分自身が属するディレクトリ（svc）。
path_svc_dir_ptr = os.path.dirname(path_script_raw)
# 変数: プロジェクトルート（svc の一つ上）のパスを算出。
path_project_root_ptr = os.path.dirname(path_svc_dir_ptr)

# 判定コメント: 検索パスの先頭にプロジェクトルートが存在するか確認。
if path_project_root_ptr not in sys.path:
    # 命令分離: 検索パスの先頭に物理挿入。
    # 【目的】パッケージとして実行する際や、解析ツールの解決エラーを根絶するため。
    sys.path.insert(0, path_project_root_ptr)

try:
    # 命令分離: Win32、Excel操作、GUI、進捗、対話UIの階層インポート。
    # 【Ruff】E402 (Module level import not at top of file) を抑制。
    from core import hc_w32 as w32  # noqa: E402
    from core import hc_xlc as xlc  # noqa: E402
    from ui import hc_gui as gui  # noqa: E402
    from ui import hc_prg as prg  # noqa: E402
    from ui import hc_trm as trm_ui  # noqa: E402
except ImportError:
    # 救済コメント: 階層解決に失敗した場合の直接インポート試行。
    import hc_w32 as w32  # type: ignore # noqa: E402
    import hc_xlc as xlc  # type: ignore # noqa: E402
    import hc_gui as gui  # type: ignore # noqa: E402
    import hc_prg as prg  # type: ignore # noqa: E402
    import hc_trm as trm_ui  # type: ignore # noqa: E402


# ==============================================================================
# メイン処理: trim_cells (原子レベル分解版)
# ==============================================================================


def trim_cells(target_hwnd: Optional[int] = None) -> None:
    """
    Method Name : trim_cells
    Arguments   : target_hwnd (Optional[int]) : 親 Excel の物理ハンドル
    Return      : None
    概要: [VBA呼出] 不要な空白を物理スキャンし、対話 UI を独立プロセスとして起動。
    """
    # 1. 接続環境の原子捕捉。
    # 【目的】Excel の Application, Workbook, Worksheet ポインタを安全に取得するため。
    # 命令分離: Excel 接続情報の取得。
    tuple_context = xlc.get_ctx(hwnd_val=target_hwnd)

    # 変数: ポインタの原子分解。
    ptr_a = tuple_context[0]
    ptr_w = tuple_context[1]
    ptr_s = tuple_context[2]

    # 判定コメント: 接続情報の健全性チェック。
    if ptr_a is None:
        return
    if ptr_s is None:
        return

    # 変数: 親 HWND の物理解決。
    val_xl_h = 0
    # 判定。
    if target_hwnd is not None:
        # 命令分離: 引数のハンドルを優先使用。
        val_xl_h = int(target_hwnd)
    else:
        # 命令分離: アプリポインタから Win32 ID を原子抽出。
        val_xl_h = xlc.get_h(ptr_a)

    # 変数: 進捗 UI ポインタ。
    p_ui = None

    try:
        # --- 走査領域の精密解析 ---
        # 【目的】全セルスキャンを避け、データ実体が存在する範囲に限定するため。
        # 変数: 使用済み矩形捕捉。
        rect_used = ptr_s.used_range

        # 変数: 物理座標および行列サイズの原子抽出。
        val_y1 = rect_used.row
        val_x1 = rect_used.column
        # 変数: 行数。
        ptr_rows_col = rect_used.rows
        val_yn = ptr_rows_col.count
        # 変数: 列数。
        ptr_cols_col = rect_used.columns
        val_xn = ptr_cols_col.count

        # 判定コメント: 走査すべきデータが物理的に存在しない。
        if val_yn < 1:
            return

        # --- 進捗窓の物理具現化 ---
        # 命令分離: ProgressWin インスタンス生成。
        p_ui = prg.ProgressWin(
            "トリミング物理スキャン", parent_hwnd=val_xl_h, wb_for_status=ptr_w
        )

        # 【重要】Win32 所有権リンクの物理確立（Owned Window）。
        if val_xl_h != 0:
            # 変数: 窓ハンドル取得。
            ptr_win_api = p_ui.win_handle
            if ptr_win_api is not None:
                # 命令分離: 所有者セット執行。
                # 【目的】Excel 前面に進捗を固定しつつ、操作を妨げない。
                w32.set_owner(ptr_win_api, val_xl_h)

        # --- 物理吸引スキャンフェーズ ---
        # 変数: 二次元データコンテナ。
        arr_data = None

        # 命令分離: 最適化コンテキスト適用。
        with xlc.Opt(ptr_a):
            # 命令分離: 物理チャンク読込エンジンの執行。
            # 【目的】COM 限界による不応答を回避しながらシート情報をメモリへ転送。
            arr_data = xlc.read_chunk(
                ptr_s, val_y1, val_x1, val_yn, val_xn, p_ui, "全セルを精密スキャン中..."
            )

        # 変数: 指摘箇所レジストリ。
        list_errors = []

        # 命令分離: UI 表示の更新。
        p_ui.update(None, None, "不要な空白を原子特定中...", custom_text="演算中")

        # 機能部位: 原子レベル判定ループ。
        for r_idx in range(len(arr_data)):
            # 変数: 行データ。
            row_vals = arr_data[r_idx]

            for c_idx in range(len(row_vals)):
                # 変数: ターゲットセルの実値。
                val_raw = row_vals[c_idx]

                # 判定コメント: 文字列型データのみを物理的な検査対象とする。
                if isinstance(val_raw, str):
                    # 命令分離: 前後空白の物理トリミング。
                    val_fixed = val_raw.strip()

                    # 判定コメント: 物理的な不一致（空白存在）を検知。
                    if val_raw != val_fixed:
                        # 判定コメント: 内容があることを二次確認。
                        if len(val_raw) > 0:
                            # 変数: シート上の絶対座標。
                            py = val_y1 + r_idx
                            px = val_x1 + c_idx

                            # 命令分離: セル範囲の物理捕捉。
                            ptr_cell = ptr_s.range((py, px))
                            # 変数: アドレス取得。
                            str_addr_raw = ptr_cell.address
                            # 命令分離: $ 記号を抹消したラベルの生成。
                            str_addr = str_addr_raw.replace("$", "")

                            # 命令分離: 指摘情報の原子登録。
                            dict_error_v = {
                                "addr": str_addr,
                                "original": val_raw,
                                "fixed": val_fixed,
                                "y": py,
                                "x": px,
                            }
                            list_errors.append(dict_error_v)

        # 判定コメント: 指摘箇所が皆無であった場合の早期終了。
        if len(list_errors) == 0:
            if p_ui is not None:
                # 命令分離: クローズ。
                p_ui.close()
            # 命令分離: VBA 私書箱への統計報告。
            xlc.save_status(ptr_w, "指摘箇所は物理的に検出されませんでした。")
            return

        # ---------------------------------------------------------
        # 3. サブプロセス起動シーケンス (ハングアップ回避)
        # ---------------------------------------------------------
        # 【目的】Tkinter mainloop が Python 通信スレッドを固めないようにするため。
        # 変数: 一時通信ファイルの生成。
        val_fd, path_pkl = tempfile.mkstemp(suffix=".pkl", prefix="hc_trm_")
        # 命令分離: 低層 OS ハンドルの物理クローズ。
        os.close(val_fd)

        # 変数: ペイロード辞書の構築。
        dict_payload = {
            "errors": list_errors,
            "book": ptr_w.name,
            "sheet": ptr_s.name,
            "parent_hwnd": val_xl_h,
        }

        # 命令分離: 物理シリアライズ執行。
        with open(path_pkl, "wb") as f_out:
            pickle.dump(dict_payload, f_out)

        # 変数: Python 実行環境の特定。
        str_exe_p = sys.executable
        # 変数: コンソール抑制ランナー (pythonw.exe) への置換。
        str_runner_p = str_exe_p.replace("python.exe", "pythonw.exe")

        # 変数: 自分自身のスクリプト絶対パス。
        str_script_p = os.path.abspath(__file__)

        # 命令分離: 独立プロセスの物理フォーク。
        # 【重要】CREATE_NO_WINDOW (0x08000000) により DOS 窓を抑制。
        subprocess.Popen(
            [str_runner_p, str_script_p, path_pkl],
            shell=False,
            creationflags=0x08000000,
        )

        # 命令分離: 解析完了のステータス報告。
        str_rep_msg = (
            f"トリミング解析完了。指摘: {len(list_errors)} 件を対話画面で精査。"
        )
        xlc.save_status(ptr_w, str_rep_msg)

    except Exception as ex_fatal:
        # 異常情報の原子報告。
        xlc.save_status(ptr_w, f"ERROR: スキャン例外 Detail: {ex_fatal}")

    finally:
        # 資源解放。
        if p_ui is not None:
            # 命令分離: クローズ執行。
            p_ui.close()

        # 命令分離: 本来の親 Excel への物理フォーカス還流エンジンの執行。
        w32.focus_parent(val_xl_h)


# ==============================================================================
# クラス: TrimController (サブプロセス用・原子レベル記述)
# ==============================================================================


class TrimController:
    """概要: サブプロセス側での対話制御および Excel 物理修正命令の執行。"""

    def __init__(self, ctx_map: Dict) -> None:
        """【目的】データ復旧と xlwings 物理接続。"""
        # 命令分離: xlwings 動的インポート。
        import xlwings as xw

        # 変数: コンテキスト情報の物理展開。
        self.errors = ctx_map.get("errors", [])
        self.nm_wb = ctx_map.get("book", "")
        self.nm_sh = ctx_map.get("sheet", "")
        self.val_h = ctx_map.get("parent_hwnd", 0)

        # 変数: 現在のインデックス。
        self.ptr = 0
        # 変数: 修正済みフラグ配列。
        self.done = [False] * len(self.errors)

        # 命令分離: 物理ブック・シート接続の捕捉。
        self.wb = xw.books[self.nm_wb]
        self.sh = self.wb.sheets[self.nm_sh]

        # 変数: UI オブジェクトポインタ。
        self.ui = None

    def set_ui(self, ui_obj: Any) -> None:
        """概要: 相互参照のセット。"""
        # 命令分離: 代入。
        self.ui = ui_obj

    def _highlight(self, idx: int, on: bool) -> None:
        """概要: ターゲットセルの物理強調および自動スクロール。"""
        try:
            # 変数: 指摘情報抽出。
            it = self.errors[idx]
            # 変数: 対象レンジ捕捉。
            cell_obj = self.sh.range((it["y"], it["x"]))

            # 判定コメント: 強調 ON 命令。
            if on:
                # 命令分離: 背景色の物理セット（グレー系）。
                cell_obj.color = (200, 200, 200)
                # 命令分離: セルの物理選択。
                cell_obj.select()
                # 命令分離: ウィンドウの物理スクロール執行。
                ptr_win = self.wb.app.api.ActiveWindow
                # 演算: 視認性を高めるためのオフセット行算出。
                val_sr = max(1, it["y"] - 10)
                ptr_win.ScrollRow = val_sr
            else:
                # 判定: 修正済みフラグの確認。
                if self.done[idx]:
                    # 命令分離: 薄緑セット（修正済み）。
                    cell_obj.color = (230, 255, 230)
                else:
                    # 命令分離: カラー属性の物理抹消。
                    cell_obj.color = None
        except Exception:
            # 判定コメント: COM 接続断絶時は沈黙。
            pass

    def on_next(self) -> None:
        """概要: 次の指摘箇所への物理ジャンプ。"""
        # 命令分離: 現状の強調解除。
        self._highlight(self.ptr, False)
        # 演算: インデックスの循環加算。
        val_next = self.ptr + 1
        self.ptr = val_next % len(self.errors)
        # 命令分離: 新しい箇所の物理強調。
        self._highlight(self.ptr, True)

        # 変数: ターゲット要素抽出。
        e_it = self.errors[self.ptr]
        # 命令分離: UI ラベルの物理同期。
        self.ui.update_index(self.ptr, e_it["addr"], e_it["original"], e_it["fixed"])

    def on_prev(self) -> None:
        """概要: 前の指摘箇所への物理ジャンプ。"""
        # 命令分離: 強調解除。
        self._highlight(self.ptr, False)
        # 演算: 循環減算。
        val_prev = self.ptr - 1
        self.ptr = val_prev % len(self.errors)
        # 命令分離: 物理強調。
        self._highlight(self.ptr, True)

        # 変数: 要素抽出。
        e_it = self.errors[self.ptr]
        # 命令分離: UI 表示の同期。
        self.ui.update_index(self.ptr, e_it["addr"], e_it["original"], e_it["fixed"])

    def on_fix_current(self) -> None:
        """概要: 現在表示されているセルのみを物理修正。"""
        try:
            # 変数: 要素抽出。
            it_cur = self.errors[self.ptr]
            # 変数: セルポインタ。
            ptr_c = self.sh.range((it_cur["y"], it_cur["x"]))
            # 命令分離: 修正済み文字列の物理書込。
            ptr_c.value = it_cur["fixed"]
            # 命令分離: 済フラグのセット。
            self.done[self.ptr] = True
            # 命令分離: 自動的に次へ遷移。
            self.on_next()
        except Exception:
            pass

    def on_fix(self) -> None:
        """概要: 全未処理箇所の一括物理修正執行。"""
        try:
            # 命令分離: 最適化。
            with xlc.Opt(self.wb.app):
                # 物理巡回ループ。
                for i_it in range(len(self.errors)):
                    # 判定。
                    if not self.done[i_it]:
                        # 変数: 抽出。
                        e_it = self.errors[i_it]
                        # 命令分離: セル値の原子上書き。
                        self.sh.range((e_it["y"], e_it["x"])).value = e_it["fixed"]
            # 命令分離: 終了シーケンスへ。
            self.on_cancel()
        except Exception:
            self.on_cancel()

    def on_mark(self) -> None:
        """概要: 指摘箇所を黄色背景で物理マークして終了。"""
        try:
            # 命令分離: 最適化適用。
            with xlc.Opt(self.wb.app):
                for e_m in self.errors:
                    # 変数: レンジ捕捉。
                    ptr_m_cell = self.sh.range((e_m["y"], e_m["x"]))
                    # 命令分離: 黄色背景色の物理セット。
                    ptr_m_cell.color = (255, 255, 180)
            # 終了。
            self.on_cancel()
        except Exception:
            self.on_cancel()

    def on_cancel(self) -> None:
        """概要: クリーンアップと窓破棄エンジンの物理執行。"""
        try:
            # 機能部位: 一時強調色の物理抹消。
            for e_c in self.errors:
                # 変数: セル。
                ptr_c_ref = self.sh.range((e_c["y"], e_c["x"]))
                # 命令分離: クリア。
                ptr_c_ref.color = None
        except Exception:
            pass
        # 命令分離: UI 破棄命令の送信。
        self.ui.close()


# ==============================================================================
# メインエントリ (サブプロセス具現化ハンドラ)
# ==============================================================================

if __name__ == "__main__":
    """概要: 対話 UI プロセスの原子レベル物理具現化。"""
    # 変数: 引数リストから Pickle パスを物理スキャン。
    path_captured = None
    for arg_it in sys.argv[1:]:
        if arg_it.endswith(".pkl"):
            # 判定。
            if os.path.exists(arg_it):
                path_captured = arg_it
                break

    # 具現化物理フェーズ。
    if path_captured is not None:
        try:
            # 【目的】COM サーバーとの物理干渉（デッドロック）を防止。
            # 命令分離: 標準出力の抹消。
            sys.stdout = open(os.devnull, "w")
            # 命令分離: 標準エラーの抹消。
            sys.stderr = open(os.devnull, "w")

            # 命令分離: データの物理復元（Unpickle）。
            with open(path_captured, "rb") as f_in_bin:
                d_ctx = pickle.load(f_in_bin)

            # 命令分離: 物理一時資源の抹消。
            try:
                os.remove(path_captured)
            except Exception:
                pass

            # --- 物理具現化シーケンス ---
            # 命令分離: コントローラの生成。
            p_logic = TrimController(d_ctx)

            # 変数: 指摘総数。
            val_err_count = len(d_ctx["errors"])
            # 変数: 親 Excel ハンドル。
            val_h_parent = d_ctx["parent_hwnd"]

            # 命令分離: ポップアップ UI の具現化。
            p_popup = trm_ui.TrimPopupWin(
                p_logic, val_err_count, parent_hwnd_ptr=val_h_parent
            )

            # 命令分離: 相互紐付けの執行。
            p_logic.set_ui(p_popup)

            # 命令分離: 初期箇所の物理強調 ON。
            p_logic._highlight(0, True)

            # 変数: 先頭要素。
            e_init = d_ctx["errors"][0]
            # 命令分離: UI 表示の初期同期。
            p_popup.update_index(0, e_init["addr"], e_init["original"], e_init["fixed"])

            # 【重要】Tk mainloop の物理起動。
            # 命令分離: シングルトンルート経由でのイベント待機。
            gui.get_root().mainloop()

        except Exception:
            # 異常時は物理自決。
            sys.exit(1)
