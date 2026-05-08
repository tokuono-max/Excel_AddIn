# -*- coding: utf-8 -*-
"""
Pythonバージョン: 3.10+
モジュール名: hc_undo
作成日: 2025-12-05
更新日: 2026-01-26
バージョン: 1.0.3
概要:
    Excel 操作の「元に戻す（Undo）」機能を司るサービスモジュール。
    破壊的な加工（集約・削除・変換等）の直前の物理状態を Pickle 形式のキャッシュから最小単位レベルで復元する。
    本バージョンにて、プロジェクト規定に従い「原子」表現を「最小単位」へ完全移行。
    プロパティアクセスを最小単位（1命令1行）に解体し、記述密度と堅牢性を最大化。
    最新の xlc.save_status (Ver 1.0.12) および hc_prg (Ver 1.0.6) に適合。
    非モーダル親子関係（Owner）を完遂し、復元中も Excel の操作性を物理的に維持。

改訂履歴:
    1.0.3: プロパティアクセスの最小単位分解を徹底、最新基盤 v1.0.12 への適合。
    1.0.2: 用語規定（最小単位）の適用、およびステータス復元処理の追加。
    1.0.1: 未使用インポート（Any, hc_cst, hc_gui）の抹消。
"""

import os
import sys
import re
from typing import Optional

# ==============================================================================
# 階層化インポート解決セクション (Ruff/Pylance対応)
# ==============================================================================
# 【目的】svc フォルダ内からプロジェクトルートの基盤モジュールを確実に参照するため。

# 変数: 現在実行中のファイルの物理絶対パスを取得。
path_script_raw = os.path.abspath(__file__)
# 変数: svc ディレクトリの特定。
path_svc_dir_ptr = os.path.dirname(path_script_raw)
# 変数: プロジェクトルート（svc の一つ上）のパスを算出。
path_project_root_ptr = os.path.dirname(path_svc_dir_ptr)

# 判定コメント: 検索パスの先頭にプロジェクトルートが存在するか確認。
if path_project_root_ptr not in sys.path:
    # 命令分離: 検索パスの先頭に物理挿入。
    # 【目的】パッケージとして実行する際や、解析ツールの解決エラーを根絶するため。
    sys.path.insert(0, path_project_root_ptr)

try:
    # 命令分離: 基盤および UI モジュールの階層インポート。
    # 【Ruff】E402 (Module level import not at top of file) を抑制。
    from core import hc_w32 as w32  # noqa: E402
    from core import hc_xlc as xlc  # noqa: E402
    from core import hc_sys as hsys  # noqa: E402
    from ui import hc_prg as prg  # noqa: E402
except ImportError:
    # 救済コメント: 階層解決に失敗した場合の直接インポート試行。
    import hc_w32 as w32  # type: ignore # noqa: E402
    import hc_xlc as xlc  # type: ignore # noqa: E402
    import hc_sys as hsys  # type: ignore # noqa: E402
    import hc_prg as prg  # type: ignore # noqa: E402


# ==============================================================================
# 公開関数: exec_undo (Undo 執行・最小単位分解版)
# ==============================================================================


def exec_undo(target_hwnd: Optional[int] = None) -> None:
    """
    Method Name : exec_undo
    Arguments   : target_hwnd (Optional[int]) : 呼び出し元 Excel の物理ハンドル
    Return      : None
    概要: [VBA呼出] 物理キャッシュから直前の状態を最小単位レベルで復元。
    """
    # ---------------------------------------------------------
    # ステップ 1: 実行コンテキストの厳格な捕捉
    # ---------------------------------------------------------
    # 【目的】復元対象の Application, Workbook, Worksheet を物理特定するため。
    # 命令分離: Excel 接続情報の最小単位取得エンジンの実行。
    tuple_context = xlc.get_ctx(hwnd_val=target_hwnd)

    # 変数: ポインタの解体。
    # 【重要】多重代入を避け、一行ずつ確実に抽出。
    ptr_a = tuple_context[0]
    ptr_w = tuple_context[1]
    ptr_s = tuple_context[2]

    # 判定コメント: Excel 自体が捕捉できない場合は即座に終了（最小単位検証）。
    if ptr_a is None:
        return

    # 判定コメント: 処理対象ブックの有効性確認。
    if ptr_w is None:
        return

    # 判定コメント: 処理対象シートの有効性確認。
    if ptr_s is None:
        return

    # --- 親ウィンドウハンドルの物理解決 ---
    # 変数: 物理 HWND 格納変数。
    val_xl_h_id = 0

    # 判定コメント: 呼出元からの指定有無を確認。
    if target_hwnd is not None:
        # 命令分離: 物理数値へのキャスト執行。
        val_xl_h_id = int(target_hwnd)
    else:
        # 命令分離: Excel インスタンスから動的に Win32 ID を最小単位抽出。
        val_xl_h_id = xlc.get_h(ptr_a)

    # ---------------------------------------------------------
    # ステップ 2: 物理キャッシュの検索と解析
    # ---------------------------------------------------------
    # 【目的】このブック・シートに紐付いた最新のバックアップを物理特定するため。
    # 変数: ブック名称の取得。
    str_wb_name_v = ptr_w.name
    # 変数: シート名称の取得。
    str_sh_name_v = ptr_s.name
    # 変数: 現在のプロセスIDの取得。
    val_current_pid = os.getpid()

    # 命令分離: 固体識別検索キーの最小単位構築。
    # 【重要】マルチインスタンス環境下での混線を物理的に防止する。
    str_raw_key_val = f"{val_xl_h_id}_{val_current_pid}_{str_wb_name_v}_{str_sh_name_v}"

    # 命令分離: 安全なキー文字列への物理置換（ファイル名禁止文字の除去）。
    str_undo_solid_key = re.sub(r'[\\/:*?"<>|]', "_", str_raw_key_val)

    # 命令分離: 物理キャッシュのロード執行。
    # 【目的】Pickle 形式で永続化されたデータをメモリへ復元するため。
    dict_undo_payload = hsys.CacheManager.load(str_undo_solid_key)

    # 判定コメント: 復元可能な履歴が存在しない場合の安全な終了。
    if dict_undo_payload is None:
        # 変数: 通知テキストの構築。
        str_not_found_msg = "元に戻すための物理キャッシュ情報が見つかりませんでした。"
        # 命令分離: VBA 側へステータス情報を物理報告（save_status 執行）。
        xlc.save_status(ptr_w, str_not_found_msg)
        return

    # --- 復元モードの最小単位判定 ---
    # 変数: ヘッダ集約由来のキャッシュかを確認。
    bool_is_structure_undo = "num_rows" in dict_undo_payload

    # 変数: 進捗通知 UI ポインタの初期化。
    p_prog_ui = None

    try:
        # ---------------------------------------------------------
        # ステップ 3: 物理復元シーケンスの最小単位執行
        # ---------------------------------------------------------
        if bool_is_structure_undo:
            # 【目的】ヘッダの高度な構造復旧は、専門機能を持つサービスへ最小単位委譲。
            # 命令分離: モジュールの動的ロード（循環参照の物理防止）。
            from svc import hc_hd_rs

            # 命令分離: 物理復旧関数の呼び出し。
            # 【重要】target_hwnd を伝播し、非モーダル親子関係を維持する。
            hc_hd_rs.restore_header_logic(target_hwnd=val_xl_h_id)

            # 命令分離: VBA 側へ完了沈黙通知（"UI" 信号）。
            # 【重要】xlc 1.0.12 のフィルタリングにより、金庫（成功メッセージ）は保護される。
            xlc.save_status(ptr_w, "UI")

        else:
            # 【目的】一般的なセル加工（日付変換、削除等）を物理的に一括復元。
            # 変数: 行列データの最小単位抽出。
            list_original_matrix = dict_undo_payload.get("data", [])

            # 判定コメント: データの物理実体チェック。
            if not list_original_matrix:
                # 最小単位リターン。
                return

            # 命令分離: 進捗画面（ProgressWin）の物理具現化。
            p_prog_ui = prg.ProgressWin(
                "物理状態の復元中", parent_hwnd=val_xl_h_id, wb_for_status=ptr_w
            )

            # 判定コメント: 所有権リンクの物理確立。
            if val_xl_h_id != 0:
                # 変数: 窓ハンドル捕捉。
                ptr_win_v = p_prog_ui.win_handle
                if ptr_win_v is not None:
                    # 命令分離: Win32 所有権の設定執行。
                    # 【目的】Excel の前面を維持しつつ、操作を妨げない。
                    w32.set_owner(ptr_win_v, val_xl_h_id)

            # 変数: アプリ API ポインタの捕捉。
            ptr_xl_app_api = ptr_a.api
            # 命令分離: Excel インタラクティブ・ロックの物理奪取。
            ptr_xl_app_api.Interactive = False

            # 命令分離: パフォーマンス最適化コンテキスト（画面計算停止）の適用。
            with xlc.Opt(ptr_a):
                # 命令分離: UI メッセージ更新。
                # 【重要】hc_prg 1.0.6 内で is_save=False が適用され、金庫汚染を防止する。
                str_sync_msg = "シート内容を以前の状態へ物理同期中..."
                p_prog_ui.update(None, None, str_sync_msg, custom_text="書込中")

                # 命令分離: 物理チャンク書戻しの執行。
                # 【目的】大規模データ時も OS ハングを物理回避して同期するため。
                xlc.write_chunk(
                    ptr_s,
                    1,
                    1,
                    list_original_matrix,
                    p_prog_ui,
                    "データの物理展開を執行中...",
                )

                # 変数: 使用済み矩形捕捉。
                ptr_used_rect = ptr_s.used_range
                # 命令分離: 全列幅の物理オートフィット執行。
                ptr_used_rect.columns.autofit()

            # --- 最終物理報告文章の構築 ---
            str_rep_fin = "加工直前の物理状態へ正常に復元されました。"
            # 命令分離: ステータス保存執行（金庫への永続同期）。
            xlc.save_status(ptr_w, str_rep_fin)

        # ---------------------------------------------------------
        # ステップ 4: 物理キャッシュの最小単位での破棄
        # ---------------------------------------------------------
        # 【重要】Undo 成功後、使用済みの一次履歴を物理抹消（1世代仕様）。
        # 命令分離: キャッシュ消去エンジンの物理執行。
        hsys.CacheManager.delete(str_undo_solid_key)

    except Exception as ex_undo_fatal:
        # 命令分離: 異常時クローズ。
        if p_prog_ui is not None:
            # 命令分離: 窓資源破棄。
            p_prog_ui.close()

        # 変数: 詳細テキストの構築。
        str_err_text = f"ERROR: 元に戻す処理中に物理的な例外。\\n詳細: {ex_undo_fatal}"
        # 命令分離: エラー情報の物理報告保存。
        xlc.save_status(ptr_w, str_err_text)

    finally:
        # ---------------------------------------------------------
        # ステップ 5: 物理環境の復旧とフォーカス還流
        # ---------------------------------------------------------
        try:
            # 変数: アプリ API の再捕捉。
            ptr_xl_api_fin = ptr_a.api
            # 命令分離: インタラクティブ権限の最小単位復旧。
            ptr_xl_api_fin.Interactive = True
            # 命令分離: 画面描画の最小単位再開。
            ptr_xl_api_fin.ScreenUpdating = True
        except Exception:
            # 判定コメント: 切断時は沈黙。
            pass

        # 命令分離: Z順固定の物理解除（ kill_topmost 執行）。
        w32.kill_topmost(val_xl_h_id)

        # 資源解放。
        if p_prog_ui is not None:
            # 命令分離: クローズ執行。
            p_prog_ui.close()

        # 【重要】ステータス表示の消失を物理解消するための最終復元命令。
        # 【目的】Python プロセス切断の直前に、最新の復元報告を Excel バーへ再掲させる。
        xlc.restore_status(ptr_w)

        # 命令分離: 本来の親 Excel への物理フォーカス還流エンジンの執行。
        # 【重要】ダイアログ消失直後の OS アクティブウィンドウ制御スタックを正常化する。
        w32.focus_parent(h_target_ptr=val_xl_h_id)


# ==============================================================================
# メインエントリポイント (物理デバッグ用)
# ==============================================================================

if __name__ == "__main__":
    """概要: モジュール単体での物理動作テスト用エントリポイント。"""
    # 動作補足: 直接実行時は OS 上のアクティブ Excel を自動ターゲットとして認識。
    try:
        # 命令分離: Undo 執行のテストを最小単位レベルで起動。
        exec_undo()
    except Exception as ex_direct_fatal:
        # 命令分離: エラー表示。
        print(f"Direct execution physically failed: {ex_direct_fatal}")
