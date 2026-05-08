# -*- coding: utf-8 -*-
"""
リボン公開名（customUI の tag / core.ribbon_invoke.invoke の action）と
svc_server が受け付ける内部 action 名の対応。

単一の dict に集約し、ルート hc_main.py（ブリッジ）・ribbon_invoke の許可 action の二重定義を避ける。
値は svc/svc_server.py の _ACTION_MAP のキーと一致させること。

ribbon_invoke.invoke の finally で notify_wait_form_ready を呼ぶ対象もここから派生する
（load/save/merge/split は READY_UI 側が WaitForm を閉じる）。
DEFER は Qt 初回表示（ui_server）と svc_server 完了で閉じる（data_agg は依頼のみ即 return するため svc は失敗時のみ）。
"""

from __future__ import annotations

# リボン control.Tag（= invoke の action） -> svc_server._ACTION_MAP のキー
RIBBON_PUBLIC_TO_SVC_ACTION: dict[str, str] = {
    "load_csv": "csv_ld",
    "save_csv": "csv_sv",
    "merge_csv": "csv_mg",
    "split_csv": "csv_sp",
    "normalize_header": "hd_nr",
    "insert_shuka_header": "hd_in",
    "undo_last_action": "undo",
    "check_duplicates": "dupli",
    "delete_empty_rows": "row_dl",
    "delete_empty_cols": "col_dl",
    "convert_date_ymd": "dt_ymd",
    "convert_date_ymd_hm": "dt_hm",
    "trim_spaces": "trm_ex",
    "show_help": "help",
    "check_for_updates": "update_check",
    "run_data_agg": "data_agg",
}

RIBBON_INVOKE_ACTION_KEYS: frozenset[str] = frozenset(RIBBON_PUBLIC_TO_SVC_ACTION.keys())
RIBBON_TARGET_SVC_ACTION_KEYS: frozenset[str] = frozenset(
    RIBBON_PUBLIC_TO_SVC_ACTION.values()
)

# READY_UI / notify_ui_ready 系で砂時計・WaitForm を閉じる公開 action（invoke finally では閉じない）
RIBBON_ACTIONS_READY_UI_CLOSES_WAITFORM: frozenset[str] = frozenset(
    ("load_csv", "save_csv", "merge_csv", "split_csv")
)

# invoke finally では早めに閉じず、Qt 初回 Show（ui_server）＋必要なら svc ハンドラ完了で閉じる
RIBBON_ACTIONS_DEFER_WAITFORM_DISMISS_TO_UI: frozenset[str] = frozenset(
    (
        "normalize_header",
        "insert_shuka_header",
        "check_duplicates",
        "delete_empty_rows",
        "delete_empty_cols",
        "convert_date_ymd",
        "convert_date_ymd_hm",
        "trim_spaces",
        "undo_last_action",
        "show_help",
        "run_data_agg",
    )
)

# ribbon_invoke.invoke の finally で notify_wait_form_ready を呼ぶ公開 action（現状は空になりうる）
RIBBON_INVOKE_FINALLY_NOTIFY_WAITFORM: frozenset[str] = (
    RIBBON_INVOKE_ACTION_KEYS
    - RIBBON_ACTIONS_READY_UI_CLOSES_WAITFORM
    - RIBBON_ACTIONS_DEFER_WAITFORM_DISMISS_TO_UI
)

# svc_server: ハンドラ終了後に notify（UI 未表示の早期 return や設定エラー経路の救済）
# data_agg を含めない: 成功時は依頼送出のみで即 return するため、閉じは ui_server の初回 Show に任せる
SVC_ACTIONS_NOTIFY_WAITFORM_AFTER_HANDLER: frozenset[str] = frozenset(
    (
        "hd_nr",
        "hd_in",
        "dupli",
        "row_dl",
        "col_dl",
        "dt_ymd",
        "dt_hm",
        "trm_ex",
        "undo",
        "help",
        "update_check",
    )
)
