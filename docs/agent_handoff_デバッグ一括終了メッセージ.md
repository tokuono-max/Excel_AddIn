# Agent 引き継ぎ：デバッグ一括実行の終了メッセージ（1・2）

**実装日**: 2026-05-28  
**状態**: 実装済み

## 内容

- シナリオモード（`_mode == 0`）・マスタモード（`_mode == 1`）の連続一括が**正常完了**したとき、`QMessageBox.information` で終了メッセージを表示。
- 文言は `config/ui_data_agg.json` の `SCREENS.DEBUG`（`MSG_RUN_ALL_*_DONE`, `DIALOG_RUN_ALL_DONE_TITLE`）。
- 中断時は従来どおりログのみ（ダイアログなし）。

## 変更ファイル

| ファイル | 内容 |
|----------|------|
| `config/ui_data_agg.json` | DEBUG 用ダイアログキー追加 |
| `ui_qt/ui_data_agg_debug.py` | `_continuous_initial_steps`, `_show_continuous_run_done_dialog`, 完了時呼び出し |

## 保留

本番一括の進捗キャンセル（旧 3 番）は未着手。
