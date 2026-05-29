# Agent 引き継ぎ：本番データ集約 — 進捗キャンセル

**実装日**: 2026-05-28  
**状態**: 実装済み（v0.4.8: 即時ログ・50ms ポール・抽出/結合内チェック）

## 仕様（利用者決定）

- キャンセル範囲: **フォルダ走査 + 集約計算**（Excel 書込み途中は対象外）
- 途中集約データ: **破棄**（マスタへ書かない）
- 新規シート: 作成後・書込み前に中止 → **削除**（`batch_sheet_pending_delete`）
- イベントログ: 一括サマリ **失敗**（`ok=False`, `error=cancelled`）
- 通知: 進捗 `CANCEL` + `MESSAGES.STATUS_CANCEL` → `write_batch_done_notify`
- 結合ブロック前後など **force ポール** あり

## 変更ファイル

| ファイル | 内容 |
|----------|------|
| `svc/data_agg_cancel.py` | 例外・IPC パス・ポール |
| `svc/svc_data_agg_scan.py` | `scan_folder(cancel_check=...)` |
| `svc/svc_data_agg.py` | `_run_batch` 順序変更、`compute_batch_table_rows(cancel_check=...)` |
| `config/ui_data_agg.json` | `MESSAGES.STATUS_CANCEL` |
| `tests/test_data_agg_batch_cancel.py` | ユニット |

## ログ（キャンセル時）

| タイミング | ログ例 |
|------------|--------|
| ボタン押下 | `[DATA_AGG] progress cancel clicked path=...`（`hc_csv.log`） |
| ワーカー検知 | `[DATA_AGG] batch cancel detected ...` / `[DATA_AGG_DIAG] batch_run cancel detected ...` |
| 進捗終了 | `[UI_PROGRESS_DIAG] CANCEL pickle ...` |

進捗ラベルは押下直後に「中止しています…」。

## 手動確認

1. 一括実行 → 走査中にキャンセル  
2. 集約計算中にキャンセル  
3. マスタ未更新・親画面に中止メッセージ  
