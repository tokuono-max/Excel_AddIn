# 進捗バー RUN 時の pct / done/total 監査（2026-06-04）

## UI 共通（0.1.21+）

`ui_dialog_progress.ProgressDialog`: **RUN 時に pickle の `pct` キーがあれば `done/total` より優先**（0–99）。
DONE 時は従来どおり 100% で閉じる。

## 修正済み

| 機能 | モジュール | 内容 |
|------|------------|------|
| CSV保存 | `svc_csv_sv` 1.3.8 | 進捗 1.3.7 + 性能: 軽量 valid チェック・大容量日付正規化スキップ |

## 他機能の確認結果（コードレビュー）

| 機能 | モジュール | seq | 明示 pct | done=total で 99% 張り付き | 異常時終了 |
|------|------------|-----|----------|---------------------------|------------|
| CSV読込 | `svc_csv_ld` | あり | あり（`calc_progress_pct`） | UI修正で緩和 | DONE/ERROR あり |
| CSV結合 | `svc_csv_mg` | `_progress_write_monotonic` | あり | 要実機（多くは pct 併記） | ERROR/DONE |
| CSV分割 | `svc_csv_sp` | なし | あり（`min(99,...)`） | **UI修正で緩和**（pct 併記） | DONE |
| 重複削除等 | `svc_trm_ex` | なし | あり（99 明示） | **UI修正で緩和** | 失敗時 DONE |
| 日付変換等 | `svc_dt_ymd` / `svc_row_dl` / `svc_col_dl` | DONE 時 999 | あり | 要実機 | DONE |
| Undo | `svc_undo` | — | — | — | 進捗別経路 |

### 追加対応が必要になり得るもの

- **`svc_csv_sp`**: `seq` 未付与（従来 `-1` で常に採用のため致命的ではないが、mg と同様 monotonic 化は将来改善可）
- **長時間ブロックで pickle 更新なし**: 保存系は `to_csv` 中は 75% 表示のまま（完了まで待つ仕様）

## 再現テスト（手動）

1. CSV保存（数千行以上）→ バーが 0→〜49→50→75→100、完了ダイアログ
2. 保存中にタスクマネージャで pythonw 落とさない限り 99% で固まり続けない（75% まで進む）
