# データ集約メイン画面 起動性能ベースライン（改善前）

改善後の `hc_csv.log` / `create_dialog elapsed_ms` と比較用。

## 計測日・環境

- 日時: 2026-06-22 10:16 頃
- ログ: `%TEMP%\csv_tool\hc_csv.log`

## 改善前の数値

| 指標 | 値 | ログ根拠 |
|------|-----|----------|
| `run_data_agg` (svc) | **~1503 ms** | `[SVC_SERVER] exec done action=data_agg ms=1503` |
| **`create_dialog` 全体** | **31762 ms (~32秒)** | `[UI_DATA_AGG] create_dialog ok ... elapsed_ms=31762` |
| `create_dialog enter` | 10:16:36.491 | `[DATA_AGG_UI] create_dialog enter` |
| `create_dialog ok` | 10:17:07.224 | 上記 elapsed_ms |

## 内訳（推定）

- **~28–30秒**: `_DataAggMainWindow.__init__` 内の同期 `_on_scan` → `scan_folder`
- **~2–3秒**: Excel COM（`prepare_dialog` / `showEvent`）

## 改善後の目標

| 指標 | 目標 |
|------|------|
| `create_dialog elapsed_ms` | **≤ 5000 ms** |
| 走査完了 | 表示後バックグラウンド |

## 比較方法

1. 同一フォルダでリボン「データ集約」を起動
2. `hc_csv.log` の `[UI_DATA_AGG] create_dialog ok ... elapsed_ms=` を確認
3. 本ファイルの **31762** と比較
4. `[DATA_AGG_UI] folder_scan_async start/done` で走査完了を確認
