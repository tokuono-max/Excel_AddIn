# CSV 抽出高速化 — 実装済み（2026-06-23）

## 概要
- **A**: `xlsx_workbook_scope` 内で CSV を1回読み `csv_dfs`（Polars DF）に保持
- **B**: link/join 列一括読取 + `postprocess_link_rule_value_batch`
- **C**: 主キー縦/横反復の DF 列 slice 一括
- **D**: `infer_schema_length=0` / `try_parse_dates=False`
- **E**: バッチスコープ内 legacy 再読込禁止（`DataAggCsvReadError`）
- **F**: CSV 読込中の progress_hook（`_extract_cell_rules_series_fast_map` / `_extract_cell_rule_series_fast`）
- バッチ中の CSV は **ファイル再読込フォールバックなし**（失敗時 `DataAggCsvReadError`）

## 変更ファイル
- `svc/svc_data_agg_extract.py` (v0.1.12)
- `svc/data_agg_value_post.py` — `postprocess_link_rule_value_batch`
- `svc/svc_data_agg.py` — `precache_csv_matrix_for_file` 呼び出し
- `tests/test_data_agg_csv_extract_perf.py`

## テスト
```bash
pytest tests/test_data_agg_csv_extract_perf.py
pytest tests/test_data_agg_batch_stability.py tests/test_data_agg_link_write_and_match_keys.py
```
