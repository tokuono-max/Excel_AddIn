# 本番一括 table_ms 高速化（2026-06-23 実装済み）

## 概要
`merged_rows` → `table_rows` の行単位ループ＋進捗同期が `table_ms` のボトルネック（実測 148s / 168s compute）だったため、チャンク一括組み立てに置換。

## 変更
- `svc/svc_data_agg.py`
  - `_table_assembly_chunk_size()` … 既定 2000、`DATA_AGG_TABLE_ASSEMBLY_CHUNK_SIZE`
  - `_append_merged_rows_to_table_chunked()` … 共通ヘルパ
  - ファイル単位一覧組立（旧 3902 行付近）
  - 結合プール一覧組立（旧 4245 行付近）
- `svc/data_agg_batch_compute.py`
  - `_batch_hook` 内のキャンセル二重読込削除（`_ph` → `_poll_cancel` に委譲）
- `tests/test_data_agg_table_assembly_chunked.py`

## 環境変数
- `DATA_AGG_TABLE_ASSEMBLY_CHUNK_SIZE` … チャンク行数（下限 100、既定 2000）

## 検証
```bash
pytest tests/test_data_agg_table_assembly_chunked.py tests/test_data_agg_csv_extract_perf.py tests/test_data_agg_batch_stability.py tests/test_data_agg_link_write_and_match_keys.py tests/test_data_agg_join_search.py -q
```

実機: Test_01 本番一括で `compute_batch_timing` の `table_ms` が 10 秒未満を目標。
