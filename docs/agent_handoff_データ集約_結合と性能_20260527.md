# エージェント向け：データ集約 — 結合代入・性能（実装済み 2026-05-27）

**実装版**: `svc/svc_data_agg.py` **0.4.4**

## 実装内容

### 結合キー検索（`join_defs`）

1. **二段パス**: 全ファイル走査・マージ後に `_apply_join_key_search_across_file_passes` で結合。`table_rows` は結合後に出力。
2. **スライス k ↔ `__iter_index`**: `n_prim > 1 && n_join > 1` では一致行のうち `__iter_index == k` のみへ主キー・連携を書込み。
3. **単行ホスト (`n_prim == 1`)**: `iter_index == k` に絞り、複数一致時は最小 `iter_index` 1 行のみ。
4. **プール分離**: `_join_host_needs_cross_file_pool` で横断判定。同一ファイル内は `_join_search_pool_scope` でホスト file_path に限定。横断時は **当時点の global_pool のうち Excel 出力対象行**（`_TableRowEmitContext.should_emit`）のみで比較索引を構築し、前段結合・連携で積み上げた列値を照合する。
5. **出力行フィルタ**: `_row_should_emit_to_table` で結合ホスト専用ファイル行（錨列なし）を除外。
6. **`match_keys` 併用**: 結合後に file_pass 単位で `join_on_match_keys` を実行（従来経路維持）。
7. **長不一致警告**: `[DATA_AGG_WARN] join_values length mismatch`

### テスト

- `tests/test_data_agg_link_write_and_match_keys.py` … 横断結合・ペア iter・同一ファイル AS30 結合・回帰
- `tests/test_data_agg_scenario_odn_integration.py` … ODN-164 実データ（手元にファイルがある場合）

### 実データ確認（2026-05-27 自動実行）

**ODN-164**（`C:\Project\データ集約テストファイル\ODN164\`）

| 項目 | 結果 |
|------|------|
| 出力行数 | **3597**（従来 7195 の倍化なし） |
| 機器番号 / MAC | 3597 件すべて非空 |
| PT番号 / 製番 | 1182 件（MAC 一致分のみ） |
| join_events | 0（match_keys 空のため） |

**ODN375** … 専用 xlsx のパスは環境依存。`DATA_AGG_ODN375_XLSX` を設定して integration テストを実行。

### 未実装（別タスク）

- CSV 繰り返しの 1 回読み
- xlsx used range 限定
- ネットワーク向けローカルミラー

## 確認シナリオ

- `ODN-164出荷履歴試験用.json`
- `ODN375_ALL_1NCM90024.json`

手動: `HC_DIAG_DATA_AGG_JOIN=1` で `[DATA_AGG_JOIN_DUMP]` を確認。
