# エージェント向け：結合キー代入先行の修正（実装済み 2026-05-28）

**実装版**: `svc/svc_data_agg.py` **0.4.5**

## 変更内容

`_narrow_join_matched_rows_for_write` に `cross_file` を追加。

| 条件 | 代入行 |
|------|--------|
| `cross_file=True` または `n_prim==1` | 値一致した行すべて |
| 同一ファイルかつ `n_prim>1` かつ `n_join>1` | `__iter_index==k` のみ |

`_apply_join_key_search_write` / `_apply_join_key_search_link_write` から `_apply_join_key_search_across_file_passes` の `cross` を渡す。

## テスト

- `tests/test_data_agg_link_write_and_match_keys.py` … `test_cross_file_join_writes_all_mac_matches_ignore_iter` 追加、`test_join_search_link_empty_on_matched_rows_overwrite` 期待値更新（n_prim==1 は一致行すべて）
- 既存: `test_paired_join_respects_iter_index`, `test_cross_file_join_writes_to_anchor_row_only`
- `tests/test_data_agg_scenario_odn_integration.py` … `test_odn164_pt_seq_join_on_matching_mac`（紐づけ C5/P5/J5 と MAC 一致行の PT・製番）、`test_odn375_mac_loc_rmt_join_on_matching_device`

## 確認

```powershell
python -m pytest tests/test_data_agg_link_write_and_match_keys.py tests/test_data_agg_join_search.py tests/test_data_agg_scenario_odn_integration.py -v
```

ODN-164 実データ: `C:\Project\データ集約テストファイル\ODN164\`
