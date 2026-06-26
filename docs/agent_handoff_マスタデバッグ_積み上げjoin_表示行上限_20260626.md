# エージェント向け：マスタデバッグ — 表示行上限内の積み上げ join（2026-06-26）

## 合意サマリ（一言）

**マスタデバッグは「結果一覧に表示されている行」（表示行上限 `MASTER_DEBUG_DISPLAY_ROWS` まで）だけを正とする。各ステップでは前ステップの表示結果を土台に、当ステップで新たに必要なファイルだけディスクから読み、join 列を積み上げる。前ステップで読んだファイルの再読込はしない。表示されていない行・データは無視してよい（本番一括の全件一致は求めない）。**

---

## 実装要点（2026-06-26）

| フラグ / 関数 | 役割 |
|---------------|------|
| `master_preview_stacked_join` | UI が prior table_rows seed 時に立てる |
| `join_search_seed_from_table_rows` | seed が表示行由来である印 |
| `master_preview_stacked_join_active()` | 上記を判定（`data_agg_master_preview_perf.py`） |
| `_filter_file_paths_for_master_preview_stacked_host()` | paths を当ステップホストのみに限定 |
| `preview_extract_item_allowlist = [mi_idx]` | 積み上げ時は当項目列だけ extract |
| `master_preview_join_read_full_files = False` | 積み上げ時は full scan しない |
| pool extend スキップ | stacked 時は seed プールにファイル行を足さない |

### 触るファイル

- `svc/data_agg_master_preview_perf.py`
- `svc/svc_data_agg.py` — path filter, compute_batch pool
- `ui_qt/ui_data_agg_debug.py` — `_mpv_apply_join_item_debug_diag`
- `tests/test_master_preview_join_priority.py`

### 成功条件

- 表示行上限内で join 列が埋まる
- 積み上げステップの `paths_after_filter` / `file_passes` がホストのみ
- 初回結合項目（prior 表示行なし）は従来どおり side+host を読む
