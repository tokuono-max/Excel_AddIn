# エージェント向け：データ集約 2 件（実装済み 2026-05-28）

**実装版**: `svc/svc_data_agg.py` **0.4.6**

## 1. マスタデバッグ ODN164 — 2つ目シナリオ空白

- **原因**: `filter_file_paths_for_master_preview` が複数 `file_pattern` を AND → paths=0
- **対策**: パターン2種類以上は **OR**（和集合）。1種類のみは従来どおりその pattern のみ
- **本番**: `preview_master_mode` 時のみ呼び出し → **本番一括に影響なし**

## 2. 一括完了時に集約データシートをアクティブ

- **対策**: `_run_batch` で `sheet_out = sheet` 初期化、`_activate_output_sheet()` で `sheet_out.activate()`
- **本番**: 書込みロジックは不変。新規シート出力時に完了後の前面シートが集約先になる（UX）

## テスト

- `tests/test_data_agg_regression_preview.py` … `test_filter_file_paths_for_master_preview_*`
- 結合・統合: `tests/test_data_agg_scenario_odn_integration.py`
