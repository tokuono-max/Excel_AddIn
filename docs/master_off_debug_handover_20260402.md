# マスタOFF デバッグ引継ぎメモ（2026-04-02）

## 目的
- シナリオ2/3の表示崩れ（品名列の上書き・未反映）を解消する
- 連続実行時の体感遅延を低減する
- `data_agg_diag.log` で挙動追跡可能にする

## 現在の到達点（最新）
- シナリオ3の結果は反映される状態まで改善済み（ユーザー確認済み）
- 終端で表示が消える問題は解消済み
- ただし速度は依然として重い区間が残る（主に progress 側 compute）

## 変更の要点（実装済み）

### 1) 表示崩れ/未反映対策
- 途中表示は `show_merged_current=False` を基本とし、最終表示タイミングを制御
- `active_slots=0` 項目遷移時の空再描画を抑止（`value_grid_keep` ログ）
- 直近完了項目の rows を再利用するフォールバック追加
  - `_last_master_completed_mi_idx`
  - `_master_off_progress_rows_by_mi`
- 描画対象列を `mi_idx` と分離
  - `_master_off_display_mi_idx`
  - 終了時に `mi_idx=10` でも `display_mi=9` で描画整合

### 2) 遅延対策（部分）
- `master_off_progress` で不要な再計算を削減
  - `no_active_slots` の無駄計算回避
  - `end_of_item` 再利用ロジック追加
- `colvals` 取得は高速化と正確性の折衷
  - 前段ステップは progress 由来の列値利用
  - 後段（late_step）は extract 経路へ戻す（シナリオ3反映優先）

### 3) 追跡ログ追加
- `master_step_prepare_start`
- `master_off_progress start/cache=miss/cache=hit/reuse=.../skip=...`
- `master_off_extract_start/end`
- `master_off_colvals_from_progress`（`skip=late_step`含む）
- `master_off_render`（`display_mi` 含む）
- `master_off_finalize_request`
- `master_off_finalize_apply_end_of_continuous`

## 最新ログの読み取り（@data_agg_diag.log 1-43）
- 最終反映:
  - `master_off_finalize_apply_end_of_continuous mi_idx=10 ... last_completed_mi=9`
  - `master_off_render mi_idx=10 display_mi=9 ...`
- シナリオ3の反映経路:
  - `master_off_colvals_from_progress skip=late_step ... step_idx=2`
  - `master_off_extract ... col_head=['PSU-0', 'PSU-1']`
- 重い区間（残課題）:
  - `master_off_progress cache=miss mi_idx=9 step_idx=2 rows=16 elapsed_ms=9356`
  - `master_off_progress cache=miss mi_idx=9 step_idx=3 rows=18 elapsed_ms=9809`

## 既知の残課題（優先順）
1. **速度**: `step_idx=2/3` の `master_off_progress cache=miss` が 9-10秒級
2. `mismatch_first5` が 4-5 のケースが残る（見た目が最終的に正しくても途中差分が大きい）
3. 連続実行中の `compute_batch` コストが高く、項目数23で効いている

## 次環境での推奨作業（根本改善）

### A. progress の 1回計算キャッシュ化（最優先）
- 目標: 同一 `mi_idx` で `step_idx=2/3` の再計算を避ける
- 方針案:
  - `progress_rows_final` を `mi_idx` 単位で1回計算してキャッシュ
  - 各 step はマスク表示のみ（再計算なし）
  - キャッシュキー候補:
    - `mi_idx`
    - `id(self._scenario_for_dry_run)`
    - `len(scan_paths)` + 先頭数件ハッシュ
    - `active_slot_indices` の内容

### B. 途中 mismatch の削減
- `master_off_render` の `gcell/prog_rows` 差分が大きいケースを監視
- 必要なら「最終表示まで gcell を完全無効」に寄せる（現状はほぼその方針）

### C. 検証手順（再現性）
1. マスタOFF + 連続実行
2. 品名項目（mi=9）で step0→1→2→3 を実行
3. 次項目（mi=10, active_slots=0）へ遷移
4. 最終表示確認
5. `data_agg_diag.log` で以下を確認
   - `display_mi=9`
   - `skip=late_step` + `extract ... PSU-0/PSU-1`
   - `cache=miss step_idx=2/3 elapsed_ms`

## 関連ファイル
- `ui_qt/ui_data_agg_debug.py`（主変更）
- `svc/data_agg_master_preview.py`（step preview シナリオ構築）
- `svc/svc_data_agg.py`（`compute_batch_table_rows`, `filter_file_paths_for_master_preview`）
- ログ: `%TEMP%\csv_tool\data_agg_diag.log`

## 補足
- 目的は「見え方維持」のまま高速化。機能追加よりも、再計算抑制と表示参照整合を優先。
- 次環境で最初に着手すべきは **A（progress final cache）**。
