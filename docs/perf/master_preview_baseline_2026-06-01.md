# マスタデバッグ（mode=1）性能ベースライン

計測日: 2026-06-01（ユーザー実機ログ `hc_csv.log` / `hc_csv_diag.log`）  
シナリオ: ODN375（19 ファイル）、ODN164（20 ファイル）  
状態: MAC RMT 行順修正 **後**・本ドキュメントの colvals/backfill 最適化 **前**

## ODN375（mi=9 → 23、連続実行）

| 区間 | 時刻 | 所要時間 |
|------|------|----------|
| 品名 step0 開始 → MAC RMT 完了 | 22:18:49.522 → 22:19:08.184 | **約 18.7 s** |

### 項目別 `mpv_progress cache=miss`（elapsed_ms）

| mi | 項目（ログより） | step | elapsed_ms | rows |
|----|------------------|------|------------|------|
| 9 | 品名 | 0 | 266 | 0 |
| 9 | 品名 | 1 | 831 | 18 |
| 9 | 品名 | 3 | 1357 | 324 |
| 12 | 出荷番号 | 0 | 1633 | 324 |
| 12 | 出荷番号 | 1 | 1311 | 324 |
| 22 | MAC LOC | 0 | 1547 | 324 |
| 22 | MAC LOC | 1 | 1243 | 324 |
| 23 | MAC RMT | 0 | 1259 | 324 |
| 23 | MAC RMT | 1 | 1279 | 324 |

### 項目別 `mpv_colvals_from_progress`（step0）

| mi | elapsed_ms | row_count | 備考 |
|----|------------|-----------|------|
| 22 | **1548** | 324 | 直後に progress miss 1547ms（二重 compute 相当） |
| 23 | 1 | 324 | cache=hit_step |
| 12 | **1634** | 324 | 同上 |

### その他

- `mpv_progress_backfill`: mi=9 完了後 **2 回**（22:18:53.408, 53.997）
- `mpv_grid_render`（mi=23）: `prog_rows=324`, `sync_grid_rows=500`
- `mpv_final_table_rows mi=23`: **324** 行（修正後・正）
- `mpv_join_search seed`: **なし**（行順修正確認済み）

## ODN164（mi=21 PT 番号）

| 処理 | elapsed_ms |
|------|------------|
| mpv_progress step0 | 5549 |
| mpv_colvals_from_progress step0 | **5551** |
| mpv_progress step1 | 8400 |
| cross_join pool | 11000 |

## 比較時のログキーワード

- `[DATA_AGG_DIAG] mpv_colvals_from_progress`
- `[DATA_AGG_DIAG] mpv_progress cache=miss`
- `[DATA_AGG_DIAG] mpv_extract_end`
- `[DATA_AGG_DIAG] mpv_progress_backfill` / `caller=mpv_progress_backfill`
- `[DATA_AGG_PROBE] compute_batch` + `caller=`
- `[DATA_AGG_DIAG] mpv_grid_render`（`prog_rows` / `sync_grid_rows`）

改善後は同一操作で再計測し、上表と突き合わせる。

---

## 2026-06-01 実装した最適化（本 baseline の直後）

1. **step0 列プレビュー**: 未キャッシュ時は `extract` のみ（`_mpv_progress_batch_rows` を起動しない）
2. **`_mpv_can_colvals_from_progress`**: `step_idx<=1` で常に True にしていた条件を削除（キャッシュがあるときのみ）
3. **backfill**: 連続実行中（`_continuous_busy` / `_master_step_loop_busy`）はスキップ
4. **グリッド行数**: `sync_grid_rows` を `prog_rows` に合わせる（500 枠の空行削減）

### 改善見込み（再計測で確認）

| 対象 | ベースライン | 見込み |
|------|--------------|--------|
| MAC LOC step0 colvals | 1548 ms | extract のみ **~430 ms** 級 |
| 出荷番号 step0 colvals | 1634 ms | 同上 **~1.2 s 短縮/項目** |
| ODN375 mi=9→23 合計 | ~18.7 s | backfill 2 回削減 + 上記で **数秒短縮** |
| mi=23 grid | sync 500 行 | sync **324** 行 |
| ODN164 PT step0 colvals | 5551 ms | extract **~400 ms** 級（要再計測） |

---

## 再計測結果（2026-06-01 22:26〜22:27・最適化後）

### ODN375 連続（mi=9 step0 → mi=23 完了）

| 指標 | 改善前 | 改善後 | 差分 |
|------|--------|--------|------|
| 区間所要時間 | **18.7 s** | **12.3 s** | **−6.4 s（約 34%）** |

### step0 列プレビュー（colvals 経路）

| mi | 項目 | 改善前 | 改善後 |
|----|------|--------|--------|
| 12 | 出荷番号 | colvals **1634 ms** | extract **15 ms** |
| 22 | MAC LOC | colvals **1548 ms** + progress 1547 ms | extract **511 ms** のみ（step0 で progress miss なし） |
| 23 | MAC RMT | colvals 1 ms（キャッシュ） | extract **520 ms**（strategy=extract_first） |
| 21 | PT（ODN164） | colvals **5551 ms** | extract **391 ms**（**−93%**） |

### その他確認

- `mpv_progress backfill=skip reason=continuous_run`（mi=9）— backfill **0 回**
- `sync_grid_rows=324` = `prog_rows=324`（ODN375）
- `mpv_final_table_rows mi=23`: **324 行**
- `mpv_join_search seed`: **なし**
- ボトルネック残: ODN164 PT step1 `cache=miss` **8550 ms**（改善前 8400 ms と同程度）

---

## ODN164 PT 横断結合 — 追加最適化（実装済み・要再計測）

### 分析（22:27 ログ）

| 処理 | 時間 | 割合 |
|------|------|------|
| `mpv_progress` step1 mi=21 | **8550 ms** | 支配的 |
| `join_pass` / cross_join index | **118 ms** | **約 1%** |
| step0 extract（済） | 391 ms | — |

cross_join の索引構築は既に速い。**20 ファイル × 22 項目の再抽出**が遅い。

### 対策 `preview_extract_item_allowlist`

横断結合ホスト（PT）では match_keys・結合列・link 先のみ Excel 読取。

### 見込み

| 指標 | 22:27 実測 | 見込み |
|------|------------|--------|
| PT step1 | 8550 ms | **3000〜4500 ms**（**約 45〜65% 短縮**） |
| 抽出項目/ファイル | 〜22 | **〜5〜8** |

ログ: `mpv_cross_join_extract_allowlist`, `master_preview_extract_allowlist`
