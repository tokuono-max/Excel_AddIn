# エージェント向け：マスタデバッグ高速化（2026-05-29）

## 目的

- マスタ項目デバッグ（モード1）の体感速度を改善する
- **シナリオ単位の積み上げ表示**（結果サマリ・ログ・結合プレビュー）を維持する
- 集約結果は **表示専用**（本番マスタ・他機能へ流用しない。終了時破棄でよい）

## スナップショット（ロールバック用）

| 項目 | 内容 |
|------|------|
| コミット | `9c44108` — `chore: マスタデバッグ高速化検討前の作業ツリースナップショット` |
| 用途 | 以降の最適化を試す前の正。`git reset --hard 9c44108` で戻せる |
| 注意 | その後の高速化実装は **未コミット** の可能性あり。作業前に `git status` を確認 |

---

## 実装済み（2026-05-29）

### 1. 項目内一括 compute（one-shot）

**対象**: 結合キー探索 **なし**・同一項目に **cell/csv ソースが2つ以上**

| ファイル | 内容 |
|----------|------|
| `svc/data_agg_master_preview.py` | `master_preview_one_shot_eligible()` |
| `core/core_env.py` | `data_agg_master_progress_one_shot_enabled()` |
| `ui_qt/ui_data_agg_debug.py` | 一括 compute・先読み・バックフィル・colvals 拡張 |

**挙動**:

- 最終ステップ（`n_pick == len(active)`）で `use_max_sources_for_current_item=True` により1回だけ全 active ソースを `compute_batch`
- 先読み: 結合探索なし時は **次ステップ** ではなく **項目内フル** をバックグラウンド計算
- フル計算後: `n_pick = 1 .. n_act-1` をバックグラウンドで段階 compute し `_mpv_progress_rows_step_cache` を埋める
- `mpv_colvals_from_progress`: フル行キャッシュ済みなら late step も progress 列切り出し（extract 回避）

**無効化**:

```text
DATA_AGG_MASTER_ONE_SHOT=0
```

**先読み**（従来どおり）:

```text
DATA_AGG_MASTER_OFF_PREFETCH=1
# または HC_DIAG_DATA_AGG_MASTER_PREFETCH=1
```

### 2. テスト

- `tests/test_data_agg_master_preview_one_shot.py`（新規）
- 回帰: `tests/test_data_agg_regression_preview.py` ほか

---

## 制約（変えてはいけない UX）

要求定義 §3.1.3 および現行 UI の前提:

| 維持するもの | 実装の要点 |
|--------------|------------|
| ステップごとにサマリ・ログを追記 | `_upsert_summary_row_at`、`_log_append_master_scenario_row` |
| 未到達シナリオ分は結合表に出さない | `scenario_for_stepped_preview` の `n_pick` / 未来項目 `sources=[]` |
| セッションで結果を積み重ね | 項目をまたいでも既定でクリアしない |
| 本番マスタへ書かない | `preview_master_mode`、ドライランのみ |

**高速化の原則**: 「表示の段階性」と「裏の再計算回数」は切り離せる。段階表示を維持したまま、同一入力の `compute_batch` 回数を減らす。

---

## ODN164 / ODN375 で one-shot が効かない理由

ユーザー実測ログ（2026-05-29）より:

| 条件 | ODN164（例: PT mi=21） | ODN375（例: MAC RMT mi=23） |
|------|------------------------|-----------------------------|
| `join_defs` | **あり** → one-shot **対象外** | あり |
| `active_slots` | **1**（シナリオ1つのみ） | 多くは 1 |
| 主ボトルネック | `mpv_progress cache=miss` **約16秒/step** | **約21秒/step** |
| join パス | 数十 ms 程度（主因ではない） | 約1.5秒（全体の一部） |

→ **今回の one-shot は ODN 系ではほぼ無効**。別経路の最適化が必要。

---

## ログ分析サマリ（2026-05-29）

### ODN164 — PT番号（mi_idx=21）

参照: `%TEMP%\csv_tool\hc_csv.log`, `data_agg_diag.log`

| 区間 | 目安 | 備考 |
|------|------|------|
| step 0 `mpv_progress` miss | 約3秒 | `paths 10→5`、`items=22` |
| step 0 `mpv_extract` | 約1.3秒 | `col_count=550`（progress 列切り出しは 0 件で extract に落下） |
| step 1 `mpv_progress` miss | **約16秒** | 10ファイル × 22項目まで抽出 |
| `join_search_done` | 約69ms | pool=5500 |
| `value_grid_rebuild` | 500×22 列をステップ中に複数回 | UI コスト |

### ODN375 — MAC RMT 等（mi_idx=23）

| 区間 | 目安 | 備考 |
|------|------|------|
| `mpv_progress` miss / step | **約21秒** | 18〜19ファイル × 24項目 |
| `join_search_done` | 約1.5秒 | pool=324 等 |
| 品名 `item_timing` | ファイルあたり約220〜280ms × ファイル数 | `prim_count=18`, `source_count=3` |

### 二重処理

1ステップでよくあるパターン:

1. `_mpv_progress_batch_rows` → フル `compute_batch`（結合プレビュー表）
2. 続けて `_mpv_extract_colvals`（当該列・同じファイル走査）

`mpv_colvals_from_progress` が `col_count=0` のとき extract にフォールバックしているケースあり。

---

## 残課題・次に効く改善（優先順）

前提: **マスタデバッグのデータは表示のみ・終了時破棄可** → 本番 `compute_batch` より積極的な省略が可能。

### A. 完了項目の表示キャッシュ（最優先・ODN 向け）

**現状**: PT（mi=21）の各 step で、項目 0〜20 も含め **毎回フル `compute_batch`**（`j < mi` はフルソースのため）。

**案**:

- マスタ項目離脱時（`_capture_master_leave_item`）に、その時点の **結合行スナップショット**（または列ベクトル）をセッションキャッシュ
- 次項目では列 `0..mi-1` はキャッシュから合成、列 `mi` のみ step 分を compute / extract
- ウィンドウ閉鎖でキャッシュ破棄

**厳しさの確認**（実装前に要合意）:

- **厳格**: 各ステップで全列が本番同等である必要がある → キャッシュ＋差分検証が必要
- **緩和**: 当該シナリオ列＋サマリが合えばよい → フル `compute_batch` をステップごとに呼ばない設計が可能

### B. `active_slots=1` 項目の軽量経路

PT・機器番号など:

- 結合プレビュー表を **毎ステップ 22列×全ファイル** で再計算しない
- **前項目キャッシュ ＋ 当該列 extract** で結果タブを更新（サマリ積み上げは現状維持）

### C. プレビュー行数上限

ログ上 `max_primary_rows=550`, `max_table_rows=500`:

- `config/ui_data_agg.json` の `MASTER_PREVIEW_READ_ROWS` / `MAX_VALUE_ROWS` をデバッグ向けに引き下げ（例: 50〜100）だけでも I/O・描画が減る

### D. extract と progress の一本化

- `mpv_colvals_from_progress` が空になる原因を修正
- step 0 で `compute_batch` 後に同じ列の `mpv_extract` を呼ばない

### E. マスタプレビューでのファイル並列

- 本番は `DATA_AGG_FILE_PARALLEL_WORKERS=auto` 有効
- `preview_master_mode` では現状 **並列オフ**（`svc/svc_data_agg.py` の `use_file_parallel`）
- 表示専用なら有効化の余地あり（`xlsx_workbook_scope` のスレッド安全性を要確認）

### F. join あり向けのデバッグ専用省略（難度高）

- 現在項目＋結合依存項目だけプールに載せる
- または項目完了時点の join プールをキャッシュし差分更新

### G. UI 間引き

- 連続実行中の `_render_mpv_grid`（500行×N列）を項目完了／全体完了に寄せる（一部は `finalize` / `_finish_continuous_run` で制御済み）

---

## 診断ログの見方

| ログ | パス |
|------|------|
| 一般 | `%TEMP%\csv_tool\hc_csv.log` |
| data_agg 詳細 | `%TEMP%\csv_tool\data_agg_diag.log`（`HC_LOG_DIAG` / 診断フラグ時） |

**注目キーワード**:

| パターン | 意味 |
|----------|------|
| `mpv_progress cache=miss` … `elapsed_ms=` | 結合プレビュー再計算（重い） |
| `mpv_progress cache=hit_step` / `prefetch=done` | キャッシュ／先読み命中 |
| `mpv_colvals_from_progress` / `mpv_extract` | 列値取得経路 |
| `compute_batch join_search_done` | join パス完了（ODN では全体の一部） |
| `item_timing` … `prim_count=` | 行数・ソース数が大きいと重い |
| `value_grid_rebuild` … `prog_rows=500` | UI 再描画コスト |

---

## 関連ファイル

| 種別 | パス |
|------|------|
| UI（主） | `ui_qt/ui_data_agg_debug.py` |
| プレビュー組み立て | `svc/data_agg_master_preview.py` |
| 結合核 | `svc/svc_data_agg.py`（`compute_batch_table_rows`, `preview_master_mode`, `filter_file_paths_for_master_preview`） |
| 環境 | `core/core_env.py` |
| 要求 | `docs/etc/データ集約ツール要求定義書.md` §3.1.3 |
| 旧引継ぎ | `docs/master_off_debug_handover_20260402.md` |
| ODN フィルタ | `docs/agent_handoff_データ集約_2件_20260528.md` |

---

## 検証手順（再現）

### one-shot が効くケース（結合なし・複数シナリオ）

1. 結合 `join_defs` のない項目で active シナリオが2つ以上
2. マスタデバッグでステップ実行
3. `data_agg_diag.log` で `prefetch=done` / `cache=hit_step`、最終 step で `one_shot` 相当の1回計算を確認

### ODN164

1. ODN164 シナリオ・走査パスをメインからデバッグ起動
2. PT番号（mi≈21）で step 0→1
3. `mpv_progress cache=miss` の `elapsed_ms` を記録
4. `join_search_done` が全体の何％か確認（通常は小さい）

### ODN375

1. 同様に MAC RMT 等の項目で step 実行
2. ファイル数×品名 `item_timing` の累積を確認

### 回帰テスト

```powershell
cd C:\Project\Python\Excel_AddIn
python -m pytest tests/test_data_agg_master_preview_one_shot.py tests/test_data_agg_regression_preview.py tests/test_data_agg_scenario_odn_integration.py -q
```

---

## 過去メモとの関係

- `docs/master_off_debug_handover_20260402.md` の **A. progress の1回計算キャッシュ化** は、2026-05-29 実装で **結合なし・複数シナリオ** に部分対応（`use_max_sources` + 先読み + バックフィル）
- **ODN 系**は同メモの課題が未解消。本ファイルの **A. 完了項目の表示キャッシュ** が次の本命

---

## 未コミット作業

- 高速化実装（`master_preview_one_shot_*` 一式）がコミット済みかは `git log -1` で確認すること
- スナップショット `9c44108` の **後** にコミットする場合は、メッセージ例:  
  `perf: マスタデバッグ one-shot 先読み（結合なし複数シナリオ向け）`
