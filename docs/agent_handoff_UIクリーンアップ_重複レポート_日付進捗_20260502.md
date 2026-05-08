# Agent 引き継ぎ: UI クリーンアップ集約 + 重複レポート表 + EXCEL_FRONT_FOLLOW 修正 + 日付進捗前面

**作成日:** 2026-05-02  
**目的:** 以下を一括で実装する Agent 向けの要件・根拠・変更箇所をまとめる。

---

## 背景（なぜ機能間で影響が出るか）

- Qt UI サーバは **単一プロセス**。`EXCEL_FRONT_FOLLOW`（`_front_follow_dialog`）、モードレス一覧、Excel の `Interactive`／ロックは **グローバル共有**。
- `apply_window_config` は `destroyed → stop_front_follow` を付けるが、**モードレスの hide では `destroyed` が来ない**機能があり、フックが残る。
- レポート等を閉じた後も **フォアグラウンドフックが生きたまま**、次機能の `ensure_front` が **既に削除された C++ ウィジェット**を触ると、`Internal C++ object (...) already deleted` や **Excel 操作不能・完了通知が一瞬／出ない**につながる（`hc_csv_diag.log` の `ensure_front error` 参照）。

**設計方針:** 機能終了時に **共有リソースをクリア**し、**正常／キャンセル／異常の各終了経路を 1 か所の teardown に集約**する（冪等・二重呼び可を目安）。

---

## タスク A: 重複検出レポート一覧（`ui_qt/ui_dupli.py` / `DupliReportDialog`）

### 要件

1. **テーブルヘッダを左寄せ**（データ集約メインの `horizontalHeader().setDefaultAlignment(AlignLeft | AlignVCenter)` と同系）。
2. **モード A（通常の 3 列: 行・座標・重複内容）**  
   - 先頭列〜最後の一つ前まで: 内容に合わせた幅（`ResizeToContents`）。  
   - **最終列（重複内容）**: ウィンドウリサイズに **追従して伸縮**（`Stretch`）。  
   - 既存の `setStretchLastSection(True)` だけでは「最終列だけ伸びる」が、列モードを明示した方が安定。
3. **モード B**（`_link_col == 2`）は既存の `MODE_B_WINDOW` / `VALUE_COL_MAX_WIDTH` 等のブロックを維持し、**モード A 用の列設定と競合しないよう順序に注意**（ヘッダ左寄せは両モード共通で可）。

### 実装メモ

- 設定参照: `config/ui_dupli.json` の `SCREENS.REPORT.COLUMNS`（`summary` が重複内容列）。
- テーブル生成ブロック: `QTableWidget` 作成〜`resizeColumnToContents` ループの後、`if not _is_mode_b:` で `QHeaderView` の `setSectionResizeMode` を設定。
- バージョン・History を 1 行更新。

---

## タスク B: EXCEL_FRONT_FOLLOW 残骸・削除済みウィジェット対策（`ui_qt/ui_common.py`）

### 要件

1. **`ensure_front(w, parent_hwnd)`**  
   - `bring_to_front(Excel)` の **前**に `shiboken6.Shiboken.isValid(w)` を確認。無効なら `stop_front_follow()` して `return`。  
   - 既存の `ui_dialog_progress.py` の `Shiboken.isValid` 利用と整合。
2. **`ensure_front` の `except`**  
   - 例外メッセージに `already deleted` または `Internal C++ object` が含まれる場合、`stop_front_follow()` を呼ぶ。
3. **`_handle_foreground_event`**  
   - `QTimer.singleShot(ensure_front)` の **前**に、`_front_follow_dialog` が非 `None` かつ `not Shiboken.isValid(d)` なら `stop_front_follow()` して `return`。

### 実装メモ

- `stop_front_follow` は既に **複数回呼び可**（docstring 済み）。  
- History を 1 行更新（`__version__` はプロジェクト流儀に合わせる）。

---

## タスク C: 重複レポート閉鎖時にフック停止（`ui_qt/ui_dupli.py`）

### 要件

- **`DupliReportDialog.closeEvent` の先頭**（遅延 hlclr 分岐に入る前で可）で、ベストエフォート:

  ```python
  try:
      from ui_qt.ui_common import stop_front_follow
      stop_front_follow()
  except Exception:
      pass
  ```

- データ集約メインの `_teardown_before_hide_main` と同趣旨（hide では `destroyed` が来ない前提の補完）。

---

## タスク D: 日付変換・日付時刻変換の進捗の前面挙動（設定 JSON）

### 現象

- 進捗が Excel より手前に出にくい。  
- 他アプリを前面にしたとき **Excel が前に出る**ように感じる。

### 要件

- `config/ui_dt_ymd.json` および `config/ui_dt_hm.json` の **`SCREENS.PROGRESS.WINDOW`** に例として次を追加（既存キーとマージ）:

  - `"CENTER_ON_EXCEL": true`
  - `"EXCEL_FRONT_FOLLOW": true`

- **`EXCEL_KEEP_FOREGROUND` は付けない**（`apply_window_config` 内のポーリングで、他アプリ前面時にまで Excel→ダイアログを繰り返し前面化し得るため）。

### 実装メモ

- 進捗は `ui_dt_ymd.create_dialog` / `ui_dt_hm.create_dialog` で `create_progress_dialog` に `_deep_merge(main, progress)` した `progress_cfg` を渡している。JSON の `PROGRESS.WINDOW` に書けば `apply_window_config` に届く。
- モジュール先頭の Version/History を 1 行ずつ更新してよい。

---

## タスク E（横断）: 終了経路の集約とクリーンアップ

### 方針（ドキュメント化のみでよいが、触れるファイルを直すなら）

- 各モードレス／進捗で **単一の `_teardown` または `finalize`** に  
  `stop_front_follow`（必要なら）・`_remove_from_modeless`・Excel ロック解除・タイマ停止を集約。
- **`reject` / `closeEvent` / キャンセル**は可能な限りその **1 か所**を経由（データ集約 `_DataAggMainWindow` を参照実装にできる）。
- **冪等**にする。

### 参照実装

- `ui_qt/ui_data_agg.py`: `_teardown_before_hide_main`（`stop_front_follow` + `_remove_from_modeless`）。

---

## 検証手順（開発モード）

1. **重複チェック** → レポート表示 → ヘッダ左寄せ・列幅・リサイズで最終列が追従するか。  
2. レポートを閉じた直後に **別機能**（空白行削除・日付変換など）を実行し、**完了通知・Excel 操作**が正常か。`hc_csv.log` に `ensure_front error` / `already deleted` が出ないこと。  
3. **日付変換・日付時刻変換**の進捗が Excel より手前に出やすいか。ブラウザ等を前面にしたとき **Excel だけが勝手に最前面にならない**か（完全制御は OS 依存のため、主観＋ログで確認）。

---

## 主要ファイル一覧

| ファイル | 変更内容 |
|----------|----------|
| `ui_qt/ui_common.py` | `ensure_front` / `_handle_foreground_event` の Shiboken ガードと例外時 `stop_front_follow` |
| `ui_qt/ui_dupli.py` | レポート表ヘッダ・列モード；`DupliReportDialog.closeEvent` で `stop_front_follow` |
| `config/ui_dt_ymd.json` | `SCREENS.PROGRESS.WINDOW` に `EXCEL_FRONT_FOLLOW` / `CENTER_ON_EXCEL` |
| `config/ui_dt_hm.json` | 同上 |
| `ui_qt/ui_dt_ymd.py` / `ui_dt_hm.py` | 任意: Version/History 1 行 |

---

## ログで成功を見る目安

- `[EXCEL_FRONT_FOLLOW] stop_front_follow` がレポート閉鎖付近で出る（タスク C）。  
- `[EXCEL_FRONT_FOLLOW] ensure_front error ... already deleted` が **再発しない**（タスク B/C）。

---

## 注意

- `dist/CSV_Tool/config/` とリポジトリ直下 `config/` の両方を配布物で使う場合は、**同様の JSON 変更が必要か**プロジェクト運用を確認すること。
