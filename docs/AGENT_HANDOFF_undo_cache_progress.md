# Agent 手渡し資料: Undo 進捗・キャッシュ削除・終了通知

**実装メモ（2026-05-03）:** `svc_undo` 1.7.6 でキャッシュ削除を進捗フェーズ `PHASE_UNDO_CACHE_DELETE`（pct 96）に含め、削除後に `_undo_progress_done`。`_undo_progress_done` の `done_delay_ms` を 200ms、`sleep` を 0.12s に短縮。`config/ui_undo.json` / `dist/CSV_Tool/config/ui_undo.json` に `MESSAGES.PHASE_UNDO_CACHE_DELETE` を追加。

## 1. 背景（ユーザー課題）

- **Undo 成功後**、ヘルプや他機能のモーダルが **Excel の背後に隠れる**ことがある（失敗経路では再現しにくい）。
- **進捗が「完了」して閉じてから**、**終了通知（UNDO_DONE）が出るまで**に **明らかな空白**があり、体感が悪い。
- ログ・計測上、その空白の多くは **`CacheManager.delete(str_undo_key)` 前後**や、`suspend_sheet_updates` の **`with` 終了後**の Excel 負荷に重なる可能性が示唆されている。

参考ログ（ユーザー提供）: `before_cache_delete` と `after_cache_delete` の間が約 1.6 秒程度取れる例あり。

## 2. すでに入っている関連変更（確認用）

| 内容 | 場所 |
|------|------|
| ~~復元成功時 `finally` の `bring_to_front` をスキップ~~ **1.7.7 で撤去**（常に `bring_to_front`） | `svc/svc_undo.py` |
| Undo 進捗 IPC の `excel_lock: False`（他機能進捗に寄せた変更） | `svc/svc_undo.py` `_submit_undo_progress_ui` |
| 前面追従の実験（0.2.72–73）は撤回済み | `ui_qt/ui_common.py` 0.2.74 |

**切り分け結果:** ヘルプの `config/ui_help.json` で **`TOPMOST: true`** または **`EXCEL_FRONT_FOLLOW: false`** にすると前面問題は緩和する。根本の「Undo 後だけ」のトリガーは成功経路の負荷・ウィンドウ順序の組み合わせ。

## 3. 今回 Agent に依頼したい主題

### 3.1 目的（優先）

**キャッシュ削除を「進捗の一部」として扱い、進捗が閉じる前にユーザーに見せる。**

- いま: 復元本体 → `_undo_progress_done`（`DONE` + 短い `sleep`）→ **`CacheManager.delete`** → `_show_undo_done_dialog`
- 望ましいイメージ: 復元本体の進捗のあと、**「キャッシュ削除中」等のフェーズを進捗 Pickle に書き、UI の進捗がまだ開いている間に `delete` を実行** → 完了後に `_undo_progress_done`（または最終フェーズのあと `DONE`）→ 進捗閉鎖 → `_show_undo_done_dialog`

**ユーザー理解の確認:** 「進捗バーにキャッシュ削除も入れる」＝ **Undo 全体の進捗として削除完了まで `DONE` にしない／100% にしない**、で合っている。

### 3.2 副次（任意・別 PR 可）

- **`_undo_progress_done` 内の `done_delay_ms`（350）と `time.sleep(0.38)`** を、体感 **~300ms 台**に下げるか、UI 側遅延と重複しないよう整理する。
- **代替案:** `delete` を **`_show_undo_done_dialog` 終了後**に移す（通知までの待ちから削除を外す）。仕様上「OK までキーが残る」ことの影響を要確認。

## 4. 実装時に触る想定ファイル

| ファイル | 内容 |
|----------|------|
| `svc/svc_undo.py` | `exec_undo` データ復元成功分岐: `_undo_progress_done` と `CacheManager.delete` の順序・間に `_undo_progress_phase` 相当の書き込み。構造復元分岐（`is_structure_undo`）も同様に要検討。 |
| `config/ui_undo.json` | 新フェーズ用の `MESSAGES` キー（例: `PHASE_UNDO_CACHE_DELETE`）と `_header` の説明追記。 |
| `dist/CSV_Tool/config/ui_undo.json` | 配布物をビルド手順に従い同期する場合 |

参考:

- `_undo_progress_done` は `done_delay_ms` と `time.sleep(0.38)` を含む: `svc_undo.py` 内 `_undo_progress_done`
- 進捗 UI は `ProgressDialog` が Pickle の `RUN` / `DONE` と `done_delay_ms` を解釈: `ui_qt/ui_dialog_progress.py`

## 5. 技術メモ

- **データ復元**は `phase_total=8` で `_submit_undo_progress_ui`（`svc_undo`）。最終フェーズの番号・`pct` と、新フェーズ（削除）の整合を取ること。
- **構造復元**は `phase_total=3`。同様に削除をどこに挿すか決めること。
- `_undo_progress_done` を **`delete` の前**に呼ばないこと（現状は `delete` の前に `_undo_progress_done` が来ている箇所があるので、**削除を進捗に含めるなら `DONE` は削除完了後**）。
- 進捗が閉じるまでの間、**同じ Undo キーで二重実行されないか**（リボン連打等）は既存仕様と照合すること。

## 6. 受け入れ条件（テスト観点）

1. 大きいシートで Undo 実行時、**進捗に「キャッシュ削除」相当のメッセージが表示される**。
2. **進捗が閉じてから終了通知までの無言時間**が、現状より短く感じられる、または説明できる（削除が進捗に含まれる）。
3. Undo 成功後、**再度同じ操作で Undo できない**（1 回だけ戻す）既存仕様を壊さない。
4. 構造復元・データ復元の**両分岐**で動作確認。

## 7. ログ（任意改善）

`finally` の `bring_to_front` は成功・失敗問わず常に実行（1.7.7）。前面まわりを追う場合は `hc_csv_tool.diag.undo` 向けに 1 行ログを足すと切り分けしやすい。

---

**作成日:** 2026-05-03（会話コンテキストに基づく）  
**想定読者:** 実装 Agent（本リポジトリを編集する）
