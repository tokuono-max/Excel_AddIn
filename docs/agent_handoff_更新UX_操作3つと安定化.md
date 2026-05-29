# エージェント向け資料：更新 UX（操作3＋UAC別）と安定化

**目的**: bin 更新を安定化し、アプリ側の操作者判断は **3 回**（更新／Excel 終了／完了後再起動）。**UAC は別操作**。経過は詳細表示。**patch 優先維持**。

**実装メモ（2026-05）**:
- full `app\bin` → `_copy_merge_tree`（`hc_updater.py`）
- インタラクティブ「すぐに更新」でアプリの管理者 Yes/No を出さない（UAC は `maybe_apply_pending_bootstrap_update`）
- `updater_last_result.json` で成功も次回 1 回通知可能

**テスト**: `tests/test_hc_updater_full_apply.py`, `tests/test_update_apply_notify.py`
