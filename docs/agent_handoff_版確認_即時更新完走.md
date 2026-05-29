# Agent 引き継ぎ: 版確認「すぐに更新」の即時シーケンス完走

## 実装済み（2026-05-25 追記）

- `core/update_process_cleanup.py`
  - `mutex_blocks_pending_apply` / `should_relax_svc_mutex_for_interactive_defer`
  - リボン即時更新（`skip_apply_confirm` + `hc_svc_server` + defer）時は **自プロセスの svc mutex をゲートから除外**
- `bootstrap/update_bootstrap.py` … 上記を `pending_apply` の mutex 待ちに適用

## 実装済み（2026-05-24）

- `core/packaged_update.py`
  - `run_interactive_bin_apply_now` … リボン／起動プロンプト共通の即時 apply
  - `_apply_pending_update_with_retry` … `concurrent_apply` 時リトライ
  - `_apply_succeeded_for_interactive` … `applied` または `deferred_to_updater` のみ成功
- 表示版: `display_name_for_install_scope` + `_try_sync_display_version` + `notify_installed_apps_list_changed`（bootstrap / hc_updater 既存）

## 完走の定義

1. 「すぐに更新」確定 → `pending` 予約
2. 同セッションで `apply_pending_update` → 進捗 →（既定）`hc_updater` + `PROGRESS_DEFER_DONE`
3. Excel 全終了後 bin 適用 → `VERSION.txt` 更新
4. Uninstall レジストリ `DisplayName`/`DisplayVersion` + `SHChangeNotify`（hc_updater またはインライン完了時）

## ログ確認

`{install}\logs\hc_update.log` で `interactive_apply: deferred_to_updater` または `applied=true`。`skipped': 'concurrent_apply'` のみで終わらないこと。
