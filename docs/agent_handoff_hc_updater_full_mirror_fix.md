# エージェント向け資料：hc_updater full 更新の自己ロック失敗修正

**目的**: full bin 更新時に `hc_updater` が `app\bin` 内から実行されたまま `_mirror_tree` で同ディレクトリを削除しようとして失敗する問題を修正する。

**実装（2026-05）**: 方針 A — full の `app\bin` は `_copy_merge_tree`、addin は `_mirror_tree` 維持。full でも `__delete_list.txt` を適用。

**対象**: `hc_updater.py`, `tests/test_hc_updater_full_apply.py`

**再現ログ**: `WinError 183` on `app\bin` after `mirror moved_running_updater_exe` while `executable=...\app\bin\python.exe`.
