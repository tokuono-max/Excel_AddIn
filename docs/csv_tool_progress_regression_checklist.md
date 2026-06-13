# CSV Tool 進捗 UI 回帰チェックリスト（読込・保存・結合・分割）

`ProgressDialog` / `core.csv_tool_progress_ui` / 各 `svc_csv_*` の進捗・砂時計改善後に確認する。

## 前提

- Excel を**再起動**してから試す（常駐 `ui_server` にコード変更を反映）
- ログ: `%TEMP%\csv_tool\hc_csv.log`

## 自動テスト

```powershell
python -m pytest tests/test_csv_tool_progress_ui.py tests/test_csv_ld_ui_lifecycle.py tests/test_csv_ld_perf_helpers.py tests/test_progress_wait_cursor.py tests/test_progress_dialog_pct.py -q
```

## 1. CSV読込

| # | 操作 | 期待 |
|---|------|------|
| 1 | 小ファイル読込 | バーが段階表示。「完了」は100%付近。仕上げ中…→完了 |
| 2 | ファイル選択直後 | 進捗画面上で砂時計（Excel を触らなくても） |
| 3 | 連続読込 | 完了通知が Excel 手前。ACK→restore 順 |

## 2. CSV保存

| # | 操作 | 期待 |
|---|------|------|
| 4 | 保存先確定直後 | 進捗が即表示。砂時計が進捗上で見える |
| 5 | 小シート保存 | バーが一気に100%にならない（creep 有効時） |
| 6 | 完了通知 | 従来どおり表示・OK で閉じる |

## 3. ファイル結合

| # | 操作 | 期待 |
|---|------|------|
| 7 | 結合実行開始 | 進捗表示・砂時計 ON |
| 8 | 小ファイル数結合 | バー段階表示（creep） |
| 9 | 完了通知 | 結合完了ダイアログ。Excel 復帰 |

## 4. ファイル分割

| # | 操作 | 期待 |
|---|------|------|
| 10 | 分割保存開始 | 進捗表示・砂時計 ON |
| 11 | 少数ファイル分割 | バー段階表示 |
| 12 | 重複確認で「分割実施」 | 保存が開始され分割画面に戻らない |

## 5. ログ（任意）

- `[CSV_TOOL] immediate progress shown feature=ld|sv`
- `CURSOR_WAIT_ON: Run(Main.ForceCursorOnProgress)`

## 環境変数（共通）

| 変数 | 既定 | 説明 |
|------|------|------|
| `HC_CSV_TOOL_PROGRESS_POLL_MS` | 40 | 進捗ポーリング間隔 |
| `HC_CSV_TOOL_PROGRESS_BAR_CREEP_PCT` | 2 | バー 1 ティックの増分（%） |
| `HC_CSV_LD_*` | （同上） | 後方互換の別名 |
