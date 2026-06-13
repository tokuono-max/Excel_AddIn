# CSV読込 回帰チェックリスト

`svc_csv_ld` / `ProgressDialog` / `DoneDialog` の UI 責務整理後、**毎回**次を確認する。

## 前提

- Excel を**再起動**してから試す（常駐 `ui_server` にコード変更を反映）
- ログ: `%TEMP%\csv_tool\hc_csv.log` / `hc_csv_diag.log`

## 1. 連続読込（小・中・大）

| # | 操作 | 期待 |
|---|------|------|
| 1 | 10行程度の CSV を読込 | 進捗バーが一気に100%にならない |
| 2 | 続けて 5万行程度を読込 | 同上。完了通知が Excel **手前** |
| 3 | 続けて 80万行程度を読込 | 進捗が段階更新。完了通知が手前 |

## 2. 完了通知

| # | 操作 | 期待 |
|---|------|------|
| 4 | 完了通知が出たら OK を押さずに次の読込を開始 | ファイル選択が開く（WaitForm タイムアウトなし） |
| 5 | 完了通知の OK を押す | ダイアログが閉じる |

## 3. 他アプリ前面化

| # | 操作 | 期待 |
|---|------|------|
| 6 | 進捗表示中にブラウザ等を Excel 前に出す | **前面に留まれる**（進捗は Excel オーナーだが TOPMOST でない） |
| 7 | 完了通知表示中に別アプリを前面に出す | 同上 |

## 4. ログ（任意）

```
[CSV_LD] progress close ack ok path=...progress_csv_ld_closed_...
[EXCEL_RESTORE] restored hwnd=...
```

- ACK が `EXCEL_RESTORE` **より先**に出ること
- `EXCEL_RESTORE` が1回（正常終了時は sheet_id 付きで1回）

## 自動テスト

```powershell
python -m pytest tests/test_csv_ld_ui_lifecycle.py tests/test_csv_ld_perf_helpers.py tests/test_progress_dialog_pct.py tests/test_csv_tool_progress_ui.py -q
```

詳細な手動確認は [csv_tool_progress_regression_checklist.md](csv_tool_progress_regression_checklist.md) を参照。
