# CSV結合（csv_mg）機能仕様書

**文書ID**: SPEC-CSV-MG  
**対象モジュール**: svc_csv_mg / ui_csv_mg  
**作成日**: 2026-03-05  
**更新日**: 2026-03-05  

---

## 1. 概要

### 1.1 目的

複数の CSV ファイルを 1 つの Excel シートへ結合して出力する機能を定義する。  
ユーザーがファイル一覧を選択・並べ替えし、ヘッダの扱いを指定したうえで「結合開始」により、指定シートの末尾（または先頭）にデータを追記する。

### 1.2 呼び出し経路

| 段階 | 対象 | 説明 |
|------|------|------|
| 1 | VBA（リボン等） | ユーザー操作で Python ブリッジを呼び出す |
| 2 | hc_main.invoke(action='merge_csv', target_hwnd=..., sheet_id=...) | ブックを解決し、svc_server へ依頼 |
| 3 | _call_svc_server(ACTION_CSV_MG, book_ptr, sheet_id) | IPC で svc_server に req を投入 |
| 4 | svc_server | book を解決し、svc.svc_csv_mg.merge_csv(book, sheet_id) を実行 |
| 5 | svc_csv_mg.merge_csv | Qt UI サーバ起動 → 結合画面表示 → 結果待ち → 結合処理 |

### 1.3 アーキテクチャ上の前提

- **2 プロセス分離**: サービス層（svc_csv_mg）と UI 層（ui_server + ui_csv_mg）は別プロセス。IPC は Pickle ファイル（req_*.pkl / res_*.pkl / ready_*.pkl / progress_*.pkl）で行う。
- **Excel コンテキスト**: 処理対象のブック・シートは、svc_server が `book` を解決し、`sheet_id`（GUID）でシートを特定する。
- **UI サーバ**: `ui_qt.ui_server.py` が req をポーリングし、`ui_qt.ui_csv_mg.create_dialog` で結合画面を表示する。

---

## 2. 機能一覧

|  No | 機能 | 説明 |
|----:|------|------|
|  1 | 結合画面の表示 | Excel を親にしたモーダルダイアログで、ファイル一覧・ヘッダオプション・操作ボタンを表示する |
|  2 | ファイルの追加 | 「ファイル選択」「フォルダ選択」で CSV を一覧に追加する |
|  3 | 一覧の編集 | 行の上下移動・削除・オールクリアで結合順を変更する |
|  4 | ヘッダ取り扱いの指定 | 3 種類の結合モード（後述）をラジオで選択する |
|  5 | 結合実行 | 「結合開始」で選択 CSV を読み、指定シートへ追記する |
|  6 | 進捗表示 | 読込・Excel 書き込み・列幅調整・完了を進捗ウィンドウで表示する |
|  7 | 完了通知 | 結合結果（ファイル数・行数）をポップアップで表示し、結合画面を再表示する |
|  8 | 行数超過時の扱い | 結合後の総行数が Excel 上限を超える場合は警告し、結合画面へ戻す |

---

## 3. 結合モード（ヘッダの取り扱い）

| モード ID | 表示名（既定） | 動作 |
|-----------|----------------|------|
| **mode_append** | 最初のファイルヘッダを追加 | 先頭ファイルの 1 行目のみをヘッダとして 1 回出力し、2 ファイル目以降の 1 行目はデータとして結合する |
| **mode_replace** | 各ファイルのヘッダ付きで結合 | 各 CSV の 1 行目をヘッダとして都度出力し、続くデータを結合する |
| **mode_preview** | 全てヘッダなしで結合 | 全ファイルの行をヘッダ扱いせず、そのまま連結して出力する |

---

## 4. 処理フロー

### 4.1 全体シーケンス

```
[ユーザー] リボン「CSV結合」クリック
    → hc_main.invoke（merge_csv）
    → ensure_svc_server / _call_svc_server(csv_mg, book, sheet_id)
    → svc_server: merge_csv(book, sheet_id) 実行

[svc_csv_mg]
    1. ensure_qt_ui_server()
    2. req 投入（action=csv_mg, module=ui_qt.ui_csv_mg, result_path, ready_path, sheet_id）
    3. READY_UI を別スレッドで監視 → 受信時に notify_ui_ready（砂時計解除等）
    4. RESULT を同期で監視
    5. status=OK かつ files あり → 進捗 pkl を準備し _merge_files_to_sheet() 実行
    6. 結合完了 → 完了ポップアップ用に _submit_done_ui（done_then_merge）投入
```

### 4.2 結合画面（UI）の動作

- **表示**: Excel をオーナーにしたモーダルダイアログ。設定は `config/ui_csv_mg.json` の MAIN / WINDOW / RIBBON / RADIO / TABLE 等を参照。
- **Excel ロック**: 設定 `COMMON.EXCEL.LOCK_WHEN_OPEN = true` のとき、表示中は Excel の子ウィンドウ操作を無効化する。
- **結果の返し方**: ユーザーが「結合開始」を押し、一覧が 1 件以上なら `get_result()` で `status=OK`, `files=[パス列]`, `radio=モードID` を返す。キャンセルまたは 0 件の場合は `status=CANCEL` または 0 件メッセージで終了。

### 4.3 結合処理（_merge_files_to_sheet）の流れ

1. **CSV 読込**: 各ファイルを pandas で読み、エンコーディングは utf-8-sig → utf-8 → cp932 の順で試行。1 行目をヘッダ、2 行目以降をデータとして保持。
2. **総行数・モード計算**: 選択モードに従い「出力する行」を決定し、総行数 `total_rows` を算出。
3. **ブック・シート解決**: xlwings でアクティブアプリの `book` を取得し、`sheet_guid` で `core_xlc.find_sheet_by_guid` によりシートを特定。未特定時はアクティブシートを使用。
4. **書込開始行**: シートが空なら 1 行目、データがあれば最終行の次から。
5. **最大行数チェック**: `last_row + total_rows > EXCEL_MAX_ROWS`（既定 1,048,000 行）の場合は進捗 pkl に `OVER_LIMIT` とメッセージ・ return_merge_payload を書き、結合を中止。UI 側で警告表示のうえ結合画面を再表示する。
6. **チャンク書込**: 一定行数ずつ（設定可能。既定は列数に応じたチャンク）で `sht.range(...).value = chunk` を実行。進捗 pkl に phase=Excel書き込み、done/total/pct/current_file を更新。
7. **ブロック単位の付加情報**  
   - 各ファイル先頭行のセルにコメントを付与（シート名-行番号 ファイル名、追加行数）。  
   - ジャンプ用に Workbook.Names を追加（シート名 開始行 ファイル名(拡張子除く) → その行の A 列）。
8. **列幅調整**: 書込範囲のセル数が `AUTOFIT_MAX_CELLS`（既定 100,000）以下なら `Columns.AutoFit` を実行。超過時はスキップ（砂時計長時間化防止）。
9. **ステータス保存**: `core_stat.set_status_info` でシートの HC_STATUS_INFO に「シート名：結合ファイル名＋…」形式を設定。ステータスバーにも同一文字列を表示。
10. **進捗 DONE**: progress pkl に status=DONE, phase=完了, pct=100 を書き、進捗ウィンドウを閉じる。
11. **完了ポップアップ**: `_submit_done_ui(done_then_merge)` で ui_csv_mg に依頼。0 行のファイルも含む明細を表示し、OK で結合画面を再表示（テーブルクリアまたは維持は payload で指定）。

---

## 5. 定数・制限値

| 項目 | 値 | 説明 |
|------|-----|------|
| EXCEL_MAX_ROWS | 1,048,000 | 結合後のシート総行数がこれを超える場合は結合を行わず警告する |
| AUTOFIT_MAX_CELLS | 100,000 | 書込範囲のセル数がこれを超える場合は列幅の AutoFit をスキップする |
| 文字コード試行順 | utf-8-sig, utf-8, cp932 | CSV 読込時のエンコーディング検出順 |

---

## 6. 設定ファイル（config/ui_csv_mg.json）

- **MAIN**: 結合画面のタイトル、説明文、リボンボタン（ファイル選択・フォルダ選択・▲▼・削除・オールクリア）、結合開始/キャンセル、メッセージ、ファイル追加/フォルダ選択ダイアログのタイトル・フィルタ。
- **MAIN.RADIO**: ヘッダ取り扱いのラジオ（ENABLED, DEFAULT, ITEMS: mode_append / mode_replace / mode_preview の id, label, tooltip）。
- **MAIN.TABLE**: ファイル一覧テーブルの列定義・行高・表示行数等。
- **COMMON.EXCEL.LOCK_WHEN_OPEN**: 結合画面表示中の Excel 操作ロックの有無。
- **WINDOW / MAIN.WINDOW**: リサイズ可否、最小化/最大化ボタン、CENTER_ON_EXCEL、SHOW_IN_TASKBAR、DEFAULT_WIDTH / HEIGHT、TOPMOST、EXCEL_FRONT_FOLLOW 等。
- **SCREENS.PROGRESS**: 進捗ウィンドウのタイトル・WINDOW 設定。
- **SCREENS.DONE**: 完了通知のタイトル、メッセージ書式、OK ボタン等。
- **SCREENS.DUPLICATE**: 重複追加確認ダイアログ（同一ファイルを再度追加する場合等）のメッセージ・ボタン。

---

## 7. IPC ファイル構成

| 種別 | パス例 | 役割 |
|------|--------|------|
| 要求 | %TEMP%\csv_tool\...\req_*.pkl | svc → ui_server。action, module, parent_hwnd, sheet_id, result_path, ready_path 等 |
| 結果 | ipc_root/results/res_{sheet_id}_*.pkl | ui_server → svc。status, files, radio |
| READY | ipc_root/ready/ready_{sheet_id}_*.pkl | UI 表示完了通知。status=READY_UI で砂時計解除等 |
| 進捗 | ipc_root/progress/progress_*.pkl | 結合処理の phase, done, total, pct, current_file。UI がポーリング表示 |
| 完了再表示 | res_done_then_merge_*.pkl | 完了ポップアップ後に結合画面を再表示するための req |

---

## 8. エラー・例外

- **UI エラー**: create_dialog や exec の失敗時、result に status=ERROR, message, traceback を返す。svc 側はログに記録し、結合は実行しない。
- **最大行数超過**: 上記のとおり OVER_LIMIT を進捗に書き、return_merge_payload で結合画面を再表示。
- **シャットダウン要求**: 処理中に shutdown.flag が立った場合は結合を中断し、進捗に CANCEL を書く。
- **COM/書込例外**: 例外をログに記録し、進捗に status=ERROR を書き、必要に応じて done_then_merge で画面を戻す。

---

## 9. 関連モジュール

| モジュール | 役割 |
|------------|------|
| hc_main | merge_csv エントリ。_ensure_book, _call_svc_server(ACTION_CSV_MG, book, sheet_id) |
| svc.svc_server | req 受信・book 解決・merge_csv(book, sheet_id) 呼び出し |
| svc.svc_csv_mg | merge_csv, _merge_files_to_sheet, 進捗/完了の IPC 投入、READY/RESULT 監視 |
| svc.svc_host | ensure_ui_server（Qt UI サーバ起動） |
| ui_qt.ui_server | req ポーリング、create_dialog 呼び出し、result/ready 書き出し |
| ui_qt.ui_csv_mg | 結合画面・進捗・完了・重複確認の create_dialog、get_result |
| core.core_xlc | find_sheet_by_guid, yield_to_excel 等 |
| core.core_stat | set_status_info, get_status_info（HC_STATUS_INFO） |
| core.core_cursor | notify_ui_ready（READY_UI 受信時の砂時計解除等） |

---

## 10. 改訂履歴

| 日付 | 版 | 内容 |
|------|-----|------|
| 2026-03-05 | 1.0 | 初版作成 |
