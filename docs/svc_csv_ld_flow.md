# svc_csv_ld 画面推移（時系列）

CSV読込機能の、ユーザーから見える画面・通知の流れを時系列で示します。

**時間計測（区間 A/B・`pick_to_done_ms` 等）**: **`docs/csv_ld_perf_measurement.md`** を参照。

---

## 全体の流れ（簡略）

```
[Excel] リボン「CSV読込」クリック
    → [1] ファイル選択ダイアログ（Qt）
    → [2] （任意）分割確認メッセージ（Win32）
    → [3] 進捗画面（Qt）表示～閉じる
    → [4] 完了通知（ステータスバー / HC_NOTIFY_RETV）
```

---

## 時系列（詳細）

| 時刻 | 発生場所 | 画面・処理 | 補足 |
|------|----------|------------|------|
| **T0** | Excel / VBA | ユーザーがリボン「CSV読込」をクリック | VBA が RunPython 等で Python を呼び出す |
| **T1** | svc_server | `load_csv()` 開始。ui_server へ req 投入（action=csv_ld） | 砂時計が出る場合あり |
| **T2** | ui_server | req を取得。ready_path に READY_UI を書き込み | まだダイアログは出ていない |
| **T3** | ui_server | **ファイル選択ダイアログ表示**（ui_qt.ui_csv_ld） | タイトル・フィルタは config/ui_csv_ld.json から取得 |
| **T4** | svc_server | _watch_ready が READY_UI を検知 → `notify_ui_ready()` | 砂時計解除 |
| **T5** | ユーザー | ファイルを選んで「開く」または「キャンセル」 | キャンセル時はここで終了（以降なし） |
| **T6** | ui_server | ダイアログ終了。result_path に結果を書き込み（status=OK/CANCEL, path=...） | |
| **T7** | svc_server | _watch_result が結果を取得。OK かつ path ありなら `_do_load_csv()` 開始 | |
| **T8** | svc_server | 行数が MAX_ROWS_PER_SHEET 超のときのみ **分割確認 MessageBox**（Win32）表示 | 「～枚のシートに分割して読み込みますか？」 |
| **T9** | svc_server | progress_path に phase 1 を書き、進捗表示用 req を投入。**2.5 秒スリープ** | 進捗窓が表示・ポーリング開始するまでの待ち |
| **T10** | ui_server | 進捗用 req を取得 → **進捗画面表示**（ui_qt.ui_csv_ld, モデルレス） | 表示内容: 「1/4 ファイル解析中」0% |
| **T11** | svc_server | CSV 読込・Excel 書込ループ。progress_path を phase 2 で更新 | 「2/4 Excelへ書き込み中」と done/total・pct |
| **T12** | svc_server | 書込完了後、phase 3 を書き **列幅調整（オートフィット）** 実行 | 「3/4 列幅調整中」95% |
| **T13** | svc_server | progress_path に **DONE** を書き込み | |
| **T14** | ui_server | 進捗画面が DONE を検知。「**完了**」100% 表示 → **1.5 秒後に進捗画面を閉じる** | ui_common.ProgressDialog の仕様 |
| **T15** | svc_server | HC_STATUS_INFO / HC_NOTIFY_RETV を設定。ステータスバーに「CSV読込終了｜…」を表示 | 完了通知（ポップアップは VBA 側で制御） |
| **T16** | Excel | load_csv から制御が戻る。必要に応じて VBA が完了メッセージを表示 | |

---

## 画面ごとの表示内容

### 1. ファイル選択ダイアログ（T3～T5）

- **表示元**: ui_qt.ui_csv_ld（config/ui_csv_ld.json の MAIN.TITLE / MAIN.FILTER）
- **表示**: 「読み込む CSV ファイルを選択してください」＋ CSV ファイル選択
- **終了**: ユーザーが「開く」または「キャンセル」

### 2. 分割確認メッセージ（T8、条件付き）

- **表示元**: Win32 MessageBox（svc_csv_ld 内）
- **条件**: ファイル行数が MAX_ROWS_PER_SHEET（100万行）を超えるとき
- **表示**: 「総行数 ○○ 行を検知しました。全 ○ 枚のシートに分割して読み込みますか？」
- **終了**: 「はい」で続行、「いいえ」で処理中止

### 3. 進捗画面（T10～T14）

- **表示元**: ui_qt.ui_csv_ld の action "progress"（config/ui_csv_ld.json の SCREENS.PROGRESS、ui_common.ProgressDialog）
- **表示内容の遷移**:
  - 1/4 ファイル解析中（0%）
  - 2/4 Excelへ書き込み中（done/total と pct で更新）
  - 3/4 列幅調整中（95%）
  - 4/4 完了（100%）→ 2.5 秒表示後に閉じる
- **終了**: DONE 検知から 2.5 秒後に自動で閉じる

### 4. 完了通知（T15～）

- **ステータスバー**: 「CSV読込終了｜シート名：… ｜ 行数：… ｜ …」
- **HC_NOTIFY_RETV**: ポップアップ用テキスト（VBA が参照して表示する想定）

---

## プロセス・スレッドの関係

- **svc_server**: load_csv → _watch_result（ブロック）→ _do_load_csv → _execute_jit_import。進捗は pkl に書き、進捗窓は持たない。
- **ui_server**: req を順に処理。csv_ld req でファイルダイアログ、progress req で進捗窓を表示。進捗窓は progress_path の pkl を 200ms 間隔でポーリング。
- **_watch_ready**: 別スレッド。READY_UI を検知したら notify_ui_ready() で砂時計解除。

---

## 参照コード

| 処理 | ファイル |
|------|----------|
| エントリ・req 投入・待ち | svc/svc_csv_ld.py `load_csv`, `_watch_result`, `_watch_ready` |
| ファイル選択 UI | ui_qt/ui_csv_ld.py `create_dialog` |
| 進捗 UI・DONE で 2.5 秒後に閉じる | ui_qt/ui_common.py `ProgressDialog`, `_tick`, `_close_after_done` |
| 進捗 pkl の書き込み | svc/svc_csv_ld.py `_execute_jit_import`, `_progress_write` |

---

## 動作確認

### 事前（Python のみ）

- プロジェクトルートで: `python -c "from svc import svc_csv_ld; svc_csv_ld.load_csv(None, ''); print('OK')"` → 即終了・例外なしで OK。

### 本番（Excel から）

1. **Excel でアドインを有効にし、ブックを開く。**
2. **対象シート（GUID が設定されたシート）をアクティブにする。**
3. **リボンから「CSV読込」を実行。**
4. **確認項目**
   - ファイル選択ダイアログが表示される（タイトルは config/ui_csv_ld.json の MAIN.TITLE）。
   - 1 件 CSV を選んで「開く」→ 進捗画面（「CSV読込 進捗」）が表示され、完了まで更新される。
   - 進捗は「ファイル解析中」→「Excelへ書き込み中」→「列幅調整中」→「完了」の順で変わる。
   - 完了後、進捗画面が閉じ、ステータスバーに「CSV読込終了｜…」が出る。
   - キャンセル時: ファイル選択で「キャンセル」→ 何も書き込まれず終了。
5. **ログ**: 不具合時は `core_log` の設定先（例: `%TEMP%\csv_tool\hc_csv.log`）の時系列を確認。
