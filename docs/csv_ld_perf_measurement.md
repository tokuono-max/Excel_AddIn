# CSV 読込（`load_csv`）の時間計測ポイント

`hc_main` 経路と `bridge_requests` 経路の比較、および「リボン〜最初の UI」「ファイル確定〜出力完了」の切り分けに使う。

## 計測の開始手順（既存機能）

1. **環境変数**  
   - **`HC_LOG_PERF=1`** を **ユーザー環境変数**に設定する（Excel は**すべて終了した状態**で設定し、**新規に起動した Excel** から有効になる）。  
   - 詳細・解除例は **`docs/environment_variables.md`** の §6 参照。

2. **（任意）診断ログ**  
   - **`HC_LOG_DIAG=1`** にすると `[CSV_LD_TRACE]` などが **`hc_csv_diag.log`** に出る。運用ログだけで足りなければ併用。

3. **操作**  
   - ブックを開き、**リボン「CSV 読込」**を実行（冷起動直後と、数回操作後の **ウォーム時**で別々に取ると比較しやすい）。

4. **ログの場所**  
   - **計測（VBA・`hc_main`）**: `%TEMP%\csv_tool\hc_csv_perf.log`  
   - **運用（`[CSV_LD]`・`[CSV_LD_UI]`・`[MAIN]`）**: `%TEMP%\csv_tool\hc_csv.log`  
   - いずれも **LIFO** のため、**時刻列**で並べ替えて読むか、検索でキーワードを辿る。

5. **読むときのキーワード**  
   - 区間 A/B・`pick_to_done_ms` の対応はこの文書の **「ログ行の対応表」「推奨の読み方」** を参照。

6. **VBA 計測の注意（既存実装）**  
   - `before_runpython_safe` には **`BeginWaitForRibbon`（WaitForm）** の時間が含まれる。  
   - **`load_csv`（リボン）**は **`before_bridge_submit` / `after_bridge_submit`** のみ（`RunPython` なし）。**必ず `RibbonPerfEnd` が呼ばれる**。  
   - `GetTickCount64` 由来のため **分解能はおおむね 15ms オーダー**のことが多い。

### PowerShell（ユーザー環境に 1 回設定する例）

```powershell
[Environment]::SetEnvironmentVariable("HC_LOG_PERF", "1", "User")
```

解除例:

```powershell
[Environment]::SetEnvironmentVariable("HC_LOG_PERF", $null, "User")
```

（設定後は Excel を一度終了し、新しいセッションで起動する。）

## 前提

- **運用ログ**: `%TEMP%\csv_tool\hc_csv.log`（常時）
- **計測ログ**: `%TEMP%\csv_tool\hc_csv_perf.log`（**`HC_LOG_PERF=1`** のときのみ）
- ログは **LIFO（新しい行が先頭付近）** のため、時系列で読むときは時刻列に注意。

## 区間の定義

| 区間 | 意味 | 含まないもの（切り分け） |
|------|------|---------------------------|
| **A** | リボン押下〜**最初の UI（ネイティブファイル選択が開く直前）** | ユーザーがファイルを探して選ぶ時間 |
| **B** | **ファイル確定（OK）直後**〜**`load_csv` 一連フロー終了**（Excel 出力・進捗終了など含む） | 区間 A およびファイル選択中の待ち |
| **全体** | `svc_csv_ld.load_csv` 開始〜`load_csv_flow_done` | 区間 A+B+ユーザー選択待ちが **同一 `elapsed_since_load_ms`** に含まれる |

## ログ行の対応表

### VBA（`HC_LOG_PERF=1`）

| キーワード | 区間 |
|------------|------|
| `RibbonInvoke` / `phase=click_enter` / `elapsed_since_click_ms=0` | **全体比較の始点**（リボン） |
| `phase=before_runpython_safe` | RunPython 直前（`load_csv` 以外のリボン） |
| `phase=before_bridge_submit` | `load_csv` 用: `bridge_requests` へ JSON 書き込み直前 |
| `phase=after_bridge_submit` | `load_csv` 用: JSON 書き込み直後（`RibbonPerfEnd` の直前） |
| `phase=before_xlwings_runpython` | xlwings 呼び出し直前 |
| `phase=after_xlwings_runpython` | 短寿命 Python から戻った直後（`hc_main` 経路） |

**`load_csv`（リボン）**では **`after_xlwings_runpython` は出ない**。区間の補助に **`[MAIN] forwarded`**（`hc_csv.log`）と上記 **`after_bridge_submit`** を使う。

### `hc_main`（`hc_csv_perf.log`）

| キーワード | 区間 |
|------------|------|
| `invoke phase=enter` / `action='load_csv'` | `hc_main` 処理開始 |
| `call_svc phase=enter` … `after_return_early` 等 | `svc_server` へ依頼を書くまで |

### 常駐 `hc_main`（`hc_csv.log`）

| キーワード | 区間 |
|------------|------|
| `[MAIN] forwarded public_action=load_csv svc_action=csv_ld ...` | JSON を受理し **`svc_requests` へ pkl 書き込み後**（`bridge_requests` 経路の目印） |

### `svc_server` / `svc_csv_ld`（`hc_csv.log`）

| キーワード | 区間 |
|------------|------|
| `[CSV_LD] 開始` / `[CSV_LD_TRACE] phase=enter` | **`load_csv` エントリ**（`elapsed_since_load_ms` の 0 に近い基準） |
| `[CSV_LD] phase=after_ensure_ui_server` | UI サーバ確保後 |
| `[CSV_LD] ui_ipc ok req=...` | `ui_server` へ `req` 投入済み |
| `[CSV_LD] phase=result_ok` | ファイル選択結果を svc が取得（**まだ `_do_load_csv` 前**） |
| `[CSV_LD] phase=pick_confirmed` | **区間 B 始点**（ファイル確定直前の内部アンカーと同時刻付近） |
| `[CSV_LD] phase=load_csv_flow_done` **`pick_to_done_ms=`** | **区間 B の長さ（ms）**。キャンセル時は **`pick_to_done_ms=-1`** |

### `ui_qt.ui_csv_ld`（`hc_csv.log`）

| キーワード | 区間 |
|------------|------|
| `[CSV_LD_UI] phase=native_file_dialog_open` | **区間 A の終点目安**（ネイティブ「開く」ダイアログを出す直前） |

### `ui_server`（`hc_csv.log` / 診断時 `hc_csv_diag.log`）

| キーワード | 区間 |
|------------|------|
| `[UI_CSV_LD] create_dialog ok ... elapsed_ms=...` | `ui_server` 内でダイアログ生成まで（ファイル選択 **exec** 前後は実装依存のため、区間 A は **`[CSV_LD_UI] native_file_dialog_open`** を推奨） |

## 推奨の読み方

1. **区間 A（リボン〜最初の UI）**  
   - `hc_csv_perf.log` の **`RibbonInvoke` `elapsed_since_click_ms`** と、同一操作の **`[CSV_LD_UI] phase=native_file_dialog_open`** の **ログ時刻差**（または VBA の ms と Python 時刻の突合）。  
   - `bridge_requests` 経路の比較時は **`[MAIN] forwarded`** 時刻をリボン側マークと揃える。

2. **区間 B（ファイル確定〜出力完了）**  
   - **`[CSV_LD] phase=load_csv_flow_done` の `pick_to_done_ms`** をそのまま使う（**ユーザーがダイアログを開いている時間は含まない**）。

3. **全体（従来の「遅い」体感に近いもの）**  
   - **`[CSV_LD] phase=load_csv_flow_done` の `elapsed_since_load_ms`**（`load_csv` 開始からの経過。ファイル選択待ちを含む）。

## RunPython 経路の記録ベースライン（比較用・2026-04-10）

リボンを **`bridge_runner` 経路に切り替える前**にユーザー環境で取得した値のメモ。**同一マシン・同程度のデータ量**で再計測し、差分を見る。

| 項目 | 目安（当時のログより） |
|------|-------------------------|
| `hc_csv_perf`: `after_xlwings_runpython − before_xlwings_runpython` | 約 **6〜13 秒**（初回は大きめ） |
| `hc_csv.log`: `pick_to_done_ms`（`load_csv_flow_done`） | 例 **13660 / 19859 / 22220 ms**（データ量・環境依存） |

切り替え後は **`after_bridge_submit − before_bridge_submit`**（VBA）と **`[MAIN] forwarded`** 時刻を、上記 **`before/after_xlwings_runpython`** 相当区間と突き合わせる。

## 関連ドキュメント

- 画面遷移の時系列: **`docs/svc_csv_ld_flow.md`**
- VBA 計測・環境変数: **`docs/environment_variables.md`**（`HC_LOG_PERF` 等）

## 変更履歴

| 日付 | 内容 |
|------|------|
| 2026-04-10 | 「計測の開始手順」: `HC_LOG_PERF`・ログパス・VBA 注意・PowerShell 例。 |
| 2026-04-10 | 初版。`pick_to_done_ms` / `pick_confirmed` / `[CSV_LD_UI] native_file_dialog_open` / `[BRIDGE] forwarded` を追記。 |
| 2026-04-11 | フェーズ E: 運用ログのキーワードを **`[MAIN] forwarded`** に更新（旧 `[BRIDGE] forwarded` は廃止）。 |
| 2026-04-10 | リボン `load_csv` を bridge 経路にした前提で VBA マーク・ベースライン表を追記。 |
