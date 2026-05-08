# 起動時間改善案（Excel 起動 ～ csv_ld ファイル読込まで）

## ログから読み取ったタイムライン（概算）

| 時刻 | イベント | 経過/遅延 |
|------|----------|------------|
| 10:15:05.43 | Excel セッション開始・アドイン読込 | - |
| 10:15:05.54 | アドイン走査開始 | - |
| 10:15:06.91 | アドイン走査完了 | **約 1.4 秒** |
| 10:15:06.95 | warmup 予約（WaitAndInit → 1秒後に RunInitEvents） | - |
| 10:15:07.07 | RunPython 解釈・ExecuteWindows | **約 1 秒**（OnTime 遅延） |
| 10:15:08.87 | ensure_svc_server / ensure_ui_server コマンド送信 | - |
| 10:15:11.06 | 初回 Python プロセスで svc_host 読込 (pid=22864) | **約 2.2 秒**（Python 起動） |
| 10:15:11.52 | SVC_SERVER  spawn | - |
| 10:15:13.72 | SVC_SERVER BOOT (pid=22784) | **約 2.2 秒**（サーバ起動） |
| 10:15:13.84 | QT_UI_SERVER spawn | - |
| 10:15:15.86 | UI_SERVER BOOT | **約 2 秒**（UI サーバ起動） |
| （ユーザが「CSV読込」クリック） | | |
| 10:15:38.40 | 2回目 RunPython: import hc_main; invoke(load_csv) | - |
| 10:15:43.25 | hc_main 側で svc_host MODULE_LOAD (pid=10108) | **約 4.8 秒**（2回目 Python 起動） |
| 10:15:43.83 | ensure_svc_server 完了・依頼書出し | - |
| 10:15:43.99 | SVC_SERVER exec start（依頼取り込み） | - |
| 10:15:47.17 | load_csv START（svc_csv_ld） | **約 3.2 秒**（依頼～ハンドラ開始） |
| 10:15:47.29 | READY_UI（ファイル選択ダイアログ準備） | - |
| 10:15:48.68 | UI_READY done（ユーザ操作待ち約 1.35 秒） | - |
| 10:15:59.83 | ファイル選択完了・_do_load_csv 開始 | **約 11 秒**（ユーザ選択時間含む） |
| 10:15:59.86 | progress phase1 書出し → **2.5 秒 sleep** | **2.5 秒**（固定待ち） |
| 10:16:02.42 | sleep 終了・読込ループ開始 | - |

（注）2026-04-06 以降の実装では、アドイン起動時の **xlwings `RunPython` は Workbook_Open 内 1 回**（`excel_startup_workbook_open_full`）に集約され、約 1 秒後の `InitPythonServer` は成功時 **2 回目の RunPython を行わない**ため、上表の「2 回目の起動経路」は旧ログ向けの記述です。

---

## 実装済み（2026-04-06）：起動時 RunPython の一本化と mutex 待ち

- **VBA**: `Workbook_Open` で `excel_startup_workbook_open_full(Application.hwnd)` を 1 回だけ実行（svc / ui / bridge / `register_book`）。成功時（`Err.Number = 0`）に `Main.MarkWorkbookOpenFullPythonDone` を立て、遅延の `InitPythonServer` では RunPython をスキップ。
- **手動初期化**: `Manual_Init` は `ResetWorkbookOpenFullPythonDone` のうえで `InitEvents` し、従来どおり `excel_startup_after_excel_idle` を通す。
- **Python**（`svc_host`）: `ensure_*` 後の mutex 待ちを **約 5 秒**に延長し、ループ終了後 **約 0.4 秒の猶予**でもう一度だけ判定。間に合った場合は INFO（`mutex observed after grace wait`）、それでも無ければ従来どおり WARNING。

---

## 改善案一覧

### 1. 【高効果】リボンクリック時の Python 再起動を避ける（4.8 秒短縮目標）

**現象**  
「CSV読込」クリックのたびに `RunPython("import hc_main; hc_main.invoke(action='load_csv', ...)")` で **毎回新しい Python プロセス** が立ち上がり、約 4.8 秒かかっている。

**案 A: ブリッジ常駐プロセス（設計変更）**  
- ウォームアップ時に「ブリッジ用」Python プロセスを 1 本だけ起動し、ファイルまたは名前付きパイプで「action + 引数」を受け取る。
- VBA は RunPython の代わりに、そのプロセスへ依頼を送る（例: 特定フォルダへ `req_*.pkl` を書き、ブリッジがポーリングして hc_main と同等の依頼を svc_server に送る）。
- クリックのたびの Python 起動がなくなり、4.8 秒が 0.1 秒オーダーになる可能性がある。  
- **難点**: VBA から「別プロセスに依頼する」手段が必要（ファイル監視＋ブリッジ側で svc_server へ依頼する形なら、既存の req/res と似た構成で実装可能）。

**案 B: hc_main の遅延インポート（実装コスト低）**  
- `hc_main.py` のトップレベルで `import xlwings as xw` と `from svc.svc_host import ensure_svc_server` を必須にしているため、クリックのたびに重いモジュールが読まれる。
- `load_csv` の先頭で `import xlwings as xw` と `ensure_svc_server` を遅延 import し、`_ensure_book` や `_call_svc_server` で必要時のみ読む。
- 加えて、`core.core_log` も遅延化できるなら、起動直後の import チェーンを軽くする。  
- **効果**: プロセス起動は避けられないが、モジュール読込時間を数百 ms 程度短縮できる可能性がある。

---

### 2. 【高効果】svc_server の依頼取り込み～ハンドラ開始の 3.2 秒短縮

**現象**  
`exec start` (43.99) から `load_csv START` (47.17) まで約 3.2 秒。  
`_process_one` 内の `_load_handler("csv_ld")`（ウォームアップ済みなら軽い）と **`_attach_book`（xlwings で Excel ブック取得）** が主因と推測される。

**案 A: _attach_book の COM 呼び出し最適化**  
- `xw.apps` の列挙や `hwnd` 取得が COM 経由で重い可能性がある。
- キャッシュ: 同一 `excel_hwnd` に対して短時間なら前回の `book` を返す（プロセス内で TTL 付きキャッシュ）。
- 並列化: 依頼の書出しと並行して、別スレッドで「次の想定 hwnd の book 取得」を先行させることは、Excel COM のスレッド制約に注意しつつ検討。

**案 B: ウォームアップで csv_ld を確実にキャッシュ**  
- `config/svc_warmup.json` または `HC_SVC_WARMUP_ACTIONS` に `csv_ld` を必ず含め、初回クリック時の `_load_handler("csv_ld")` の import コストをゼロにする（既に実施済みなら、_attach_book 重点でよい）。

**案 C: アイドルポール間隔の短縮**  
- `HC_SVC_IDLE_POLL_SEC` の既定 0.1 秒を 0.05 秒にすると、依頼書出しから取り込みまでの最大遅延が半減する。  
- **トレードオフ**: CPU 使用率がわずかに増える。

---

### 3. 【中効果】進捗ウィンドウ表示待ちの 2.5 秒短縮

**現象**  
`svc_csv_ld.py` の `PROGRESS_WINDOW_STARTUP_WAIT_SEC = 2.5` により、phase1 書出し後に毎回 2.5 秒 sleep している。

**改善案**  
- 定数を 1.0～1.5 秒に下げ、環境変数で上書き可能にする（例: `HC_PROGRESS_WINDOW_STARTUP_WAIT_SEC`）。
- 進捗ウィンドウ側で「ポーリング開始」や「初回 RUN 検知」をログに出し、実際に何秒で準備できるか計測してから既定値を決めると安全。
- **効果**: 1～1.5 秒短縮可能。

---

### 4. 【中効果】HC_RETURN_EARLY 時の 1 秒 sleep 短縮

**現象**  
`hc_main._call_svc_server` で、依頼を書出したあと `time.sleep(1.0)` してから return している（サーバが req を拾うまでの余裕）。

**改善案**  
- サーバのポール間隔が 0.1 秒（または 0.05 秒）なら、0.3～0.5 秒でも十分なことが多い。
- 環境変数 `HC_RETURN_EARLY_WAIT_SEC` を導入し、既定を 0.5 秒にする。  
- **効果**: 0.5 秒短縮。

---

### 5. 【中効果】アドイン起動時のアドイン走査（約 1.4 秒）の遅延／軽量化

**現象**  
`Workbook_Open` で `LogInstalledAddins` が全アドインを走査し、約 1.4 秒かかっている。

**案 A: 走査の遅延実行**  
- 走査を `Workbook_Open` では行わず、`Application.OnTime` で 2～3 秒後（または「最初のリボンクリック時」）に実行する。
- 起動直後の体感は軽くなる。ログの時系列は変わるだけなので、デバッグ時は「遅延ログ」であることを把握すればよい。

**案 B: 走査内容の縮小**  
- 有効なアドインの数だけログに書き、名前・パスは「CSV Tool 自身」と「問題調査用の 1～2 個」だけにする。
- ループとログ書出しが減り、数百 ms 単位で短縮できる可能性がある。

---

### 6. 【中効果】WaitAndInit の 1 秒遅延短縮

**現象**  
`Main.WaitAndInit` が `Application.OnTime(Now + TimeSerial(0,0,1), "Main.RunInitEvents")` で 1 秒後に実行している。

**改善案**  
- 0.5 秒に短縮する（`TimeSerial(0,0,1)` → `TimeSerial(0,0,0)+0.5/24/3600` など）。
- または設定可能にして、環境が重い場合のみ 1 秒に戻す。  
- **効果**: 0.5 秒短縮。Excel が完全に Ready になる前に実行すると不安定になる可能性はあるため、短くしすぎないよう注意。

---

### 7. 【中効果】ensure_svc_server / ensure_ui_server の mutex 待機間隔

**現状**  
- `ensure_svc_server`: mutex 確認まで `time.sleep(0.05)`（既に 0.02 にしている箇所は ui_server 側）。
- `ensure_ui_server`: 0.02 秒でポーリング。

**改善案**  
- `ensure_svc_server` の待機も 0.02 秒に揃える（`svc_host.py` の 459 行付近）。  
- **効果**: サーバ起動直後の「初回依頼までの待ち」が最大で数十 ms 短縮される。

---

### 8. 【低～中効果】初回ウォームアップの 2 プロセス起動（合計約 4 秒）

**現象**  
ensure_svc_server で SVC_SERVER を spawn し、その直後に ensure_ui_server で UI_SERVER を spawn。それぞれ約 2 秒ずつかかっている。

**案 A: 並列起動**  
- 現状は同一プロセス（pid=22864）内で `ensure_svc_server()` 完了後に `ensure_ui_server()` を呼んでいる。  
- 先に両方 spawn してから、mutex を 0.02 秒間隔で両方見にいくようにする。  
- 実装: まず `spawn_svc_server()` と `spawn_ui_server()` を連続で呼び、その後 `is_svc_server_running()` と `is_ui_server_running()` をループで待つ。  
- **効果**: 2+2=4 秒が、並列のため max(2,2)≈2 秒程度になる可能性がある。

**案 B: モジュール読込の軽量化**  
- `svc_server.py` のトップレベル import（`ctypes`, `multiprocessing` 等）や、`_run_warmup` で読むモジュールを必要最小限にし、起動時の import 時間を削る。  
- **効果**: 数百 ms 単位の短縮が期待できる。

---

### 9. 【低効果】SVC_SERVER のウォームアップ対象の見直し

**現状**  
ウォームアップで csv_ld, csv_mg, csv_sv, csv_sp などを一括でハンドラキャッシュしている。

**改善案**  
- 実際に多く使う機能だけに絞る（例: 最初は `csv_ld` のみ）。  
- 初回クリックが csv_ld の場合、既にキャッシュ済みで、ウォームアップ時間の短縮と初回クリック時の _load_handler の短縮の両方に効く。  
- 他のアクションは初回実行時に遅延ロードされる。

---

## 優先度と実装しやすさ（目安）

| 優先度 | 改善案 | 期待効果 | 実装コスト |
|--------|--------|----------|------------|
| 高 | 進捗ウィンドウ 2.5 秒 → 1.0～1.5 秒（案 3） | 1～1.5 秒 | 低 |
| 高 | HC_RETURN_EARLY 後の sleep 1→0.5 秒（案 4） | 0.5 秒 | 低 |
| 高 | _attach_book のキャッシュ／最適化（案 2） | 1～2 秒 | 中 |
| 高 | ブリッジ常駐でリボンクリック時の Python 起動を廃止（案 1-A） | 約 4.8 秒 | 高 |
| 中 | アドイン走査の遅延／縮小（案 5） | 0.5～1.4 秒 | 低 |
| 中 | WaitAndInit 1 秒→0.5 秒（案 6） | 0.5 秒 | 低 |
| 中 | ensure_svc_server と ensure_ui_server の並列起動（案 8-A） | 約 2 秒 | 中 |
| 中 | アイドルポール 0.1→0.05 秒（案 2-C） | 最大 0.1 秒 | 低 |
| 低 | hc_main 遅延 import（案 1-B） | 数百 ms | 低 |
| 低 | ウォームアップ対象の見直し（案 9） | 数百 ms | 低 |

---

## まとめ

- **即効性が高い**: 案 3（PROGRESS_WINDOW_STARTUP_WAIT_SEC）、案 4（HC_RETURN_EARLY_WAIT_SEC）、案 6（WaitAndInit 短縮）、案 5（アドイン走査の遅延／縮小）。いずれも設定変更または小さいコード変更で 0.5～2 秒程度の短縮が見込める。
- **体感が大きい**: 案 1-A（ブリッジ常駐）で「クリック～ダイアログ」の約 4.8 秒を削減。案 2（_attach_book とポール間隔）で「依頼～処理開始」の 3.2 秒を短縮。
- **設計変更を伴う**: 案 1-A は VBA と Python の連携方法の見直しが必要。まずは上記の低コスト案から適用し、計測しながら段階的に導入するのがよい。
