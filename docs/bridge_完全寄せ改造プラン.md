# bridge 完全寄せ改造プラン

**作成日**: 2026-04-10  
**目的**: リボン機能を **`bridge_requests` → 常駐プロセス → `svc_server`** に統一し、**短寿命 `RunPython` + `hc_main.invoke` を廃止**する。完了後、**現行 `hc_main.py` を削除**し、**現 `svc/bridge_runner.py` をルートの `hc_main.py` にリネーム・配置**して「トップの常駐入口」を一本化する。

**関連**: `CSV_Tool_xml.txt`（リボン `tag`）、`svc/svc_server.py`（`_ACTION_MAP`）、`VBA/Main.bas`、`svc/svc_host.py`（起動・mutex）、`docs/environment_variables.md`。

### 文字コード（運用ルール）

- **VBA モジュール（`.bas` 等）**: リポジトリ上の正本は **Shift-JIS（CP932）** で保存する（日本語コメント・文字列の破損防止）。
- **`bridge_requests` の JSON（VBA ↔ 常駐 `hc_main`）**: **UTF-8** で書き出す（`ADODB.Stream`、`Charset=utf-8`）。Python 側は **UTF-8 優先**し、旧ファイル向けに **cp932 フォールバック**で読む。

---

## 1. ゴール状態（As-Is → To-Be）

| 項目 | 現状 | ゴール |
|------|------|--------|
| リボン（多くの機能） | `RunPython` → `hc_main.invoke` → `svc` または in-process `svc_*` | **JSON → 常駐ブリッジ → `svc_req_*.pkl` → `svc_server`** |
| Excel 起動時の常駐化 | `VBA` → `RunPython` → **`svc_host.excel_startup_workbook_open_full`**（**`hc_main` は経由しない**） | **変更なし**（本プランの主戦場はリボン経路） |
| トップレベル常駐スクリプト | `svc/bridge_runner.py` | **プロジェクトルート `hc_main.py`**（役割は現ブリッジと同一・**`invoke` は持たない**） |
| 旧 `hc_main.py`（903 行級） | xlwings 受付・`_call_svc_server`・各 `_invoke_impl_*` | **削除**（必要ロジックは `core` / `svc_host` / `svc_server` へ移管） |
| VBA の `hc_main` 文字列 | `RunPythonSafe`・終了時 `clear_registry` 等 | **ゼロ**（ユーザー決定 Q3） |
| ログ・mutex・環境変数の接頭辞 | `BRIDGE` / `HC_BRIDGE_*` | **`MAIN` / `HC_MAIN_*` に揃える**（ユーザー決定 Q5。旧名は互換読み取りを推奨） |

### 1.1 新ルート `hc_main.py` の役割（設計方針）

- **常駐ブリッジ専用**（現 `bridge_runner` の `main` ループと同等）。**レジストリ掃除・タイムアウト表・`invoke` は載せない**。
- 旧 `hc_main` の処理は **「ブリッジへの統合」ではなく**「**`invoke` 層の削除** + **残りは適切な層へ移す**」と捉える（ブリッジは配送専用のまま）。

---

## 2. ユーザー決定事項の反映

| ID | 内容 |
|----|------|
| Q1 | 新ルート `hc_main.py` は **常駐ブリッジ専用**（推奨 A）。 |
| Q2 | 旧 `hc_main` の invoke 以外は **`core` / `core_env` / `svc_server` / 専用小モジュール**へ **保守的に分散**（ルート `hc_main.py` には載せない）。 |
| Q3 | 移行完了後、**VBA から `hc_main` 文字列は完全に排除**（起動は `svc_host` のみ、終了処理は別モジュール直呼び）。 |
| Q4 | 子プロセス起動は **実装都合でよい**（推奨: 現状同様 `pythonw` + **ルート `hc_main.py` の絶対パス**で最小差分）。 |
| Q5 | **`HC_MAIN_*` / `[MAIN]`** 等へリネーム。**`HC_BRIDGE_*` は当面フォールバックで読む**ことを推奨（運用ブレイク防止）。 |

---

## 3. フェーズ別ロードマップ（推奨順）

### フェーズ A — リボンを bridge に全面移行（旧 `hc_main` はまだ残す）

1. **VBA**（フェーズ A 着手分は反映済み）  
   - `SubmitSvcRequestViaBridge(action, …)` で JSON の `action` をリボン `tag` と一致させて出力。`SubmitLoadCsvViaBridge` は `load_csv` 向けラッパ。  
   - `RibbonInvokeFromControl` は **全 `act` を bridge 経由**（`RunPythonSafe` は非リボン用に残置）。

2. **`bridge_runner`（当面は `svc` 配下のまま）**  
   - `_ACTION_MAP` を **全リボン action → `svc_server` の action 名**で拡張（下表参照）。

3. **`hc_main` のリファクタ（橋渡し）**  
   - `check_duplicates` ～ `convert_date_ymd_hm` を **`_call_svc_server` 経由**に寄せ、**`svc_server` 経路と挙動を一致**させる（bridge 単体テストと hc_main 経路の二重保守を避ける）。

4. **WaitForm / `NotifyUiReady`**  
   - **action ごと**に「誰が `HC_WaitForm.NotifyUiReady` を呼ぶか」を棚卸し。  
   - `csv_ld` / `csv_sv` / `csv_mg` / `csv_sp` は **READY_UI 等の既存経路**。  
   - `core.ribbon_public_to_svc.RIBBON_INVOKE_FINALLY_NOTIFY_WAITFORM`（旧 hc_main 内集合と同一）相当は **`svc` 完了時または `core_cursor.notify_wait_form_ready`** を必ず通すよう確認。

5. **`run_data_agg`**  
   - JSON に **`payload`（または既存 kwargs と同形）** を載せられるよう **bridge JSON スキーマを拡張**し、`svc_server` の pickle に引き継ぐ。

### フェーズ B — VBA から `hc_main` 文字列の除去

1. **リボンから `RunPythonSafe` を呼ばない**（`SubmitSvcRequestViaBridge` のみ）。`RunPythonSafe` 自体は非リボン用手動・マクロ用に **残置可**（`core.excel_session.invoke_action` 経由）。  
2. **終了・クリア**  
   - `Main.bas` の `import hc_main; hc_main.clear_registry()` を **`core` 配下の明示関数**（例: `core.excel_session.clear_internal_registry`）の **`RunPython` 1 行**へ変更。  
3. **その他** `import hc_main` を **VBA / Python / テスト**から grep し、すべて代替へ。

#### フェーズ B — 完了記録（2026-04-11 反映）

| 項目 | 状態 |
|------|------|
| リボン | `RibbonInvokeFromControl` → `SubmitSvcRequestViaBridge`（JSON UTF-8）のみ。`RunPythonSafe` は呼ばない。 |
| `RunPythonSafe` | `from core.excel_session import invoke_action; invoke_action(...)`（VBA 文字列に `import hc_main` なし）。 |
| 終了・レジストリ掃除 | `from core.excel_session import clear_internal_registry; clear_internal_registry()`。 |
| `import hc_main` の残存 | **原則なし**（`excel_session` → **`core.ribbon_invoke`**）。`tools/patch_main_bas_bridge.py`・`VBA/Old/*` に旧文字列のみ。 |
| `xlwings.bas` ブートストラップ | 引き続き **`hc_main.py` の存在**でプロジェクトルートを決定。フェーズ C/D でルート常駐スクリプト名が変わる際に **同ブロックを更新**（モジュール内コメント参照）。 |
| `svc.svc_host` バージョン | ソースの **`__version__`** とログの版号を照合する。ずれるときは Excel を完全終了して **常駐プロセス（`hc_main` / svc_server / Qt UI）を再起動**し、読み込んだモジュールを更新する。 |

##### スモーク回帰（手動・短文）

1. アドイン読込 → `startup_full` が 1 回、`[MAIN]`（常駐 `hc_main`）起動・`[SVC_SERVER]` 起動がログに出る。  
2. リボン「CSV 読込」→ UI〜完了まで（bridge → `csv_ld`）。  
3. リボンから別 action 1 本（例: 重複チェック）。  
4. （任意）VBA イミディエイト: `Main.RunPythonSafe "check_duplicates", ExcelUtil.GetSheetIdSafe(ActiveSheet)` → `%TEMP%\csv_tool\hc_csv_diag.log` に `invoke_action` が出る。  
5. アドイン終了 → `clear_internal_registry`、bridge / svc の shutdown ログ。  

### フェーズ C — 短寿命 invoke の `core` 集約・旧 `hc_main.py` 削除

#### 実施済み（2026-04-11）— 司令塔の移設

- **`hc_invoke.py` を削除**。短寿命の **`invoke` / `register_book` / `clear_registry`** は **`core/ribbon_invoke.py`**（**1.12.0**）に集約。  
- **`core/excel_session.py`** は **`core.ribbon_invoke` を直接 import**（ルートモジュール `hc_invoke` は廃止）。  
- **リボン経路**: 全 action は既に **bridge（`hc_main`）で `svc_server` に届く**想定。短寿命経路は **VBA `RunPythonSafe` / `invoke_action`** および **`run_data_agg` の batch 等**に限定（プラン当時の棚卸しどおり）。  
- **診断**: `hc_csv_tool.diag.ribbon_invoke`（MODULE_LOAD / DEBUG）。旧名 `diag.hc_main` は廃止。

#### 準備: `hc_main` 依存棚卸し（参考・2026-04-11 時点）

| ファイル | 区分 | 備考 |
|----------|------|------|
| `core/excel_session.py` | **VBA 向け薄い入口** | `core.ribbon_invoke` を re-export。 |
| `hc_main.py`（ルート） | **常駐ブリッジ** | `bridge_runner` 相当。invoke は持たない。 |
| `core/ribbon_invoke.py` | **短寿命 invoke** | 旧 `hc_invoke.py` 相当。 |
| `tools/patch_main_bas_bridge.py` | **文字列リテラル** | 旧 VBA 置換用。 |
| `core/ribbon_public_to_svc.py` | **契約** | ブリッジ・`ribbon_invoke` が参照。 |

#### コーディング規約（移行後）

- **新規 Python コード**で **`import hc_main`（invoke 目的）とルート `hc_invoke` を直接書かない**。短寿命からの入口は **`core.excel_session`**。  
- **リボン相当**は原則 **bridge → `svc_server`**。`excel_session.invoke_action` は **RunPythonSafe 等**に限定。

---

（旧チェックリスト・完了）**ルート `hc_main.py`（ブリッジ）は維持**。`xlwings.bas` はルート `hc_main.py` でパス解決。

### フェーズ D — `bridge_runner` → ルート `hc_main.py`

#### 実施済み（2026-04-11・一部）

- **invoke 司令塔**を **`core/ribbon_invoke.py`** に集約（`excel_session` が import）。旧ルート `hc_invoke.py` は削除済み（フェーズ C）。  
- **常駐ブリッジ**を **ルート `hc_main.py`**（旧 `bridge_runner` 本体、**0.2.0**）。bootstrap は「本ファイルのディレクトリ＝プロジェクトルート」。  
- **`svc_host._resolve_bridge_path`** は **ルート `hc_main.py` 優先**、無い場合のみ **`svc/bridge_runner.py`**。`spawn_bridge` の argv 判定に **`hc_main.py`** を追加。  
- **`svc/bridge_runner.py`** は **ルート `hc_main.py` を `runpy` 実行する薄い互換ラッパ**。  

**フェーズ E 実施済み（2026-04-11）**: ログラベル **`[MAIN]`**、`HC_MAIN_*`（`HC_BRIDGE_*` フォールバック）、ブートログ **`hc_main_boot_*.log`**。  
**Mutex 移行（2026-04-11 追補）**: 新 `hc_main` は **`HC_MAIN_RUNNER` と `HC_BRIDGE_RUNNER` の両方**を Create（旧 svc_host が旧名のみ参照するため）。`svc_host.is_main_runner_running` / **`is_bridge_running`** は **いずれかの Mutex が存在すれば真**。

1. **`svc/bridge_runner.py` の内容をルート `hc_main.py` に移動**（モジュール docstring・`__name__` ログを更新）。  
2. **`svc_host`**  
   - `_resolve_bridge_path()` を **プロジェクトルートの `hc_main.py`** を指すよう変更。  
   - `spawn_bridge` 内の **`sys.argv[0]` 判定**（`bridge_runner.py`）を **`hc_main.py`** に合わせる。  
   - **mutex 名**: **`HC_MAIN_RUNNER` ＋ 旧 `HC_BRIDGE_RUNNER` 二重検知**を実装済み（`is_main_runner_running` / `is_bridge_running`）。  
3. **`core/ipc_cleanup`**  
   - `run_bridge_startup_sweeps` の import 元を **ルート `hc_main` から呼ぶ**か、**`core.ipc_cleanup` に集約**してブリッジ依存をなくす。  
4. **旧 `svc/bridge_runner.py`** は削除または **薄い再エクスポート**（非推奨ワンライナー）で移行期間のみ残すかを決定。

### フェーズ E — ログ・環境変数・ドキュメントの `MAIN` 化

#### 実施済み（2026-04-11）

1. **`core_env.LOG_MAIN_PREFIX`**（`[MAIN]`）と、`hc_main` / `svc_host` の spawn・ensure ログへの適用。  
2. **`HC_MAIN_POLL_SEC` / `HC_MAIN_MIN_FILE_AGE_SEC` / `HC_MAIN_BAD_FILE_MAX_POLLS`** — 実装は **`core_env.hc_main_*()`**。未設定時 **`HC_BRIDGE_*` フォールバック**。  
3. **ブートログ** `logs/hc_main_boot_*.log`（旧 `bridge_boot_*.log` は出力しない。運用で grep している場合はリリースノートに記載）。  
4. **`spawn_bridge`**: ルート配置の `hc_main.py` について **`project_root = bridge_py.parent`** を正す（venv 判定・`cwd`）。  
5. ドキュメント: **`environment_variables.md`**, **`IPC_TEMP_CLEANUP.md`**, **`csv_ld_perf_measurement.md`**, **`ipc_file_仕様書.md`**, 本プラン。

---

## 4. 機能別: bridge に載せられるか・追加作業

**前提**: `svc_server._ACTION_MAP` に既に存在する action は、**pickle の形（`excel_hwnd`, `book_fullname`, `book_name`, `sheet_id`）が揃えば** bridge からそのまま届け可能。

**リボン `tag` の出典**: `CSV_Tool_xml.txt`  
**`svc_server` action 出典**: `svc/svc_server.py` の `_ACTION_MAP`

| リボン `tag` | `svc_server` action | bridge 載せ可否 | 追加作業・注意 |
|--------------|---------------------|-----------------|----------------|
| `load_csv` | `csv_ld` | **済** | WaitForm は READY_UI 系。計測キーワード更新（`[MAIN] forwarded`）。 |
| `save_csv` | `csv_sv` | **可** | UI・保存ダイアログまわりの `NotifyUiReady` 確認。 |
| `merge_csv` | `csv_mg` | **可** | 同上（進捗・ファイル選択）。 |
| `split_csv` | `csv_sp` | **可** | 同上。 |
| `normalize_header` | `hd_nr` | **可** | 現状 `hc_main` は `_call_svc_server`。WaitForm は invoke `finally` → **svc 完了側で必ず通知**すること。 |
| `insert_shuka_header` | `hd_in` | **可** | 同上。 |
| `undo_last_action` | `undo` | **可** | 2 ボタン同一 `tag`。WaitForm 同上。 |
| `trim_spaces` | `trm_ex` | **可** | 既に `_call_svc_server`。 |
| `check_duplicates` | `dupli` | **可** | invoke 経路は **`_call_svc_server("dupli", …)` に統一済み**（1.11.6 系）。bridge との挙動差は pickle・WaitForm のみ確認。 |
| `delete_empty_rows` | `row_dl` | **可** | 同上（`row_dl`）。 |
| `delete_empty_cols` | `col_dl` | **可** | 同上（`col_dl`）。 |
| `convert_date_ymd` | `dt_ymd` | **可** | 同上（`dt_ymd`）。 |
| `convert_date_ymd_hm` | `dt_hm` | **可** | 同上（`dt_hm`）。 |
| `show_help` | `help` | **可** | WaitForm・ヘルプ UI の閉じ方確認。 |
| `run_data_agg` | `data_agg` | **可（要拡張）** | **`payload` 等の追加フィールド**を JSON と pickle の `kwargs` に載せるスキーマ設計が必要。 |

### 4.1 共通の追加作業（全機能で検討）

- **VBA `HC_RibbonPerf`**: `before_bridge_submit` / `after_bridge_submit` を **汎用名**にするか、action 別に残すか。  
- **`bridge_requests` フォルダ名**: `main_requests` へリネームするかは任意（**パス変更は VBA・掃除・ドキュメント全更新**）。本プランでは **当面 `bridge_requests` のまま**でも可（Q5 はログ・env・mutex が主眼）。  
- **エラー時**: `SubmitSvcRequestViaBridge` の **Err.Raise + ErrorHandler で `NotifyUiReady`**（現 `load_csv` と同型）。

---

## 5. 旧 `hc_main.py` の移管マップ（案）

| 旧 `hc_main` の塊 | 推奨移管先 |
|-------------------|------------|
| `clear_registry` 等 Excel 内部レジストリ掃除 | **`core`** 配下の専用モジュール（例: `core/excel_session.py`）。VBA は `RunPython` で当該関数のみ直呼び。 |
| タイムアウト・環境連動 | **`core_env`**（既存パターンに合わせる）。 |
| `_ALLOWED_SVC_ACTIONS` / `invoke` ディスパッチ | **削除**（`svc_server` が正）。 |
| `_call_svc_server` | **削除**（bridge が pickle 生成）。 |
| `_ensure_book` / ribbon ログ | **`svc_host` の `register_book`** 等と役割分担を整理。**短寿命プロセス専用のものは削除**。 |
| `INVOKE_ACTIONS` 定数の参照元 | **VBA・テスト**を **リボン tag 一覧**または **単一設定ファイル**に寄せる。 |

---

## 6. 変更が集中するファイル（チェックリスト）

- `VBA/Main.bas` — bridge 一般化、`RunPythonSafe` 削減/削除、`clear_registry` 呼び先  
- `CSV_Tool_xml.txt` — 原則変更なし（`tag` が契約）  
- `svc/bridge_runner.py` → **ルート `hc_main.py`**（フェーズ D）  
- `svc/svc_host.py` — spawn パス、mutex、ログラベル、`is_bridge_running` 相当  
- `core/ipc_cleanup.py` — 起動スイープの呼び出し元  
- `docs/environment_variables.md`, `docs/IPC_TEMP_CLEANUP.md`, `docs/csv_ld_perf_measurement.md`  
- **テスト** — `bridge_runner` / `hc_main` import の置換  
- `VBA/xlwings.bas` — `hc_main.py` パス探索ロジックの有無確認  

---

## 7. リスクと緩和

| リスク | 緩和 |
|--------|------|
| 環境変数リネームで既存ユーザー設定が無効 | **`HC_BRIDGE_*` をフォールバック**で読む期間を設ける。 |
| mutex 名変更で二重起動 | 移行版で **旧 mutex 名も短時間ポーリング**、または **メジャー版で一括切替**をリリースノートに明記。 |
| WaitForm が閉じない | **機能別に結合テスト**（リボン 1 クリック〜画面終了まで）。 |
| `run_data_agg` の payload 不整合 | **JSON スキーマをドキュメント化**し、単体テストで pickle 内容を検証。 |

---

## 8. 完了定義（Definition of Done）

1. すべてのリボン `tag` が **bridge（将来のルート `hc_main.py`）経由**で `svc_server` に到達する。  
2. **VBA に `import hc_main` 文字列が無い**。  
3. **リポジトリに旧 `hc_main.py`（invoke 版）が無い**。  
4. **常駐ブリッジはルート `hc_main.py` のみ**が正とする。  
5. ログ・ドキュメントで **`[MAIN]` / `HC_MAIN_*`** が主表記で、旧 `BRIDGE` は「非推奨・互換」の説明がある。  

---

## 9. 変更履歴

| 日付 | 内容 |
|------|------|
| 2026-04-10 | 初版。議論内容・機能別表・フェーズ・旧 hc_main 移管方針を反映。 |
| 2026-04-11 | フェーズ B 完了記録・スモーク手順・`svc_host` バージョン／常駐再起動メモ・`xlwings.bas` 更新タイミングを追記。 |
| 2026-04-11 | フェーズ C 準備: `hc_main` 依存棚卸し表・移行期間コーディング規約・C の次アクション。機能表の dupli/row_dl 等を invoke 統一済みに更新。 |
| 2026-04-11 | `core/ribbon_public_to_svc.py` 新設: リボン→svc action を `bridge_runner` / `hc_main` で共有。`svc_server.SVC_SERVER_ACTION_KEYS`・契約テスト追加。 |
| 2026-04-11 | WaitForm: `RIBBON_INVOKE_FINALLY_NOTIFY_WAITFORM` / `RIBBON_ACTIONS_READY_UI_CLOSES_WAITFORM` を `ribbon_public_to_svc` に追加。`hc_main` は参照のみ。 |
| 2026-04-11 | `hc_main` 1.11.9: invoke 実装を `_invoke_simple_svc` / `_invoke_csv_family` に集約（フェーズ C 向け重複削減）。 |
| 2026-04-11 | **フェーズ D（一部）**: invoke → `hc_invoke.py`、常駐ブリッジ → ルート `hc_main.py`（0.2.0）。`svc_host` 0.4.23。`svc/bridge_runner.py` は互換ラッパ。 |
| 2026-04-11 | **フェーズ E**: `[MAIN]` ログ、`HC_MAIN_*` + `HC_BRIDGE_*` フォールバック、`hc_main_boot_*.log`、`hc_main` 0.2.1、`svc_host` 0.4.24、上記ドキュメント。 |
| 2026-04-11 | **フェーズ C 完了** + **Mutex**: `core/ribbon_invoke.py`（`hc_invoke.py` 削除）、`excel_session` 更新。`HC_MAIN_RUNNER` / 旧 `HC_BRIDGE_RUNNER` 併用。`hc_main` 0.2.2、`svc_host` 0.4.25。診断 `hc_csv_tool.diag.ribbon_invoke`。 |
