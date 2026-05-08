# IPC 一時フォルダの起動時スイープ（`%TEMP%\csv_tool`）

## 1. 目的

- **古いリクエストファイル**が残ったまま次回起動すると、`ui_server` / 常駐 **`hc_main.py`**（旧 `bridge_runner`）/ `svc_server` がそれを拾い、**意図しないダイアログ表示や処理再実行**につながることがある。
- 本機能は、各常駐プロセスが **単一インスタンス用ミューテックスの取得に成功した直後**に、指定フォルダ内の滞留ファイルを掃除する。基本は **TTL スイープ**（最終更新時刻がしきい値より古いものだけ削除）だが、**`bridge_requests` の `*.json` のみ**は **`hc_main` 起動時**に **age に関わらず全削除**したうえで、同じく TTL スイープを **補助**として続けて実行する（長時間常駐やエッジ向け）。
- **二重起動と判定されたプロセス**（mutex が既に存在）はスイープを行わない。実行中の別インスタンスのキューを消さないため。

## 2. 実装場所

| モジュール | 役割 |
|------------|------|
| `core/ipc_cleanup.py` | TTL 削除の共通実装・各プロセス用エントリポイント |
| `ui_qt/ui_server.py` | `run_ui_server_startup_sweeps` |
| ルート `hc_main.py`（互換: `svc/bridge_runner.py` ラッパ） | `run_bridge_startup_sweeps` |
| `svc/svc_server.py` | `run_svc_server_startup_sweeps` |

## 3. フォルダ別の扱い

| フォルダ（IPC ルート配下） | いつ掃除するか | パターン | 既定 TTL | 備考 |
|----------------------------|----------------|----------|----------|------|
| `requests` | `ui_server` 起動時 | `req_*.pkl` | 24h | claim 済み `.work.pkl` も `req_*.pkl` に合致するため対象 |
| `bridge_requests` | **`hc_main` 起動時** | `*.json` | 24h（補助） | **起動直後に `*.json` を全削除**（前セッション残留の誤再処理防止）。続けて TTL で古い残りを掃除。JSON は **UTF-8**（`hc_main` は utf-8-sig / utf-8 / cp932 の順で解釈） |
| `svc_requests` | `svc_server` 起動時 | `svc_req_*.pkl` | 24h | `hc_main` / ブリッジからの svc 依頼 |
| `svc_results` | `svc_server` 起動時 | `svc_res_*.pkl` | **3600s（1h）** | `hc_main._cleanup_old_res_files` と同系（早期 return で読まれず残りやすい） |
| `control` | 上記各プロセス起動時 | `*_starting.flag` のみ | **600s（10m）** | **`shutdown.flag` / `svc_shutdown.flag` は削除しない**（パターン不一致のため触れない） |

### 3.1 `control` について

- 削除対象は **`ui_server_starting.flag`**, **`bridge_starting.flag`**, **`svc_server_starting.flag`** のような **`*_starting.flag`** のみ。
- **最終更新から TTL を超えたもの**だけ削除（クラッシュ等で `finally` が走らず残ったガードの掃除）。
- 終了要求用の **`shutdown.flag`**（UI）および **`svc_shutdown.flag`**（svc）は **本スイープでは削除しない**。

## 4. 環境変数

詳細は `docs/environment_variables.md` の **§4.1.1** も参照。

| 変数名 | 既定 | 意味 |
|--------|------|------|
| `HC_IPC_DISABLE_STARTUP_SWEEP` | 無効 | `1` / `true` 等で **全スイープを無効化**（調査用）。 |
| `HC_IPC_SWEEP_QUEUE_TTL_SEC` | `86400` | `requests` / `bridge_requests` / `svc_requests` の TTL（秒）。 |
| `HC_IPC_SWEEP_SVC_RESULTS_TTL_SEC` | `3600` | `svc_results` の TTL（秒）。`hc_main` の古い `svc_res` 掃除と揃える想定。 |
| `HC_IPC_SWEEP_STARTING_FLAG_TTL_SEC` | `600` | `*_starting.flag` の TTL（秒）。 |

## 5. 本スイープの対象外（参考）

以下は **本機能では削除しない**。別途、読取後の `unlink` や手動削除が必要になる場合がある。

- `results` / `ready` / `result`（単数）配下のタイムスタンプ付き pickle
- `logs` 配下のブートログ（例: **`hc_main_boot_*.log`**。旧 **`bridge_boot_*.log`** はフェーズ E 以降は作成しない）
- `csv_tool` 直下の `hc_csv.log` など運用ログ

将来、TTL 定期スイープや読後削除を広げる場合は本ドキュメントを更新する。

## 6. 変更履歴

| 日付 | 内容 |
|------|------|
| 2026-04-10 | 初版。`core/ipc_cleanup` と各常駐プロセス起動時フック。 |
| 2026-04-11 | フェーズ E: 常駐入口表記を **`hc_main`** に、`logs` のブートログ名を追記。 |
