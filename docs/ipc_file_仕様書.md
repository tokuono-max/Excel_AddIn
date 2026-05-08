# ipc_file 仕様書（IPC ファイル通信）

**対象**: ui_qt/ipc_file.py — UI サーバと svc 層のファイル（Pickle）通信  
**作成日**: 2026-03-09  
**目的**: リクエスト・結果・制御用パス、Pickle 読書、Mutex、要求投入の役割と API を定義する。

---

## 1. 機能概要

**ipc_file** は、Qt UI サーバ（別プロセス）と svc 層の間で、**ファイル（Pickle）を用いた最小 IPC** を提供する。

- **req_*.pkl**: svc が ui_server へ依頼を渡すために書き出す。ui_server が get_request_dir 配下を監視し、pop_next_request で claim（リネーム）してから処理する。
- **res_*.pkl**: ui_server が処理結果を書き出し、svc が result_path を監視して受け取る。
- **ready_*.pkl**: 初回描画完了などの早期通知を ui_server が書き、svc が ready_path を監視する場合に使用。
- **control 配下**: シャットダウンフラグ等。ジオメトリ保存は ui_common が get_control_dir 配下の geometry を使用。
- **基準フォルダ**: 直近のファイル読込/保存で使ったフォルダを last_folder.txt で保持し、CSV 読込・保存・分割で共通利用する。

環境変数 **HC_QT_IPC_DIR** が設定されていればそのパスを IPC ルートに使用。未設定時は **%TEMP%\\csv_tool** に固定し、双方のプロセスで一致させる。

**起動時スイープ**: `ui_server` 起動直後（mutex 取得成功時）に、`requests` 配下の古い `req_*.pkl` を TTL 削除する。詳細・環境変数は **`docs/IPC_TEMP_CLEANUP.md`**（実装: **`core.ipc_cleanup`**）を参照。

---

## 2. ディレクトリ構成（IPC ルート配下）

| パス | 説明 |
|------|------|
| **requests/** | req_*.pkl を格納。ui_server が pop_next_request で取り出し、.work.pkl にリネームして処理。 |
| **results/** | res_*.pkl を格納。svc が result_path に結果を書き、呼び出し元が読みに来る。 |
| **ready/** | ready_*.pkl（READY_UI 等）を格納。 |
| **control/** | shutdown.flag、geometry/ 等。 |
| **control/geometry/** | ui_common の save_geometry / restore_geometry が screen_key ごとの pkl を保存。 |
| **progress/** | 進捗用 pkl（progress_hd_nr_*.pkl 等）を格納。svc が書き、UI がポーリングして読む。 |
| **last_folder.txt** | 基準フォルダの絶対パス（1 行）。 |
| **logs/** | 子プロセスのブート用ログ等（例: **`hc_main_boot_*.log`**。実装は `svc_host.spawn_bridge` 等）。本仕様の API 外だが IPC ルート配下に作成される。 |

---

## 3. 公開 API

### 3.1 パス取得

| 関数 | 戻り値 | 説明 |
|------|--------|------|
| **get_ipc_root()** | Path | IPC ルート（HC_QT_IPC_DIR または %TEMP%\\csv_tool）。必ず mkdir して返す。 |
| **get_request_dir()** | Path | get_ipc_root() / "requests"。mkdir 済み。 |
| **get_server_log_path()** | Path | get_ipc_root() / "ui_server.log"。 |
| **get_control_dir()** | Path | get_ipc_root() / "control"。mkdir 済み。 |
| **get_shutdown_flag_path()** | Path | get_control_dir() / "shutdown.flag"。 |

### 3.2 基準フォルダ

| 関数 | 説明 |
|------|------|
| **get_last_folder()** | last_folder.txt の内容を返す。無い・空・ディレクトリでない場合は空文字。 |
| **set_last_folder(dir_path)** | 有効なディレクトリパスを last_folder.txt に書き込む。空なら何もしない。 |

### 3.3 シャットダウン

| 関数 | 説明 |
|------|------|
| **write_shutdown_flag()** | shutdown.flag を作成（ベストエフォート）。 |
| **clear_shutdown_flag()** | shutdown.flag を削除（ベストエフォート）。 |

### 3.4 Pickle 読書

| 関数 | 説明 |
|------|------|
| **write_pickle(path, data)** | data を pickle 化し、同一ディレクトリに一時ファイルで書いてから os.replace で原子的に path に置換。中途半端な 0 バイトファイルを残さない。 |
| **read_pickle(path)** | path.read_bytes() を pickle.loads して返す。 |

### 3.5 リクエスト投入・取得

| 関数・クラス | 説明 |
|--------------|------|
| **UiRequest** | parent_hwnd, result_path, ready_path, sheet_id, log_path, action, module, req_dict を持つ dataclass。to_dict / from_dict で辞書と相互変換。 |
| **submit_request(req)** | get_request_dir() に req_&lt;timestamp&gt;_&lt;pid&gt;.pkl を生成して書き込む。存在・サイズ確認し、失敗時は RuntimeError。戻り値はその Path。 |
| **pop_next_request()** | get_request_dir() 内の req_*.pkl を更新時刻でソートし、先頭を _claim_request_file で .work.pkl にリネーム。リネームできた Path を返す。取り出せなければ None。 |
| **_claim_request_file(path)** | path を .work.pkl に原子的にリネーム。成功時は新しい Path、失敗時は None。 |

### 3.6 Mutex（多重起動防止）

| 関数 | 説明 |
|------|------|
| **create_single_instance_mutex(name)** | Windows の CreateMutexW を実行。name 未指定時は _MUTEX_NAME（Global\\HC_QT_UI_SERVER）。戻り値は (handle, already_running)。already_running は GetLastError() == ERROR_ALREADY_EXISTS のとき True。 |
| **release_mutex(handle)** | ReleaseMutex と CloseHandle を実行。失敗しても無視。 |

---

## 4. 注意事項

- **req の取り出し**: ui_server は pop_next_request で「リネーム」により claim するため、同一 req が二重に処理されない。書き込み直後に path.exists() および st_size を確認し、作成失敗時は submit_request が例外を上げる。
- **write_pickle**: 監視側が「存在した瞬間」に読みに来る場合を考慮し、一時ファイルに書き終えてから replace する atomic 書き込みを行う。

---

## 5. 参照

| 項目 | 内容 |
|------|------|
| **モジュール** | ui_qt/ipc_file.py |
| **利用元** | ui_server（pop_next_request, read_pickle, get_request_dir）、各 svc（write_pickle, get_ipc_root, get_request_dir, get_control_dir 等）、ui_common（write_pickle, read_pickle, get_control_dir, get_shutdown_flag_path） |
| **デグレ防止** | docs/共通モジュール変更時_デグレ防止.md（ipc_file 変更時：各機能の UI 表示・結果やり取りの確認） |
