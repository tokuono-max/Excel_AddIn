# app 仕様書（Qt アプリケーションエントリ）

**対象**: ui_qt/app.py — Qt UI の最小エントリ（同一プロセス・デバッグ用）  
**作成日**: 2026-03-09  
**目的**: QApplication の単一生成・DPI 初期化・結合ダイアログ起動の役割と API を定義する。

---

## 1. 機能概要

**app** は、Qt UI を**同一プロセス**で動かすときの最小エントリを提供する。

- **QApplication の単一化**: get_qapp() で 1 回だけ QApplication を生成し、setQuitOnLastWindowClosed(True) を設定する。
- **DPI 初期化**: Qt プロセス側でのみ DPI 認識を初期化する（SetProcessDpiAwarenessContext: PER_MONITOR_AWARE_V2）。core / svc では行わない。
- **結合ダイアログの起動**: run_merge_dialog(parent_hwnd) で CSV 結合設定ダイアログ（CsvCsvMergeDialog）を表示し、結果辞書を返す。同一プロセス・デバッグ用途に限定する。

**既定運用**は、別プロセス runner（pythonw）による UI サーバ起動であり、OLE 待機抑止のため Excel と同一プロセスで Qt を動かさない。本モジュールは **USE_QT_SUBPROCESS=False** などのデバッグ用途に限って使用する。

---

## 2. 依存関係

| 対象 | 内容 |
|------|------|
| **PySide6** | QApplication |
| **ui_qt** | ui_csv_mg.CsvCsvMergeDialog |
| **Win32** | ctypes（SetProcessDpiAwarenessContext） |

---

## 3. 公開 API

| 関数 | 説明 |
|------|------|
| **get_qapp()** | グローバルな QApplication を 1 回だけ生成して返す。初回時は _init_dpi_awareness() を実行したあと、QApplication(sys.argv) を生成し、setQuitOnLastWindowClosed(True) を設定する。 |
| **run_merge_dialog(parent_hwnd)** | get_qapp() でアプリを取得し、結合設定ダイアログ（ui_csv_mg.CsvMergeDialog）を生成して exec() する。戻り値は dlg.get_result()（accepted / paths / header_mode 等を含む辞書）。 |

### 内部

| 関数 | 説明 |
|------|------|
| **_init_dpi_awareness()** | SetProcessDpiAwarenessContext(-4)（PER_MONITOR_AWARE_V2）を実行。失敗時は無視。 |

---

## 4. 運用上の注意

- **本番の UI 表示**: 通常は svc_host / ui_server が別プロセスで UI を起動するため、app を直接 import して run_merge_dialog を呼ぶ経路は使わない。
- **デバッグ時**: 同一プロセスで結合ダイアログだけを表示して動作確認する場合に、app.get_qapp() と app.run_merge_dialog(hwnd) を使用する。

---

## 5. 参照

| 項目 | 内容 |
|------|------|
| **モジュール** | ui_qt/app.py |
| **結合ダイアログ** | ui_qt/ui_csv_mg.py（CsvMergeDialog） |
| **通常起動** | svc_host.ensure_ui_server 経由で別プロセス UI サーバが起動し、ipc_file で req/res をやり取りする。 |
