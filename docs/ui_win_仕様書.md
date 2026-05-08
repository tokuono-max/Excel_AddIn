# ui_win 仕様書（ウィンドウ制御）

**対象**: ui_qt/ui_win.py — Qt UI サーバ側のウィンドウ制御（挙動）  
**作成日**: 2026-03-09  
**目的**: Excel HWND に対するロック/解除・前面追従・矩形取得の役割と API を定義する。

---

## 1. 機能概要

**ui_win** は、Qt UI 層における「ウィンドウ制御（挙動）」を集約するモジュールである。次の責務を持つ。

- **Excel 実質モーダル**: ダイアログ表示中に Excel の操作を無効化し、閉じたあとで解除する。子 HWND を再帰列挙して EnableWindow を実行する。**有効化時はルート（Excel トップ）HWND も含める**。ルートを含めないとリボンのみ有効でシート・スクロール・セル編集が効かない事象を防ぐ。
- **前面追従**: Excel が前面になったときに、指定ダイアログを 1 回だけ raise/activate する軽量な補助（ping-pong を起こさない）。
- **矩形取得**: Excel HWND のウィンドウ矩形を物理ピクセルで取得する。

本モジュールは **設定（cst）を読まない**。表示仕様の解釈は ui_common の責務。Win32 API の宣言・直呼びは **core.core_w32** に委譲し、重複実装を避ける。

---

## 2. 依存関係

| 対象 | 内容 |
|------|------|
| **core** | core_w32（enum_child_windows, enable_windows, get_window_rect, get_window_pid, start_foreground_hook, stop_foreground_hook） |
| **PySide6** | QtCore, QtWidgets（QWidget） |

---

## 3. 公開 API

| 関数 | 説明 |
|------|------|
| **enable_excel_window(hwnd, enabled)** | Excel のトップ HWND を起点に子を再帰列挙（最大 20000 件）し、EnableWindow を実行。**enabled=True のときは to_enable にルートを先頭に加え**、ルート＋子すべてを有効化する。False のときは子のみ無効化。 |
| **get_excel_rect(parent_hwnd)** | Excel HWND のウィンドウ矩形を取得する。core_w32.get_window_rect を呼び、物理ピクセルで (left, top, right, bottom) 相当の値を返す。失敗時は None。 |
| **start_front_follow(dialog, parent_hwnd)** | Excel のプロセス ID を取得し、EVENT_SYSTEM_FOREGROUND で前面ウィンドウが Excel と同じ PID になったときに dialog.raise_() と activateWindow() を 1 回だけ実行するフックを開始する。 |
| **stop_front_follow()** | 前面追従用のフックを停止する。 |

---

## 4. 利用元

| 利用元 | 用途 |
|--------|------|
| **ui_csv_ld, ui_csv_sv, ui_csv_sp, ui_csv_mg, ui_hd_nr** | ダイアログ表示時・閉じ時に enable_excel_window(False/True) を呼び、Excel 操作のロック・解除に使用。 |
| **ui_common** | center_on_excel 等で get_excel_rect の代わりに自前の get_excel_rect を持つ。enable_excel_window は ui_common 内のダイアログから自モジュール実装を呼ぶ。機能側は ui_win の enable_excel_window を明示的に import して使用することを推奨。 |

---

## 5. 参照

| 項目 | 内容 |
|------|------|
| **モジュール** | ui_qt/ui_win.py |
| **Win32 実装** | core/core_w32.py |
| **デグレ防止** | docs/共通モジュール変更時_デグレ防止.md（ui_win 変更時：各機能の Excel ロック/解除の確認） |
