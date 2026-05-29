# 更新通知・二重起動抑制（実装済み）

## 目的

Excel 起動後、新版確認（Qt）を開いたままにすると **遅れて 2 枚目**（Tk / 予約確認など）が出る問題を抑止する。

## 実装

### Python（配布に含まれる）

| ファイル | 内容 |
|----------|------|
| `core/startup_session_gate.py` | 同一 HWND で起動 UI の重複を判定 |
| `svc/svc_host.py` | `maybe_apply` / `maybe_check` をゲートでスキップ可能に |

判定:

- 他 RunPython が `in_progress` … 1 回目がダイアログ待ちのあいだ 2 回目を抑止
- `init_bridge` が `startup_full` 完了直後 … 遅延 InitPythonServer の二重抑止

無効化: 環境変数 `HC_STARTUP_UI_GATE_DISABLE=1`

ログ: `logs/hc_update.log` の `startup: skip_duplicate_update_ui`

### VBA（xlam 側・要エクスポート）

`tools/patch_vba_startup_gate_cp932.py` … **CP932 + CRLF** で `VBA/Main.bas` をパッチ。

1. xlam から `Main.bas` / `ThisWorkbook.cls` をエクスポート
2. `python tools/patch_vba_startup_gate_cp932.py`
3. xlam に再インポート

## 受け入れ

- 新版確認を放置しても **2 枚目が出ない**
- **1 回目の Qt 確認**は残る

## テスト

`pytest tests/test_startup_session_gate.py`
