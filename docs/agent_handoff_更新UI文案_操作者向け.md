# Agent 引き継ぎ: 更新 UI 文案（操作者向け）— 実装済み

## 概要

新バージョン検出から更新完了まで、操作者向けの日本語に統一。本文なしの画面も `ui_update_check.json` で `""` を明示。

## 変更ファイル

- `config/ui_update_check.json`
- `core/packaged_update.py`（確認ダイアログフォールバック）
- `bootstrap/update_bootstrap.py`（進捗 UI 初期表示・`_ui_update_message` の空文字対応）
- `hc_updater.py`（`状態:` プレフィックス・フォールバック文案）

## 操作者に見える順序

1. MessageBox: お使いの版 / 新しい版 → すぐに更新 / 後で
2. 進捗: 準備中 → 準備完了 + Excel 終了
3. hc_updater: Excel 終了待ち（本文 `""`）→ 開始 → 取得 → 展開 → 適用 → 完了
4. MessageBox: Excel を起動し直してください。

## 本文が空のキー

- `UPDATER_INITIAL_MESSAGE`: `""`
- `UPDATER_PHASE_WAIT_MESSAGE`: `""`
