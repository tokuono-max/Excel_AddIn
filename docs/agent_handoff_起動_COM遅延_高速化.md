# データ集約メイン画面 起動高速化 Phase 2（COM / WaitForm）

## 実装内容（0.4.47）

`create_dialog(main)` の起動順を変更:

```
__init__ → prepare → write_waitform_ready → show → QTimer(0) 非同期 _pulse
```

- 起動時の **同期 `_pulse`（guid_scan 含む）を削除**
- **`.ready` を `prepare_done` 直後**に移動（VBA WaitForm を COM 前に閉じる）
- **`show()` 後に `QTimer.singleShot(0, _deferred_create_pulse)`** — guid_scan は維持
- `showEvent` のパルス（即時 + 90/200/450 ms）は変更なし

## 計測（改善前・Phase 1 後）

| 指標 | 値 |
|------|-----|
| 走査非同期化前 `create_dialog` | ~31,762 ms |
| 走査非同期化後 | ~9,881 ms |
| `pulse_after_get_ctx` step_ms | ~5,010 ms（WaitForm 待ちの主因） |

## 実装内容（0.4.48 / Phase 3）

```
__init__ → prepare → write_waitform_ready → show
  → showEvent: _schedule_excel_unlock_pulse_chain（QTimer(0) + 90/200/450ms、1本のみ）
```

- `showEvent` の**同期** `_pulse` を廃止（`show_done` ~3.9秒の主因を除去）
- `_schedule_excel_unlock_pulse_chain` で create/showEvent の二重 pulse を防止
- `create_dialog` の `_deferred_create_pulse` を削除（showEvent が同チェーンを予約）
- teardown で `_excel_unlock_pulse_chain_scheduled` をリセット

## 計測（Phase 2 後）

| 指標 | 値 |
|------|-----|
| `create_dialog ok` | ~6,458 ms |
| `waitform_written` | ~1,410 ms |
| `show_done` step_ms | ~3,868 ms（Phase 3 で短縮見込み） |

## リスク

- 起動直後 ~1秒、Excel リボン操作感（showEvent + 非同期 pulse で緩和）
- guid_scan 削除はしていない（マルチブック誤結合回避）
