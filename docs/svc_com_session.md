# svc_server Excel COM セッション（B+ 常駐）

**対象**: 開発者・運用（ログ解析）  
**関連コード**: `core/excel_com_session.py`, `svc/svc_server.py`, `svc/svc_host.py`

## 方針（B+）

| 項目 | 動作 |
|------|------|
| 通常操作（成功） | `svc_server` **常駐**。handler キャッシュ・warmup を再利用 |
| マルチ Excel | `_book_cache_by_hwnd` で HWND ごとに Book を保持。切替えで再起動しない |
| COM エラー | `com_recycle scheduled reason=com_session_error` → プロセス終了 → 次回 spawn |
| リボン前 | **事前再起動なし**（`prepare_com_session_before_request` は no-op） |
| Excel 終了 | `com_monitor` が死んだ HWND をキャッシュから除去（プロセスは維持） |
| 全 Excel 終了 | lifecycle monitor が `request_shutdown_all` → 正常終了 |

旧 A+（操作ごとの com_recycle・同一 HWND でも事前 kill）は **廃止**。

## 起動と warmup

1. 初回リボン（または `ensure_python_hosts_ready`）で `svc_server` spawn（1 回）
2. `config/svc_warmup.json` の `warmup_actions` で handler を事前 import（**初回のみ**）
3. 2 回目以降のリボンは `SVC_SERVER already running` のみ

warmup リストは **よく使う・初回が遅い action** に絞る（`data_agg` 等の重い module は初回利用時に load）。

## ログで確認するキーワード

### 正常（常駐）

```
SVC_SERVER spawned          # セッション初回のみ
warmup: action=csv_ld       # 初回のみ
SVC_SERVER already running  # 2 回目以降のリボン
exec start action=... pid=XXXXX  # 同一 pid が続く
attach_book phase=cache_hit hwnd=...
com_monitor pruned_dead_hwnds count=1  # Excel 終了時の掃除のみ
```

### 異常（再起動が多すぎる）

```
com_recycle scheduled reason=after_com_session  # 旧 A+。B+ では成功時に出ない
com_recycle restart begin                       # 旧事前再起動。B+ では出ない
recovery restart begin                          # 救済用 restart_svc_server（通常未使用）
spawned                                         # リボンごとに出る
```

### COM 救済（エラー時のみ）

```
com_recycle scheduled reason=com_session_error
exiting pid=...
# 次のリボンで 1 回 spawned → 以降再び常駐
```

## 診断用 IPC

| ファイル | 用途 |
|----------|------|
| `control/svc_last_com_hwnd.txt` | 最後に COM 接続した HWND（診断用）。`com_monitor` 掃除時に 0 でクリア |

再起動判定には **使わない**（B+）。

## 手動確認（回帰）

1. **単一 Excel**: リボン 10 回 → `spawned`/`warmup` は初回のみ、pid 固定
2. **マルチ Excel**: Book1 ↔ Book2 交互 → `cache_hit` / `xlc_ctx`、再起動なし
3. **Book 終了**: 片方閉じる → `pruned_dead_hwnds`、他方は継続
4. **COM エラー**: 発生時のみ `com_session_error` → 1 回 spawn 後に復帰

## 参照

- `docs/解析_リボンからファイル選択画面までの遅延.md` — 遅延要因と warmup
- `config/svc_warmup.json` — warmup 対象 action
- `docs/environment_variables.md` — `HC_SVC_WARMUP_ACTIONS` フォールバック
