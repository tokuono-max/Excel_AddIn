# エージェント向け引継ぎ：COM 切断・csv_ld 連続失敗（EXE）（2026-07-03）

## 1. 現状（2026-07-03 夕方更新）

| 項目 | 状態 | 備考 |
|------|------|------|
| **症状の特定** | ✅ 確定 | 連続 csv_ld / ld→sv で `_attach_book` 即死。`hc_svc_server` 手動キル後は成功 |
| **直接原因** | ✅ 確定 | B+ 常駐 + 死んだ Book キャッシュ再利用 + `com_error` 後の `SystemError` で回復未到達 |
| **B+ 導入との関係** | ✅ 確定 | 旧 A+ は操作ごと `com_recycle`。B+（0.1.19〜）で同一 pid 累積が顕在化 |
| **前面化** | ✅ 無関係 | 2 回目失敗は UI 前の `_attach_book` |
| **バグ修正（操作境界）** | ✅ 実施 | `svc_server` 0.1.21 — cache 破棄・skip_cache・PyErr_Clear・com_recycle 検出拡張 |
| **根本対策（UI 待ち後）** | ✅ 実施 | `svc_server` 0.1.22 — `resolve_fresh_book_after_ui_wait`。`csv_ld` 1.3.38 / `csv_sv` 1.3.14 |
| **Python 検証** | ✅ 連続 2 回成功 | 同一 pid・`resolve_enter`・`cache_hit` なし（18:10〜18:12 ログ） |
| **EXE 検証** | ⏳ 未実施 | ビルド後に同一手順で確認 |

---

## 2. 症状（ユーザー報告）

- **配布 EXE** で **csv_ld 連続 2 回** または **ld 直後 sv** が `_attach_book` で `com_error (-2147220995)`。
- **Python 開発**では同一手順が通ることが多い（修正後は連続 ld も同一 pid で成功）。
- **手動で `hc_svc_server.exe` を終了**すると Excel を開いたまま次操作が通る → **プロセス内キャッシュ汚染**が主因。

---

## 3. 確定した失敗パス（100% 再現時）

```
1 回目 csv_ld 成功（内部で [COM_NG] → book_reattached のことも）
    ↓
_book_cache_by_hwnd に死にかけの Book ラッパーが残る（B+ は com_recycle しない）
    ↓
2 回目 _attach_book → _validate_book_alive で com_error
    ↓
例外状態漏れ → logger が SystemError → get_excel_context_from_hwnd 未到達
    ↓
com_recycle も不発 → 以降も不安定
```

**分類**: アーキテクチャ全否定ではなく、**B+ のキャッシュ前提の穴 + 回復経路の実装バグ**。

---

## 4. 実装済み対策（二段構え）

### 4.1 操作境界（`svc_server` 0.1.21）

| 内容 | ファイル |
|------|----------|
| attach_book action は `skip_cache` + 入口/終了で `invalidate_attached_book_cache` | `svc/svc_server.py` |
| `_validate` 後 `PyErr_Clear`、COM エラー時 1 回リトライ | 同上 |
| `is_com_session_error` が `SystemError` の `__cause__` を辿る | `core/excel_com_session.py` |
| `book_resolved_via_hwnd` 後 `_store_attached_book` | `svc_csv_ld.py`, `svc_csv_sv.py` |

### 4.2 UI 待ち後の根本対策（`svc_server` 0.1.22）

**方針**: `[COM_NG]` を「正常に近いソフト失敗」とみなさない。**`find_sheet_by_guid` の前に必ず HWND から Book を取り直す**。万一失敗したときだけ `get_excel_context_from_hwnd` 等で保険。

| API / 変更 | 役割 |
|------------|------|
| `resolve_fresh_book_after_ui_wait()` | UI 確定後・GUID 解決前の **必須** Book 再取得 |
| `csv_ld._resolve_book_and_sheet` | 上記を先に呼び、`book_reattached` 事後救済を整理 |
| `csv_sv._resolve_book_and_sheet` | `_reattach_book` 経由で同様（先に fresh） |

**期待ログ（正常時）**:

```
book_fresh_after_ui_wait hwnd=...
find_sheet_by_guid → COM_NG なし
```

`book_reattached` は **保険経路**（`book_resolved_via_hwnd`）に集約。

---

## 5. ログで確認するキーワード

### 正常（修正後・連続操作）

```
exec start action=csv_ld pid=XXXX   # 1・2 回目とも同一 pid 可
attach_book phase=resolve_enter     # cache_hit ではない
book_fresh_after_ui_wait            # do_load 前（根本対策）
load_csv_flow_done / save_csv_flow_done
```

### 異常（再発時）

```
attach_book phase=cache_hit         # attach_book action で出たら要調査
[COM_NG] find_sheet_by_guid         # 出ても book_fresh 後なら要調査
exec failed / SystemError
com_recycle scheduled               # 失敗時は出るべき
```

---

## 6. 検証手順

### 6.1 Python（実施済み 2026-07-03）

1. Excel 再起動 → `start_excel_dev.bat`
2. csv_ld ×2（svc キルなし）
3. 確認: 同一 pid、2 回とも `exec done`、`resolve_enter`、`cache_hit` なし

### 6.2 EXE（未実施）

1. 再ビルド・再インストール（`svc_server` ≥ 0.1.22）
2. csv_ld ×2 → ld→sv
3. §5 の正常キーワードを確認

### 6.3 切り分け（参考）

| 操作 | 意味 |
|------|------|
| 1 回目後に `hc_svc_server` キル → 2 回目成功 | プロセス内汚染（修正対象） |
| キルなしで 2 回目成功 | 修正効果あり |

---

## 7. 関連コード

| ファイル | 役割 |
|----------|------|
| `svc/svc_server.py` | `_attach_book`, `invalidate_attached_book_cache`, `resolve_fresh_book_after_ui_wait` |
| `core/excel_com_session.py` | `action_attach_book_fresh_resolve`, `is_com_session_error` |
| `svc/svc_csv_ld.py` | `_resolve_book_and_sheet`, `load_csv` |
| `svc/svc_csv_sv.py` | `_reattach_book`, `_resolve_book_and_sheet` |
| `core/core_xlc.py` | `find_sheet_by_guid`（`[COM_NG]` はここでログ） |
| `docs/svc_com_session.md` | B+ 方針の正本 |

---

## 8. テスト

```powershell
cd c:\Project\Python\Excel_AddIn
.venv\Scripts\python.exe -m pytest tests/test_excel_com_session.py tests/test_svc_server_attach_book.py tests/test_csv_sv_book_reattach.py tests/test_csv_sp_book_reattach.py -q
```

---

## 9. やってはいけないこと

- `[COM_NG]` を正常フローとして放置する（fresh resolve 前に出るのは設計漏れ）
- `cache_hit` リトライだけで済ませる（操作境界 invalidate とセットが必須）
- 前面化パッチと COM 修正を同一コミットに混ぜない（トピック分離）

---

## 10. 環境

| 項目 | 値 |
|------|-----|
| `svc_server` | **0.1.22**（本引継ぎ時点） |
| `svc_csv_ld` | **1.3.38** |
| `svc_csv_sv` | **1.3.14** |
| ログ | `%TEMP%\csv_tool\hc_csv.log` |
| COM 方針 doc | `docs/svc_com_session.md` |

---

*更新: 2026-07-03（操作境界修正 + UI 待ち後 fresh Book 根本対策を反映）*
