# エージェント向け引継ぎ：UI 前面化・Z 順（csv_ld / 全機能退行）（2026-06-29）

## 1. 本日の目的と現状

| 項目 | 状態 | 備考 |
|------|------|------|
| **柱1** ld/sv ネイティブファイル選択（comdlg32） | ✅ 実装済・単体テスト PASS | `QFileDialog` ホストを出さない。pythonw FG 化の主因対策 |
| **柱2** Qt 画面の `bring_excel_first` 既定 false | ✅ 実装済・単体テスト PASS | data_agg / 重複 / 空白行等の Excel 背後退行対策 |
| **柱3** `ensure_front` AttachThreadInput 共通化 | ✅ 実装済・単体テスト PASS | ヘルプ限定だった強化を全ダイアログへ |
| **ui_server** モーダル中 ld/sv IPC 拒否 | ✅ 実装済 | sp 分割表示中の ld 割り込み防止 |
| **実機受け入れ** | ❌ **未確定・要継続** | ユーザー報告: 「破綻状態に近い」。4 本テスト未完了の可能性 |
| **コミット** | ❌ 未実施 | 前面化関連を含む広範 diff が未コミット（**41 ファイル超**）。ユーザー依頼時のみコミット |
| **設計一本化** | ⚠️ 途中 | パッチ層は減らしたが、タイマー経路は依然複数（`prepare` / `showEvent` / `apply_window_config`） |

---

## 2. 背景（症状の経緯）

### 2.1 当初の症状

- **csv_ld / csv_sv** のファイル選択で **pythonw.exe のウィンドウ**が見える、または **Excel 背後**に「開く」が隠れる。
- **2 巡目**（ld→sv→mg→sp 後に再度 ld）で悪化しやすい。
- **csv_sp** 分割画面は `bring_excel_first=False` + `TOPMOST:true` で一部改善済みだったが、他機能とポリシーが分裂。

### 2.2 退行拡大（2026-06-29 ユーザー報告）

- **データ集約・重複チェック・空白行/列削除**など、ファイル選択以外の Qt 画面も **Excel の背後**に表示。
- 「積み重ねた対策が振り出しに戻った」——**症状ごとの局所パッチが設計を壊した**状態。

### 2.3 ログで確定していたこと（`hc_csv_diag.log` 分析）

| 観測 | 意味 |
|------|------|
| `ensure_front` で `dlg=QFileDialog`, `fg_exe=pythonw.exe`, `sfw_ok=1` | API は成功しているが **FG は Qt ホスト**。OS 標準 `#32770` ではない |
| sp 表示中に ld が起動（同一 `ui_pid`） | **ui_server にモーダル排他がなかった** |
| data_agg はログ上前面成功のケースもあり | 症状は **手順・累積依存** |

---

## 3. 根本原因（設計レベル）

### 3.1 前面化ポリシーの分裂

| パターン | 機能例 | 設定 |
|----------|--------|------|
| A | csv_sp 分割・重複 | `bring_excel_first=False` + JSON `TOPMOST:true` |
| B | data_agg, mg 等 | `TOPMOST:false` + 旧既定 `bring_excel_first=True` |
| C | ld/sv（旧） | `QFileDialog.exec` + FG 監視 + `ensure_front` on QFileDialog |

**長寿命 ui_server（pythonw 1 本）**が全機能を共有するため、1 機能の対策が他機能の Z 順を壊す。

### 3.2 ld/sv 固有

- 仕様: ネイティブ OS ダイアログを Excel の子として Excel **より前**に出す（`docs/svc_csv_ld_仕様書.md`）。
- 誤り: `QFileDialog` に `SetForegroundWindow` → **pythonw が FG になる**（仕様違反）。

### 3.3 Qt 全般

- `bring_excel_first=True` は先に Excel を前面化 → 続くダイアログの `SetForegroundWindow` が失敗しやすい → **ダイアログが Excel 背後**。
- `SetForegroundWindow` 単体では不十分（Windows のフォーカス制限）。**AttachThreadInput + nudge** が必要だったがヘルプだけに限定されていた。

---

## 4. 実装方針（3 本柱・同時にいじらない）

### 柱1: ld/sv ネイティブファイル（comdlg32）

| ファイル | 内容 |
|----------|------|
| `core/core_w32.py` v2.2.0 | `win32_get_open_file_name` / `win32_get_save_file_name`, `qt_name_filter_to_win32` |
| `ui_qt/ui_fil.py` v0.4.0 | Excel 向けは comdlg32 のみ。`QFileDialog` / `ensure_front` / FG 監視タイマー廃止 |
| `ui_qt/ui_csv_ld.py` v1.3.16 | `ui_fil.show_open_file_dialog_for_excel` 経由 |
| `ui_qt/ui_csv_sv.py` v1.4.6 | `ui_fil.show_save_file_dialog_for_excel` 経由 |

**やらないこと**: `QFileDialog.exec` を前面化の錨にする、`bring_excel_first=True` を ld に戻す。

### 柱2: Qt モーダル／モードレスの既定ポリシー

| ファイル | 内容 |
|----------|------|
| `ui_qt/ui_common.py` v0.2.88〜0.2.89 | `ensure_front` / `prepare_dialog` の既定 `bring_excel_first=False`。プロパティ `_hc_ensure_front_bring_excel_first=True` のときのみ Excel 先前面化 |
| `ui_qt/ui_data_agg.py` | `_schedule_deferred_excel_owner_front` で `bring_excel_first=False` |

**csv_sp 成功パターン**（必要時のみ）: `bring_excel_first=False` + JSON `TOPMOST:true`（`ui_qt/ui_csv_sp.py` 参照）。

### 柱3: ensure_front 強化の一本化

| ファイル | 内容 |
|----------|------|
| `ui_qt/ui_common.py` v0.2.89 | `_strengthen_widget_foreground`: `bring_excel_first=False` または SFW 失敗時に **全ダイアログ**へ `AttachThreadInput` + `nudge_top_level_to_foreground` |
| `ui_qt/ui_server.py` v1.4.64 | modeless 二重 `ensure_front` タイマーを**削除**（ui_common に一本化） |
| `ui_qt/ui_server.py` v1.4.63 | csv_sp 分割 / 同名確認 / csv_mg 結合表示中は csv_ld/csv_sv ファイル選択 IPC を `CANCEL` |

---

## 5. 前面化の正本（次担当が触る場所）

```
[リボン] → svc → ui_server (長寿命 pythonw)
                    ↓
         create_dialog → prepare_dialog_excel_center_before_show
                    ↓
              ensure_front (既定 bring_excel_first=false)
                    ↓
         _strengthen_widget_foreground (attach + nudge)
```

| 経路 | 用途 | 触るべきか |
|------|------|------------|
| `ui_fil` + comdlg32 | ld/sv ファイル選択のみ | ld/sv 問題時のみ |
| `ui_common.ensure_front` | **全 Qt 画面** | **ここを正本にする** |
| `ui_common.prepare_dialog_excel_center_before_show` | exec/show 前の 1 回準備 | ちらつき・配置も担当 |
| 各画面 `showEvent` の遅延 ensure | data_agg, dupli 等 | **増やさない**。既存削除は要検討 |
| `apply_window_config` の QTimer | 旧経路 | 二重化注意 |
| JSON `TOPMOST` | 最後の手段 | csv_sp 型のみ |

---

## 6. 変更ファイル一覧（前面化関連・抜粋）

未コミット diff に含まれる。`git diff --stat HEAD` で全体確認。

```
core/core_w32.py              # comdlg32 ラッパ
core/core_env.py              # HC_UI_NATIVE_FILE_DIAG
ui_qt/ui_fil.py               # v0.4.0
ui_qt/ui_common.py            # v0.2.89（__version__ 行は 0.2.87 のまま未整合あり）
ui_qt/ui_server.py            # v1.4.64（__version__ 行は 1.4.62 のまま未整合あり）
ui_qt/ui_csv_ld.py            # v1.3.16
ui_qt/ui_csv_sv.py            # v1.4.6
ui_qt/ui_data_agg.py          # deferred front
ui_qt/ui_csv_sp.py            # bring_excel_first プロパティ（既存）
config/ui_csv_sp.json         # TOPMOST 等
tests/test_ui_fil_native_dialog_prep.py
tests/test_ui_common_ensure_front_default.py
tests/test_ui_server_file_pick_block.py
```

**注意**: 同じ diff に **data_agg マスタデバッグ・起動高速化・svc リファクタ**等が混在している。コミット時は **トピック別に分割**を推奨。

---

## 7. テスト

### 7.1 自動テスト（実施済み）

```powershell
cd c:\Project\Python\Excel_AddIn
.venv\Scripts\python.exe -m pytest `
  tests/test_ui_fil_native_dialog_prep.py `
  tests/test_ui_common_ensure_front_default.py `
  tests/test_ui_server_file_pick_block.py -q
```

期待: **10 passed**（2026-06-29 時点）。

### 7.2 実機受け入れ（未完了・必須）

**Excel 完全終了 → 再起動**後に実施。

| # | 手順 | 合格条件 |
|---|------|----------|
| 1 | 1 巡目 **csv_ld** のみ | pythonw **不可視**。OS「開く」が Excel の前 |
| 2 | ld→sv→mg→sp 後、**2 巡目 ld** | 同上 |
| 3 | **データ集約** メイン | Excel の前 |
| 4 | **重複チェック** レポート | Excel の前 |
| 5 | **空白行/列** 確認画面 | Excel の前 |

### 7.3 診断ログ

環境変数:

```
HC_UI_FG_DIAG=1
HC_UI_NATIVE_FILE_DIAG=1
```

ログ: `%TEMP%\csv_tool\hc_csv.log`, `hc_csv_diag.log`

| 見る行 | 正常時の目安 |
|--------|----------------|
| `[UI_FG]` / `[EXCEL_FRONT_FOLLOW] ensure_front_snap` | ld 時 `dlg=QFileDialog` **出ない**（comdlg32 化後） |
| `after_sfw` の `fg_exe` | ld 時 **pythonw.exe にならない** |
| `ensure_front:after_bring_excel` | 既定ポリシーでは **出ない**（bring_excel_first=false） |

---

## 8. 次担当がやってはいけないこと

1. **症状ごとに `ensure_front` タイマーを足す**（破綻の直接原因）。
2. **ld/sv で `QFileDialog.exec` に戻す**。
3. **`prepare` の既定を `bring_excel_first=True` に戻す**（全体退行）。
4. **柱1 と 柱2 を同時に大きくいじる**（切り分け不能になる）。
5. **41 ファイルまとめて 1 コミット**（レビュー・ロールバック不能）。

---

## 9. 次にやること（推奨順）

1. **実機 5 本テスト**（§7.2）の結果を記録。
2. 失敗した機能が 1 つだけなら **その機能だけ**調査（ログ + 該当 `ui_*.py` の `showEvent`）。
3. まだ足りない場合のみ JSON で `TOPMOST: true`（csv_sp パターン）を **その画面だけ**検討。
4. `ui_common.py` / `ui_server.py` の **ヘッダ Version と `__version__` の不整合**を修正。
5. 前面化トピックを **独立ブランチ・独立コミット**に分離。
6. 長期: `showEvent` / `apply_window_config` の遅延 `ensure_front` を監査し、`prepare` + `ensure_front` に集約。

---

## 10. 関連ドキュメント

| ドキュメント | 内容 |
|--------------|------|
| `docs/svc_csv_ld_仕様書.md` | ファイル選択・前面表示の仕様 |
| `docs/ui_common_仕様書.md` | ensure_front API（**ヘッダは旧記述の可能性**—実装は bring_excel_first 既定 false） |
| `docs/incident_csv_ld_ui_20260305.md` | 過去の ld UI 不具合事例 |
| `docs/agent_handoff_UIクリーンアップ_重複レポート_日付進捗_20260502.md` | EXCEL_FRONT_FOLLOW / ensure_front 削除済みウィジェット |
| `docs/environment_variables.md` | `HC_UI_FG_DIAG` 等 |

---

## 11. 会話・判断の要約

- 「退行を追いかけているだけ」というユーザー指摘は正しい。**設計統一が先**。
- B1/B2 切り分け: sp 単独が犯人とは限らない。ui_server 累積とポリシー分裂が主因候補。
- data_agg の Excel 背後は **手順依存**（B2 では前面成功のログもあり）。
- Ask モードで根本分析 → Agent モードで柱1実装 → 他機能退行 → 柱2/3 で `ui_common` 一本化。

---

*作成: 2026-06-29 / 前面化・Z 順退行のエージェント引継ぎ*
