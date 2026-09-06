# エージェント向け資料：更新確認 VER_HISTORY と進捗の差分／フル表示（2026-08-20／2026-08-29／2026-09-06 更新）

対象: `core/changever.py`, `core/packaged_update.py`, `ui_qt/ui_help.py`, `ui_qt/ui_update_check.py`, `hc_updater.py`, `config/ui_help.json`, `config/ui_update_check.json`

## 版（このリリースで配るもの）

| 系統 | 版 | 今回配るか |
|------|-----|------------|
| **APL（bin）** | **1.1.9.4**（`VERSION.txt`） | **配る** |
| **config** | catalog の `config.latest_version` を **現行より上げる** | **配る**（`ui_help.json` の `VER_HISTORY`・変更履歴 UI、`ui_update_check.json` の窓 600×480） |
| **bootstrap** | **1.0.9**（`BOOTSTRAP_VERSION.txt`） | **配らない** |

`catalog.bin.latest_version` = `1.1.9.4`。

---

## 1. 目的

- **ヘルプ「変更履歴」**: インストール済み **`config/ui_help.json` の `VER_HISTORY` すべて**。
- **更新確認**（すぐに更新／後で）: **catalog の `config.payload`（`cfg_*.zip`）内の `ui_help.json`** を適用せず読み、**今の 4 桁セットより新しく、catalog の `set_version` 以下**の節だけ出す。zip が読めなければ版番号のみ。確認画面の「お使いの版／新しい版」も **4 桁セット**。
- 確認ダイアログは **Qt が本線**。出せないときだけ Win32 MessageBox（はい／いいえ）。Qt が出たあとに重ねない。
- Excel 終了後の進捗（`hc_updater`）で、**差分更新中／フル更新中** と本文の文頭を使い分ける。

**`CHANGEVER.txt` / catalog `release_notes` は廃止**（残っていても読まない）。

この改修は **改修入りのクライアント（1.1.10.6 以降）から次の版への確認**で効く。

---

## 2. 操作者に見える動き

### 2.1 ヘルプ → 変更履歴

1. リボン「ヘルプ」
2. 「変更履歴」→ ヘルプ前面にスクロール専用副画面（タイトル「変更履歴」）
3. 「戻る」で副画面だけ閉じ、ヘルプに戻る

表示内容はインストール済み `ui_help.json` の `VER_HISTORY`（**BIN + BOOTSTRAP すべて**）。

### 2.2 更新確認（次版）

起動時またはリボン「更新確認」。新しい版があるとき:

```
お使いの版: 1.1.9.5
新しい版: 1.1.10.6

変更内容:
1.1.10.6
- …
```

- 出すのは **今のセット版より新しく、行き先 `set_version` 以下** の節だけ（`installed_set < section <= latest_set`）。
- 正本は **配布 cfg zip 内の `ui_help.json`**（インストール済み JSON は使わない）。
- **bin 確認には BIN 節だけ。** BOOTSTRAP は bootstrap 単独確認のときだけ。

### 2.3 進捗 UI（`hc_updater`）

差分／フルの進捗文言分岐は従来どおり（履歴テキストは出さない）。

---

## 3. 配布

1. リポジトリの **`config/ui_help.json`** の `VER_HISTORY` を編集（新しい版を配列先頭へ）。
2. **config zip を作り直し**、`catalog.config.latest_version` を上げる。
3. 共有への `CHANGEVER.txt` 手置きは不要。

---

## 4. 変更ファイル（履歴 UI 関連）

| ファイル | 内容 |
|----------|------|
| `config/ui_help.json` | `VER_HISTORY` / `SCREENS.VER_HISTORY` / ヘルプの「変更履歴」ボタン |
| `core/changever.py` | JSON 読取・差分整形・閲覧整形 |
| `ui_qt/ui_help.py` | 変更履歴副画面 |
| `core/packaged_update.py` | 更新確認は cfg zip の VER_HISTORY。版表示は 4 桁セット。Qt ready 後は Win32 に落とさない |
| `tests/test_changever.py` | JSON ベースのテスト |
| （削除）`CHANGEVER.txt` / `installer/CHANGEVER.sample.txt` | 廃止 |

```
python -m pytest tests/test_changever.py tests/test_packaged_update_ver_history_cfg_zip.py tests/test_hc_updater_progress_text.py -q
```
