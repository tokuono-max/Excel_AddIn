# エージェント向け資料：更新確認 VER_HISTORY と進捗の差分／フル表示（2026-08-20／2026-08-29 更新）

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

- 更新確認（すぐに更新／後で）に、**`config/ui_help.json` の `VER_HISTORY`** から差分履歴を出す。
- ヘルプの **「変更履歴」** 副画面で、同 JSON の記載すべてをいつでも参照できる。
- Excel 終了後の進捗（`hc_updater`）で、**差分更新中／フル更新中** と本文の文頭を使い分ける。

**`CHANGEVER.txt` / catalog `release_notes` は廃止**（残っていても読まない）。

---

## 2. 操作者に見える動き

### 2.1 ヘルプ → 変更履歴

1. リボン「ヘルプ」
2. 「変更履歴」→ ヘルプ前面にスクロール専用副画面（タイトル「変更履歴」）
3. 「戻る」で副画面だけ閉じ、ヘルプに戻る

表示内容は `VER_HISTORY` の **BIN + BOOTSTRAP すべて**。

### 2.2 更新確認（差分）

起動時またはリボン「更新確認」。新しい版があるとき本文末尾に:

```
変更内容:
1.1.9.4
- …
```

- 出すのは **今の版より新しく、配布版以下** の節だけ（`installed < section <= latest`）。
- 正本はインストール済み **`config/ui_help.json`**（config 更新で届く）。
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
| `core/packaged_update.py` | 更新確認を VER_HISTORY に一本化 |
| `tests/test_changever.py` | JSON ベースのテスト |
| （削除）`CHANGEVER.txt` / `installer/CHANGEVER.sample.txt` | 廃止 |

```
python -m pytest tests/test_changever.py tests/test_hc_updater_progress_text.py -q
```
