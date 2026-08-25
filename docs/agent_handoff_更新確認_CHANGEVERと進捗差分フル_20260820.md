# エージェント向け資料：更新確認 CHANGEVER と進捗の差分／フル表示（2026-08-20）

対象: `core/changever.py`, `core/packaged_update.py`, `ui_qt/ui_update_check.py`, `hc_updater.py`, `config/ui_update_check.json`

## 版（このリリースで配るもの）

| 系統 | 版 | 今回配るか |
|------|-----|------------|
| **APL（bin）** | **1.1.9.4**（`VERSION.txt`） | **配る** |
| **config** | catalog の `config.latest_version` を **現行より上げる** | **配る**（`ui_update_check.json` の窓 600×360） |
| **bootstrap** | **1.0.8**（`BOOTSTRAP_VERSION.txt`） | **配らない**（1 番のバックアップ廃止は 1.1.8.4 / 1.0.8 で済み） |

`catalog.bin.latest_version` = `1.1.9.4`。`catalog.bootstrap.latest_version` は現場の **1.0.8 のまま**。

---

## 1. 目的

- 更新確認（すぐに更新／後で）に、配布ルートの **`CHANGEVER.txt`** から差分履歴を出す。
- Excel 終了後の進捗（`hc_updater`）で、**差分更新中／フル更新中** と本文の文頭を使い分ける。

自動更新のバックアップ廃止（1 番）は **対象外**（済）。

---

## 2. 操作者に見える動き

### 2.1 確認 UI（APL）

起動時またはリボン「更新確認」。窓 **600×360**（新 config 適用後）、本文はスクロール。1 行は幅まかせ（日本語おおよそ 40 字）。

```
新しいバージョンがあります。

お使いの版: 1.1.9.4
新しい版: （次の APL 版）

変更内容:
（次の APL 版）
- …
```

- 出すのは **今の版より新しく、配布版以下** の節だけ。
- 履歴ブロックは最大 **8 行**（版番号行＋箇条書き。見出し「変更内容:」は含まない）。超えたら `（続きは CHANGEVER.txt）`。
- ファイルが無い／読めないときは、今どおり版番号だけ。
- **bin 確認には `[1.x.x.x]`（APL）の節だけ。** `[bootstrap x.y.z]` は bootstrap 単独確認のときだけ。

### 2.2 いつ見えるか（確認計画）

`1.1.8.4` → `1.1.9.4` の **その 1 回では、履歴も新進捗も出ない。**

| 何 | 1.1.8.4 → 1.1.9.4 | 1.1.9.4 が入ったあとの **次の** APL 更新 |
|--|--|--|
| 確認の履歴 | 出ない（読む APL が旧） | 出る（CHANGEVER.txt があれば） |
| 確認窓 600×360 | config が先に入れば大きくなるが、本文は旧 UI | 新 UI |
| 進捗の差分／フル | 出ない（適用するのは旧 `hc_updater`） | 出る |

別 PC で履歴・進捗を見るには、**1.1.9.4 を入れたあと**、catalog の bin を仮にさらに上げて 2 回目の更新を走らせる。

### 2.3 進捗 UI（`hc_updater`、Excel 終了後）

Excel 終了待ち・完了案内は従来どおり。  
「更新中」だった区間だけ分岐する。

| | 状態 | 本文 |
|--|------|------|
| 差分 | 差分更新中 | 差分 更新処理を開始しています。 など |
| フル | フル更新中 | フル 更新ファイルを取得しています。 など |
| 差分・適用 | 差分更新中 | 差分パッケージをインストールしています |
| フル・適用 | フル更新中 | フルパッケージをインストールしています |

差分からフルへ切り替わったあとはフル側。準備画面（bootstrap の「差分パッケージを構築しています」）は従来どおり。

---

## 3. 配布でやること（別 PC）

### 3.1 `CHANGEVER.txt` はビルドでは作らない

Nuitka / `pack.bat` / インストーラは **`CHANGEVER.txt` を自動生成しない。** zip 作りとは別に、配布前に手で用意する。

- `installer/CHANGEVER.sample.txt` は **リポジトリ内の雛形**だけ。APL はこれを読まない。
- 配布物は **`catalog.json` と同じフォルダの `CHANGEVER.txt`**（UTF-8）。zip の中には入れない。
- 作り方: 雛形をコピーして `CHANGEVER.txt` に改名するか、同じ形式で新規作成する。
- 以降のリリースでは、同じファイルの **先頭に新しい `[版]` を足す**（古い節は残してよい）。
- 雛形への追記はチーム用の控えであり、配布の必須作業ではない。

### 3.2 手順

1. **APL `1.1.9.4`** をビルドし、`catalog.bin.latest_version` を `1.1.9.4` にする。
2. **config zip を作り直す**（中に新しい `ui_update_check.json`）。`catalog.config.latest_version` を現行より上げる。`config.min_bin_version` は `1.1.9.4` 以下（推奨: `1.1.9.4`。新 JSON を旧 APL に先に入れない）。
3. **bootstrap zip は触らない。** `BOOTSTRAP_VERSION.txt` は **1.0.8**。catalog の bootstrap 版も 1.0.8 のまま。
4. **§3.1 のとおり `CHANGEVER.txt` を手で用意し**、`catalog.json` と同じフォルダに置く。
5. catalog に次を足す（省略時は同フォルダの `CHANGEVER.txt` を探す）。

```json
"release_notes": {
  "relative_path": "CHANGEVER.txt"
}
```

`bin.release_notes_url` は今も **未使用**。履歴は `release_notes.relative_path` / `CHANGEVER.txt` だけ。

6. `CHANGEVER.txt` の書き方（APL 節のみでよい）:

```
[1.1.9.4]
- 更新確認に変更内容を表示
- 進捗で差分／フルを区別
```

bootstrap を将来上げるときだけ `[bootstrap 1.0.x]` を足す。

---

## 4. 変更ファイル

| ファイル | 内容 |
|----------|------|
| `VERSION.txt` | `1.1.9.4` |
| `BOOTSTRAP_VERSION.txt` | `1.0.8`（上げない） |
| `core/changever.py` | 読み取り・節解析・8 行整形 |
| `core/packaged_update.py` | 確認文面に履歴ブロックを連結 |
| `ui_qt/ui_update_check.py` | CONFIRM をスクロール可能な本文に |
| `config/ui_update_check.json` | 窓 600×360、進捗の差分／フルキー |
| `hc_updater.py` | `updater_busy_title` / `updater_busy_body` |
| `installer/catalog.sample.json` | `release_notes.relative_path` |
| `installer/CHANGEVER.sample.txt` | 雛形 |
| `tests/test_changever.py` | 解析・範囲・省略 |
| `tests/test_hc_updater_progress_text.py` | 進捗文言 |
| `docs/インストールと運用（利用者・運用向け）.md` | catalog の `release_notes` / CHANGEVER |

---

## 5. テスト

```
python -m pytest tests/test_changever.py tests/test_hc_updater_progress_text.py -q
```

---

## 6. 実装上の注意

- **`CHANGEVER.txt` はビルド成果物ではない。** 配布前に手で用意し、catalog と同じフォルダへ置く（§3.1）。
- 履歴は **zip 内ではなく配布ルート**。更新前の APL が次の更新で読む。
- 確認 UI は APL。進捗の差分／フルは `hc_updater.exe`（bin）。窓サイズは **config**。
- 進捗文言は JSON が古くても `hc_updater` 内のフォールバックで出る。窓 600×360 は **新 config が必要**。
- `{changelog}` を format に載せると `CHANGEVER.txt` 内の `{` で壊れるので、本文の後ろに連結している。
