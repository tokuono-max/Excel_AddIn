# エージェント向け資料：更新基盤改善（3件）

対象リポジトリ: CSV Tool / Excel Add-in（`core/packaged_update.py`, `bootstrap/update_bootstrap.py`, `config/ui_update_check.json`, `svc/svc_host.py` 等）

---

## 現状の整理（実装前に読む）

### 更新の判定

- **bin**: `catalog.bin.latest_version` とインストール済み bin を比較 → `needs_bin_update`。対話で Yes 後 `_queue_pending_bin_update`。
- **config**: `check_for_updates` 内でサイレント適用（`needs_config_update`）。
- **bootstrap**: `needs_bootstrap_update` は計算されるが、**OTA では `_queue_pending_bin_update` にしか載らず**、かつ **bin 更新キュー時に catalog の `bootstrap.full` を無条件でコピー**（版の大小は未使用）。bootstrap 単独のキューはない。

### 適用経路

- Excel 再起動後、`svc_host` → `maybe_apply_pending_bootstrap_update` → `bootstrap.update_bootstrap.apply_pending_update`（**hc_main プロセス内**）。
- **差分 zip**: `release/_lib/make_diff_zip.py` は **bsdiff4-manifest-v1**（`manifest.json` + `patches/` + `files/`）。
- **`apply_pending_update` の `_apply_zip`（patch）**は **レガシー型**（展開直後に `app/bin` または `addin` が存在）のみ想定 → manifest 型は `E_PATCH_MANIFEST_INVALID` → **フルへフォールバック**（`bootstrap_update.log` に `fallback_patch_to_full=true reason=E_PATCH_MANIFEST_INVALID`）。

### 既存の正解パス（参照実装）

- `core/packaged_update.py` の **`_materialize_patch_zip_for_worker`**（約 1189 行〜）: manifest 差分を **インストール済みファイルをベースに** materialize し、**レガシー型の一時 zip** を生成。戻り値 `(staged_zip, keep_dir, stats, error)`。

---

## １．bin と bootstrap の更新判定分離

### 目的

- **無駄な bootstrap 入れ替えを避ける**（catalog と同一版ならコピー/swap しない、またはスキップ）。
- **bootstrap だけ新しい**場合も OTA で届けられるようにする（任意だが推奨）。

### 要件（提案）

1. **`_queue_pending_bin_update`（`core/packaged_update.py`）**
   - `catalog.bootstrap` / `bootstrap.full` を解決したうえで、**`needs_bootstrap_update` と同等の判定**（`read_installed_bootstrap_version` vs `catalog.bootstrap.latest_version`）が **偽**なら、`bootstrap.new` のコピーと `pending_swap` を**行わない**（ログに `bootstrap queue skipped reason=already_latest` 等）。
   - **真**のときのみ現状どおりコピー。

2. **`check_for_updates_interactive`**
   - **`needs_bootstrap_update` かつ `not needs_bin_update`** のとき:
     - 専用メッセージ（`config/ui_update_check.json` にテンプレ追加）で **Yes/No**。
     - Yes → **bootstrap 専用 pending** を書く（下記スキーマ）。

3. **`apply_pending_update`（`bootstrap/update_bootstrap.py`）**
   - **`pending` に bin 用の patch/full が無い／スキップフラグ**のときは、bootstrap swap のみ（または swap 後に即完了）し、**bin 適用フェーズをスキップ**。
   - `pending.json` 拡張例（案）:
     - `apply_scope`: `"bin+bootstrap" | "bootstrap_only" | "bin_only"` のいずれか
     - または `skip_bin_apply: bool` + 既存 `mode` / `patch` / `full` の欠如で判定

4. **テスト**: `tests/test_packaged_update_bin_apply.py` 等に、キュー条件・pending 形状のユニットテストを追加。

### 受け入れ条件

- bootstrap 版が catalog と同じで bin だけ更新する場合、**不要な `bootstrap.new` コピーが発生しない**（ログで確認可能）。
- catalog 上 bootstrap のみ新しい環境で、**ユーザーが Yes したときだけ** bootstrap が更新される（単体テストまたは手動シナリオ）。

---

## ２．Excel 再起動後の「更新開始通知」と操作者の可否

### 目的

- 予約済み更新の適用が始まる直前に、**「これから更新を開始する。続行するか？」** を明示し、**No なら当起動では適用しない**（後で再開できるよう pending は残す方針を推奨）。

### 要件（提案）

1. **表示タイミング**: `apply_pending_update` の先頭で `_ProgressUi` を出す**前**、または進捗ウィンドウ表示直後の最初の操作として、**モーダル Yes/No**（tk `messagebox` または既存の `_message_box` 相当が使えるなら統一）。
2. **文言**: `config/ui_update_check.json` にテンプレを追加（例: `PENDING_APPLY_CONFIRM_TITLE` / `PENDING_APPLY_CONFIRM_TEMPLATE`）。内容に **bin 目標版・差分/フル（pending の `mode`）・bootstrap 有無** を含める。
3. **No のとき**:
   - `apply_pending_update` を **中断**し、`pending.json` は削除しない（または `state=deferred` を検討）。
   - ログに `user_decision=pending_apply_no`。
4. **Yes のとき**: 既存どおり bootstrap swap → bin 適用へ。

### 注意

- **二重表示**: `maybe_apply_pending_bootstrap_update` が起動シーケンスで複数回呼ばれないよう、`svc_host` 側の呼び出し回数と整合を取る（既に startup で二重の可能性あり → 調査してガード）。

### 受け入れ条件

- 再起動直後、**進捗バーが動く前**に確認ダイアログが出る。
- No 選択時、**当セッションでは bin/bootstrap のファイル置換が行われない**。

---

## ３．差分作成・適用の改善（１案：manifest 差分を Excel 内適用で解釈する）

### 問題

- `make_diff_zip` 出力（manifest 型）が `apply_pending_update` の patch 経路と**非互換**のため、常にフルへフォールバックしている。

### 改善案（推奨）

1. **`_materialize_patch_zip_for_worker` のロジックを共通化**
   - `core/packaged_update.py` から **`materialize_manifest_patch_zip` のような公開関数**（または `core/patch_manifest.py` に移動）に切り出し。
   - `apply_pending_update` の **`mode == "patch"`** 分岐で、**`_apply_zip` に入る前に**:
     - `patch_zip` を materialize し、**得られたレガシー zip パス**を `_apply_zip` に渡す（または `_apply_zip` 内の先頭で分岐）。
   - **クリーンアップ**: materialize が作った `keep_dir` は適用完了後に `shutil.rmtree`（`finally` で確実に）。

2. **依存関係**
   - 実行バイナリ（hc_main）に **bsdiff4** がバンドルされているか確認。無い場合は Nuitka 依存に追加するか、**stdlib のみのフォールバックは不可**のためビルド手順を更新。

3. **ログ**
   - `bootstrap_update.log` に `patch materialize ok stats=...` または `materialize failed err=...` を出す。

4. **回帰テスト**
   - 小さな fixture zip（manifest 最小）で materialize → 一時 zip に `app/bin` が含まれること。

### 受け入れ条件

- 実機で `bootstrap_update.log` に **`apply_mode_final=patch`** が記録されるケースがある（差分が catalog にあり、インストールが `from_*` 範囲内のとき）。
- 従来どおり manifest が無いレガシー patch zip は **そのまま** `_apply_zip` で動く。

---

## 参照ファイル一覧

| 領域 | パス |
|------|------|
| 更新確認・キュー | `core/packaged_update.py` |
| 適用・進捗 UI | `bootstrap/update_bootstrap.py` |
| 文言 | `config/ui_update_check.json` |
| 起動時呼び出し | `svc/svc_host.py`, `core/packaged_update.py`（`maybe_check_updates_on_startup`） |
| 差分生成 | `release/_lib/make_diff_zip.py`, `release/full.bat` |
| pending 永続化 | `core/update_state.py` |

---

## 作業順序の提案

1. **３（差分適用）** — ログ上の問題を直し、patch の意味を復活させる。
2. **２（再起動後確認）** — UX 固定。
3. **１（bin/bootstrap 分離）** — pending スキーマと UI 分岐が増えるため、３で触った `apply_pending_update` とまとめてテストしやすい。

（プロジェクトの優先度に応じて入れ替え可。）

---

*作成日: 2026-05-01（エージェント実装依頼用ドラフト）*
