# エージェント向け資料：更新運用方針（2026-05-09 合意）

対象: `core/packaged_update.py`, `bootstrap/update_bootstrap.py`, `installer/CSV_Tool_Setup.iss`, `docs/インストールと運用（利用者・運用向け）.md`

---

## 1. この文書の目的

本書は、更新基盤の実装・運用で迷いやすい論点を「今回の合意内容」で固定するためのハンドオフ資料である。  
特に以下を固定する:

- インストールモード継承（Current user / All users）
- `bin/config/bootstrap` の更新差異
- `min_bin_version` の具体的な扱い
- full zip バックアップ世代管理
- 復旧方式（半自動）

---

## 2. 合意済みポリシー（実装前提）

### 2.1 インストールモードと更新

- 更新は**初回インストールモードを継承**する。
- 更新処理中にモード変更はしない。
- モード変更が必要な場合は**再インストール**で対応する。
- モード識別は **Uninstall キー（`{AppId}_is1`）配下の独自値 `InstallScope`** を参照する（新規キーは増やさない）。
- `InstallScope` の **書き込みは通常側のみ**（`Software\Microsoft\Windows\CurrentVersion\Uninstall\{AppId}_is1`）。**`Software\WOW6432Node\...` には書かない**。参照側 `core/packaged_update.py` の `_resolve_install_scope` は通常側を優先で読むため、通常側だけで必要十分。
- 書き込みタイミングはインストーラ `[Code]` の `ssPostInstall` 後（`WriteInstallScopeToUninstallKey`）。`[Registry]` セクションでは書かない（Inno の Uninstall キー再生成で値が消えるため）。
- インストーラ `[Code]` の `ssInstall` で反対モード残骸（`_is1` キー / `HC_*` 環境変数）を `CleanupCrossModeRemnants` が掃除する。Current user 起動時の HKLM 残骸は権限不足のため警告ログのみ。
- DisplayName はモード別: Current user → `CSV Tool (User)`、All users → `CSV Tool (All User)`。
- `InstallScope=current` のときは更新時に管理者確認ダイアログ/UAC を出さない。
- `InstallScope=all` のときは更新時に管理者確認ダイアログを表示し、承認時のみ UAC 経路に進む。

### 2.2 操作者の操作量

- 方針: **操作者の操作は最小化し、必要情報だけ通知する**。
- `config` は自動適用（失敗時のみ通知）。
- `bin` は必要時のみ確認ダイアログ。
- 失敗時は「原因要約 + ログパス」を通知し、詳細はログ参照。

### 2.3 カタログ解決順

- 優先順は現行維持:  
  `HC_CATALOG_PATH` → `config/catalog_path.txt` → `HC_DEPLOY_ROOT\catalog.json`

### 2.4 起動時チェックの排他

- 起動時に `pending` が存在する場合は、**予約適用フローのみ実行**する。
- 同一起動内で通常のバージョンチェック（`catalog` 比較）は実行しない。
- `pending` が defer された場合も、同一起動内では通常チェックを再開しない（次回起動で再評価）。

---

## 3. 更新系の役割分担（実装準拠）

- `bin` 更新:
  - `patch` を優先
  - 不可/失敗時に `full` へフォールバック
- `config` 更新:
  - `config.payload`（full）をサイレント適用
- `bootstrap` 更新:
  - `bootstrap.full` を予約し、次回起動時に自己差し替え

---

## 4. `min_bin_version` の具体仕様

`config` 更新では、`catalog.config.min_bin_version` が設定されている場合、インストール済み `bin` が下限未満なら適用しない。

判定イメージ（`core/packaged_update.py`）:

- `installed_bin < min_bin_version` → `config` 更新スキップ
- `installed_bin >= min_bin_version` → `config` 更新候補

具体例:

1) スキップされる例
- installed `bin`: `1.0.6`
- `min_bin_version`: `1.0.7`
- 結果: `config` は適用しない（`bin` 更新を先行）

2) 適用される境界例
- installed `bin`: `1.0.7`
- `min_bin_version`: `1.0.7`
- 結果: 条件を満たすため `config` 適用対象

注: `min_bin_version` は「低すぎる bin を弾くための互換下限」である。

---

## 5. full zip バックアップ運用（今回合意）

### 5.1 運用ルール

- 更新時に、復旧用として full zip をバックアップ保存する。
- 保存先は **`HC_INSTALL_ROOT\update\archive\full\`** に固定する。
- バックアップ名は **版ベース単一保持**（`full_prev_<version>.zip`）とする。
- 「更新成功」は **`bin + bootstrap` の双方成功**と定義し、成功時に古いバックアップを削除する。
- 実効的には「直近バックアップを常に 1 世代保持」する運用となる。

### 5.2 初回インストール時の扱い

- **初回インストーラーへの修正は不要**（今回スコープ外）。
- 初回は「戻すべき前世代」が無いため、バックアップ未作成でよい。
- バックアップ運用は「初回後の最初の更新」から開始する。

### 5.3 保持先（推奨）

- `HC_INSTALL_ROOT\update\archive\full\`
- メタ情報: `retain.json`
  - `previous_version`
  - `zip_path`
  - `sha256`
  - `created_at`

---

## 6. 復旧方式（今回合意）

方針は**半自動**。

- 自動で行う:
  - 復旧候補 zip 特定
  - 復旧可否の確認表示
  - 復旧実行フロー開始
- 操作者が行う:
  - 最終確認（Yes/No）
- 完全自動ロールバックは現時点では採用しない。

理由:

- 失敗要因（権限・ロック・環境差）のばらつきが大きく、誤自動復旧リスクがあるため。

---

## 7. 実装時の確認チェックリスト

1. 更新でモード変更を発生させていないか（継承のみか）。
2. Program Files 配下で昇格分岐が期待どおりか。
3. `min_bin_version` 未満で `config` が必ずスキップされるか。
4. 更新成功後にバックアップ世代が 1 世代へ収束するか。
5. 復旧導線が半自動（確認あり）で動作するか。
6. UAC 拒否時は「非昇格で続行」せず、今回適用を defer して次回再確認になるか。

### 7.1 受け入れテスト（最小）

標準セットは **uac-strict-7** を採用する:

1. Current user インストールで更新成功
2. All users インストールで更新成功
3. `min_bin_version` 未満で config 更新がスキップ
4. bin 適用失敗時に半自動復旧ダイアログを即時表示
5. Program Files 配下で UAC 経路を確認
6. patch 失敗時に full フォールバック
7. 起動時排他（pending 優先）と defer 時の通常チェック抑止を確認

---

## 8. 非目標（今回やらない）

- 初回インストーラーの挙動変更
- 完全自動ロールバック
- カタログ優先順の変更

