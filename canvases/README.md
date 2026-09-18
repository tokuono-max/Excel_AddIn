# canvases/

リポジトリ内の Canvas **保管・Git 用**フォルダ。

## 重要（表示されないとき）

Cursor の Canvas プレビューが検出するのは **次の実ディレクトリだけ**です（ジャンクション不可）:

`%USERPROFILE%\.cursor\projects\c-Project-Python-Excel-AddIn\canvases\`

- ここは **実フォルダ**（ジャンクションにしない）
- リポジトリの `Excel_AddIn\canvases\` は Git 用のコピー
- 編集したら両方がずれないよう、必要ならコピーで同期する

## 開き方

1. `Ctrl+Shift+P` → **Open Canvas**
2. **Codebase Improvement Review** を選ぶ（最新の全体一覧）
3. エディタに `.tsx` ソースだけが出たら、Canvas プレビュー側のタブ／カードを開く

## 置くファイル

- `*.canvas.tsx` のみ（サブフォルダ不可）
- `node_modules/`・`*.canvas.status.json` はコミットしない
