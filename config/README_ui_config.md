# 機能別画面設定ファイル（外部のみ・救済なし）

このフォルダの JSON が**唯一の設定源**です。core_cst の画面定義は参照しません。

- **共通仕様は 2 種類**。**JSON の共通仕様**（キー名・構造・ルール）は [docs/共通仕様_JSON定義.md](../docs/共通仕様_JSON定義.md)、**機能面の共通仕様**（シートプロパティ・VBA 実行権復帰等）は [docs/共通仕様_機能.md](../docs/共通仕様_機能.md) を参照すること。
- **共通モジュール（core_cst, ui_common 等）を変更した場合**: [docs/共通モジュール変更時_デグレ防止.md](../docs/共通モジュール変更時_デグレ防止.md) に従い、他モジュール・他機能への影響確認を行うこと（共通仕様_機能.md からも参照）。

## 方針

- **外部だけ**: 画面パラメータはこのファイルのみを参照する。
- **救済しない**: ファイルが無い・JSON が壊れている場合は、**エラー種別を表示して終了**する（フォールバックなし）。

## ファイル名

- **CSV結合**: `ui_csv_mg.json`（必須）
- **CSV読込**: `ui_csv_ld.json`（必須・ダイアログタイトル・フィルタ）
- **CSV保存**: `ui_csv_sv.json`（必須・ダイアログタイトル・フィルタ）
- **CSV分割**: `ui_csv_sp.json`（必須・分割画面・進捗・完了・ワーニング）
- **ウィンドウ遅延（全機能共通）**: `ui_window_timing.json`（任意。無い・不正時は `core.ui_window_timing` のコード既定。各キーに `_comment` で意味・既定を記載）
- 将来の機能: `ui_<機能キー小文字>.json`

## 配置場所

- プロジェクトルートの `config/` フォルダ  
  例: `C:\Project\Python\Excel_AddIn\config\ui_csv_mg.json`

## ファイルヘッダ（任意）

ルートに **`_header`** を置くと、メタ情報として記録できます。実行時には参照されません。

| 項目 | 内容 |
|------|------|
| File | ファイル名（例: ui_csv_mg.json） |
| Function | 機能画面名（例: CSV結合） |
| Created | 作成日 |
| Updated | 更新日 |
| Setting_details | 設定項目および設定可能な値の一覧（配列または文字列） |

## 画面定義（WINDOW）

各機能の JSON で **WINDOW** を指定すると、ダイアログのサイズ・位置・タスクバー表示などを制御できます。  
（CSV結合は MAIN / COMMON と WINDOW をマージして適用。CSV読込・CSV保存は WINDOW をそのまま参照。）

| キー | 説明 | 例 |
|------|------|-----|
| **DEFAULT_WIDTH** | 初期幅（ピクセル）。0 は未指定（OS 既定）。 | 600 |
| **DEFAULT_HEIGHT** | 初期高さ（ピクセル）。0 は未指定。 | 400 |
| **RESIZABLE** | ユーザーがリサイズできるか。 | true / false |
| **CENTER_ON_EXCEL** | Excel ウィンドウを基準に中央に配置するか。 | true / false |
| **SHOW_IN_TASKBAR** | タスクバーにアイコンを出すか（false で Excel の子として非表示）。 | true / false |
| **SHOW_MINIMIZE** | 最小化ボタンを表示するか。 | true / false |
| **SHOW_MAXIMIZE** | 最大化ボタンを表示するか。 | true / false |
| **STARTUP_POSITION** | 起動位置。`center` / `remember_last` 等。 | "center" |
| **TOPMOST** | `WindowStaysOnTopHint` で最前面表示（他アプリの上に出やすい）。 | true / false |
| **EXCEL_FRONT_FOLLOW** | Excel が前景のときにダイアログを前面へ追従するか（`start_front_follow`）。 | true / false |

- **CSV読込（ui_csv_ld）**: DEFAULT_WIDTH / DEFAULT_HEIGHT を 0 より大きくすると、ファイル選択ダイアログのサイズを変更できます。指定時は Qt 描画のダイアログを使用します（OS ネイティブの場合はサイズは反映されません）。
- **CSV結合（ui_csv_mg）**: WINDOW はトップレベル・COMMON・MAIN の順でマージされ、結合画面・進捗・完了・重複確認に適用されます。

## 書き方

- `ui_csv_mg.json` は **WINDOW / COMMON / MAIN** および **MAIN.SCREENS**（PROGRESS, DONE, DUPLICATE）を含む**フル定義**が必要です。
- `ui_csv_ld.json` / `ui_csv_sv.json` は **WINDOW**（任意）と **MAIN**（TITLE, FILTER）、**SCREENS.PROGRESS** を定義します。
- `ui_csv_sp.json` は **WINDOW / MAIN**（TITLE, DESC, TABLE, DIALOG_BUTTONS, DIALOGS.FOLDER）および **SCREENS**（PROGRESS, DONE, WARNING）を定義します。
- キー名・構造の詳細は **docs/共通仕様_JSON定義.md** にまとめています。
- 編集・保存後、**次にその画面を開いたとき**に反映されます（Excel の終了は不要です）。

## エラー時

- **設定ファイルが見つかりません**: パスを確認してください。
- **設定ファイルの形式が正しくありません（JSON エラー）**: 行・列とメッセージを確認し、カンマ・括弧・引用符を修正してください。
- **設定ファイルの読み込みに失敗しました**: ファイルの権限・文字コード（UTF-8）を確認してください。
- いずれもダイアログでエラー種別を表示したうえで処理を終了し、救済は行いません。
