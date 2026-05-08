# VBA–Python 実行権とシートプロパティ解析

## 1. シートプロパティの一覧

**Python 側で参照・設定しているシートのカスタムプロパティ（CustomProperties）は次の 4 種類のみ。**

| キー名 | 定義箇所 | 設定者 | 用途 |
|--------|----------|--------|------|
| **HC_BOOK_NAME** | core_stat.KEY_BOOK_NAME | svc_csv_ld, hc_csv_ld（set_prop） | ブック名（Excel で一意とみなす名前） |
| **HC_GUID_B64** | core_stat.KEY_GUID_B64 | core_stat.set_guid / 各 svc | シート識別用の一意値（Base64） |
| **HC_STATUS_INFO** | core_stat.KEY_STATUS_INFO | core_stat.set_status_info / 各 svc | ステータスバーに表示する文字列 |
| **HC_NOTIFY_RETV** | core_stat.KEY_NOTIFY_RETV | core_stat.set_notify_retv（特別な時のみ呼び出し） | Python 復帰後に VBA の MsgBox で表示する値（空なら表示しない） |

※ 仕様書等で **HC_NOTIFY_REV** と書かれている場合がありますが、実装上のキー名は **HC_NOTIFY_RETV**（RETV）です。

**上記以外のシートプロパティは、本プロジェクトの Python コードでは使用していません。**  
（core_cst の VBA_NOTIFY_NAME / INFO_NAME は上記キー名の別名であり、core_xlc の get/set_sheet_prop はキー名を引数で受け取る汎用 API です。）

---

## 2. Python から VBA に実行権が戻るタイミング

### 2.1 呼び出し経路

1. **VBA**  
   `Main.RunPythonSafe("load_csv", sId)`  
   → `RunPython "import hc_main; hc_main.invoke(action='load_csv', target_hwnd=..., sheet_id='...')"`
2. **xlwings**  
   `RunPython` は別プロセスで上記 Python コマンドを実行し、**そのプロセスが終了するまで VBA は待機**（同期）。
3. **hc_main**  
   `load_csv()` → `_call_svc_server(ACTION_CSV_LD, ...)`  
   - 依頼内容を pickle で書き出し、  
   - `res_path` ができるまで `time.sleep(0.05)` を繰り返してポーリング、  
   - 結果を読んだら return。
4. **実行権が VBA に戻るタイミング**  
   **RunPython で起動した Python プロセスが終了した直後**。  
   つまり `hc_main.invoke(...)`（load_csv 相当の action）が return し、プロセスが終了した時点で `RunPython` が return し、続けて VBA で `CheckAndNotifyVBA(sId)` と `HC_Bridge.RestoreStatBar` が実行される。

### 2.2 実装上のポイント

- 実処理（ファイル読込・UI など）は **svc_server 側の別プロセス**で実行される。
- VBA が起動するプロセスは **依頼の書き出し → 結果ファイルのポーリング待ち** が主で、`time.sleep(0.05)` で OS に制御を返すため、その間にある程度 Excel が応答しやすくなる設計になっている。
- 「実行権を VBA に戻す」＝ **VBA がブロックしている RunPython が return する瞬間**＝上記 Python プロセス終了時、と解釈してよい。

### 2.3 Excel が OS から無応答とみなされる可能性

**現状の方式では、Excel が OS から「無応答」と判定される可能性はある。**

- **理由**  
  VBA の `RunPython` は xlwings 側で **`WScript.Shell.Run(..., WaitOnReturn:=True)`** により **子プロセス（Python）の終了まで同期待ち** している。  
  その間、**VBA の実行はブロック**され、Excel のメインスレッドは **Windows のメッセージ（描画・入力等）を処理しない**。
- **OS の挙動**  
  メインスレッドが数秒以上メッセージを処理しないと、OS がアプリを「応答なし」と判断し、タイトルバーに「応答していません」と出すことがある。
- **いつ起こりうるか**  
  Python 側の処理（ファイル選択のためダイアログを開いている時間や、読込・保存の実行時間）が長いほど、上記の「数秒」を超えやすく、無応答とみなされやすい。
- **設計意図との関係**  
  「Python に飛んだ後、実行権をいったん VBA に戻し、Python 機能は別に起動する」という設計意図であれば、**RunPython を同期で待たず、非同期で起動してすぐ VBA に戻る**方式にしないと、無応答リスクはなくならない。現状は「VBA が Python プロセス終了まで待つ」ため、実行権が戻るのは **処理がすべて終わった後** である。

### 2.4 Python 側のみでの対策（早期復帰）

**VBA／xlwings を変更せず、Python 側だけで無応答を抑える対応を入れている。**

- **HC_RETURN_EARLY=1（既定）**  
  `hc_main._call_svc_server` で、依頼を `svc_req_*.pkl` に書き出したあと、**結果ファイル (res_*.pkl) を待たずに約 1 秒待って return する**。  
  これにより **Python プロセスが早く終了**し、RunPython が return して VBA に実行権が戻るため、Excel が OS から無応答とみなされにくくなる。
- **実処理**  
  依頼は従来どおり `svc_server` が別プロセスで受け取り、処理・進捗・完了通知はすべて `svc_server` および UI 側で行う。ユーザーから見た動作は変わらない。
- **従来どおり結果を待ちたい場合**  
  環境変数 `HC_RETURN_EARLY=0` にすると、従来どおり結果が返るまで待つ（Python 側で例外も受け取れる）。テストやデバッグ時に利用可能。

### 2.5 コード上の対応関係

| 段階 | ファイル・処理 |
|------|----------------|
| VBA が Python を起動して待機 | Main.bas `RunPython sCmd` |
| Python でエントリ実行 | hc_main.invoke(action='load_csv', …) |
| 依頼送信・結果待ち | hc_main._call_svc_server()（res_path ができるまでループ） |
| 実処理 | svc/svc_server → svc_csv_ld 等（別プロセス） |
| Python プロセス終了 | load_csv() return → プロセス終了 |
| VBA に制御復帰 | RunPython が return → 次行の CheckAndNotifyVBA 等が実行 |
| 早期復帰時 | hc_main は依頼書き出し＋約 1 秒後に return。svc_server が別プロセスで処理・UI 表示 |

---

## 3. Python が HC_NOTIFY_RETV を設定する仕様と現状

### 3.1 設計仕様

**通常は Python 側では HC_NOTIFY_RETV を設定しない。特別な時だけ設定する。**

- VBA は RunPython 復帰後に **CheckAndNotifyVBA** で `HC_NOTIFY_RETV` を参照し、**値があれば** MsgBox 表示、**値がなければ**何もしない（パス）。
- したがって、Python が「特別な時」だけ `set_notify_retv` を呼べばよく、毎回設定する必要はない。

### 3.2 実装（修正済み）

**通常の完了時には set_notify_retv を呼ばないように修正済み。**

| ファイル | 対応 |
|----------|------|
| **svc/svc_csv_ld.py** | 読込完了時の `set_notify_retv` 呼び出しを削除 |
| **svc/svc_csv_sv.py** | 保存完了時の `set_notify_retv` 呼び出しを削除 |
| **svc/hc_csv_ld.py** | 読込完了時の `set_notify_retv` 呼び出しを削除 |
| **svc/hc_csv_sv.py** | 保存完了時の `set_notify_retv` 呼び出しを削除 |

特別な時（エラー通知・ユーザー設定で VBA MsgBox を出したい場合など）のみ、呼び出し元で `core_stat.set_notify_retv` を呼ぶ形とする。

---

## 4. 他シートプロパティの意図した使用法

| プロパティ | 意図した用途 | 設定箇所 | 参照箇所 | 確認結果 |
|------------|--------------|----------|----------|----------|
| **HC_STATUS_INFO** | ステータスバーに表示する文字列。VBA の RestoreStatBar 等で参照される。 | csv_ld, csv_sp, csv_mg, hd_in, hd_nr, undo, hc_csv_ld, hc_csv_sv（処理中・完了・エラー時） | csv_mg（既存値に追記）, undo（復元前の状態取得）, VBA 側でステータスバー復元 | 意図どおり。処理状態・完了・エラーメッセージを一貫して set_status_info で保存している。 |
| **HC_GUID_B64** | シートを一意に識別するための Base64 GUID。 | csv_ld, hc_csv_ld（新規シート作成時、未設定のときのみ set_guid） | core_xlc.find_sheet_by_guid, csv_ld/hc_csv_ld（get_guid=="" で未設定判定） | 意図どおり。新規シートにのみ刻印し、VBA から渡された sheet_id でシートを特定する際に利用される。 |
| **HC_BOOK_NAME** | ブック名（Excel で一意とみなす名前）の保存。 | csv_ld, hc_csv_ld（新規シート作成時、set_prop(..., KEY_BOOK_NAME, book.name)） | （Python 側では参照していない。VBA 等で必要に応じて参照可能） | 意図どおり。core_stat.KEY_BOOK_NAME でキー名を統一済み。 |

上記以外のシートプロパティは本プロジェクトの Python コードでは使用していない。

---

**共通仕様（機能面）**: 本解析の「仕様」としてのまとめは **docs/共通仕様_機能.md** に記載する。

---

## 5. 参照コード位置

| 項目 | ファイル | 行目目安 |
|------|----------|----------|
| シートプロパティ定数 | core/core_stat.py | KEY_* 42–45（KEY_BOOK_NAME 含む） |
| set_prop / get_prop | core/core_stat.py | 98–116, 163–196 |
| CustomProperties 読書 | core/core_xlc.py | get_sheet_prop 53–62, set_sheet_prop 65–75 |
| KEY_BOOK_NAME 設定 | svc/svc_csv_ld.py, hc_csv_ld.py | 新規シート作成時 |
| set_notify_retv（特別時のみ） | core_stat.set_notify_retv。通常完了時は各 svc からは呼ばない | - |
| VBA 実行権戻り後の処理 | VBA/Main.bas | RunPythonSafe 内 RunPython の次: CheckAndNotifyVBA, RestoreStatBar |
| HC_NOTIFY_RETV 参照・MsgBox・削除 | VBA/Main.bas | CheckAndNotifyVBA 347–396 |
