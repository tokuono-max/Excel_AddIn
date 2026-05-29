# 開発PC用: `HC_*` 環境切替バッチ

配布確認と通常開発で **`HC_INSTALL_ROOT` / `HC_PACKAGED_DEPLOYMENT` を切り替えやすくする**ため、このフォルダのバッチが **現在の cmd ウィンドウにだけ**変数をセット／クリアする。

**Excel は起動しない。** 同じ PC でスタートメニューやショートカットから Excel を開くと、そのプロセスは **直前にこのバッチで整えた環境**を引き継ぐ（**新規起動の Excel** に限る）。

- **`start_excel.bat`**（引数なし）… メニューで **`1`**＝配布（既定 **`C:\Program Files\Excel_Addin\CSV_Tool`**）、**`2`**＝開発。**`start_excel 1 "別パス"`** でルート上書き。**`start_excel packaged`** / **`start_excel dev`** も可。
- **`start_excel_packaged_test.bat`** … **`HC_PACKAGED_DEPLOYMENT=1`** と **`HC_INSTALL_ROOT`** をセット。既定ルートは **`C:\Program Files\Excel_Addin\CSV_Tool`**。ローカル **`dist\CSV_Tool`** で試すときは **`set CSV_TOOL_PACKAGED_ROOT=dist\CSV_Tool`** を付けてから実行するか、第1引数にパスを渡す。

運用方針の全体像は **`docs/Exe化（開発者向け）.md` セクション 2.6** を参照。

**手動で配布フォルダを作り Nuitka 成果物をコピーする手順**（コピー範囲・数百 MB の扱い）は同書 **セクション 2.7** を参照。

## 使い方（cmd.exe）

リポジトリルートで:

```bat
call tools\dev\start_excel_dev.bat
```

または

```bat
call tools\dev\start_excel_packaged_test.bat
```

その **同じ黒いウィンドウ** で、続けてスタートメニューから Excel を起動する。**別のバッチから呼ぶときは必ず `call`**（`call` なしだと、呼び出し元に環境が戻らない場合がある）。

## CMD と PowerShell

| シェル | 説明 |
|--------|------|
| **コマンドプロンプト（`cmd.exe`）** | **推奨**。上記のとおり **`call tools\dev\....bat`** で **`HC_*` がこのウィンドウに残る**。 |
| **PowerShell** | **`.\tools\dev\....bat` だけ**だと **別プロセスの cmd** が一瞬動いて終わるため、**PowerShell の `$env:HC_*` には通常残らない**。環境をセットした **cmd を開いたまま**使うなら、例: **`cmd /k "cd /d C:\Project\Python\Excel_AddIn && call tools\dev\start_excel_dev.bat"`** のあと、その **残った cmd** から Excel を起動する、または **`cmd.exe` を直接開いて**同じ手順を実行する。 |

**補足**

- バッチ先頭の **`cd /d "%~dp0..\.."`** により、**内部でリポジトリルートへ移動**する。引数の相対パス（例: `dist\CSV_Tool`）は **そのリポジトリルート基準**で解決される。
- **`.bat` を PowerShell の構文として書かない**（`set`・`if` などは cmd 用）。

## 前提

- **`setlocal` は使わない**ため、**同じ cmd で `call` したあと**、各バッチ末尾の **`--- Verification ---`** で **`HC_INSTALL_ROOT`** / **`HC_PACKAGED_DEPLOYMENT`** を表示する（期待値は行の上に併記）。必要なら続けて **`set HC`** で一覧してもよい。
- **このウィンドウから後から起動する Excel** が、ここで整えた環境を継承する。既に起動済みの Excel には影響しない。
- **配布ツリー**は `docs/Exe化（開発者向け）.md` のセクション 5 に沿った構成（**`app\bin\hc_main.exe`** 等、**VBA ソース `*.bas` / `*.cls` は Shift-JIS（CP932）で保存**）。**`start_excel_packaged_test.bat`** は **`hc_main.exe`** が無いとエラーにする。**`xlwings.conf`** の **`INTERPRETER_*`** は別途、配布／開発に合わせる。

## `start_excel_dev.bat`

- リポジトリルートへ **`cd`** し、**`HC_INSTALL_ROOT` / `HC_PACKAGED_DEPLOYMENT` を空にする**。
- **Excel は起動しない。**

## `repair_venv_python_dlls.bat`

- **`.venv\Scripts\pythonw.exe`** が **`python312.dll` が見つからない`** と出るとき、ベース Python（`pyvenv.cfg` の `home`）から **`python312.dll` / `python3.dll`** を `.venv\Scripts` にコピーする。
- xlwings は RunPython 時に **`cd` で Scripts に移ってから `pythonw` を起動**するため、DLL が Scripts に無いと失敗することがある（`python.exe` だけでは再現しない場合あり）。
- 例: `call tools\dev\repair_venv_python_dlls.bat`

## `start_excel.bat`

- **引数なし** … `choice` で **1＝配布** / **2＝開発**（いずれも **`HC_*` のみ変更**）。
- **`start_excel 1`** / **`start_excel 2`** … メニューなし。
- **`start_excel packaged`** / **`start_excel dev`** も可。

## `start_excel_packaged_test.bat`

- **第1引数を省略**すると **`C:\Program Files\Excel_Addin\CSV_Tool`**（既定）。**`dist\CSV_Tool`** を既定にしたいときは **`set CSV_TOOL_PACKAGED_ROOT=dist\CSV_Tool`**（リポジトリルート基準の相対パス可）。
- 第1引数に **インストールルート**を渡すこともできる。例:  
  `call tools\dev\start_excel_packaged_test.bat "dist\CSV_Tool"`  
  `call tools\dev\start_excel_packaged_test.bat "C:\Temp\CSV_Tool"`
- **`HC_PACKAGED_DEPLOYMENT=1`** と **`HC_INSTALL_ROOT`** をセット。**Excel は起動しない。**
- **`%HC_INSTALL_ROOT%\app\bin\hc_main.exe`** が無い場合はエラー終了。
- 文字化けする場合は **UTF-8（BOM なし）または ANSI・CRLF** で保存し直す。

## 環境変数の確認（`HC_*`）

### 同じ cmd で

```bat
set HC
```

### PowerShell（**その PowerShell プロセス**）

バッチでセットした値は **通常ここには載らない**（上表参照）。手で触る場合:

```powershell
Get-ChildItem Env: | Where-Object { $_.Name -like 'HC_*' }
```

### Excel 内（**この cmd で環境を整えたあとに起動した Excel**）

```vb
?Environ("HC_INSTALL_ROOT")
?Environ("HC_PACKAGED_DEPLOYMENT")
```

### ユーザー／マシンに永続登録した値（`setx`・インストーラ）

```powershell
[Environment]::GetEnvironmentVariable("HC_INSTALL_ROOT", "User")
[Environment]::GetEnvironmentVariable("HC_INSTALL_ROOT", "Machine")
```

## 注意

- **モードを切り替えたいときは**、まず **Excel を終了**し、このバッチで **`HC_*` を変えたうえで** Excel を **もう一度起動**すると分かりやすい。
- 配布本番PCでの設定は **インストーラ**側（`docs/Exe化（開発者向け）.md` ・ `docs/インストーラ化（開発者向け）.md` ・ `docs/インストールと運用（利用者・運用向け）.md`）が担当する想定。
