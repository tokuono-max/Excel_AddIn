リリース用バッチ（簡易ガイド）

[前提]
- リポジトリ直下で実行される想定です。
- VERSION.txt の先頭行が既定のリリース版になります（形式: X.Y.Z.N / 例: 1.2.3.4）。
- pack 生成時は X.Y.Z.N を自動分解し、bin=3桁(X.Y.Z)・config=1桁(N)・set_version=X.Y.Z.N を生成します。
- 生成先は dist\releases\<版>\ です。

[ファイル一覧]
- full.bat   : 薄いラッパ → 正本は tools\release\full.bat（ASCII+CRLF、git 管理）
- nuitka.bat : nuitka ビルド（full / bridge / svc / ui / runner）
- pack.bat   : full・bin full・cfg zip・catalog.json を生成
- diff.bat   : 旧 bin full と比較して差分 zip を生成
- cfg.bat    : config 単独更新向けの生成（bin full は既存を流用）

[文字化け防止（重要）]
- `.bat` 本文は **ASCII のみ + CRLF**。日本語の echo は入れない（cmd / Cursor UTF-8 で必ず壊れる）。
- 正本: `tools\release\full.bat`。再生成: `python tools\release\_write_full_bat.py`
- 検証: `python tools\release\assert_bat_ascii_crlf.py`
- 従来どおり `release\full.bat` でも起動できる（ラッパが正本を呼ぶ）。

[引数早見表（1行表）]
- `nuitka.bat [Mode]` : `Mode=full|bridge|svc|ui|runner`（省略時 `full`）
- `pack.bat [ReleaseVersion] [BinVersion] [ConfigVersion]` : 省略時は `VERSION.txt` を自動分解（Bin/Config 手動指定は互換用途）
- `diff.bat [ReleaseVersion] [OldBinFullZip]` : 省略時は `ReleaseVersion=VERSION.txt`、`OldBinFullZip=dist\releases` から自動選択
- `cfg.bat <ReleaseVersion> [BinLatestVersion] [BaseBinFullZip]` : 第1引数必須、`BaseBinFullZip` 省略時は `dist\releases` から自動選択
- `full.bat [ReleaseVersion] [OldBinFullZip]` : `nuitka full -> pack -> diff` を順に実行（省略時は `VERSION.txt` と自動選択）。起動時に `NUITKA_JOBS` と `require_uninstall_reinstall`（ON/OFF、Enter=OFF）を対話入力する。非対話は `CSV_TOOL_FULL_NONINTERACTIVE=1`（フラグは `CSV_TOOL_REQUIRE_UNINSTALL_REINSTALL=on` で上書き、未指定は OFF）

[引数仕様（重要）]
1) nuitka.bat [Mode]
   - Mode: full / bridge / svc / ui / runner
   - 省略時: full

2) pack.bat [ReleaseVersion] [BinVersion] [ConfigVersion]
   - ReleaseVersion: 出力フォルダ名・zip名の版
   - BinVersion    : （互換用途）bin_<ver>_full.zip 内 VERSION.txt に書く版
   - ConfigVersion : （互換用途）cfg_<ver>.zip 内 config\VERSION.txt に書く版
   - 省略時:
     - ReleaseVersion 省略 -> ルート VERSION.txt
     - BinVersion 省略     -> ReleaseVersion の先頭3桁（X.Y.Z）
     - ConfigVersion 省略  -> ReleaseVersion の4桁目（N）

3) diff.bat [ReleaseVersion] [OldBinFullZip]
   - ReleaseVersion: 差分生成先の版（省略時は VERSION.txt）
   - OldBinFullZip : 比較元 bin_*_full.zip のフルパス
   - 省略時:
     - OldBinFullZip 省略 -> dist\releases 配下から最新の bin_*_full.zip を自動選択

4) cfg.bat <ReleaseVersion> [BinLatestVersion] [BaseBinFullZip]
   - ReleaseVersion : 必須（config 更新版）
   - BinLatestVersion: catalog.json の bin.latest_version に設定する値
   - BaseBinFullZip : 流用する既存 bin_*_full.zip のフルパス
   - 省略時:
     - BinLatestVersion 省略 -> BaseBinFullZip 内の版（または既定値）を使用
     - BaseBinFullZip 省略  -> dist\releases 配下から最新の bin_*_full.zip を自動選択

5) full.bat [ReleaseVersion] [OldBinFullZip]
   - 起動時: NUITKA_JOBS（空 Enter は既定）
   - 1/3 nuitka.bat full
   - 2/3 pack.bat [ReleaseVersion]
   - 3/3 diff.bat [ReleaseVersion] [OldBinFullZip]
   - 完了後対話: require_uninstall_reinstall（ON / OFF、空 Enter は OFF）を catalog.json へ書く
   - 非対話: CSV_TOOL_FULL_NONINTERACTIVE=1（フラグ既定 OFF）
   - 省略時:
     - ReleaseVersion 省略 -> VERSION.txt
     - OldBinFullZip 省略 -> 自動選択

[よく使う実行例]
1) フル一括（通常）
   release\full.bat

2) 版を指定してフル一括
   release\full.bat 1.2.3.4

3) pack のみ（VERSION.txt の版で作成）
   release\pack.bat

4) pack のみ（版を明示）
   release\pack.bat 1.2.3.4

5) pack のみ（bin と config の内部版を分ける）
   release\pack.bat 1.2.3.5 1.2.3.4 1.2.3.5

6) 差分のみ（旧版 zip を明示）
   release\diff.bat 1.2.3.5 "C:\releases\bin_1.2.3.4_full.zip"

7) config 単独（bin は 1.2.3 を維持、config を 5 へ）
   release\cfg.bat 1.2.3.5 1.2.3 "C:\releases\bin_1.2.3_full.zip"

8) config 単独（bin zip 自動選択）
   release\cfg.bat 1.2.3.5 1.2.3

[運用のコツ]
- 引数に空白を含むパス（例: C:\work folder\...）は必ず "..." で囲ってください。
- `.xlam` を配布物に含める場合は、リポジトリ直下 `addin\` に配置してください（`nuitka.bat` 実行時に `dist\CSV_Tool\addin\` へ自動反映）。
- まず `release\pack.bat` で full/bin/cfg/catalog を作り、その後 `release\diff.bat` を実行すると切り分けしやすいです。
- `cfg.bat` は ReleaseVersion（第1引数）が必須です。省略するとエラー終了します。
- `diff.bat` は差分が無い場合、patch zip を作らずに終了します（異常ではありません）。
- `diff.bat` は `release\_lib\make_diff_zip.py`（`bsdiff4`）を使います。実行 Python は原則 `.\.venv\Scripts\python.exe`（無ければ `python`）です。
- Nuitka の `--remove-output` は既定で **OFF**（`*.build` 削除のタイミング失敗を回避）。有効化する場合だけ `set NUITKA_REMOVE_OUTPUT=1` を付けて実行します。
- 既定 OFF でも、各ビルド後に `tools\nuitka\nuitka_flatten_dist_into_parent.bat` / `merge_nuitka_stage_into_bin.bat` で出力直下の **`*.build` を削除してから** `app\bin` へマージするため、`bin_*_full.zip` に中間生成物が混ざりません。

[出力]
- bootstrap_<ver>_full.zip
- bin_<ver>_full.zip
- cfg_<ver>.zip
- （差分がある場合）bin_<old>_<new>_d.zip
- catalog.json

[差分 zip と実行時]
- `diff.bat` が作る **bin_<旧>_<新>_d.zip** は `catalog.json` の **`bin.patch`** に登録される。
- patch zip には `manifest.json`（`patch_format=bsdiff4-manifest-v1`）が入り、`.exe/.dll/.pyd` は bsdiff パッチ、その他はコピーで格納される。
- **エンドユーザ PC**では `core.packaged_update` が **インストール版が `from_min_version`～`from_max_version` に入るとき `bin.patch` を優先**し、使えない場合は **`bin.full`** にフォールバックして適用する（Excel 全終了後の PowerShell ワーカー）。
- **初回 Inno がどの zip を展開するか**／**各 zip がどの経路で使われるか**の一覧は **`docs\インストールと運用（利用者・運用向け）.md` §4.0.1** を参照。

[catalog.json 反映時の注意]
- 共有に置く際は、この `catalog.json` をそのまま利用できます。
- relative_path と実ファイル配置を一致させてください。
- zip の中身を変更した場合は sha256 の再計算が必要です。（ファイル名変更は対象外）

[共有置き場の例]
- 共有ルート（例）: \\mcom\oec1\work\H05095_小野\releases\CSV_Tool
- 配置イメージ:
  - catalog.json
  - releases\<版>\full_<版>.zip
  - releases\<版>\bin_<版>_full.zip
  - releases\<版>\cfg_<版>.zip
  - releases\<版>\bin_<旧>_<新>_d.zip （差分がある場合）

※ catalog.json の relative_path は上記実ファイル配置と必ず一致させてください。
