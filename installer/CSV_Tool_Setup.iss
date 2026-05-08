; =============================================================================
; CSV Tool — 薄いインストーラ（Inno Setup 6.1 以降推奨）
; =============================================================================
; 【役割】共有（UNC またはローカル）の配布ルートにある catalog.json を読み、
;         初回導入の bin/config/bootstrap 各 zip を決定して {app} へ展開する。
;         HKCU\Environment の HC_*、{app}\xlwings.conf を設定する。本体は EXE に同梱しない。
; 【詳細ドキュメント】docs\インストーラ化（開発者向け）.md §2 ビルド手順
;
; 【ビルド（推奨）】リポジトリから:
;   installer\build_csv_tool_setup.bat
;   installer\build_csv_tool_setup.bat "\\サーバー\共有\CSV_Tool"
; 第 1 引数で SHAREPAYLOAD を指定すると、.iss 内の既定値を上書きしてコンパイルされる。
;
; 【ビルド（手動）】Inno の ISCC から:
;   "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" CSV_Tool_Setup.iss
;   "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" /DSHAREPAYLOAD=\\server\share\CSV_Tool CSV_Tool_Setup.iss
;
; 【埋め込みされる値】コンパイル時に確定するもの:
;   - SHAREPAYLOAD … 配布ルート（catalog.json がある場所）。互換で ...\current 指定も可
;   - DefaultDirName … ウィザードの「既定のインストール先」（ユーザーは通常、別パスに変更可能）
;   変更したら必ず再コンパイルすること。
;
; 【前提】エンドユーザーの PC から、インストール時に SHAREPAYLOAD が参照できること（VPN・UNC・権限）。
;
; 【NTFS ACL】ssPostInstall で payload 展開後、{app} に Authenticated Users (S-1-5-11) の
;   Modify (OI)(CI) 継承を /T で付与する。一般ユーザーが logs・config 書き込みと
;   packaged_update の install_root 直下一時フォルダ作成に必要（Program Files 既定 ACL では拒否される）。
;
; 【HC_DEPLOY_ROOT】catalog.json が置かれた配布ルート。
;   既定は SHAREPAYLOAD と同一、ただし互換で SHAREPAYLOAD が ...\current の場合は親を採用。
;
; 【HKCU と UAC】[Registry] は HKCU。別ユーザーで昇格のみすると Environment が意図したプロファイルに付かない
;   ことがある。可能なら実ユーザーの操作でセットアップを実行すること（Inno: UsedUserAreasWarning）。
; =============================================================================

#define MyAppName "CSV Tool"
; リポジトリルートの VERSION.txt と揃えること
#define MyAppVersion "1.0.4"
#define MyDefaultInstallDir "{autopf}\Excel_Addin\CSV_Tool"

; 配布ルート（catalog.json 配置先）。互換で ...\current を指定してもよい
; ビルド時に上書き: ISCC の /DSHAREPAYLOAD=... または build_csv_tool_setup.bat の第 1 引数
#ifndef SHAREPAYLOAD
#define SHAREPAYLOAD "\\mcom\oec1\work\H05095_小野\releases\CSV_Tool"
#endif

[Setup]
; 製品識別子（再インストール・アップグレード判定に使用。製品ごとに固定し変更しない）
AppId={{B5E8C4A2-1D3F-4E6B-9C0D-1A2B3C4D5E6F}
AppName={#MyAppName}
AppPublisher=大井電気(株)
AppVersion={#MyAppVersion}

; ウィザードで最初に表示される既定インストール先（x64 の Program Files 系 + 相対パス）
DefaultDirName={code:GetDefaultInstallDir}
DefaultGroupName={#MyAppName}
; 生成 EXE の出力先（この .iss から見た相対: リポジトリの dist\）
OutputDir=..\dist
OutputBaseFilename=CSV_Tool_Setup
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
; 既定の Program Files 配下などに入れるため管理者権限を要求
PrivilegesRequired=admin
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
DisableProgramGroupPage=yes
; HKCU\Environment を更新したことを OS に通知し、新規プロセスで反映されやすくする
ChangesEnvironment=yes
UninstallDisplayIcon={app}\app\bin\hc_main.ico
; 薄いインストーラは [Files] でペイロードを持たないため、初回は {app} が空でも警告しない
UsePreviousAppDir=no

[Languages]
Name: "japanese"; MessagesFile: "compiler:Languages\Japanese.isl"

[Registry]
; ユーザー環境変数（§1.3–1.4・docs\environment_variables.md）
Root: HKCU; Subkey: "Environment"; ValueType: string; ValueName: "HC_INSTALL_ROOT"; ValueData: "{app}"; Flags: uninsdeletevalue
Root: HKCU; Subkey: "Environment"; ValueType: string; ValueName: "HC_PACKAGED_DEPLOYMENT"; ValueData: "1"; Flags: uninsdeletevalue
; 配布ルート（catalog.json の既定参照先）。値の算出は [Code] GetDeployRoot
Root: HKCU; Subkey: "Environment"; ValueType: string; ValueName: "HC_DEPLOY_ROOT"; ValueData: "{code:GetDeployRoot}"; Flags: uninsdeletevalue

[Icons]
Name: "{group}\Uninstall {#MyAppName}"; Filename: "{uninstallexe}"

[UninstallDelete]
; zip 展開で配置したため [Files] 管理外。アンインストール時に {app} ごと削除
Type: filesandordirs; Name: "{app}"

[Code]
{ 配布ルート解決:
  - SHAREPAYLOAD 直下に catalog.json があればそれを採用
  - 互換: SHAREPAYLOAD が ...\current の場合は親フォルダを採用 }

var
  InstallProgressPage: TOutputProgressWizardPage;
  SetupIniLoaded: Boolean;
  SetupIniPathCache: String;
  SetupIniDeployRoot: String;
  SetupIniInstallDir: String;

function TrimAndNormalizePath(const S: String): String;
begin
  Result := Trim(S);
  while (Length(Result) > 0) and ((Result[Length(Result)] = '\') or (Result[Length(Result)] = '/')) do
    SetLength(Result, Length(Result) - 1);
end;

function SetupIniPath: String;
begin
  if SetupIniPathCache <> '' then
  begin
    Result := SetupIniPathCache;
    Exit;
  end;
  SetupIniPathCache := AddBackslash(ExtractFileDir(ExpandConstant('{srcexe}'))) + 'setup.ini';
  Result := SetupIniPathCache;
end;

procedure EnsureSetupIniLoaded;
var
  IniPath: String;
begin
  if SetupIniLoaded then
    Exit;
  SetupIniLoaded := True;

  IniPath := SetupIniPath;
  if not FileExists(IniPath) then
  begin
    Log('setup.ini not found; using defaults. path=' + IniPath);
    Exit;
  end;

  SetupIniDeployRoot := TrimAndNormalizePath(GetIniString('Setup', 'DeployRoot', '', IniPath));
  SetupIniInstallDir := Trim(GetIniString('Setup', 'InstallDir', '', IniPath));

  if SetupIniDeployRoot <> '' then
    Log('setup.ini DeployRoot override detected: ' + SetupIniDeployRoot)
  else
    Log('setup.ini DeployRoot is empty; fallback to built-in default');

  if SetupIniInstallDir <> '' then
    Log('setup.ini InstallDir override detected: ' + SetupIniInstallDir)
  else
    Log('setup.ini InstallDir is empty; fallback to built-in default');
end;

function GetDefaultInstallDir(Param: String): String;
begin
  EnsureSetupIniLoaded;
  if SetupIniInstallDir <> '' then
    Result := SetupIniInstallDir
  else
    Result := ExpandConstant('{#MyDefaultInstallDir}');
end;

procedure ProgressUpdate(const Status: String; Position, Max: Integer);
begin
  if InstallProgressPage <> nil then
  begin
    InstallProgressPage.SetText('初回導入 payload を適用中です。しばらくお待ちください。', Status);
    InstallProgressPage.SetProgress(Position, Max);
  end;
end;

function GetDeployRoot(Param: String): String;
var
  S, C: String;
begin
  EnsureSetupIniLoaded;
  if SetupIniDeployRoot <> '' then
  begin
    Result := SetupIniDeployRoot;
    Exit;
  end;

  S := ExpandConstant('{#SHAREPAYLOAD}');
  while (Length(S) > 0) and ((S[Length(S)] = '\\') or (S[Length(S)] = '/')) do
    SetLength(S, Length(S) - 1);
  C := AddBackslash(S) + 'catalog.json';
  if FileExists(C) then
    Result := S
  else
    Result := ExtractFileDir(S);
end;

function RobocopyOk(Code: Integer): Boolean;
begin
  { robocopy 終了コード: 0〜7 は成功扱い（意味は公式ドキュメント参照）、8 以上は失敗 }
  Result := (Code >= 0) and (Code < 8);
end;

function FindTextFrom(const S, Token: String; StartPos: Integer): Integer;
var
  Tail: String;
  P: Integer;
begin
  if StartPos < 1 then
    StartPos := 1;
  Tail := Copy(S, StartPos, MaxInt);
  P := Pos(Token, Tail);
  if P > 0 then
    Result := StartPos + P - 1
  else
    Result := 0;
end;

function JsonExtractStringValueNear(const JsonText, KeyToken: String; StartPos: Integer): String;
var
  KeyPos, ColonPos, Q1, Q2: Integer;
begin
  Result := '';
  KeyPos := FindTextFrom(JsonText, KeyToken, StartPos);
  if KeyPos = 0 then
    Exit;
  ColonPos := FindTextFrom(JsonText, ':', KeyPos + 1);
  if ColonPos = 0 then
    Exit;
  Q1 := FindTextFrom(JsonText, '"', ColonPos + 1);
  if Q1 = 0 then
    Exit;
  Q2 := FindTextFrom(JsonText, '"', Q1 + 1);
  if (Q2 = 0) or (Q2 <= Q1) then
    Exit;
  Result := Trim(Copy(JsonText, Q1 + 1, Q2 - Q1 - 1));
end;

function JsonExtractBinLatestVersion(const JsonText: String): String;
var
  BinPos: Integer;
begin
  Result := '';
  BinPos := FindTextFrom(JsonText, '"bin"', 1);
  if BinPos = 0 then
    Exit;
  Result := JsonExtractStringValueNear(JsonText, '"latest_version"', BinPos);
end;

function JsonExtractConfigLatestVersion(const JsonText: String): String;
var
  ConfigPos: Integer;
begin
  Result := '';
  ConfigPos := FindTextFrom(JsonText, '"config"', 1);
  if ConfigPos = 0 then
    Exit;
  Result := JsonExtractStringValueNear(JsonText, '"latest_version"', ConfigPos);
end;

function JsonExtractBootstrapLatestVersion(const JsonText: String): String;
var
  BootstrapPos: Integer;
begin
  Result := '';
  BootstrapPos := FindTextFrom(JsonText, '"bootstrap"', 1);
  if BootstrapPos = 0 then
    Exit;
  Result := JsonExtractStringValueNear(JsonText, '"latest_version"', BootstrapPos);
end;

function JsonExtractSetVersion(const JsonText: String): String;
begin
  Result := JsonExtractStringValueNear(JsonText, '"set_version"', 1);
end;

function JsonExtractBinFullRelativePath(const JsonText: String): String;
var
  FullPos, BinPos: Integer;
begin
  Result := '';
  BinPos := FindTextFrom(JsonText, '"bin"', 1);
  if BinPos > 0 then
  begin
    FullPos := FindTextFrom(JsonText, '"full"', BinPos);
    if FullPos > 0 then
      Result := JsonExtractStringValueNear(JsonText, '"relative_path"', FullPos);
  end;
end;

function JsonExtractConfigPayloadRelativePath(const JsonText: String): String;
var
  ConfigPos, PayloadPos: Integer;
begin
  Result := '';
  ConfigPos := FindTextFrom(JsonText, '"config"', 1);
  if ConfigPos > 0 then
  begin
    PayloadPos := FindTextFrom(JsonText, '"payload"', ConfigPos);
    if PayloadPos > 0 then
      Result := JsonExtractStringValueNear(JsonText, '"relative_path"', PayloadPos);
  end;
end;

function JsonExtractBootstrapFullRelativePath(const JsonText: String): String;
var
  BootstrapPos, FullPos: Integer;
begin
  Result := '';
  BootstrapPos := FindTextFrom(JsonText, '"bootstrap"', 1);
  if BootstrapPos > 0 then
  begin
    FullPos := FindTextFrom(JsonText, '"full"', BootstrapPos);
    if FullPos > 0 then
      Result := JsonExtractStringValueNear(JsonText, '"relative_path"', FullPos);
  end;
end;

function PathIsAbsolute(const P: String): Boolean;
begin
  Result :=
    ((Length(P) >= 2) and (P[2] = ':')) or
    ((Length(P) >= 2) and (P[1] = '\\') and (P[2] = '\\'));
end;

function ResolvePayloadPath(const DeployRoot, RelativeOrAbs: String): String;
begin
  if PathIsAbsolute(RelativeOrAbs) then
    Result := RelativeOrAbs
  else
    Result := AddBackslash(DeployRoot) + RelativeOrAbs;
end;

function PsQuote(const S: String): String;
begin
  { 実運用の Windows パスでは単一引用符は稀なため、単純に '...' で囲む }
  Result := '''' + S + '''';
end;

function ExpandZipToTemp(const ZipPath, TempDir: String): Boolean;
var
  Cmd, Params: String;
  Rc: Integer;
begin
  Cmd := ExpandConstant('{sys}\WindowsPowerShell\v1.0\powershell.exe');
  Params := '-NoProfile -ExecutionPolicy Bypass -Command "Expand-Archive -LiteralPath ' +
    PsQuote(ZipPath) + ' -DestinationPath ' + PsQuote(TempDir) + ' -Force"';
  Result := Exec(Cmd, Params, '', SW_HIDE, ewWaitUntilTerminated, Rc) and (Rc = 0);
end;

function FindExpandedRootByMarker(const TempDir, MarkerRelativePath: String): String;
var
  Candidate: String;
begin
  Result := '';
  Candidate := AddBackslash(TempDir) + MarkerRelativePath;
  if FileExists(Candidate) then
  begin
    Result := TempDir;
    Exit;
  end;
  Candidate := AddBackslash(TempDir) + 'current\' + MarkerRelativePath;
  if FileExists(Candidate) then
  begin
    Result := AddBackslash(TempDir) + 'current';
    Exit;
  end;
end;

procedure DeployPayloadZip(
  const PayloadName, PayloadZip, Dest, MarkerRelativePath, TempSuffix: String;
  const StepExtract, StepVerify, StepCopy, StepDone, StepMax: Integer
);
var
  TempDir, SrcRoot, Cmd, Params: String;
  ResultCode: Integer;
begin
  TempDir := AddBackslash(ExpandConstant('{tmp}')) + 'csv_tool_setup_' + TempSuffix;
  if DirExists(TempDir) then
    DelTree(TempDir, True, True, True);
  if not ForceDirectories(TempDir) then
    RaiseException('作業フォルダ作成に失敗しました: ' + TempDir);

  try
    Log('initial_deploy: ' + PayloadName + ' start zip=' + PayloadZip);
    ProgressUpdate(PayloadName + ' zip を展開中...', StepExtract, StepMax);
    if not ExpandZipToTemp(PayloadZip, TempDir) then
      RaiseException(PayloadName + ' payload zip の展開に失敗しました: ' + PayloadZip);

    ProgressUpdate(PayloadName + ' 展開結果を検証中...', StepVerify, StepMax);
    SrcRoot := FindExpandedRootByMarker(TempDir, MarkerRelativePath);
    if SrcRoot = '' then
      RaiseException(PayloadName + ' 展開結果に ' + MarkerRelativePath + ' が見つかりません。');

    ProgressUpdate(PayloadName + ' を {app} へコピー中...', StepCopy, StepMax);
    Cmd := ExpandConstant('{cmd}');
    Params := '/c robocopy "' + SrcRoot + '" "' + Dest + '" /E /COPY:DAT /DCOPY:DAT /R:2 /W:2 /NFL /NDL /NJH /NJS /np';
    if not Exec(Cmd, Params, '', SW_HIDE, ewWaitUntilTerminated, ResultCode) then
      RaiseException(PayloadName + ' robocopy を起動できませんでした。');
    if not RobocopyOk(ResultCode) then
      RaiseException(PayloadName + ' robocopy 失敗 (exit ' + IntToStr(ResultCode) + ')');
    ProgressUpdate(PayloadName + ' のコピー完了を確認中...', StepDone, StepMax);
    Log('initial_deploy: ' + PayloadName + ' ok');
  finally
    DelTree(TempDir, True, True, True);
  end;
end;

procedure DeployByCatalog(const DeployRoot, Dest: String);
var
  CatalogPath, BinRel, ConfigRel, BootstrapRel: String;
  BinZip, ConfigZip, BootstrapZip: String;
  Raw: AnsiString;
begin
  ProgressUpdate('catalog.json を読み取り中...', 1, 14);
  CatalogPath := AddBackslash(DeployRoot) + 'catalog.json';
  if not LoadStringFromFile(CatalogPath, Raw) then
    RaiseException('catalog.json を読み取れません: ' + CatalogPath);

  ProgressUpdate('導入対象 payload を判定中...', 2, 14);
  BinRel := JsonExtractBinFullRelativePath(String(Raw));
  ConfigRel := JsonExtractConfigPayloadRelativePath(String(Raw));
  BootstrapRel := JsonExtractBootstrapFullRelativePath(String(Raw));

  if BinRel = '' then
    RaiseException('catalog.json に bin.full.relative_path がありません。');
  if ConfigRel = '' then
    RaiseException('catalog.json に config.payload.relative_path がありません。');
  if BootstrapRel = '' then
    RaiseException('catalog.json に bootstrap.full.relative_path がありません。');

  BinZip := ResolvePayloadPath(DeployRoot, BinRel);
  ConfigZip := ResolvePayloadPath(DeployRoot, ConfigRel);
  BootstrapZip := ResolvePayloadPath(DeployRoot, BootstrapRel);

  if not FileExists(BinZip) then
    RaiseException('catalog が指す bin zip が見つかりません: ' + BinZip);
  if not FileExists(ConfigZip) then
    RaiseException('catalog が指す config zip が見つかりません: ' + ConfigZip);
  if not FileExists(BootstrapZip) then
    RaiseException('catalog が指す bootstrap zip が見つかりません: ' + BootstrapZip);

  DeployPayloadZip('bin', BinZip, Dest, 'app\bin\hc_main.exe', 'payload_bin', 3, 4, 5, 6, 14);
  DeployPayloadZip('config', ConfigZip, Dest, 'config\VERSION.txt', 'payload_config', 7, 8, 9, 10, 14);
  DeployPayloadZip('bootstrap', BootstrapZip, Dest, 'bootstrap\update_bootstrap.exe', 'payload_bootstrap', 11, 12, 13, 14, 14);
end;

function NormalizeDisplayVersion(const RawVersion: String): String;
var
  I, StartPos, Count: Integer;
  Part: String;
begin
  Result := '';
  StartPos := 1;
  Count := 0;
  for I := 1 to Length(RawVersion) + 1 do
  begin
    if (I > Length(RawVersion)) or (RawVersion[I] = '.') then
    begin
      Part := Trim(Copy(RawVersion, StartPos, I - StartPos));
      StartPos := I + 1;
      if Part = '' then
        Continue;
      if Result <> '' then
        Result := Result + '.';
      Result := Result + Part;
      Count := Count + 1;
      if Count >= 3 then
        Break;
    end;
  end;
end;

function NormalizeSetVersion(const RawVersion: String): String;
var
  I, StartPos, Count: Integer;
  Part: String;
begin
  Result := '';
  StartPos := 1;
  Count := 0;
  for I := 1 to Length(RawVersion) + 1 do
  begin
    if (I > Length(RawVersion)) or (RawVersion[I] = '.') then
    begin
      Part := Trim(Copy(RawVersion, StartPos, I - StartPos));
      StartPos := I + 1;
      if Part = '' then
      begin
        Result := '';
        Exit;
      end;
      if StrToIntDef(Part, -1) < 0 then
      begin
        Result := '';
        Exit;
      end;
      if Result <> '' then
        Result := Result + '.';
      Result := Result + IntToStr(StrToIntDef(Part, 0));
      Count := Count + 1;
    end;
  end;
  if Count <> 4 then
    Result := '';
end;

function NormalizeBootstrapVersion(const RawVersion: String): String;
var
  I, StartPos, Count: Integer;
  Part: String;
begin
  Result := '';
  StartPos := 1;
  Count := 0;
  for I := 1 to Length(RawVersion) + 1 do
  begin
    if (I > Length(RawVersion)) or (RawVersion[I] = '.') then
    begin
      Part := Trim(Copy(RawVersion, StartPos, I - StartPos));
      StartPos := I + 1;
      if Part = '' then
      begin
        Result := '';
        Exit;
      end;
      if StrToIntDef(Part, -1) < 0 then
      begin
        Result := '';
        Exit;
      end;
      if Result <> '' then
        Result := Result + '.';
      Result := Result + IntToStr(StrToIntDef(Part, 0));
      Count := Count + 1;
    end;
  end;
  if Count <> 3 then
    Result := '';
end;

procedure TrySyncDisplayValues(const RootKey: Integer; const SubKey, DisplayName, DisplayVersion: String; var AnyWritten: Boolean);
begin
  if not RegKeyExists(RootKey, SubKey) then
  begin
    Log('Display sync skipped (key not found): root=' + IntToStr(RootKey) + ' key=' + SubKey);
    Exit;
  end;
  try
    if RegWriteStringValue(RootKey, SubKey, 'DisplayName', DisplayName) then
    begin
      Log('DisplayName synced: root=' + IntToStr(RootKey) + ' key=' + SubKey + ' value=' + DisplayName);
      AnyWritten := True;
    end;
  except
    Log('DisplayName write failed: root=' + IntToStr(RootKey) + ' key=' + SubKey);
  end;
  try
    if RegWriteStringValue(RootKey, SubKey, 'DisplayVersion', DisplayVersion) then
    begin
      Log('DisplayVersion synced: root=' + IntToStr(RootKey) + ' key=' + SubKey + ' value=' + DisplayVersion);
      AnyWritten := True;
    end;
  except
    Log('DisplayVersion write failed: root=' + IntToStr(RootKey) + ' key=' + SubKey);
  end;
end;

procedure SyncDisplayVersionFromCatalog(const DeployRoot: String);
var
  CatalogPath, LatestBin, LatestConfig, SetVersionRaw, DisplayVersion, SubKey, WowSubKey, SelectedKey: String;
  Raw: AnsiString;
  AnyWritten: Boolean;
begin
  CatalogPath := AddBackslash(DeployRoot) + 'catalog.json';
  if not LoadStringFromFile(CatalogPath, Raw) then
  begin
    Log('DisplayVersion sync skipped: catalog.json not readable: ' + CatalogPath);
    Exit;
  end;

  SetVersionRaw := JsonExtractSetVersion(String(Raw));
  DisplayVersion := NormalizeSetVersion(SetVersionRaw);
  if DisplayVersion = '' then
  begin
    LatestBin := JsonExtractBinLatestVersion(String(Raw));
    LatestConfig := JsonExtractConfigLatestVersion(String(Raw));
    if (LatestBin <> '') and (LatestConfig <> '') then
      DisplayVersion := NormalizeSetVersion(NormalizeDisplayVersion(LatestBin) + '.' + LatestConfig);
  end;
  if DisplayVersion = '' then
  begin
    Log('DisplayVersion sync skipped: invalid set_version and fallback compose failed');
    Exit;
  end;

  AnyWritten := False;
  SubKey := 'Software\Microsoft\Windows\CurrentVersion\Uninstall\{B5E8C4A2-1D3F-4E6B-9C0D-1A2B3C4D5E6F}_is1';
  WowSubKey := 'Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{B5E8C4A2-1D3F-4E6B-9C0D-1A2B3C4D5E6F}_is1';

  { Inno が管理するキーを優先して 1 箇所のみ更新する（重複生成を避ける） }
  SelectedKey := '';
  if RegKeyExists(HKLM, SubKey) then
    SelectedKey := SubKey
  else if RegKeyExists(HKLM, WowSubKey) then
    SelectedKey := WowSubKey
  else if RegKeyExists(HKCU, SubKey) then
    SelectedKey := SubKey
  else if RegKeyExists(HKCU, WowSubKey) then
    SelectedKey := WowSubKey;

  if SelectedKey <> '' then
  begin
    if RegKeyExists(HKLM, SelectedKey) then
      TrySyncDisplayValues(HKLM, SelectedKey, 'CSV Tool', DisplayVersion, AnyWritten)
    else if RegKeyExists(HKCU, SelectedKey) then
      TrySyncDisplayValues(HKCU, SelectedKey, 'CSV Tool', DisplayVersion, AnyWritten);
  end;

  if not AnyWritten then
    Log('DisplayName/DisplayVersion sync failed: uninstall key not found or not writable');
end;

procedure SyncBootstrapVersionFromCatalog(const DeployRoot, DestRoot: String);
var
  CatalogPath, BootstrapVersionRaw, BootstrapVersion: String;
  Raw: AnsiString;
  Lines: TArrayOfString;
begin
  CatalogPath := AddBackslash(DeployRoot) + 'catalog.json';
  if not LoadStringFromFile(CatalogPath, Raw) then
  begin
    Log('bootstrap version sync skipped: catalog.json not readable: ' + CatalogPath);
    Exit;
  end;

  BootstrapVersionRaw := JsonExtractBootstrapLatestVersion(String(Raw));
  BootstrapVersion := NormalizeBootstrapVersion(BootstrapVersionRaw);
  if BootstrapVersion = '' then
  begin
    Log('bootstrap version sync skipped: bootstrap.latest_version missing/invalid');
    Exit;
  end;

  if not ForceDirectories(DestRoot + '\bootstrap') then
  begin
    Log('bootstrap version sync failed: bootstrap dir create failed: ' + DestRoot + '\bootstrap');
    Exit;
  end;

  SetArrayLength(Lines, 1);
  Lines[0] := BootstrapVersion;
  SaveStringsToFile(DestRoot + '\bootstrap\VERSION.txt', Lines, False);
  Log('bootstrap version synced: ' + BootstrapVersion + ' path=' + DestRoot + '\bootstrap\VERSION.txt');
end;

procedure TryDeleteUninstallKey(const RootKey: Integer; const SubKey: String);
begin
  if RegKeyExists(RootKey, SubKey) then
  begin
    if RegDeleteKeyIncludingSubkeys(RootKey, SubKey) then
      Log('Uninstall cleanup deleted key: root=' + IntToStr(RootKey) + ' key=' + SubKey)
    else
      Log('Uninstall cleanup failed to delete key: root=' + IntToStr(RootKey) + ' key=' + SubKey);
  end
  else
    Log('Uninstall cleanup skipped (key not found): root=' + IntToStr(RootKey) + ' key=' + SubKey);
end;

procedure TryRemoveEmptyParentDirOfApp;
var
  ParentDir: String;
begin
  ParentDir := ExtractFileDir(ExpandConstant('{app}'));
  if ParentDir = '' then
  begin
    Log('Uninstall parent dir cleanup skipped: empty parent path');
    Exit;
  end;
  if not DirExists(ParentDir) then
  begin
    Log('Uninstall parent dir cleanup skipped: not found: ' + ParentDir);
    Exit;
  end;
  if RemoveDir(ParentDir) then
    Log('Uninstall parent dir cleanup removed: ' + ParentDir)
  else
    Log('Uninstall parent dir cleanup skipped: directory not empty or locked: ' + ParentDir);
end;

function InitializeSetup: Boolean;
var
  P, DeployRoot, CatalogPath: String;
begin
  EnsureSetupIniLoaded;
  if SetupIniDeployRoot = '' then
  begin
    P := ExpandConstant('{#SHAREPAYLOAD}');
    if not DirExists(P) then
    begin
      MsgBox('Payload folder not found or not reachable:'#13#10 + P + #13#10#13#10 +
        'Check VPN, UNC path, and permissions. You can recompile with /DSHAREPAYLOAD=\\server\\share\\...',
        mbError, MB_OK);
      Result := False;
      Exit;
    end;
  end;

  DeployRoot := GetDeployRoot('');
  if (DeployRoot = '') or (not DirExists(DeployRoot)) then
  begin
    MsgBox(
      '配布ルートを解決できませんでした。'#13#10 +
      'DeployRoot=' + DeployRoot + #13#10 +
      'SHAREPAYLOAD=' + ExpandConstant('{#SHAREPAYLOAD}') + #13#10 +
      'setup.ini=' + SetupIniPath,
      mbError, MB_OK);
    Result := False;
    Exit;
  end;

  CatalogPath := AddBackslash(DeployRoot) + 'catalog.json';
  if not FileExists(CatalogPath) then
  begin
    MsgBox('catalog.json が見つかりません。'#13#10 + CatalogPath + #13#10#13#10 +
      'このインストーラは catalog.json を使って初回導入を行います。', mbError, MB_OK);
    Result := False;
    Exit;
  end;

  Result := True;
end;

procedure InitializeWizard;
begin
  InstallProgressPage :=
    CreateOutputProgressPage('CSV Tool セットアップ', 'インストール後半の処理を実行しています。');
end;

// {app}\addin\xlwings.conf を生成（互換のため {app}\xlwings.conf にも出力）
procedure WriteXlwingsConf(const DestRoot: String);
var
  Runner: String;
  AddinDir: String;
  Lines: TArrayOfString;
begin
  Runner := DestRoot + '\app\bin\hc_xlwings_short_runner.exe';
  AddinDir := DestRoot + '\addin';
  if not DirExists(AddinDir) then
    ForceDirectories(AddinDir);
  SetArrayLength(Lines, 10);
  Lines[0] := '# -*- coding: utf-8 -*-';
  Lines[1] := '# Generated by CSV_Tool_Setup installer ({app}-local paths)';
  Lines[2] := '';
  Lines[3] := 'INTERPRETER_WIN = "' + Runner + '"';
  Lines[4] := 'INTERPRETER = "' + Runner + '"';
  Lines[5] := 'USE_UDF_SERVER = False';
  Lines[6] := 'DEBUG_UDFS = False';
  Lines[7] := '';
  Lines[8] := 'USE_PACKAGED_RUNPYTHON = True';
  Lines[9] := '';
  SaveStringsToFile(AddinDir + '\xlwings.conf', Lines, False);
  SaveStringsToFile(DestRoot + '\xlwings.conf', Lines, False);
end;

// Program Files 配下でも一般ユーザーが logs / config / 一時ディレクトリを書けるようにする
procedure GrantAppRootUserWritableAcl(const AppRoot: String);
var
  ResultCode: Integer;
  Params: String;
begin
  if (AppRoot = '') or (not DirExists(AppRoot)) then
  begin
    Log('GrantAppRootUserWritableAcl: skip (empty or missing): ' + AppRoot);
    Exit;
  end;
  { *S-1-5-11 = Authenticated Users（ロケール非依存）。Modify + オブジェクト/コンテナ継承、既存ツリーへ反映 }
  Params := '/c icacls "' + AppRoot + '" /grant *S-1-5-11:(OI)(CI)M /T';
  if not Exec(ExpandConstant('{cmd}'), Params, '', SW_HIDE, ewWaitUntilTerminated, ResultCode) then
  begin
    Log('GrantAppRootUserWritableAcl: Exec failed to start cmd for: ' + AppRoot);
    MsgBox(
      'インストール先への書き込み権限付与（icacls）を開始できませんでした。'#13#10 +
      '一般ユーザーでの更新ログ・config 自動更新が失敗する可能性があります。',
      mbError, MB_OK);
    Exit;
  end;
  if ResultCode <> 0 then
  begin
    Log('GrantAppRootUserWritableAcl: icacls exit code=' + IntToStr(ResultCode) + ' path=' + AppRoot);
    MsgBox(
      'インストール先への書き込み権限付与に失敗しました（icacls の終了コード: ' + IntToStr(ResultCode) + '）。'#13#10 +
      '管理者にフォルダ ACL の確認を依頼してください。'#13#10 + AppRoot,
      mbError, MB_OK);
  end
  else
    Log('GrantAppRootUserWritableAcl: ok path=' + AppRoot);
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  DeployRoot: String;
  Dest: String;
begin
  if CurStep <> ssPostInstall then
    Exit;

  DeployRoot := GetDeployRoot('');
  Dest := ExpandConstant('{app}');

  if InstallProgressPage <> nil then
  begin
    InstallProgressPage.SetProgress(0, 14);
    InstallProgressPage.Show;
  end;

  { 初回導入は catalog.json で決めた bin/config/bootstrap zip を展開して適用する }
  try
    try
      DeployByCatalog(DeployRoot, Dest);
    except
      Log('DeployByCatalog failed: ' + GetExceptionMessage);
      MsgBox(
        'catalog.json から初回導入 payload を取得できませんでした。'#13#10 +
        'インストールを中断します。'#13#10#13#10 +
        '配布ルート: ' + DeployRoot,
        mbError, MB_OK
      );
      RaiseException('catalog based initial install failed.');
    end;

    if not FileExists(Dest + '\app\bin\hc_main.exe') then
      RaiseException(
        'Install folder does not contain app\bin\hc_main.exe.'#13#10 +
        'Verify the payload tree matches docs\\Exe化（開発者向け）.md section 5.'
      );
    if not FileExists(Dest + '\config\VERSION.txt') then
      RaiseException('Install folder does not contain config\VERSION.txt.');
    if not FileExists(Dest + '\bootstrap\update_bootstrap.exe') then
      RaiseException('Install folder does not contain bootstrap\update_bootstrap.exe.');

    { DisplayVersion は catalog.set_version（X.Y.Z.N）を優先して同期。無い場合は bin/config から合成 }
    SyncDisplayVersionFromCatalog(DeployRoot);
    { bootstrap 版（X.Y.Z）を catalog.bootstrap.latest_version から同期 }
    SyncBootstrapVersionFromCatalog(DeployRoot, Dest);
    WriteXlwingsConf(Dest);
    GrantAppRootUserWritableAcl(Dest);
  finally
    if InstallProgressPage <> nil then
      InstallProgressPage.Hide;
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  SubKey, WowSubKey: String;
begin
  if CurUninstallStep <> usUninstall then
    Exit;
  SubKey := 'Software\Microsoft\Windows\CurrentVersion\Uninstall\{B5E8C4A2-1D3F-4E6B-9C0D-1A2B3C4D5E6F}_is1';
  WowSubKey := 'Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{B5E8C4A2-1D3F-4E6B-9C0D-1A2B3C4D5E6F}_is1';
  { HKCU 側の重複残存キーを掃除する（HKLM 主キーは Inno 本体が管理） }
  TryDeleteUninstallKey(HKCU, SubKey);
  TryDeleteUninstallKey(HKCU, WowSubKey);
  { インストール先ディレクトリの親（例: ...\Excel_Addin）は空の場合のみ削除する }
  TryRemoveEmptyParentDirOfApp;
end;
