; =============================================================================
; CSV Tool — 正本（ステージング）同梱インストーラ（Inno Setup 6+）
; =============================================================================
; 【役割】`tools\nuitka\build_nuitka_all.bat` で生成した **`dist\CSV_Tool`**（正本と同じレイアウトの
;         ステージング一式）を **Setup.exe 内に同梱**し、インストール時に `{app}` へ展開する。
;         共有やネットワークなしで **単体 EXE 配布**できる（オフライン検証・USB 配布など向け）。
;
; 【対比】**`CSV_Tool_Setup.iss`** は **薄いインストーラ**（共有 `current` を robocopy で `{app}` へ複製。
;         ペイロードは EXE に含めない）。運用上の第1候補は通常こちら（`docs\インストールと運用（利用者・運用向け）.md` 参照）。
;
; 【前提】本スクリプトをコンパイルする **前**にステージングを用意すること（`dist\CSV_Tool\*` が存在すること）。
; 【ビルド（推奨）】`installer\build_csv_tool.bat`
;
; 【出力】`dist\CSV_Tool_Setup.exe`（`OutputBaseFilename`）。**薄い版と同じファイル名**のため、
;         両方を交互にコンパイルすると **生成物が上書き**される。配布物を混同しないこと。
;
; 【環境変数】現状、**`HKCU\Environment` の `HC_*` は本スクリプトでは設定しない**（`CSV_Tool_Setup.iss` と
;         挙動が異なる）。詳細・今後の統一方針は **`docs\インストールと運用（利用者・運用向け）.md`** を参照。
;
; 【容量】インストーラ EXE のサイズはステージング全体に依存する。フォルダサイズ目標・削減手順は
;         **`docs\Exe化（開発者向け）.md`** セクション 5.1 を参照。
;
; 【任意】`[Tasks]` の NTFS compact（DLL 主体では効果が薄いことが多い）。
; =============================================================================

#define MyAppName "CSV Tool"
#define MyAppVersion "1.0.0"

[Setup]
AppId={{B5E8C4A2-1D3F-4E6B-9C0D-1A2B3C4D5E6F}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
DefaultDirName={autopf}\Excel_Addin\CSV_Tool
DefaultGroupName={#MyAppName}
OutputDir=..\dist
OutputBaseFilename=CSV_Tool_Setup
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
DisableProgramGroupPage=yes
UninstallDisplayIcon={app}\app\bin\hc_main.exe

[Languages]
Name: "japanese"; MessagesFile: "compiler:Languages\Japanese.isl"

[Tasks]
Name: postntfscompact; Description: "Install後に compact /C/S を実行（テキスト中心のツリー向け。DLL 主体では効果が薄いことが多い）"; GroupDescription: "Optional:"; Flags: unchecked

[Files]
; `build_nuitka_all.bat` 出力のステージングルート（config, xlwings.conf, app\...）を正本同梱で展開
Source: "..\dist\CSV_Tool\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Run]
Filename: "{sys}\compact.exe"; Parameters: "/C /S:\""{app}\"""; StatusMsg: "Compressing installed files (NTFS)..."; Flags: runhidden waituntilterminated; Tasks: postntfscompact

[Icons]
Name: "{group}\Uninstall {#MyAppName}"; Filename: "{uninstallexe}"

[Code]
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

function NormalizeConfigVersion(const RawVersion: String): String;
var
  V: Integer;
  S: String;
begin
  Result := '';
  S := Trim(RawVersion);
  if S = '' then
    Exit;
  V := StrToIntDef(S, -1);
  if V < 0 then
    Exit;
  Result := IntToStr(V);
end;

procedure TrySyncDisplayValues(const RootKey: Integer; const SubKey, DisplayName, DisplayVersion: String; var AnyWritten: Boolean);
begin
  try
    if RegWriteStringValue(RootKey, SubKey, 'DisplayName', DisplayName) then
      AnyWritten := True;
  except
  end;
  try
    if RegWriteStringValue(RootKey, SubKey, 'DisplayVersion', DisplayVersion) then
      AnyWritten := True;
  except
  end;
end;

procedure SyncDisplayValuesFromInstalledVersion;
var
  VersionPath, ConfigVersionPath, DisplayVersion, ConfigVersion, SubKey, WowSubKey: String;
  RawVersion, RawConfigVersion: AnsiString;
  AnyWritten: Boolean;
begin
  VersionPath := ExpandConstant('{app}\VERSION.txt');
  if not LoadStringFromFile(VersionPath, RawVersion) then
    Exit;
  if Trim(String(RawVersion)) = '' then
    Exit;
  DisplayVersion := NormalizeDisplayVersion(String(RawVersion));
  if DisplayVersion = '' then
    Exit;
  ConfigVersionPath := ExpandConstant('{app}\config\VERSION.txt');
  if LoadStringFromFile(ConfigVersionPath, RawConfigVersion) then
  begin
    ConfigVersion := NormalizeConfigVersion(String(RawConfigVersion));
    if ConfigVersion <> '' then
      DisplayVersion := DisplayVersion + '.' + ConfigVersion;
  end;
  AnyWritten := False;
  SubKey := 'Software\Microsoft\Windows\CurrentVersion\Uninstall\{B5E8C4A2-1D3F-4E6B-9C0D-1A2B3C4D5E6F}_is1';
  WowSubKey := 'Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{B5E8C4A2-1D3F-4E6B-9C0D-1A2B3C4D5E6F}_is1';
  TrySyncDisplayValues(HKLM, SubKey, 'CSV Tool', DisplayVersion, AnyWritten);
  TrySyncDisplayValues(HKLM, WowSubKey, 'CSV Tool', DisplayVersion, AnyWritten);
  TrySyncDisplayValues(HKCU, SubKey, 'CSV Tool', DisplayVersion, AnyWritten);
  TrySyncDisplayValues(HKCU, WowSubKey, 'CSV Tool', DisplayVersion, AnyWritten);
  if not AnyWritten then
    Log('DisplayName/DisplayVersion sync skipped: uninstall key not writable');
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
    SyncDisplayValuesFromInstalledVersion;
end;
