; 教育計画PDFマージシステム - Inno Setup Script
; Inno Setup 6.0+ required

#define MyAppName "教育計画PDFマージシステム"
#define MyAppVersion "3.2"
#define MyAppPublisher "School Tools"
#define MyAppExeName "教育計画PDFマージシステム.exe"

[Setup]
; アプリケーション情報
AppId={{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
AllowNoIcons=yes
; 出力設定
OutputDir=..\dist\installer
OutputBaseFilename=PDFMergeSystem_Setup_{#MyAppVersion}
; 圧縮設定
Compression=lzma2/ultra64
SolidCompression=yes
; UI設定
WizardStyle=modern
; 権限
PrivilegesRequired=admin
PrivilegesRequiredOverridesAllowed=dialog
; その他
DisableProgramGroupPage=yes

[Languages]
Name: "japanese"; MessagesFile: "compiler:Languages\Japanese.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; メインEXE
Source: "..\dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
; 設定ファイルテンプレート（既存があれば上書きしない）
Source: "config_template.json"; DestDir: "{app}"; DestName: "config.json"; Flags: onlyifdoesntexist
; Ghostscript検出用スクリプト（インストール後に削除）
Source: "dist\post_install.exe"; DestDir: "{tmp}"; Flags: deleteafterinstall

[Dirs]
; ログ用ディレクトリ
Name: "{localappdata}\PDFMergeSystem\logs"; Permissions: users-modify
; 一時ファイル用ディレクトリ
Name: "{localappdata}\PDFMergeSystem\temp"; Permissions: users-modify

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[UninstallDelete]
; アンインストール時に削除するファイル・フォルダ
Type: files; Name: "{app}\config.json"
Type: files; Name: "{app}\*.log"
Type: files; Name: "{localappdata}\PDFMergeSystem\.last_settings.json"
Type: filesandordirs; Name: "{localappdata}\PDFMergeSystem\logs"
Type: filesandordirs; Name: "{localappdata}\PDFMergeSystem\temp"
Type: dirifempty; Name: "{localappdata}\PDFMergeSystem"

[Run]
; Ghostscript自動検出（インストール完了後、アプリ起動前に実行）
Filename: "{tmp}\post_install.exe"; Parameters: """{app}"""; StatusMsg: "Ghostscriptを検索しています..."; Flags: runhidden waituntilterminated
; アプリ起動
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[Code]
var
  GhostscriptFound: Boolean;
  GhostscriptPath: AnsiString;

function GetPostInstallResult(): Boolean;
var
  OutputFile: String;
begin
  Result := False;
  GhostscriptFound := False;
  OutputFile := ExpandConstant('{tmp}\gs_result.txt');

  // post_install.exeの出力を確認
  if FileExists(OutputFile) then
  begin
    if LoadStringFromFile(OutputFile, GhostscriptPath) then
    begin
      if Pos('OK:', GhostscriptPath) = 1 then
      begin
        GhostscriptFound := True;
        Result := True;
      end;
    end;
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
  begin
    GetPostInstallResult();
  end;
end;

function UpdateReadyMemo(Space, NewLine, MemoUserInfoInfo, MemoDirInfo, MemoTypeInfo,
  MemoComponentsInfo, MemoGroupInfo, MemoTasksInfo: String): String;
begin
  Result := MemoDirInfo + NewLine + NewLine;
  Result := Result + 'インストール後にGhostscriptを自動検出します。' + NewLine;
  Result := Result + 'PDF圧縮機能を使用するにはGhostscriptが必要です。';
end;
