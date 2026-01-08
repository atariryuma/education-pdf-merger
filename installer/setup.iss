; 教育計画PDFマージシステム - Inno Setup Script
; Inno Setup 6.0+ required
;
; Version 3.2.4 - 一太郎変換改善版（警告ダイアログ＆リトライ機能）
; - 一太郎変換中の警告ダイアログ（非モーダル、常に最前面）
; - 変換失敗時の自動リトライ（最大3回）
; - リトライ状況をリアルタイム表示

#define MyAppName "教育計画PDFマージシステム"
#define MyAppVersion "3.2.4"
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
Type: files; Name: "{app}\*.pyc"
Type: files; Name: "{app}\*.pyo"
Type: files; Name: "{localappdata}\PDFMergeSystem\user_config.json"
Type: files; Name: "{localappdata}\PDFMergeSystem\.last_settings.json"
Type: filesandordirs; Name: "{localappdata}\PDFMergeSystem\logs"
Type: filesandordirs; Name: "{localappdata}\PDFMergeSystem\temp"
Type: filesandordirs; Name: "{localappdata}\PDFMergeSystem"

[Run]
; Ghostscript自動検出（インストール完了後、アプリ起動前に実行）
Filename: "{tmp}\post_install.exe"; Parameters: """{app}"""; StatusMsg: "Ghostscriptを検索しています..."; Flags: runhidden waituntilterminated
; アプリ起動
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[Code]
var
  GhostscriptFound: Boolean;

// プロセスが実行中かチェック
function IsAppRunning(): Boolean;
var
  ResultCode: Integer;
  Output: AnsiString;
  OutputFile: String;
begin
  Result := False;

  // tasklist の出力をファイルに保存して確認
  OutputFile := ExpandConstant('{tmp}\tasklist_output.txt');

  // tasklist は常に成功コード0を返すため、出力内容で判定する必要がある
  // /NH = ヘッダーなし、/FO CSV = CSV形式
  if Exec('cmd.exe', '/C tasklist /FI "IMAGENAME eq 教育計画PDFマージシステム.exe" /NH /FO CSV > "' + OutputFile + '"', '', SW_HIDE, ewWaitUntilTerminated, ResultCode) then
  begin
    if FileExists(OutputFile) then
    begin
      if LoadStringFromFile(OutputFile, Output) then
      begin
        // 出力に実行ファイル名が含まれていれば実行中
        Result := (Pos('教育計画PDFマージシステム.exe', String(Output)) > 0);
      end;
      DeleteFile(OutputFile);
    end;
  end;
end;

// プロセスを強制終了
function KillApp(): Boolean;
var
  ResultCode: Integer;
begin
  Result := Exec('taskkill.exe', '/F /IM "教育計画PDFマージシステム.exe"', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
end;

// インストール開始前の処理
function InitializeSetup(): Boolean;
begin
  Result := True;

  // アプリが実行中の場合
  if IsAppRunning() then
  begin
    if MsgBox('教育計画PDFマージシステムが実行中です。' + #13#10 +
              'アプリケーションを終了してからインストールを続行しますか？',
              mbConfirmation, MB_YESNO) = IDYES then
    begin
      KillApp();
      Sleep(1000);  // 終了を待つ
    end
    else
    begin
      Result := False;  // インストールをキャンセル
    end;
  end;
end;

// アンインストール開始前の処理
function InitializeUninstall(): Boolean;
begin
  Result := True;

  // アプリが実行中の場合は強制終了
  if IsAppRunning() then
  begin
    KillApp();
    Sleep(1000);  // 終了を待つ
  end;
end;

var
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
