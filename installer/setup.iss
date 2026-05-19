; 教育計画PDFマージシステム - Inno Setup Script
; Inno Setup 6.0+ required

#define MyAppName "教育計画PDFマージシステム"
#define MyAppVersion "3.6.1"
#define MyAppPublisher "教育機関向けPDFツール"
#define MyAppExeName "教育計画PDFマージシステム.exe"
#define MyAppURL "https://github.com/atariryuma/education-pdf-merger"

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

; 設定ファイル（dist/にコピーされたもの）
Source: "..\dist\config.json"; DestDir: "{app}"; Flags: confirmoverwrite



[Dirs]
; ログ用ディレクトリ
Name: "{localappdata}\PDFMergeSystem\logs"; Permissions: users-modify
; 一時ファイル用ディレクトリ
Name: "{localappdata}\PDFMergeSystem\temp"; Permissions: users-modify

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\設定ファイル"; Filename: "{app}\config.json"
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
; アプリ起動
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[Code]

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


[Messages]
WelcomeLabel1=教育計画PDFマージシステム v{#MyAppVersion} へようこそ
WelcomeLabel2=教育計画や行事計画のドキュメントをPDF化して統合するツールです。%n%n【主な機能】%n• Office文書（Word/Excel/PowerPoint）のPDF変換%n• 画像・一太郎文書のPDF変換%n• 目次・ブックマーク付きPDF統合%n• Excel自動転記機能%n• 初回セットアップウィザード%n%nセットアップを続行するには「次へ」をクリックしてください。

FinishedLabel=教育計画PDFマージシステム v{#MyAppVersion} のインストールが完了しました。%n%n初回起動時にセットアップウィザードが表示されます。%nGhostscriptは自動検出されるため手動設定は不要です。%n%n【必須環境】%n• Microsoft Office（Word/Excel/PowerPoint）%n• 一太郎（.jtdファイルを変換する場合のみ）
