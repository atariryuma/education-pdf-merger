@echo off
chcp 65001 > nul
echo ========================================
echo 教育計画PDFマージシステム v3.4.1
echo インストーラー ビルドスクリプト
echo ========================================
echo.

REM Inno Setupのパスを探す
set ISCC=
if exist "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" (
    set ISCC="C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
) else if exist "C:\Program Files\Inno Setup 6\ISCC.exe" (
    set ISCC="C:\Program Files\Inno Setup 6\ISCC.exe"
) else (
    echo [エラー] Inno Setup 6 が見つかりません。
    echo.
    echo 以下からダウンロードしてインストールしてください:
    echo https://jrsoftware.org/isdl.php
    echo.
    pause
    exit /b 1
)

echo Inno Setup: %ISCC%
echo.

REM EXEが存在するか確認
if not exist "..\dist\教育計画PDFマージシステム.exe" (
    echo [エラー] EXEファイルが見つかりません。
    echo 先にbuild.batを実行してEXEをビルドしてください。
    echo.
    pause
    exit /b 1
)

REM 出力ディレクトリ作成
if not exist "..\dist\installer" mkdir "..\dist\installer"

echo [1/2] インストーラーをビルド中...
%ISCC% setup.iss

if errorlevel 1 (
    echo.
    echo [エラー] インストーラーのビルドに失敗しました。
    pause
    exit /b 1
)

echo.
echo ========================================
echo [成功] インストーラーのビルドが完了しました！
echo.
echo 出力先: dist\installer\PDFMergeSystem_Setup_3.4.1.exe
echo ========================================
echo.
pause
