@echo off
chcp 65001 > nul
echo ========================================
echo 教育計画PDFマージシステム ビルドスクリプト
echo Version 3.2.1 - 2025年ベストプラクティス対応版
echo ========================================
echo.

REM 仮想環境をアクティベート
if exist "venv\Scripts\activate.bat" (
    call venv\Scripts\activate.bat
) else (
    echo [警告] 仮想環境が見つかりません。システムのPythonを使用します。
)

REM PyInstallerがインストールされているか確認
pip show pyinstaller > nul 2>&1
if errorlevel 1 (
    echo [情報] PyInstallerをインストールしています...
    pip install pyinstaller
)

echo.
echo [1/3] クリーンアップ中...
if exist "build" rmdir /s /q build
if exist "dist" rmdir /s /q dist

echo [2/3] ビルド中...
pyinstaller build_exe.spec --clean

echo.
if exist "dist\教育計画PDFマージシステム.exe" (
    echo ========================================
    echo [成功] ビルドが完了しました！
    echo.
    echo 出力先: dist\教育計画PDFマージシステム.exe
    echo ========================================

    REM config.jsonをdistフォルダにコピー
    echo.
    echo [3/3] 設定ファイルをコピー中...
    copy /Y config.json dist\config.json > nul
    echo config.json をコピーしました。
) else (
    echo ========================================
    echo [エラー] ビルドに失敗しました。
    echo ========================================
)

echo.
pause
