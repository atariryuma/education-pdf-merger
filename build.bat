@echo off
chcp 65001 > nul
echo ========================================
echo 教育計画PDFマージシステム ビルドスクリプト
echo Version 3.7.1 - GUI重複コード徹底削減・README/インストーラー整理
echo ========================================
echo.

REM 仮想環境をアクティベート
if exist ".venv\Scripts\activate.bat" (
    call .venv\Scripts\activate.bat
) else if exist "venv\Scripts\activate.bat" (
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
echo [1/4] クリーンアップ中...
if exist "build" rmdir /s /q build
if exist "dist" rmdir /s /q dist
if exist "__pycache__" rmdir /s /q __pycache__

echo [2/4] 構文チェック中...
python -m py_compile core\pdf_converter.py converters\office_converter.py converters\image_converter.py converters\ichitaro_converter.py
if errorlevel 1 (
    echo [エラー] 構文エラーがあります。ビルドを中止します。
    pause
    exit /b 1
)
echo 構文チェック完了！

echo [3/4] ビルド中...
pyinstaller build_installer.spec --clean

echo.
if exist "dist\教育計画PDFマージシステム.exe" (
    echo ========================================
    echo [成功] ビルドが完了しました！
    echo.
    echo 出力先: dist\教育計画PDFマージシステム.exe
    echo ========================================

    REM config.jsonをdistフォルダにコピー
    echo.
    echo [4/4] 設定ファイルをコピー中...
    copy /Y config.json dist\config.json > nul
    echo config.json をコピーしました。

    echo.
    echo ビルド情報:
    echo   - バージョン: 3.7.1
    echo   - GUI重複削減（attach_placeholder, ask_folder/file_open/file_save, poll_thread）
    echo   - README.md / README_INSTALLER.md を全面刷新
    echo   - 不要なinnosetup_installer.exeを削除（9.7MB）
    echo   - excel_tab.pyのpythoncom未定義参照バグを修正
    echo.
    for %%F in ("dist\教育計画PDFマージシステム.exe") do (
        echo   - ファイルサイズ: %%~zF bytes
    )
) else (
    echo ========================================
    echo [エラー] ビルドに失敗しました。
    echo ========================================
)

echo.
pause
