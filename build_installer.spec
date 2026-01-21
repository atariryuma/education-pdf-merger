# -*- mode: python ; coding: utf-8 -*-

"""
教育計画PDFマージシステム v3.4.1
PyInstaller ビルド設定ファイル

使用方法:
    pyinstaller build_installer.spec

ビルド後:
    dist/教育計画PDFマージシステム.exe が生成されます
"""

import os
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

# プロジェクトのルートディレクトリ
project_root = os.path.abspath('.')

# データファイルの収集
datas = [
    ('config.json', '.'),  # 設定ファイル
]

# 隠しインポート（動的インポートされるモジュール）
hiddenimports = [
    # 標準ライブラリ
    'encodings',
    'encodings.utf_8',
    'encodings.cp932',

    # プロジェクトモジュール
    'pdf_converter',
    'converters',
    'converters.office_converter',
    'converters.image_converter',
    'converters.ichitaro_converter',
    'pdf_processor',
    'document_collector',
    'pdf_merge_orchestrator',  # v3.3.1で追加
    'config_loader',
    'config_validator',  # v3.4.0で追加
    'ghostscript_detector',  # v3.4.0で追加
    'year_utils',  # v3.4.0で追加
    'constants',
    'exceptions',
    'path_validator',
    'folder_structure_detector',
    'logging_config',

    # GUI関連
    'gui',
    'gui.app',
    'gui.tabs',
    'gui.tabs.base_tab',
    'gui.tabs.pdf_tab',
    'gui.tabs.excel_tab',
    'gui.tabs.settings_tab',
    'gui.tabs.file_tab',
    'gui.utils',
    'gui.styles',
    'gui.ui_constants',
    'gui.ichitaro_dialog',
    'gui.plan_type_selection_dialog',
    'gui.setup_wizard',  # v3.4.0で追加

    # 外部ライブラリ
    'customtkinter',
    'PIL._tkinter_finder',
    'win32com',
    'win32com.client',
    'win32com.client.gencache',
    'pythoncom',
    'pywintypes',
    'pywinauto',
    'pywinauto.controls',
    'pywinauto.keyboard',
    'PyPDF2',
    'fitz',  # PyMuPDF
    'reportlab',
    'reportlab.pdfgen',
    'reportlab.lib',
    'reportlab.lib.pagesizes',
    'reportlab.lib.colors',
    'reportlab.pdfbase',
    'reportlab.pdfbase.ttfonts',
    'openpyxl',
]

# customtkinter のデータファイルを収集
customtkinter_datas = collect_data_files('customtkinter')
datas.extend(customtkinter_datas)

# Analysis オブジェクト
a = Analysis(
    ['run_app.py'],
    pathex=[project_root],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib',
        'numpy',
        'pandas',
        'scipy',
        'pytest',
        'unittest',
        'test',
        'tests',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=None,
    noarchive=False,
)

# PYZ オブジェクト
pyz = PYZ(a.pure, a.zipped_data, cipher=None)

# EXE オブジェクト
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='教育計画PDFマージシステム',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # GUIアプリケーションなのでコンソールを非表示
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # アイコンファイルがあれば指定
    version='version_info.txt',  # バージョン情報ファイル（後で作成）
)
