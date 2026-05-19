# -*- mode: python ; coding: utf-8 -*-

"""
教育計画PDFマージシステム v3.6.1
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

    # shared modules
    'shared',
    'shared.constants',
    'shared.exceptions',

    # core modules
    'core',
    'core.pdf_converter',
    'core.pdf_processor',
    'core.document_collector',
    'core.pdf_merge_orchestrator',
    'core.folder_structure_detector',
    'core.update_excel_files',

    # infrastructure modules
    'infrastructure',
    'infrastructure.config_loader',
    'infrastructure.config_validator',
    'infrastructure.ghostscript',
    'infrastructure.year_utils',
    'infrastructure.path_validator',
    'infrastructure.logging_config',

    # converters
    'converters',
    'converters.office_converter',
    'converters.image_converter',
    'converters.ichitaro_converter',

    # GUI関連
    'gui',
    'gui.app',
    'gui.tabs',
    'gui.tabs.base_tab',
    'gui.tabs.pdf_tab',
    'gui.tabs.excel_tab',
    'gui.tabs.settings_tab',
    'gui.utils',
    'gui.styles',
    'gui.ichitaro_dialog',
    'gui.plan_type_selection_dialog',
    'gui.setup_wizard',
    'gui.event_names_editor',

    # 外部ライブラリ
    'PIL._tkinter_finder',
    'win32com',
    'win32com.client',
    'win32com.client.gencache',
    'pythoncom',
    'pywintypes',
    'win32timezone',  # pywin32の日付時刻処理用（Excel操作に必要）
    'win32api',
    'win32con',
    'pywinauto',
    'pywinauto.controls',
    'pywinauto.keyboard',
    'comtypes',
    'comtypes.client',
    'comtypes.stream',
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
    icon='app.ico',
    version='version_info.txt',  # バージョン情報ファイル（後で作成）
)
