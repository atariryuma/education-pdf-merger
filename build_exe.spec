# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec file for 教育計画PDFマージシステム
"""

import os
import sys

block_cipher = None

# プロジェクトルート
project_root = os.path.dirname(os.path.abspath(SPEC))

# 追加データファイル
added_files = [
    (os.path.join(project_root, 'config.json'), '.'),
]

# 隠しインポート（PyInstallerが検出できないモジュール）
hidden_imports = [
    'win32com.client',
    'win32com.shell',
    'pythoncom',
    'pywintypes',
    'pywinauto',
    'pywinauto.application',
    'pywinauto.keyboard',
    'pywinauto.findwindows',
    'openpyxl',
    'PIL',
    'PIL.Image',
    'fitz',
    'PyPDF2',
    'reportlab',
    'reportlab.lib.styles',
    'reportlab.lib.pagesizes',
    'reportlab.platypus',
    'reportlab.pdfbase.ttfonts',
    'reportlab.pdfbase.pdfmetrics',
]

a = Analysis(
    ['run_app.py'],
    pathex=[project_root],
    binaries=[],
    datas=added_files,
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib',
        'numpy',
        'scipy',
        'pandas',
        'pytest',
        'unittest',
        'tkinter.test',
        'test',
        'http.server',
        'xmlrpc',
        'pydoc',
        'doctest',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

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
    console=False,  # GUIアプリなのでコンソール非表示
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # アイコンがある場合はここにパスを指定
)
