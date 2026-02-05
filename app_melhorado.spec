# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['app_melhorado.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('Petrobras.png', '.'),
        ('logo jsl.png', '.'),
        ('BD.xlsm', '.'),
    ],
    hiddenimports=[
        'streamlit.web.cli',
        'streamlit.web.bootstrap',
        'streamlit.runtime.scriptrunner.magic_funcs',
        'selenium',
        'selenium.webdriver',
        'selenium.webdriver.chrome',
        'selenium.webdriver.chrome.service',
        'bs4',
        'openpyxl',
        'plotly',
        'plotly.express',
        'plotly.graph_objects',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
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
    name='PTM_JSL_Sistema_v2.0',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)