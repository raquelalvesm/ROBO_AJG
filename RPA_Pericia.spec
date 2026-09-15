# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=[('templates', 'templates')],
    hiddenimports=['openpyxl', 'selenium', 'selenium.webdriver', 'selenium.webdriver.chrome', 'selenium.webdriver.chrome.webdriver', 'selenium.webdriver.chrome.service', 'selenium.webdriver.chrome.options', 'selenium.webdriver.common.by', 'selenium.webdriver.common.keys', 'selenium.webdriver.common.action_chains', 'selenium.webdriver.support', 'selenium.webdriver.support.ui', 'selenium.webdriver.support.expected_conditions', 'selenium.common.exceptions', 'selenium.webdriver.remote.webdriver', 'webdriver_manager', 'webdriver_manager.chrome', 'webdriver_manager.core', 'pyperclip', 'docx', 'psutil', 'util_processo', 'ajg_scraper'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='RPA_Pericia',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['icon.ico'],
)
