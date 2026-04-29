# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[('审查逮捕案件010101表1行4列受理合计（人）.xlsx', '.'), ('一审公诉案件020101表1行19列受理合计（人）.xlsx', '.'), ('code_mapping.json', '.'), ('license.lic', '.'), ('used_codes.json', '.')],
    hiddenimports=['openpyxl', 'pandas', 'numpy', 'dateutil', 'tzdata', 'dateutil.tz', 'dateutil.parser'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['matplotlib', 'test', 'unittest', 'numpy.tests'],
    noarchive=False,
    optimize=2,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [('O', None, 'OPTION'), ('O', None, 'OPTION')],
    name='案卡冲突排查系统_v1.0.0',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['app.ico'],
)
