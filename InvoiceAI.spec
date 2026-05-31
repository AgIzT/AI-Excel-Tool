# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_all

# qfluentwidgets / qframelesswindow 含 QSS、字体、图标等资源与大量子模块，
# PyInstaller 默认会漏，必须用 collect_all 全量收集，否则 exe 启动即缺资源。
_d, _b, _h = [], [], []
for _pkg in ("qfluentwidgets", "qframelesswindow", "darkdetect"):
    _pd, _pb, _ph = collect_all(_pkg)
    _d += _pd
    _b += _pb
    _h += _ph


a = Analysis(
    ['main_app.py'],
    pathex=[],
    binaries=_b,
    datas=[('templates', 'templates'), ('prompts', 'prompts')] + _d,
    hiddenimports=['openai'] + _h,
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
    name='InvoiceAI',
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
    manifest='app.manifest',
)
