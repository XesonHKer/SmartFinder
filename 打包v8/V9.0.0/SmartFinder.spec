# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['SmartFinder_v9_0_0.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('Icon/program_icon.png', '.'),
        ('Icon/icon-windowed.icns', '.'),
        ('pdf_compressor.py', '.'),
    ],
    hiddenimports=[],
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
    [],
    exclude_binaries=True,
    name='SmartFinder',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='Icon/icon-windowed.icns',
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='SmartFinder',
)
app = BUNDLE(
    coll,
    name='SmartFinder.app',
    icon='Icon/icon-windowed.icns',
    bundle_identifier='com.xeson.smartfinder',
    info_plist={
        'CFBundleShortVersionString': '9.0.0',
        'CFBundleVersion': '9.0.0',
    },
)
