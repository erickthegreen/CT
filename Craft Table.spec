# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['Craft Table v3.3.py'],
    pathex=[],
    binaries=[],
    datas=[('historico_registros.json', '.'), ('config_tema.json', '.'), ('tabulacoes.json', '.')],
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
    a.binaries,
    a.datas,
    [],
    name='Craft Table',
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
    icon=['Crafting_Table_JE4_BE3.ico'],
)
