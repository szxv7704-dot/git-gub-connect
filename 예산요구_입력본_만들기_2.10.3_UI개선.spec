# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['budget_bridge.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['converter', 'plan_parser', 'edufine_form', 'bimok_resolver', 'bimok_groups', 'crosscheck', 'ubis_review', 'hwpx_native_parser', 'ui_common', 'settings', 'report_export', 'program_classes', 'program_cards', 'tutorial', 'native_ui'],
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
    name='예산요구_입력본_만들기_2.10.3_UI개선',
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
)
