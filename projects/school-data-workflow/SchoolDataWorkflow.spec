# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path

project_dir = Path(SPECPATH)

a = Analysis(
    [str(project_dir / "launcher.py")],
    pathex=[str(project_dir)],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["PySide6.Qt3DCore", "PySide6.QtMultimedia", "PySide6.QtWebEngineCore"],
    noarchive=False,
    optimize=1,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="학교자료취합도우미",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    # UPX 압축된 Qt DLL은 일부 Windows 환경에서 QtWidgets 로딩 실패를 일으킬 수 있다.
    # 배포 안정성을 위해 Qt DLL을 압축하지 않는다.
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
