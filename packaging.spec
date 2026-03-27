# -*- mode: python ; coding: utf-8 -*-
# PyInstaller 配置：python -m PyInstaller --noconfirm --clean packaging.spec

a = Analysis(
    ["app.py"],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        "dxf_outline",
        "dxf_nesting_tab",
        "ezdxf",
        "ezdxf.entities",
        "shapely",
        "shapely.geometry",
        "shapely.ops",
        "shapely.affinity",
        "shapely.wkt",
    ],
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
    name="简易工程计算器",
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
