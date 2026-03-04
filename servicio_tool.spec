# PyInstaller spec for servicio_tool
# Build on Windows: pyinstaller servicio_tool.spec
# Output: dist/servicio_tool.exe (single file — no libraries needed on the other PC)

a = Analysis(
    ["servicio_tool.py"],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=["openpyxl", "rapidfuzz", "openpyxl.cell._writer"],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

# Single .exe: everything bundled inside. Copy only this exe + your data (config.json, calis.xlsx, groups\).
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="servicio_tool",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
