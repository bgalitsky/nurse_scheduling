# scheduler.spec
from PyInstaller.utils.hooks import collect_all

datas, binaries, hiddenimports = [], [], []

d, b, h = collect_all("ortools")
datas += d
binaries += b
hiddenimports += h

d, b, h = collect_all("pandas")
datas += d
binaries += b
hiddenimports += h

d, b, h = collect_all("tksheet")
datas += d
binaries += b
hiddenimports += h

d, b, h = collect_all("openpyxl")
datas += d
binaries += b
hiddenimports += h

block_cipher = None

a = Analysis(
    ["scheduler_tk_db_ortools_enterprise_v3.py"],
    pathex=["."],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports + [
        "ortools.sat.python.cp_model",
        "ortools.sat.python.cp_model_helper",
        "ortools.sat.python",
    ],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    cipher=block_cipher,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="scheduler",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,   # GUI
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name="scheduler",
)
