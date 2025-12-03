# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['excel_merger_gui.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        'merger',
        'styles',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib', 'mpl_toolkits',
        'tcl', 'tk', '_tkinter', 'tkinter', 'tkinter.filedialog',
        'pandas', 'numpy',
        'scipy', 'sklearn', 'pytest', 'setuptools',
        'PIL', 'wx', 'django', 'flask', 'requests',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='Excel表格合并工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
