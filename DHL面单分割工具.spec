# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['f:\\项目\\OX面单处理\\split_dhl_labels_gui.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['pandas', 'numpy', 'PIL', 'openpyxl', 'matplotlib', 'scipy', 'sklearn', 'IPython', 'jupyter', 'notebook', 'cryptography', 'urllib3', 'requests', 'certifi', 'charset_normalizer', 'fsspec', 'pytz', 'dateutil', 'psutil', 'jinja2', 'sqlite3', 'email', 'xmlrpc', 'pydoc'],
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
    name='DHL面单分割工具',
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
