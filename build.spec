# build.spec

import sys
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

block_cipher = None

hidden_imports = [
    'google.genai',
    'google.genai.types',
    'google.api_core',
    'google.auth',
    'google.protobuf',
    'grpc',
    'pdfplumber',
    'pdfminer',
    'pdfminer.high_level',
    'pdfminer.layout',
    'pdfminer.converter',
    'pdfminer.pdfparser',
    'pdfminer.pdfdocument',
    'pdfminer.pdfpage',
    'pdfminer.pdfinterp',
    'pdfminer.pdfdevice',
    'pdfminer.pdftypes',
    'pdfminer.utils',
    'docx',
    'docx.oxml',
    'openpyxl',
    'openpyxl.styles',
    'PIL',
    'PIL.Image',
    'tkinter',
    'tkinter.filedialog',
    'tkinter.messagebox',
    'charset_normalizer',
    'certifi',
]

datas = []
datas += collect_data_files('google.genai')
datas += collect_data_files('google.genai')
datas += collect_data_files('pdfplumber')
datas += collect_data_files('pdfminer')
datas += collect_data_files('certifi')

a = Analysis(
    ['main.py'],
    pathex=['.'],
    binaries=[],
    datas=datas,
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['matplotlib', 'numpy', 'pandas', 'scipy'],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='VisitingCardScanner',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,      # No console window — clean desktop app
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='VisitingCardScanner',
)
