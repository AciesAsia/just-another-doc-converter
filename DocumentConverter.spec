# DocumentConverter.spec
# PyInstaller build specification for DocumentConverter
#
# Usage (run from C:\PythonProjects\DocumentConverter):
#   pyinstaller DocumentConverter.spec
#
# Output: dist\DocumentConverter\  (folder with launcher.exe + all deps)

import sys
from PyInstaller.utils.hooks import collect_submodules, collect_data_files

block_cipher = None

# ── Hidden imports required for pywin32 and other libs ──────────────────────
hidden_imports = [
    # pywin32 COM automation
    'win32com',
    'win32com.client',
    'win32com.server',
    'win32api',
    'win32con',
    'win32gui',
    'pywintypes',
    'pythoncom',
    # weasyprint font/CSS subsystems
    'weasyprint',
    'weasyprint.fonts',
    'weasyprint.css',
    'weasyprint.document',
    # pdfminer
    'pdfminer',
    'pdfminer.high_level',
    'pdfminer.layout',
    # pymupdf
    'fitz',
    # reportlab
    'reportlab',
    'reportlab.pdfgen',
    'reportlab.lib',
    # markdown
    'markdown',
    # other converters
    'docx2pdf',
    'pytesseract',
    'PIL',
    'PIL.Image',
    'odf',
    'striprtf',
    'pptx',
    'openpyxl',
    'xlrd',
]

# ── Collect all submodules for packages that use dynamic imports ─────────────
hidden_imports += collect_submodules('weasyprint')
hidden_imports += collect_submodules('reportlab')
hidden_imports += collect_submodules('pdfminer')

# ── Data files (non-.py assets that must be bundled) ────────────────────────
datas = []
datas += collect_data_files('weasyprint')
datas += collect_data_files('certifi')

# ── Analysis: four source scripts, launcher is the entry point ───────────────
a = Analysis(
    ['launcher.py'],
    pathex=['.'],
    binaries=[],
    datas=datas,
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# NOTE: converter_gui.py, pdf_to_office_gui.py, converter.py are delivered
# directly into {app} by the Inno Setup installer — not bundled here.

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='DocumentConverter',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,          # no console window — same as CREATE_NO_WINDOW
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,              # replace with 'icon.ico' if you add one
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='DocumentConverter',
)
