# DocumentConverter.spec
# PyInstaller build specification for DocumentConverter
#
# Usage (run from C:\PythonProjects\DocumentConverter):
#   pyinstaller DocumentConverter.spec
#
# Output: dist\DocumentConverter\  (folder with DocumentConverter.exe + all deps)
#
# F-03: converter_gui.py, pdf_to_office_gui.py, and converter.py are no longer
#        bundled as DATA files.  launcher.py imports them directly as Python
#        modules; PyInstaller discovers and compiles them via the Analysis imports.

import sys
from PyInstaller.utils.hooks import collect_submodules, collect_data_files

block_cipher = None

# ── Hidden imports ───────────────────────────────────────────────────────────
hidden_imports = [
    # pywin32 COM automation
    'win32com',
    'win32com.client',
    'win32com.server',
    'win32api',
    'win32con',
    'win32gui',
    'win32process',
    'win32file',
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
    # pdfplumber — F-15: required for PDF→Excel table extraction
    'pdfplumber',
    # pypdf — F-15: required for merging per-sheet Excel PDFs
    'pypdf',
    # numpy — F-14: required for Chrome screenshot row analysis
    'numpy',
    # psutil — F-02: required for safe per-PID Office termination
    'psutil',
    # reportlab
    'reportlab',
    'reportlab.pdfgen',
    'reportlab.lib',
    # markdown / misc
    'markdown',
    'docx2pdf',
    'pytesseract',
    'PIL',
    'PIL.Image',
    'odf',
    'striprtf',
    'pptx',
    'openpyxl',
    'xlrd',
    'docx',
    # F-19: tkinterdnd2 for drag-and-drop
    'tkinterdnd2',
]

# ── Collect all submodules for packages that use dynamic imports ──────────────
hidden_imports += collect_submodules('weasyprint')
hidden_imports += collect_submodules('reportlab')
hidden_imports += collect_submodules('pdfminer')
hidden_imports += collect_submodules('pdfplumber')   # F-15
hidden_imports += collect_submodules('tkinterdnd2')  # F-19

# ── Data files (non-.py assets that must be bundled) ─────────────────────────
datas = []
datas += collect_data_files('weasyprint')
datas += collect_data_files('certifi')
datas += collect_data_files('tkinterdnd2')   # F-19: DnD Tcl extension binaries

# ── Analysis: launcher.py is the single entry point ──────────────────────────
# F-03: converter_gui, pdf_to_office_gui, and converter are imported as Python
#        modules directly by launcher.py — PyInstaller includes them automatically
#        via the import graph.  No DATA entries needed.
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

# F-03: DATA .py bundle entries removed — launcher imports modules directly.
# (Previous spec had 'converter_gui.py', 'pdf_to_office_gui.py', 'converter.py'
#  as DATA because the old launcher used subprocess to run them; that is gone.)

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
    console=False,            # no console window
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,                # F-31: replace with 'app_icon.ico' once icon is added to project root
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
