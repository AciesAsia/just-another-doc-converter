"""
=============================================================
  Document -> PDF Converter
  Usage:
    python converter.py                  # convert all files in input\
    python converter.py --file foo.docx  # convert a single file

  Input folder  : <app dir>\input\
  Output folder : <app dir>\output\
  Log file      : %LOCALAPPDATA%\DocumentConverter\logs\convert_log.txt

  Supported formats:
    Documents    : .docx  .doc
    Spreadsheets : .xlsx  .xls
    Presentations: .pptx  .ppt
    Web          : .html  .htm
    Comma-sep    : .csv
    Markdown     : .md
    Plain text   : .txt

  PDF Conversion (Phase 4):
    PDF -> Word  : .pdf -> .docx  (Word COM, then pymupdf, then OCR)
    PDF -> Excel : .pdf -> .xlsx  (pdfplumber tables, then pymupdf text)

  Dependencies (install once via pip install -r requirements.txt):
    pywin32    -- drives Word/Excel/PowerPoint COM (.doc .xlsx .pptx .ppt)
    reportlab  -- renders .md, .txt, .csv, .html, .xls to PDF (pure Python)
    xlrd       -- reads legacy .xls files
    pymupdf    -- PDF parsing and page rendering (pip install pymupdf)
    python-docx-- write .docx output (pip install python-docx)
    pdfplumber -- extract tables from PDFs (pip install pdfplumber)
    openpyxl   -- write .xlsx output (pip install openpyxl)
    pytesseract-- OCR for scanned PDFs (pip install pytesseract + Tesseract exe)
    Pillow     -- image handling for OCR (pip install Pillow)
    numpy      -- image row analysis in convert_html Chrome screenshot crop
    pypdf      -- multi-sheet Excel PDF merge
    psutil     -- safe per-PID Office process termination
=============================================================
"""

import argparse
import csv
import json
import os
import re
import subprocess
import sys
import textwrap
import threading
from datetime import datetime
from pathlib import Path

# ── Folder paths ──────────────────────────────────────────
# F-01: resolve path from executable location, not hardcoded developer path
if getattr(sys, 'frozen', False):
    BASE_DIR = Path(sys.executable).resolve().parent
else:
    BASE_DIR = Path(__file__).resolve().parent

INPUT_DIR  = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"

# F-01: log to %LOCALAPPDATA% — writable by standard users, no admin required
_localappdata = Path(os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local"))
LOG_DIR  = _localappdata / "DocumentConverter" / "logs"
LOG_FILE = LOG_DIR / "convert_log.txt"
LOG_DIR.mkdir(parents=True, exist_ok=True)

SUPPORTED = {".docx", ".doc", ".xlsx", ".xls", ".pptx", ".ppt",
             ".html", ".htm", ".md", ".txt", ".csv", ".pdf"}

# Output mode for PDF files — determines which conversion path to use.
# "word"  = PDF -> .docx
# "excel" = PDF -> .xlsx
# Set by the GUI; CLI defaults to Word.
PDF_OUTPUT_MODE = "word"

# F-18: read save_log flag from config.json; default True
_SAVE_LOG = True
try:
    _config_path = BASE_DIR / "config.json"
    if _config_path.exists():
        _cfg = json.loads(_config_path.read_text(encoding="utf-8"))
        _SAVE_LOG = bool(_cfg.get("save_log", True))
except Exception:
    pass


# ── Logging ───────────────────────────────────────────────
def log(message: str):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{timestamp}] {message}"
    print(line)
    # F-18: guard file writes behind save_log config key
    if _SAVE_LOG:
        try:
            with LOG_FILE.open("a", encoding="utf-8") as f:
                f.write(line + "\n")
        except OSError:
            pass


# ── psutil helpers for safe per-PID Office kill (F-02) ───
try:
    import psutil as _psutil
    _PSUTIL_AVAILABLE = True
except ImportError:
    _PSUTIL_AVAILABLE = False


def _snapshot_pids(exe_name: str) -> set:
    # F-02: record which PIDs of exe_name exist BEFORE Dispatch so we can find the new one
    if not _PSUTIL_AVAILABLE:
        return set()
    try:
        return {p.pid for p in _psutil.process_iter(['pid', 'name'])
                if p.info['name'] and p.info['name'].lower() == exe_name.lower()}
    except Exception:
        return set()


def _get_new_office_pid(exe_name: str, pids_before: set):
    # F-02: find the process spawned since the snapshot — our Office instance
    if not _PSUTIL_AVAILABLE:
        return None
    try:
        for p in _psutil.process_iter(['pid', 'name']):
            if (p.info['name'] and p.info['name'].lower() == exe_name.lower()
                    and p.info['pid'] not in pids_before):
                return p.info['pid']
    except Exception:
        pass
    return None


def _kill_office_pid(pid):
    # F-02: kill only this specific Office PID, never all instances
    if not _PSUTIL_AVAILABLE or pid is None:
        return
    try:
        proc = _psutil.Process(pid)
        proc.kill()
    except (_psutil.NoSuchProcess, _psutil.AccessDenied, Exception):
        pass


# ── Output file collision guard (F-07) ───────────────────
def _safe_dest(output_dir: Path, stem: str, suffix: str) -> Path:
    # F-07: prevent silently overwriting an existing output file
    dest = output_dir / (stem + suffix)
    if not dest.exists():
        return dest
    counter = 2
    while True:
        candidate = output_dir / f"{stem} ({counter}){suffix}"
        if not candidate.exists():
            return candidate
        counter += 1


# ── Preserve source file timestamps (F-26) ───────────────
def _copy_file_times(src: Path, dest: Path):
    # F-26: copy creation/access/modification timestamps from source to output
    try:
        import win32file
        import pywintypes
        st = src.stat()
        handle = win32file.CreateFile(
            str(dest), win32file.GENERIC_WRITE, 0, None,
            win32file.OPEN_EXISTING, win32file.FILE_ATTRIBUTE_NORMAL, None)
        win32file.SetFileTime(
            handle,
            pywintypes.Time(st.st_ctime),
            pywintypes.Time(st.st_atime),
            pywintypes.Time(st.st_mtime))
        win32file.CloseHandle(handle)
    except Exception:
        pass


# ── COM helpers ───────────────────────────────────────────
def _quit_with_timeout(app, timeout=8, our_pid=None):
    # F-02: attempt graceful Quit(); kill only our PID if it hangs — no system-wide taskkill
    def _quit():
        try:
            app.Quit()
        except Exception:
            pass
    t = threading.Thread(target=_quit, daemon=True)
    t.start()
    t.join(timeout=timeout)
    if t.is_alive() and our_pid is not None:
        _kill_office_pid(our_pid)


# ── Converters ────────────────────────────────────────────

def _word_suppress_dialogs(word):
    """Suppress all Word dialogs and background publishing notifications."""
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        word.Options.SaveNormalPrompt = False
    except Exception:
        pass
    try:
        word.Options.PrintBackground = False
    except Exception:
        pass
    try:
        word.Options.BackgroundSave = False
    except Exception:
        pass
    try:
        word.Options.DoNotPromptForConvert = True
    except Exception:
        pass
    try:
        word.Options.UpdateLinksAtOpen = False
    except Exception:
        pass


def convert_docx(src: Path, dest: Path, log_fn=None) -> bool:
    """Convert .docx to PDF via Word COM SaveAs2."""
    _log = log_fn or log  # F-20: use caller-supplied logger instead of monkey-patching
    word = None
    our_pid = None
    try:
        import win32com.client
        pids_before = _snapshot_pids('winword.exe')  # F-02: snapshot before Dispatch
        word = win32com.client.Dispatch("Word.Application")
        word.AutomationSecurity = 3  # F-08: msoAutomationSecurityForceDisable — block all macros
        our_pid = _get_new_office_pid('winword.exe', pids_before)  # F-02: identify our instance
        _word_suppress_dialogs(word)
        try:
            word.WindowState = 2   # wdWindowStateMinimize
        except Exception:
            pass
        doc = word.Documents.Open(
            str(src.resolve()),
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
            Revert=False,
        )
        doc.SaveAs2(str(dest.resolve()), FileFormat=17)  # wdFormatPDF = 17
        doc.Close(SaveChanges=False)
        _quit_with_timeout(word, timeout=8, our_pid=our_pid)  # F-02: kill only our Word
        _log(f"  [OK] {dest.name}")
        return True
    except ImportError:
        _log("  [ERROR] pywin32 not installed. Run: pip install pywin32")
        return False
    except Exception as e:
        # F-21: distinguish password-protected and corrupted files from generic errors
        err_str = str(e).lower()
        hresult  = getattr(e, 'hresult', 0)
        if "password" in err_str or "protected" in err_str or hresult == -2146822611:
            _log(f"  [SKIP] {src.name} — password-protected file, cannot convert.")
        elif "corrupt" in err_str or "not a valid" in err_str:
            _log(f"  [SKIP] {src.name} — file appears corrupted, skipping.")
        else:
            _log(f"  [FAIL] {src.name} — {e}")
        if word:
            _quit_with_timeout(word, timeout=4, our_pid=our_pid)  # F-02
        return False


def convert_doc(src: Path, dest: Path, log_fn=None) -> bool:
    """Convert legacy .doc to PDF — identical path to .docx."""
    return convert_docx(src, dest, log_fn=log_fn)  # F-20: pass log_fn through


def convert_xlsx(src: Path, dest: Path, log_fn=None) -> bool:
    """Convert .xlsx to PDF via Excel COM ExportAsFixedFormat with timeout guard.

    Strategy:
      1. Workbook-level ExportAsFixedFormat — works for most files.
      2. Sheet-by-sheet fallback — used when workbook-level fails.
         Each visible sheet exported to a temp PDF; pypdf merges them.
    A daemon thread enforces a 5-minute wall-clock timeout.
    """
    import tempfile
    import os as _os

    _log = log_fn or log  # F-20

    MAX_XLSX_MB = 15
    file_mb = src.stat().st_size / (1024 * 1024)
    if file_mb > MAX_XLSX_MB:
        _log(f"  [SKIP] {src.name} — file too large for PDF conversion "
             f"({file_mb:.1f} MB > {MAX_XLSX_MB} MB limit). "
             f"Open in Excel and set a print area before converting.")
        return False

    try:
        import win32com.client
    except ImportError:
        _log("  [ERROR] pywin32 not installed. Run: pip install pywin32")
        return False

    # F-02: snapshot Excel PIDs before the thread dispatches Excel
    pids_before = _snapshot_pids('excel.exe')

    result = {"ok": False, "err": None, "warn": None, "temps": []}  # F-36: added "temps"
    dest_str = str(dest.resolve())
    src_str  = str(src.resolve())

    def _do_export():
        import pythoncom  # F-04: STA apartment required for Office COM in a new thread
        pythoncom.CoInitialize()
        excel_app = None
        try:
            excel_app = win32com.client.Dispatch("Excel.Application")
            excel_app.AutomationSecurity = 3  # F-08: disable all macros
            excel_app.Visible = False
            excel_app.DisplayAlerts = False
            excel_app.AskToUpdateLinks = False
            try:
                excel_app.Calculation = -4135   # xlCalculationManual
            except Exception:
                pass
            try:
                excel_app.EnableEvents = False
            except Exception:
                pass
            try:
                excel_app.ScreenUpdating = False
            except Exception:
                pass
            wb = excel_app.Workbooks.Open(
                src_str,
                UpdateLinks=0,
                ReadOnly=True,
                IgnoreReadOnlyRecommended=True,
                CorruptLoad=2,   # xlRepairFile
            )
            try:
                wb.EnableAutoRecover = False
            except Exception:
                pass
            try:
                for conn in wb.Connections:
                    try:
                        conn.ODBCConnection.BackgroundQuery = False
                    except Exception:
                        pass
                    try:
                        conn.OLEDBConnection.BackgroundQuery = False
                    except Exception:
                        pass
            except Exception:
                pass

            # Stage 1: whole-workbook export
            wb_ok = False
            try:
                wb.ExportAsFixedFormat(0, dest_str)
                wb_ok = True
            except Exception as ex1:
                result["err"] = ex1

            if wb_ok:
                wb.Close(SaveChanges=False)
                result["ok"] = True
                result["err"] = None
                return

            # Stage 2: sheet-by-sheet export
            sheet_pdfs = []
            try:
                n = wb.Sheets.Count
                for i in range(1, n + 1):
                    sh = wb.Sheets(i)
                    if sh.Visible != -1:   # xlSheetVisible = -1
                        continue
                    tmp = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
                    tmp.close()
                    result["temps"].append(tmp.name)  # F-36: register for main-thread cleanup
                    try:
                        sh.ExportAsFixedFormat(0, tmp.name)
                        sheet_pdfs.append(tmp.name)
                    except Exception:
                        pass

                if sheet_pdfs:
                    try:
                        from pypdf import PdfWriter
                        writer = PdfWriter()
                        for p in sheet_pdfs:
                            writer.append(p)
                        with open(dest_str, "wb") as fout:
                            writer.write(fout)
                        result["ok"] = True
                        result["err"] = None
                    except ImportError:
                        import shutil
                        shutil.copy(sheet_pdfs[0], dest_str)
                        if len(sheet_pdfs) > 1:
                            result["warn"] = "pypdf not installed — only first sheet exported"
                        result["ok"] = True
                        result["err"] = None
                else:
                    result["err"] = Exception("No visible sheets could be exported")
            except Exception as ex2:
                result["err"] = ex2
            finally:
                try:
                    wb.Close(SaveChanges=False)
                except Exception:
                    pass
        except Exception as outer:
            result["err"] = outer
        finally:
            if excel_app is not None:
                try:
                    excel_app.Quit()
                except Exception:
                    pass
            try:
                pythoncom.CoUninitialize()  # F-04
            except Exception:
                pass

    t = threading.Thread(target=_do_export, daemon=True)
    t.start()
    t.join(timeout=300)   # 5-minute hard limit

    # F-36: clean temp files from main thread — covers the timeout case where
    #       _do_export's finally block may not have run yet
    for p in result.get("temps", []):
        try:
            if _os.path.exists(p):
                _os.unlink(p)
        except Exception:
            pass

    # F-02: find and kill only our Excel process, not all Excel instances
    our_pid = _get_new_office_pid('excel.exe', pids_before)
    if our_pid:
        _kill_office_pid(our_pid)

    if t.is_alive() or not result["ok"]:
        err = result["err"] or "Timed out after 300s"
        _log(f"  [FAIL] {src.name} — {err}")
        return False

    if result["warn"]:
        _log(f"  [WARN] {dest.name} — {result['warn']}")
    else:
        _log(f"  [OK] {dest.name}")
    return True


def convert_pptx(src: Path, dest: Path, log_fn=None) -> bool:
    """Convert .pptx/.ppt to PDF via PowerPoint COM SaveAs (ppSaveAsPDF)."""
    _log = log_fn or log  # F-20
    ppt = None
    our_pid = None
    try:
        import win32com.client
        pids_before = _snapshot_pids('powerpnt.exe')  # F-02
        ppt = win32com.client.Dispatch("PowerPoint.Application")
        ppt.AutomationSecurity = 3  # F-08: disable all macros
        our_pid = _get_new_office_pid('powerpnt.exe', pids_before)  # F-02
        try:
            ppt.Visible = False
        except Exception:
            pass
        try:
            ppt.WindowState = 2   # ppWindowMinimized
        except Exception:
            pass
        presentation = ppt.Presentations.Open(
            str(src.resolve()), ReadOnly=True, Untitled=False, WithWindow=False)
        presentation.SaveAs(str(dest.resolve()), 32)  # ppSaveAsPDF = 32
        presentation.Close()
        _quit_with_timeout(ppt, timeout=8, our_pid=our_pid)  # F-02
        _log(f"  [OK] {dest.name}")
        return True
    except ImportError:
        _log("  [ERROR] pywin32 not installed. Run: pip install pywin32")
        return False
    except Exception as e:
        # F-21: distinguish password-protected files
        err_str = str(e).lower()
        hresult  = getattr(e, 'hresult', 0)
        if "password" in err_str or "protected" in err_str or hresult == -2146822611:
            _log(f"  [SKIP] {src.name} — password-protected file, cannot convert.")
        elif "corrupt" in err_str or "not a valid" in err_str:
            _log(f"  [SKIP] {src.name} — file appears corrupted, skipping.")
        else:
            _log(f"  [FAIL] {src.name} — {e}")
        if ppt:
            _quit_with_timeout(ppt, timeout=4, our_pid=our_pid)  # F-02
        return False


def convert_html(src: Path, dest: Path, log_fn=None) -> bool:
    """Convert .html/.htm to PDF.
    Strategy:
      1. Chrome headless full-page screenshot -> split into landscape A4 pages
      2. WeasyPrint fallback
      3. reportlab plain-text fallback
    """
    import tempfile
    import os as _os

    _log = log_fn or log  # F-20

    chrome_paths = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    ]
    chrome_exe = next((p for p in chrome_paths if Path(p).exists()), None)

    # ── 1. Chrome screenshot → PDF ────────────────────────────────────────────
    if chrome_exe:
        try:
            from PIL import Image as PILImage
            from reportlab.lib.pagesizes import A4, landscape
            from reportlab.platypus import SimpleDocTemplate, Image as RLImage
            from reportlab.lib.units import mm

            html_src = src.read_text(encoding="utf-8", errors="replace")
            inject = """<style>
* { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }
</style>"""
            if "</head>" in html_src:
                html_src = html_src.replace("</head>", inject + "</head>")
            else:
                html_src = inject + html_src

            # F-09: use system temp dir instead of source/output dirs for temp files
            fd, _tmp_html_path = tempfile.mkstemp(suffix=".html", prefix="docconv_")
            _os.close(fd)
            fd, _png_path = tempfile.mkstemp(suffix=".png", prefix="docconv_")
            _os.close(fd)
            tmp_html = Path(_tmp_html_path)
            png_path = Path(_png_path)
            strips = []

            try:
                tmp_html.write_text(html_src, encoding="utf-8")

                import json as _json

                subprocess.run([
                    chrome_exe,
                    "--headless=new",
                    "--disable-gpu",
                    # F-27: --no-sandbox removed — not required for standard Windows user accounts
                    "--hide-scrollbars",
                    "--window-size=1280,15000",
                    "--screenshot=" + str(png_path.resolve()),
                    "--virtual-time-budget=8000",
                    tmp_html.resolve().as_uri(),
                ], capture_output=True, timeout=60)

                if not png_path.exists() or png_path.stat().st_size == 0:
                    raise RuntimeError("Chrome produced no screenshot")

                page_w, page_h = landscape(A4)
                margin = 10 * mm
                avail_w = page_w - 2 * margin
                avail_h = page_h - 2 * margin

                with PILImage.open(str(png_path)) as im:
                    im = im.convert("RGB")
                    img_w, img_h = im.size

                    import numpy as np
                    arr = np.array(im, dtype=np.float32)
                    row_std  = arr.std(axis=(1, 2))
                    row_dark = arr.min(axis=(1, 2)) < 240
                    content_rows = np.where((row_std > 1.5) | row_dark)[0]
                    if len(content_rows) > 0:
                        last_content_row = int(content_rows[-1]) + 30
                        img_h = min(last_content_row, img_h)
                        im = im.crop((0, 0, img_w, img_h))

                    px_per_page = int(img_w * (avail_h / avail_w))
                    n_pages = max(1, -(-img_h // px_per_page))

                    for i in range(n_pages):
                        top    = i * px_per_page
                        bottom = min(top + px_per_page, img_h)
                        strip  = im.crop((0, top, img_w, bottom))
                        strip_h = bottom - top
                        if strip_h < px_per_page:
                            padded = PILImage.new("RGB", (img_w, px_per_page), (255, 255, 255))
                            padded.paste(strip, (0, 0))
                            strip = padded
                        strip_arr = np.array(strip, dtype=np.float32)
                        white_ratio = (strip_arr.min(axis=2) >= 240).mean()
                        if white_ratio > 0.95:
                            continue
                        # F-09: strip files in system temp dir, not output dir
                        fd, _strip_path = tempfile.mkstemp(
                            suffix=f"_strip{i}.png", prefix="docconv_")
                        _os.close(fd)
                        strip_path = Path(_strip_path)
                        strip.save(str(strip_path), "PNG")
                        strips.append(strip_path)

                from reportlab.platypus import BaseDocTemplate, Frame, PageTemplate

                frames = [Frame(margin, margin, avail_w, avail_h, leftPadding=0,
                                rightPadding=0, topPadding=0, bottomPadding=0)]
                pt  = PageTemplate(id="main", frames=frames, pagesize=landscape(A4))
                doc = BaseDocTemplate(
                    str(dest),
                    pagesize=landscape(A4),
                    leftMargin=margin, rightMargin=margin,
                    topMargin=margin, bottomMargin=margin,
                    pageTemplates=[pt],
                )
                story = [RLImage(str(s), width=avail_w, height=avail_h, kind="direct")
                         for s in strips]
                doc.build(story)

                if dest.exists() and dest.stat().st_size > 0:
                    _log(f"  [OK] {dest.name}")
                    return True
                raise RuntimeError("PDF assembly produced empty file")

            finally:
                # F-09: clean all system-temp files regardless of success/failure
                for p in [tmp_html, png_path] + strips:
                    try:
                        if p.exists():
                            p.unlink()
                    except OSError:
                        pass

        except Exception as e:
            _log(f"  [WARN] Chrome path failed ({e}) — trying WeasyPrint")

    # ── 2. WeasyPrint fallback ────────────────────────────────────────────────
    try:
        from weasyprint import HTML
        import tempfile
        import os as _os2

        html_src = src.read_text(encoding="utf-8", errors="replace")
        inject = """<style>
@page { size: 297mm 210mm; margin: 10mm; }
body { width: 277mm; margin: 0 auto; box-sizing: border-box; }
* { max-width: 100% !important; box-sizing: border-box !important;
    -webkit-print-color-adjust: exact !important; }
</style>"""
        if "</head>" in html_src:
            html_src = html_src.replace("</head>", inject + "</head>")
        else:
            html_src = inject + html_src

        # F-09: use system temp dir for WeasyPrint's intermediate HTML
        fd, _tmp_path = tempfile.mkstemp(suffix=".html", prefix="docconv_")
        _os2.close(fd)
        tmp = Path(_tmp_path)
        try:
            tmp.write_text(html_src, encoding="utf-8")
            HTML(filename=str(tmp.resolve())).write_pdf(str(dest.resolve()))
        finally:
            try:
                if tmp.exists():
                    tmp.unlink()
            except OSError:
                pass

        if dest.exists() and dest.stat().st_size > 0:
            _log(f"  [OK-weasyprint] {dest.name}")
            return True
    except ImportError:
        _log("  [WARN] WeasyPrint not available — trying reportlab")
    except Exception as e:
        _log(f"  [WARN] WeasyPrint failed ({e}) — trying reportlab")

    # ── 3. reportlab plain-text fallback ─────────────────────────────────────
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.units import cm
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer

        html = src.read_text(encoding="utf-8", errors="replace")
        html = re.sub(r"<script[^>]*>.*?</script>", " ", html, flags=re.DOTALL | re.IGNORECASE)
        html = re.sub(r"<style[^>]*>.*?</style>",   " ", html, flags=re.DOTALL | re.IGNORECASE)
        html = re.sub(r"<!--.*?-->", " ", html, flags=re.DOTALL)
        html = re.sub(r"<br\s*/?>", "\n", html, flags=re.IGNORECASE)
        html = re.sub(r"</p>|</div>|</li>|</h[1-6]>|</tr>", "\n", html, flags=re.IGNORECASE)
        html = re.sub(r"<[^>]+>", "", html)
        html = (html.replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">")
                    .replace("&nbsp;", " ").replace("&quot;", '"').replace("&#39;", "'"))
        lines = [l.rstrip() for l in html.splitlines()]
        lines = [l for i, l in enumerate(lines) if l or (i > 0 and lines[i - 1])]

        styles = getSampleStyleSheet()
        style_body = ParagraphStyle("Body", parent=styles["Normal"],
                                    fontSize=10, leading=15, spaceAfter=4)
        story = []
        for line in lines:
            line = line.strip()
            if not line:
                story.append(Spacer(1, 6))
            else:
                safe = line.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                story.append(Paragraph(safe, style_body))
        doc = SimpleDocTemplate(str(dest), pagesize=A4,
                                leftMargin=2.5 * cm, rightMargin=2.5 * cm,
                                topMargin=2.5 * cm, bottomMargin=2.5 * cm)
        doc.build(story)
        _log(f"  [OK-fallback] {dest.name}")
        return True
    except Exception as e:
        _log(f"  [FAIL] {src.name} — all methods failed: {e}")
        return False


def convert_md(src: Path, dest: Path, log_fn=None) -> bool:
    """Convert .md to PDF using reportlab (pure Python, no system libraries)."""
    _log = log_fn or log  # F-20
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.units import cm
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, HRFlowable
        from reportlab.lib import colors

        md_text = src.read_text(encoding="utf-8", errors="replace")
        lines   = md_text.splitlines()

        styles = getSampleStyleSheet()
        style_h1   = ParagraphStyle("H1", parent=styles["Heading1"],
                                    fontSize=18, spaceAfter=8,
                                    textColor=colors.HexColor("#1a1a1a"))
        style_h2   = ParagraphStyle("H2", parent=styles["Heading2"],
                                    fontSize=14, spaceAfter=6,
                                    textColor=colors.HexColor("#1a1a1a"))
        style_h3   = ParagraphStyle("H3", parent=styles["Heading3"],
                                    fontSize=12, spaceAfter=4,
                                    textColor=colors.HexColor("#1a1a1a"))
        style_body = ParagraphStyle("Body", parent=styles["Normal"],
                                    fontSize=10, leading=15, spaceAfter=4,
                                    fontName="Helvetica")
        style_code = ParagraphStyle("Code", parent=styles["Code"],
                                    fontSize=9, leading=13, fontName="Courier",
                                    backColor=colors.HexColor("#f4f4f4"), spaceAfter=4)
        style_bullet = ParagraphStyle("Bullet", parent=styles["Normal"],
                                      fontSize=10, leading=15, leftIndent=16,
                                      spaceAfter=2, fontName="Helvetica")

        story = []
        in_code_block = False
        code_lines    = []

        def flush_code():
            if code_lines:
                text = "<br/>".join(code_lines)
                story.append(Paragraph(text, style_code))
                story.append(Spacer(1, 4))
                code_lines.clear()

        def md_inline(text):
            text = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            text = re.sub(r"`([^`]+)`", r'<font name="Courier" size="9">\1</font>', text)
            text = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", text)
            text = re.sub(r"__(.+?)__",     r"<b>\1</b>", text)
            text = re.sub(r"\*(.+?)\*",     r"<i>\1</i>", text)
            text = re.sub(r"_(.+?)_",       r"<i>\1</i>", text)
            text = re.sub(r"\[([^\]]+)\]\([^\)]+\)", r"\1", text)
            return text

        for line in lines:
            if line.strip().startswith("```"):
                if not in_code_block:
                    in_code_block = True
                else:
                    in_code_block = False
                    flush_code()
                continue
            if in_code_block:
                code_lines.append(
                    line.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))
                continue
            if line.startswith("# "):
                story.append(Paragraph(md_inline(line[2:].strip()), style_h1))
                story.append(HRFlowable(width="100%", thickness=0.5,
                                        color=colors.HexColor("#cccccc"), spaceAfter=6))
                continue
            if line.startswith("## "):
                story.append(Paragraph(md_inline(line[3:].strip()), style_h2))
                continue
            if line.startswith("### "):
                story.append(Paragraph(md_inline(line[4:].strip()), style_h3))
                continue
            if re.match(r"^[-*_]{3,}$", line.strip()):
                story.append(HRFlowable(width="100%", thickness=0.5,
                                        color=colors.HexColor("#cccccc"), spaceAfter=4))
                continue
            if re.match(r"^[-*+] ", line):
                story.append(Paragraph(f"• {md_inline(line[2:].strip())}", style_bullet))
                continue
            if re.match(r"^\d+\. ", line):
                story.append(Paragraph(
                    md_inline(re.sub(r"^\d+\. ", "", line).strip()), style_bullet))
                continue
            if line.strip() == "":
                story.append(Spacer(1, 6))
                continue
            story.append(Paragraph(md_inline(line.strip()), style_body))

        flush_code()

        doc = SimpleDocTemplate(str(dest), pagesize=A4,
                                leftMargin=2.5 * cm, rightMargin=2.5 * cm,
                                topMargin=2.5 * cm, bottomMargin=2.5 * cm)
        doc.build(story)
        _log(f"  [OK] {dest.name}")
        return True

    except ImportError:
        _log("  [ERROR] reportlab not installed. Run: pip install reportlab")
        return False
    except Exception as e:
        _log(f"  [ERROR] Markdown conversion failed: {e}")
        return False


def convert_txt(src: Path, dest: Path, log_fn=None) -> bool:
    """Convert .txt to PDF using reportlab (pure Python)."""
    _log = log_fn or log  # F-20
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.units import cm
        from reportlab.pdfgen import canvas as rl_canvas

        raw_lines = src.read_text(encoding="utf-8", errors="replace").splitlines()
        page_w, page_h = A4
        margin      = 2.5 * cm
        font_size   = 10
        line_height = font_size * 1.4

        c = rl_canvas.Canvas(str(dest), pagesize=A4)
        c.setFont("Courier", font_size)
        y = page_h - margin

        for raw_line in raw_lines:
            # F-34: wrap long lines instead of truncating at 120 characters
            wrapped = textwrap.wrap(raw_line, width=120) or [""]
            for line in wrapped:
                if y < margin:
                    c.showPage()
                    c.setFont("Courier", font_size)
                    y = page_h - margin
                c.drawString(margin, y, line)
                y -= line_height

        c.save()
        _log(f"  [OK] {dest.name}")
        return True
    except ImportError:
        _log("  [ERROR] reportlab not installed. Run: pip install reportlab")
        return False
    except Exception as e:
        _log(f"  [ERROR] Text conversion failed: {e}")
        return False


def convert_csv(src: Path, dest: Path, log_fn=None) -> bool:
    """Convert .csv to PDF as a formatted table using reportlab."""
    _log = log_fn or log  # F-20
    try:
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.lib.units import cm
        from reportlab.lib import colors
        from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

        with open(str(src), newline="", encoding="utf-8", errors="replace") as f:
            rows = list(csv.reader(f))

        if not rows:
            _log("  [WARN] CSV file is empty.")
            return False

        styles = getSampleStyleSheet()
        style_title = styles["Heading2"]
        style_cell  = ParagraphStyle("Cell", parent=styles["Normal"], fontSize=8)

        table_data = []
        for row in rows:
            table_data.append([Paragraph(str(cell), style_cell) for cell in row])

        page_w, page_h = landscape(A4)
        usable_w  = page_w - 4 * cm
        col_count = len(rows[0]) if rows else 1
        col_w     = usable_w / col_count

        table = Table(table_data, colWidths=[col_w] * col_count, repeatRows=1)
        table.setStyle(TableStyle([
            ("BACKGROUND",     (0, 0), (-1, 0),  colors.HexColor("#2c3e50")),
            ("TEXTCOLOR",      (0, 0), (-1, 0),  colors.white),
            ("FONTNAME",       (0, 0), (-1, 0),  "Helvetica-Bold"),
            ("FONTSIZE",       (0, 0), (-1, 0),  8),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f2f2f2")]),
            ("FONTNAME",       (0, 1), (-1, -1), "Helvetica"),
            ("FONTSIZE",       (0, 1), (-1, -1), 8),
            ("GRID",           (0, 0), (-1, -1), 0.4, colors.HexColor("#cccccc")),
            ("VALIGN",         (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING",    (0, 0), (-1, -1), 4),
            ("RIGHTPADDING",   (0, 0), (-1, -1), 4),
            ("TOPPADDING",     (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING",  (0, 0), (-1, -1), 3),
        ]))

        story = [Paragraph(src.name, style_title), Spacer(1, 8), table]
        doc = SimpleDocTemplate(str(dest), pagesize=landscape(A4),
                                leftMargin=2 * cm, rightMargin=2 * cm,
                                topMargin=2 * cm, bottomMargin=2 * cm)
        doc.build(story)
        _log(f"  [OK] {dest.name}")
        return True

    except ImportError:
        _log("  [ERROR] reportlab not installed. Run: pip install reportlab")
        return False
    except Exception as e:
        _log(f"  [ERROR] CSV conversion failed: {e}")
        return False


def convert_xls(src: Path, dest: Path, log_fn=None) -> bool:
    """Convert legacy .xls to PDF using xlrd (read) + reportlab (write)."""
    _log = log_fn or log  # F-20
    try:
        import xlrd
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.lib.units import cm
        from reportlab.lib import colors
        from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

        wb = xlrd.open_workbook(str(src))
        styles     = getSampleStyleSheet()
        style_title = styles["Heading2"]
        style_cell  = ParagraphStyle("Cell", parent=styles["Normal"], fontSize=8)

        story = []
        for sheet in wb.sheets():
            rows = [[str(sheet.cell_value(r, c)) for c in range(sheet.ncols)]
                    for r in range(sheet.nrows)]
            if not rows:
                continue
            story.append(Paragraph(f"{src.name} — Sheet: {sheet.name}", style_title))
            story.append(Spacer(1, 8))
            table_data = [[Paragraph(cell, style_cell) for cell in row] for row in rows]
            page_w, page_h = landscape(A4)
            usable_w  = page_w - 4 * cm
            col_count = len(rows[0]) if rows else 1
            col_w     = usable_w / col_count
            table = Table(table_data, colWidths=[col_w] * col_count, repeatRows=1)
            table.setStyle(TableStyle([
                ("BACKGROUND",     (0, 0), (-1, 0),  colors.HexColor("#2c3e50")),
                ("TEXTCOLOR",      (0, 0), (-1, 0),  colors.white),
                ("FONTNAME",       (0, 0), (-1, 0),  "Helvetica-Bold"),
                ("FONTSIZE",       (0, 0), (-1, 0),  8),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f2f2f2")]),
                ("FONTNAME",       (0, 1), (-1, -1), "Helvetica"),
                ("FONTSIZE",       (0, 1), (-1, -1), 8),
                ("GRID",           (0, 0), (-1, -1), 0.4, colors.HexColor("#cccccc")),
                ("VALIGN",         (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING",    (0, 0), (-1, -1), 4),
                ("RIGHTPADDING",   (0, 0), (-1, -1), 4),
                ("TOPPADDING",     (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING",  (0, 0), (-1, -1), 3),
            ]))
            story.append(table)
            story.append(Spacer(1, 16))

        if not story:
            _log("  [WARN] .xls file has no data.")
            return False

        doc = SimpleDocTemplate(str(dest), pagesize=landscape(A4),
                                leftMargin=2 * cm, rightMargin=2 * cm,
                                topMargin=2 * cm, bottomMargin=2 * cm)
        doc.build(story)
        _log(f"  [OK] {dest.name}")
        return True

    except ImportError:
        _log("  [ERROR] xlrd or reportlab not installed. Run: pip install xlrd reportlab")
        return False
    except Exception as e:
        _log(f"  [ERROR] XLS conversion failed: {e}")
        return False


# ── PDF → Word converter ──────────────────────────────────

def _auto_dismiss_word_dialog(stop_event, timeout=60):
    """Background thread: auto-clicks OK on Word's PDF conversion info dialog.

    Verifies the window belongs to WINWORD.EXE before clicking (F-23).
    """
    import time
    try:
        import win32gui, win32con, win32api, win32process
    except ImportError:
        return
    deadline = time.time() + timeout
    while not stop_event.is_set() and time.time() < deadline:
        def _cb(hwnd, ctx):
            if not win32gui.IsWindowVisible(hwnd):
                return True
            title = win32gui.GetWindowText(hwnd)
            if title not in ("", "?", "Microsoft Word"):
                return True
            # F-23: verify the window belongs to WINWORD.EXE, not another application
            try:
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                if _PSUTIL_AVAILABLE:
                    if _psutil.Process(pid).name().lower() != "winword.exe":
                        return True   # not Word — leave it alone
            except Exception:
                return True
            try:
                child = win32gui.FindWindowEx(hwnd, 0, "Button", None)
                while child:
                    if win32gui.GetWindowText(child).strip().upper() in ("OK", "&OK"):
                        win32api.SendMessage(child, win32con.BM_CLICK, 0, 0)
                        ctx.append(hwnd)
                        return False
                    child = win32gui.FindWindowEx(hwnd, child, "Button", None)
            except Exception:
                pass
            return True

        found = []
        try:
            win32gui.EnumWindows(_cb, found)
        except Exception:
            pass
        if found:
            stop_event.set()
            return
        time.sleep(0.2)


def convert_pdf_to_docx(src: Path, dest: Path, log_fn=None) -> bool:
    """Convert a PDF file to .docx.

    Strategy (three tiers):
      1. Word COM — highest fidelity (photos, layout, tables intact).
      2. pymupdf + python-docx — text-only fallback.
      3. pytesseract OCR — scanned/image-only PDF fallback.
    """
    import shutil
    _log = log_fn or log  # F-20

    # ── Tier 1: Word COM (high fidelity) ─────────────────────────────────────
    word = None
    our_pid = None
    try:
        import win32com.client
        stop_evt  = threading.Event()
        dismiss_t = threading.Thread(
            target=_auto_dismiss_word_dialog, args=(stop_evt,), daemon=True)
        dismiss_t.start()

        # F-02: remove stale output if present; no system-wide taskkill
        if dest.exists():
            try:
                dest.unlink()
            except Exception:
                pass

        pids_before = _snapshot_pids('winword.exe')  # F-02
        word = win32com.client.Dispatch("Word.Application")
        word.AutomationSecurity = 3  # F-08: disable all macros
        our_pid = _get_new_office_pid('winword.exe', pids_before)  # F-02
        word.Visible = False
        try:
            word.WindowState = 2   # wdWindowStateMinimize
        except Exception:
            pass
        _word_suppress_dialogs(word)
        doc = word.Documents.Open(
            str(src.resolve()),
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
        )
        stop_evt.set()
        doc.SaveAs2(str(dest.resolve()), FileFormat=16)  # wdFormatDocx = 16
        doc.Close(SaveChanges=False)
        _quit_with_timeout(word, timeout=8, our_pid=our_pid)  # F-02: kill only our Word
        _log(f"  [OK] {dest.name}")
        return True
    except ImportError:
        _log("  [WARN] pywin32 not installed — trying pymupdf")
    except Exception as e:
        # F-21: distinguish password-protected PDFs
        err_str = str(e).lower()
        hresult  = getattr(e, 'hresult', 0)
        if "password" in err_str or "protected" in err_str or hresult == -2146822611:
            _log(f"  [SKIP] {src.name} — password-protected file, cannot convert.")
            if word:
                _quit_with_timeout(word, timeout=4, our_pid=our_pid)
            return False
        _log(f"  [WARN] Word COM failed ({e}) — trying pymupdf")
        if word:
            _quit_with_timeout(word, timeout=4, our_pid=our_pid)  # F-02

    # ── Tier 2: pymupdf + python-docx ────────────────────────────────────────
    try:
        import pymupdf as fitz
        from docx import Document
        from docx.shared import Pt

        pdf = fitz.open(str(src.resolve()))
        doc = Document()
        doc.add_heading(src.stem, level=1)

        for page_num in range(len(pdf)):
            page = pdf[page_num]
            text = page.get_text("text")
            if not text.strip():
                continue
            if page_num > 0:
                doc.add_page_break()
                doc.add_heading(f"Page {page_num + 1}", level=2)
            for line in text.splitlines():
                line = line.strip()
                if not line:
                    continue
                para = doc.add_paragraph(line)
                para.style.font.size = Pt(11)

        pdf.close()
        doc.save(str(dest.resolve()))
        _log(f"  [OK-pymupdf] {dest.name}")
        return True
    except ImportError:
        _log("  [WARN] pymupdf or python-docx not installed — trying OCR")
    except Exception as e:
        _log(f"  [WARN] pymupdf extraction failed ({e}) — trying OCR")

    # ── Tier 3: pytesseract OCR (for scanned PDFs) ───────────────────────────
    try:
        import pymupdf as fitz
        from PIL import Image as PILImage
        import pytesseract
        from docx import Document
        from docx.shared import Pt
        import io

        try:
            pytesseract.get_tesseract_version()
        except Exception:
            _log("  [ERROR] Tesseract OCR not installed. "
                 "Download from: https://github.com/UB-Mannheim/tesseract/wiki")
            return False

        pdf = fitz.open(str(src.resolve()))
        doc = Document()
        doc.add_heading(src.stem, level=1)

        for page_num in range(len(pdf)):
            page = pdf[page_num]
            mat = fitz.Matrix(300 / 72, 300 / 72)
            pix = page.get_pixmap(matrix=mat, colorspace=fitz.csRGB)
            img  = PILImage.open(io.BytesIO(pix.tobytes("png")))
            text = pytesseract.image_to_string(img, lang="eng")
            if not text.strip():
                continue
            if page_num > 0:
                doc.add_page_break()
                doc.add_heading(f"Page {page_num + 1}", level=2)
            for line in text.splitlines():
                line = line.strip()
                if not line:
                    continue
                para = doc.add_paragraph(line)
                para.style.font.size = Pt(11)

        pdf.close()
        doc.save(str(dest.resolve()))
        _log(f"  [OK-OCR] {dest.name}")
        return True
    except ImportError:
        _log("  [ERROR] pytesseract / pymupdf / python-docx not installed. "
             "Run: pip install pymupdf pytesseract python-docx Pillow")
        return False
    except Exception as e:
        _log(f"  [FAIL] {src.name} — all PDF->Word methods failed: {e}")
        return False


# ── PDF → Excel converter ─────────────────────────────────

def convert_pdf_to_xlsx(src: Path, dest: Path, log_fn=None) -> bool:
    """Convert a PDF to .xlsx using word-position extraction for high fidelity.

    Strategy:
      1. pdfplumber extract_words() — groups words by Y-position to reconstruct rows.
      2. Fallback: raw text lines split into columns if word extraction fails.
    """
    _log = log_fn or log  # F-20
    try:
        import pdfplumber
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter
        from collections import defaultdict

        HEADER_FILL  = PatternFill("solid", fgColor="2C3E50")
        HEADER_FONT  = Font(color="FFFFFF", bold=True, size=10)
        ROW_FILL_ALT = PatternFill("solid", fgColor="F2F2F2")
        CELL_FONT    = Font(size=10)
        THIN = Border(
            left=Side(style="thin", color="CCCCCC"),
            right=Side(style="thin", color="CCCCCC"),
            top=Side(style="thin", color="CCCCCC"),
            bottom=Side(style="thin", color="CCCCCC"),
        )
        CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
        WRAP   = Alignment(wrap_text=True, vertical="top")

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Data"

        all_rows         = []
        header_row       = None
        first_data_words = None

        with pdfplumber.open(str(src.resolve())) as pdf:
            for page_num, page in enumerate(pdf.pages, 1):
                words = page.extract_words(
                    x_tolerance=3, y_tolerance=3, keep_blank_chars=False)
                if not words:
                    continue

                rows_by_y = defaultdict(list)
                for w in words:
                    y_key = round(w["top"] / 3) * 3
                    rows_by_y[y_key].append(w)

                sorted_ys = sorted(rows_by_y.keys())

                for y_idx, y in enumerate(sorted_ys):
                    row_words = sorted(rows_by_y[y], key=lambda w: w["x0"])
                    tokens    = [w["text"].strip() for w in row_words if w["text"].strip()]
                    if not tokens:
                        continue

                    is_header = (
                        header_row is None
                        and not any(t[0].isdigit() for t in tokens)
                        and any(kw in " ".join(tokens).lower()
                                for kw in ["date", "branch", "sales", "transaction",
                                           "count", "total", "avg"])
                    )

                    if is_header:
                        header_row = row_words
                    else:
                        if first_data_words is None and page_num == 1:
                            first_data_words = row_words
                        all_rows.append(tokens)

        if not all_rows:
            _log(f"  [WARN] No data extracted — trying text fallback")
            raise ValueError("no rows")

        from collections import Counter
        col_counts = Counter(len(r) for r in all_rows)
        num_cols   = col_counts.most_common(1)[0][0]

        if header_row and first_data_words:
            col_centres  = [(w["x0"] + w["x1"]) / 2 for w in first_data_words]
            final_header = [""] * len(col_centres)
            for hw in header_row:
                hc      = (hw["x0"] + hw["x1"]) / 2
                nearest = min(range(len(col_centres)),
                              key=lambda i: abs(col_centres[i] - hc))
                sep = " " if final_header[nearest] else ""
                final_header[nearest] += sep + hw["text"]
            if len(final_header) < num_cols:
                final_header += [f"Col{i+1}" for i in range(len(final_header), num_cols)]
            final_header = final_header[:num_cols]
            import re as _re
            def _clean_hdr(h):
                h = _re.sub(r'(?i)TCAvg.*', 'TC', h).strip()
                h = _re.sub(r'(?i)DailyA\s*Svagle', 'Avg Daily Sales', h).strip()
                h = _re.sub(r'(?i)Csheque', 'Cheque', h).strip()
                h = _re.sub(r'\s+', ' ', h).strip()
                return h
            final_header = [_clean_hdr(h) for h in final_header]
        else:
            final_header = [f"Column {get_column_letter(i+1)}" for i in range(num_cols)]

        for col_idx, hdr in enumerate(final_header, 1):
            cell = ws.cell(row=1, column=col_idx, value=hdr)
            cell.fill      = HEADER_FILL
            cell.font      = HEADER_FONT
            cell.alignment = CENTER
            cell.border    = THIN
        ws.freeze_panes = "A2"

        xlsx_row = 2
        for row in all_rows:
            if len(row) == 0:
                continue
            padded = (row + [""] * num_cols)[:num_cols]
            for col_idx, val in enumerate(padded, 1):
                cell = ws.cell(row=xlsx_row, column=col_idx, value=val)
                cell.font      = CELL_FONT
                cell.fill      = ROW_FILL_ALT if xlsx_row % 2 == 0 else PatternFill()
                cell.alignment = WRAP
                cell.border    = THIN
            xlsx_row += 1

        for col in ws.columns:
            max_len    = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                try:
                    max_len = max(max_len, len(str(cell.value or "")))
                except Exception:
                    pass
            ws.column_dimensions[col_letter].width = min(max_len + 4, 40)

        wb.save(str(dest.resolve()))
        _log(f"  [OK] {dest.name} — {xlsx_row-2} data rows, {num_cols} columns")
        return True

    except ImportError as e:
        _log(f"  [ERROR] Missing library: {e}")
        return False
    except Exception as e:
        _log(f"  [FAIL] {src.name} — {e}")
        return False


def close_office_apps():
    """Force-close Word, Excel, and PowerPoint to release file locks.

    Retained ONLY as an atexit emergency handler (F-24).
    Not called routinely after batches — per-PID cleanup (F-02) handles that.
    """
    apps = {
        "WINWORD.EXE":  "Word",
        "EXCEL.EXE":    "Excel",
        "POWERPNT.EXE": "PowerPoint",
    }
    for exe, name in apps.items():
        result = subprocess.run(
            ["taskkill", "/F", "/IM", exe],
            capture_output=True, text=True
        )
        if "SUCCESS" in result.stdout:
            log(f"  [INFO] Closed {name} (emergency cleanup on exit)")


def cleanup_output(output_dir: Path = None):
    """Remove converter temp files (docconv_ prefix) from the output folder.

    F-16: accepts output_dir parameter so the correct folder is scanned.
    F-22: splwow64.exe kill removed — document cleanup does not require it.
    F-39: sleep only after a failed deletion attempt, not unconditionally.
    """
    import time
    target_dir   = output_dir or OUTPUT_DIR
    # F-16: only delete files with known converter temp prefixes — not all non-PDF files
    TEMP_PATTERNS = ("docconv_", "__tmp_", "__strip", "__shot")

    for attempt in range(10):
        remaining = []
        removed   = []
        try:
            for f in target_dir.iterdir():
                if f.is_file() and any(f.name.startswith(p) for p in TEMP_PATTERNS):
                    try:
                        f.unlink()
                        removed.append(f.name)
                    except Exception:
                        remaining.append(f.name)
        except Exception:
            break
        if removed:
            log(f"  [INFO] Cleaned temp files: {', '.join(removed)}")
        if not remaining:
            break   # all gone — done
        # F-39: sleep only when files are still locked — never unconditionally
        time.sleep(2)
    if remaining:
        log(f"  [WARN] Could not remove locked temp files: {', '.join(remaining)}")


# ── Router ────────────────────────────────────────────────
CONVERTERS = {
    ".docx": convert_docx,
    ".doc":  convert_doc,
    ".xlsx": convert_xlsx,
    ".xls":  convert_xls,
    ".pptx": convert_pptx,
    ".ppt":  convert_pptx,
    ".html": convert_html,
    ".htm":  convert_html,
    ".md":   convert_md,
    ".txt":  convert_txt,
    ".csv":  convert_csv,
    # PDF source -> Office output (mode selected at runtime)
    ".pdf":  None,   # resolved dynamically in convert_file()
}


def convert_file(src: Path, pdf_mode: str = None, log_fn=None) -> bool:
    _log = log_fn or log  # F-20
    ext = src.suffix.lower()

    # PDF source files route to Word or Excel depending on mode
    if ext == ".pdf":
        mode = pdf_mode or PDF_OUTPUT_MODE
        if mode == "excel":
            dest = _safe_dest(OUTPUT_DIR, src.stem, ".xlsx")  # F-07
            _log(f"  [CONVERTING] {src.name} -> {dest.name} ...")
            in_bytes = src.stat().st_size  # F-06
            success  = convert_pdf_to_xlsx(src, dest, log_fn=_log)  # F-20
        else:
            dest = _safe_dest(OUTPUT_DIR, src.stem, ".docx")  # F-07
            _log(f"  [CONVERTING] {src.name} -> {dest.name} ...")
            in_bytes = src.stat().st_size  # F-06
            success  = convert_pdf_to_docx(src, dest, log_fn=_log)  # F-20

        # F-06: zero-byte output treated as failure
        if success:
            out_bytes = dest.stat().st_size if dest.exists() else 0
            if out_bytes == 0:
                success = False
                try:
                    dest.unlink()
                except OSError:
                    pass
                _log(f"  [FAIL] {src.name} — output is 0 bytes (conversion failed silently).")
            else:
                _log(f"  [SIZE] Input: {in_bytes:,} B  ->  Output: {out_bytes:,} B")
                _copy_file_times(src, dest)  # F-26

        if not success:
            _log(f"  [FAIL] {src.name} could not be converted.")
        return success

    converter = CONVERTERS.get(ext)
    if converter is None:
        _log(f"  [SKIP] Unsupported format: {src.name}")
        return False

    dest     = _safe_dest(OUTPUT_DIR, src.stem, ".pdf")  # F-07
    _log(f"  [CONVERTING] {src.name} -> {dest.name} ...")
    in_bytes = src.stat().st_size  # F-06
    success  = converter(src, dest, log_fn=_log)  # F-20

    # F-06: zero-byte output treated as failure
    if success:
        out_bytes = dest.stat().st_size if dest.exists() else 0
        if out_bytes == 0:
            success = False
            try:
                dest.unlink()
            except OSError:
                pass
            _log(f"  [FAIL] {src.name} — output is 0 bytes (conversion failed silently).")
        else:
            _log(f"  [SIZE] Input: {in_bytes:,} B  ->  Output: {out_bytes:,} B")
            _copy_file_times(src, dest)  # F-26

    if not success:
        _log(f"  [FAIL] {src.name} could not be converted.")
    return success


# ── Main ──────────────────────────────────────────────────
def run(single_file=None):
    INPUT_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    if single_file:
        src = Path(single_file)
        if not src.is_absolute():
            src = INPUT_DIR / src
        # F-33: path traversal guard — file must be inside INPUT_DIR
        try:
            src.resolve().relative_to(INPUT_DIR.resolve())
        except ValueError:
            log(f"[ERROR] File must be inside the input directory: {INPUT_DIR}")
            sys.exit(1)
        if not src.exists():
            log(f"[ERROR] File not found: {src}")
            sys.exit(1)
        files = [src]
    else:
        files = sorted(
            f for f in INPUT_DIR.iterdir()
            if f.is_file()
            and f.suffix.lower() in SUPPORTED
            and not f.name.startswith("~$")
        )

    if not files:
        log("[INFO] No supported files found in input\\ folder.")
        log(f"[INFO] Supported formats: {', '.join(sorted(SUPPORTED))}")
        return

    log("=" * 52)
    log("PDF Converter started")
    log(f"Input  : {INPUT_DIR}")
    log(f"Output : {OUTPUT_DIR}")
    log(f"Files  : {len(files)} found")
    log("=" * 52)

    ok = fail = 0
    for f in files:
        if convert_file(f):
            ok += 1
        else:
            fail += 1

    # F-24: close_office_apps() removed from routine batch end — retained only as atexit handler
    cleanup_output()  # F-16: uses default OUTPUT_DIR (now correctly resolved via F-01)
    log("-" * 52)
    log(f"Done.  Converted: {ok}  |  Failed: {fail}")
    log(f"PDFs saved to: {OUTPUT_DIR}")
    log("=" * 52)


if __name__ == "__main__":
    import atexit
    import signal

    # F-24: register emergency Office cleanup on unexpected exit
    def _emergency_cleanup():
        try:
            close_office_apps()
        except Exception:
            pass

    atexit.register(_emergency_cleanup)
    try:
        signal.signal(signal.SIGTERM, lambda *_: (_emergency_cleanup(), sys.exit(0)))
    except (OSError, ValueError):
        pass

    parser = argparse.ArgumentParser(description="Convert documents to PDF.")
    parser.add_argument("--file", metavar="FILENAME",
                        help="Convert a single file from the input\\ folder.")
    args = parser.parse_args()
    run(single_file=args.file)
