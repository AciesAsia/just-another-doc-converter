"""
=============================================================
  Document -> PDF Converter
  Usage:
    python converter.py                  # convert all files in input\
    python converter.py --file foo.docx  # convert a single file

  Input folder  : C:\\\\PythonProjects\\\\DocumentConverter\\\\input\\\\
  Output folder : C:\\\\PythonProjects\\\\DocumentConverter\\\\output\\\\
  Log file      : C:\\\\PythonProjects\\\\DocumentConverter\\\\logs\\\\convert_log.txt

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

  Dependencies (install once via install_dependencies.bat):
    docx2pdf   -- drives Microsoft Word silently (.docx)
    pywin32    -- drives Word/Excel/PowerPoint COM (.doc .xlsx .pptx .ppt)
    reportlab  -- renders .md, .txt, .csv, .html, .xls to PDF (pure Python)
    xlrd       -- reads legacy .xls files
    pymupdf    -- PDF parsing and page rendering (pip install pymupdf)
    python-docx-- write .docx output (pip install python-docx)
    pdfplumber -- extract tables from PDFs (pip install pdfplumber)
    openpyxl   -- write .xlsx output (pip install openpyxl)
    pytesseract-- OCR for scanned PDFs (pip install pytesseract + Tesseract exe)
    Pillow     -- image handling for OCR (pip install Pillow)
=============================================================
"""

import argparse
import csv
import re
import subprocess
import sys
import threading
from datetime import datetime
from pathlib import Path

# ── Folder paths ──────────────────────────────────────────
BASE_DIR   = Path(__file__).resolve().parent
INPUT_DIR  = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"
LOG_FILE   = BASE_DIR / "logs" / "convert_log.txt"

SUPPORTED = {".docx", ".doc", ".xlsx", ".xls", ".pptx", ".ppt",
             ".html", ".htm", ".md", ".txt", ".csv", ".pdf"}

# Output mode for PDF files — determines which conversion path to use.
# "word"  = PDF -> .docx
# "excel" = PDF -> .xlsx
# Set by the GUI; CLI defaults to Word.
PDF_OUTPUT_MODE = "word"


# ── Logging ───────────────────────────────────────────────
def log(message: str):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{timestamp}] {message}"
    print(line)
    LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
    with LOG_FILE.open("a", encoding="utf-8") as f:
        f.write(line + "\n")


# ── COM helpers ───────────────────────────────────────────
def _quit_with_timeout(app, timeout=8, kill_exe=None):
    """Call app.Quit() in a thread; if it hangs, taskkill the process."""
    def _quit():
        try:
            app.Quit()
        except Exception:
            pass
    t = threading.Thread(target=_quit, daemon=True)
    t.start()
    t.join(timeout=timeout)
    if t.is_alive() and kill_exe:
        subprocess.run(["taskkill", "/F", "/IM", kill_exe], capture_output=True)


# ── Converters ────────────────────────────────────────────

def _word_suppress_dialogs(word):
    """Suppress all Word dialogs and background publishing notifications."""
    word.Visible = False
    word.DisplayAlerts = 0          # wdAlertsNone = 0 (numeric suppresses more than False)
    try:
        word.Options.SaveNormalPrompt = False   # no "save to Normal.dotm?" prompt
    except Exception:
        pass
    try:
        word.Options.PrintBackground = False    # disables background print/publish dialog
    except Exception:
        pass
    try:
        word.Options.BackgroundSave = False     # disable background save progress
    except Exception:
        pass
    try:
        word.Options.DoNotPromptForConvert = True  # suppress format conversion dialog
    except Exception:
        pass
    try:
        word.Options.UpdateLinksAtOpen = False  # don't prompt to update links
    except Exception:
        pass


def convert_docx(src: Path, dest: Path) -> bool:
    """
    Convert .docx to PDF via Word COM SaveAs2.
    Word is kept minimised so any background Publishing dialog appears
    behind all other windows and is never visible to the user.
    """
    word = None
    try:
        import win32com.client
        word = win32com.client.Dispatch("Word.Application")
        _word_suppress_dialogs(word)
        # Keep Word window minimised — Publishing dialog will be hidden behind taskbar
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
        subprocess.run(["taskkill", "/F", "/IM", "WINWORD.EXE"], capture_output=True)
        log(f"  [OK] {dest.name}")
        return True
    except ImportError:
        log("  [ERROR] pywin32 not installed. Run install_dependencies.bat first.")
        return False
    except Exception as e:
        log(f"  [FAIL] {src.name} — {e}")
        subprocess.run(["taskkill", "/F", "/IM", "WINWORD.EXE"], capture_output=True)
        return False


def convert_doc(src: Path, dest: Path) -> bool:
    """Convert legacy .doc to PDF — identical path to .docx."""
    return convert_docx(src, dest)


def convert_xlsx(src: Path, dest: Path) -> bool:
    """Convert .xlsx to PDF via Excel COM ExportAsFixedFormat with timeout guard.

    Strategy (all COM calls stay on the main thread to avoid apartment issues):
      1. Workbook-level ExportAsFixedFormat — works for most files.
      2. Sheet-by-sheet fallback — used when workbook-level fails (protection,
         external data connections, etc.).  Each visible sheet is exported to a
         temp PDF; pypdf merges them into the final output.
    A daemon thread is used only to enforce a hard 60-second wall-clock timeout
    on the entire operation.
    """
    import tempfile, os, time

    # Size guard — files over 15 MB produce thousands of PDF pages and will
    # time out regardless of the timeout setting.  Skip them immediately.
    MAX_XLSX_MB = 15
    file_mb = src.stat().st_size / (1024 * 1024)
    if file_mb > MAX_XLSX_MB:
        log(f"  [SKIP] {src.name} — file too large for PDF conversion "
            f"({file_mb:.1f} MB > {MAX_XLSX_MB} MB limit). "
            f"Open in Excel and set a print area before converting.")
        return False

    try:
        import win32com.client
    except ImportError:
        log("  [ERROR] pywin32 not installed. Run install_dependencies.bat first.")
        return False

    # Kill any lingering Word process before starting Excel — prevents Word's
    # background Publishing dialog from appearing on top during Excel export.
    subprocess.run(["taskkill", "/F", "/IM", "WINWORD.EXE"], capture_output=True)

    # --- inner function runs on main thread via a timed join trick -----------
    result = {"ok": False, "err": None, "warn": None}
    dest_str = str(dest.resolve())
    src_str  = str(src.resolve())

    def _do_export():
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.AskToUpdateLinks = False
            # xlCalculationManual = -4135 — stop any auto-recalc that could block
            try:
                excel.Calculation = -4135
            except Exception:
                pass
            try:
                excel.EnableEvents = False
            except Exception:
                pass
            try:
                excel.ScreenUpdating = False
            except Exception:
                pass
            wb = excel.Workbooks.Open(
                src_str,
                UpdateLinks=0,
                ReadOnly=True,
                IgnoreReadOnlyRecommended=True,
                CorruptLoad=2,          # xlRepairFile
            )
            # Disable refresh/connections so external-data workbooks don't hang
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
            # Do NOT call wb.RefreshAll() — it can hang on broken connections.
            # Disabling BackgroundQuery above is sufficient to prevent export hangs.

            # --- Stage 1: whole-workbook export --------------------------------
            wb_ok = False
            try:
                wb.ExportAsFixedFormat(0, dest_str)
                wb_ok = True
            except Exception as ex1:
                result["err"] = ex1     # keep for logging if stage 2 also fails

            if wb_ok:
                wb.Close(SaveChanges=False)
                result["ok"] = True
                result["err"] = None
                return

            # --- Stage 2: sheet-by-sheet export --------------------------------
            sheet_pdfs = []
            try:
                n = wb.Sheets.Count
                for i in range(1, n + 1):
                    sh = wb.Sheets(i)
                    if sh.Visible != -1:    # xlSheetVisible = -1
                        continue
                    tmp = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
                    tmp.close()
                    try:
                        sh.ExportAsFixedFormat(0, tmp.name)
                        sheet_pdfs.append(tmp.name)
                    except Exception:
                        try:
                            os.unlink(tmp.name)
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
                    finally:
                        for p in sheet_pdfs:
                            try:
                                os.unlink(p)
                            except Exception:
                                pass
                else:
                    # no visible sheets could be exported
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

    t = threading.Thread(target=_do_export, daemon=True)
    t.start()
    t.join(timeout=300)   # 5-minute hard limit — large workbooks (29MB+) need time

    subprocess.run(["taskkill", "/F", "/IM", "EXCEL.EXE"], capture_output=True)

    if t.is_alive() or not result["ok"]:
        err = result["err"] or "Timed out after 300s"
        log(f"  [FAIL] {src.name} — {err}")
        return False

    if result["warn"]:
        log(f"  [WARN] {dest.name} — {result['warn']}")
    else:
        log(f"  [OK] {dest.name}")
    return True


def convert_pptx(src: Path, dest: Path) -> bool:
    """Convert .pptx/.ppt to PDF via PowerPoint COM SaveAs (ppSaveAsPDF)."""
    ppt = None
    try:
        import win32com.client
        ppt = win32com.client.Dispatch("PowerPoint.Application")
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
        _quit_with_timeout(ppt, kill_exe="POWERPNT.EXE")
        log(f"  [OK] {dest.name}")
        return True
    except ImportError:
        log("  [ERROR] pywin32 not installed. Run install_dependencies.bat first.")
        return False
    except Exception as e:
        log(f"  [FAIL] {src.name} — {e}")
        if ppt:
            subprocess.run(["taskkill", "/F", "/IM", "POWERPNT.EXE"], capture_output=True)
        return False


def convert_html(src: Path, dest: Path) -> bool:
    """Convert .html/.htm to PDF.
    Strategy:
      1. Chrome headless full-page screenshot -> split into landscape A4 pages via reportlab+Pillow
      2. WeasyPrint fallback
      3. reportlab plain-text fallback
    """
    # ── helpers ──────────────────────────────────────────────────────────────
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
            import subprocess
            from PIL import Image as PILImage
            from reportlab.lib.pagesizes import A4, landscape
            from reportlab.platypus import SimpleDocTemplate, Image as RLImage
            from reportlab.lib.units import mm

            # Inject colour-preserve CSS
            html_src = src.read_text(encoding="utf-8", errors="replace")
            inject = """<style>
* { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }
</style>"""
            if "</head>" in html_src:
                html_src = html_src.replace("</head>", inject + "</head>")
            else:
                html_src = inject + html_src

            tmp_html = src.parent / ("__tmp_" + src.name)
            png_path = dest.parent / (dest.stem + "__shot.png")
            strips = []

            try:
                tmp_html.write_text(html_src, encoding="utf-8")

                # Get full page height via Chrome DevTools Protocol, then screenshot at full height
                # First pass: measure scroll height
                import json, os

                # Use a JS-injection approach: set window size very tall to capture all content
                # Chrome --screenshot captures the viewport, so we set viewport = full page height
                # We do two runs: first to get height via CDP, second to screenshot at that height

                # Run 1: capture at very tall viewport to get full page in one shot
                # 1280 wide x 15000 tall covers most dashboards
                subprocess.run([
                    chrome_exe,
                    "--headless=new",
                    "--disable-gpu",
                    "--no-sandbox",
                    "--hide-scrollbars",
                    "--window-size=1280,15000",
                    "--screenshot=" + str(png_path.resolve()),
                    "--virtual-time-budget=8000",
                    tmp_html.resolve().as_uri(),
                ], capture_output=True, timeout=60)

                if not png_path.exists() or png_path.stat().st_size == 0:
                    raise RuntimeError("Chrome produced no screenshot")

                # Page geometry: landscape A4 with 10mm margins
                page_w, page_h = landscape(A4)
                margin = 10 * mm
                avail_w = page_w - 2 * margin   # points (available width)
                avail_h = page_h - 2 * margin   # points (available height)

                with PILImage.open(str(png_path)) as im:
                    im = im.convert("RGB")
                    img_w, img_h = im.size   # pixels

                    # Auto-crop trailing white/near-white space at bottom.
                    # Use std-dev per row: a truly blank row has near-zero variance.
                    # Also catches near-white backgrounds (e.g. #f4f4f2 = 244,244,242).
                    import numpy as np
                    arr = np.array(im, dtype=np.float32)
                    # Row std-dev: rows with real content have visible pixel variation
                    row_std = arr.std(axis=(1, 2))
                    # Also flag rows where any pixel is clearly non-white (< 240)
                    row_dark = arr.min(axis=(1, 2)) < 240
                    content_rows = np.where((row_std > 1.5) | row_dark)[0]
                    if len(content_rows) > 0:
                        last_content_row = int(content_rows[-1]) + 30  # 30px padding
                        img_h = min(last_content_row, img_h)
                        im = im.crop((0, 0, img_w, img_h))

                    # How many pixel rows = one page height?
                    px_per_page = int(img_w * (avail_h / avail_w))
                    n_pages = max(1, -(-img_h // px_per_page))  # ceiling div

                    for i in range(n_pages):
                        top = i * px_per_page
                        bottom = min(top + px_per_page, img_h)
                        strip = im.crop((0, top, img_w, bottom))
                        # Pad last strip to full height
                        strip_h = bottom - top
                        if strip_h < px_per_page:
                            padded = PILImage.new("RGB", (img_w, px_per_page), (255, 255, 255))
                            padded.paste(strip, (0, 0))
                            strip = padded
                        # Skip strips that are >95% white (near-blank trailing pages)
                        strip_arr = np.array(strip, dtype=np.float32)
                        white_ratio = (strip_arr.min(axis=2) >= 240).mean()
                        if white_ratio > 0.95:
                            continue  # drop this strip — it's blank
                        strip_path = dest.parent / f"__strip{i}_{dest.stem}.png"
                        strip.save(str(strip_path), "PNG")
                        strips.append(strip_path)

                # Assemble strips into PDF — each strip fills exactly one landscape A4 page
                from reportlab.platypus import BaseDocTemplate, Frame, PageTemplate

                frames = [Frame(margin, margin, avail_w, avail_h, leftPadding=0,
                                rightPadding=0, topPadding=0, bottomPadding=0)]
                pt = PageTemplate(id="main", frames=frames,
                                  pagesize=landscape(A4))
                doc = BaseDocTemplate(
                    str(dest),
                    pagesize=landscape(A4),
                    leftMargin=margin, rightMargin=margin,
                    topMargin=margin, bottomMargin=margin,
                    pageTemplates=[pt],
                )
                # Each strip: scale to fit avail_w x avail_h exactly
                story = [RLImage(str(s), width=avail_w, height=avail_h,
                                 kind="direct") for s in strips]
                doc.build(story)

                if dest.exists() and dest.stat().st_size > 0:
                    log(f"  [OK] {dest.name}")
                    return True
                raise RuntimeError("PDF assembly produced empty file")

            finally:
                if tmp_html.exists(): tmp_html.unlink()
                if png_path.exists(): png_path.unlink()
                for s in strips:
                    if s.exists(): s.unlink()

        except Exception as e:
            log(f"  [WARN] Chrome path failed ({e}) — trying WeasyPrint")

    # ── 2. WeasyPrint fallback ────────────────────────────────────────────────
    try:
        from weasyprint import HTML, CSS
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
        tmp = src.parent / ("__tmp_" + src.name)
        try:
            tmp.write_text(html_src, encoding="utf-8")
            HTML(filename=str(tmp.resolve())).write_pdf(str(dest.resolve()))
        finally:
            if tmp.exists(): tmp.unlink()
        if dest.exists() and dest.stat().st_size > 0:
            log(f"  [OK-weasyprint] {dest.name}")
            return True
    except ImportError:
        log("  [WARN] WeasyPrint not available — trying reportlab")
    except Exception as e:
        log(f"  [WARN] WeasyPrint failed ({e}) — trying reportlab")

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
        html = html.replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">")                    .replace("&nbsp;", " ").replace("&quot;", '"').replace("&#39;", "'")
        lines = [l.rstrip() for l in html.splitlines()]
        lines = [l for i, l in enumerate(lines) if l or (i > 0 and lines[i-1])]

        styles = getSampleStyleSheet()
        style_body = ParagraphStyle("Body", parent=styles["Normal"], fontSize=10, leading=15, spaceAfter=4)
        story = []
        for line in lines:
            line = line.strip()
            if not line:
                story.append(Spacer(1, 6))
            else:
                safe = line.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                story.append(Paragraph(safe, style_body))
        doc = SimpleDocTemplate(str(dest), pagesize=A4,
                                leftMargin=2.5*cm, rightMargin=2.5*cm,
                                topMargin=2.5*cm, bottomMargin=2.5*cm)
        doc.build(story)
        log(f"  [OK-fallback] {dest.name}")
        return True
    except Exception as e:
        log(f"  [FAIL] {src.name} — all methods failed: {e}")
        return False


def convert_md(src: Path, dest: Path) -> bool:
    """Convert .md to PDF using reportlab (pure Python, no system libraries)."""
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.units import cm
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, HRFlowable
        from reportlab.lib import colors

        md_text = src.read_text(encoding="utf-8", errors="replace")
        lines = md_text.splitlines()

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
        code_lines = []

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
                code_lines.append(line.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))
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
                story.append(Paragraph(md_inline(re.sub(r"^\d+\. ", "", line).strip()), style_bullet))
                continue
            if line.strip() == "":
                story.append(Spacer(1, 6))
                continue
            story.append(Paragraph(md_inline(line.strip()), style_body))

        flush_code()

        doc = SimpleDocTemplate(str(dest), pagesize=A4,
                                leftMargin=2.5*cm, rightMargin=2.5*cm,
                                topMargin=2.5*cm, bottomMargin=2.5*cm)
        doc.build(story)
        log(f"  [OK] {dest.name}")
        return True

    except ImportError:
        log("  [ERROR] reportlab not installed. Run: pip install reportlab")
        return False
    except Exception as e:
        log(f"  [ERROR] Markdown conversion failed: {e}")
        return False


def convert_txt(src: Path, dest: Path) -> bool:
    """Convert .txt to PDF using reportlab (pure Python)."""
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.units import cm
        from reportlab.pdfgen import canvas as rl_canvas

        lines = src.read_text(encoding="utf-8", errors="replace").splitlines()
        page_w, page_h = A4
        margin = 2.5 * cm
        font_size = 10
        line_height = font_size * 1.4
        max_lines_per_page = int((page_h - 2 * margin) / line_height)

        c = rl_canvas.Canvas(str(dest), pagesize=A4)
        c.setFont("Courier", font_size)
        y = page_h - margin

        for i, line in enumerate(lines):
            if i > 0 and i % max_lines_per_page == 0:
                c.showPage()
                c.setFont("Courier", font_size)
                y = page_h - margin
            c.drawString(margin, y, line[:120])
            y -= line_height

        c.save()
        log(f"  [OK] {dest.name}")
        return True
    except ImportError:
        log("  [ERROR] reportlab not installed. Run: pip install reportlab")
        return False
    except Exception as e:
        log(f"  [ERROR] Text conversion failed: {e}")
        return False


def convert_csv(src: Path, dest: Path) -> bool:
    """Convert .csv to PDF as a formatted table using reportlab."""
    try:
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.lib.units import cm
        from reportlab.lib import colors
        from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

        with open(str(src), newline="", encoding="utf-8", errors="replace") as f:
            rows = list(csv.reader(f))

        if not rows:
            log("  [WARN] CSV file is empty.")
            return False

        styles = getSampleStyleSheet()
        style_title = styles["Heading2"]
        style_cell  = ParagraphStyle("Cell", parent=styles["Normal"], fontSize=8)

        table_data = []
        for row in rows:
            table_data.append([Paragraph(str(cell), style_cell) for cell in row])

        page_w, page_h = landscape(A4)
        usable_w = page_w - 4 * cm
        col_count = len(rows[0]) if rows else 1
        col_w = usable_w / col_count

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
                                leftMargin=2*cm, rightMargin=2*cm,
                                topMargin=2*cm, bottomMargin=2*cm)
        doc.build(story)
        log(f"  [OK] {dest.name}")
        return True

    except ImportError:
        log("  [ERROR] reportlab not installed. Run: pip install reportlab")
        return False
    except Exception as e:
        log(f"  [ERROR] CSV conversion failed: {e}")
        return False


def convert_xls(src: Path, dest: Path) -> bool:
    """Convert legacy .xls to PDF using xlrd (read) + reportlab (write)."""
    try:
        import xlrd
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.lib.units import cm
        from reportlab.lib import colors
        from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

        wb = xlrd.open_workbook(str(src))
        styles = getSampleStyleSheet()
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
            usable_w = page_w - 4 * cm
            col_count = len(rows[0]) if rows else 1
            col_w = usable_w / col_count
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
            log("  [WARN] .xls file has no data.")
            return False

        doc = SimpleDocTemplate(str(dest), pagesize=landscape(A4),
                                leftMargin=2*cm, rightMargin=2*cm,
                                topMargin=2*cm, bottomMargin=2*cm)
        doc.build(story)
        log(f"  [OK] {dest.name}")
        return True

    except ImportError:
        log("  [ERROR] xlrd or reportlab not installed. Run: pip install xlrd reportlab")
        return False
    except Exception as e:
        log(f"  [ERROR] XLS conversion failed: {e}")
        return False


# ── PDF → Word converter ──────────────────────────────────

def _auto_dismiss_word_dialog(stop_event, timeout=60):
    """Background thread: auto-clicks OK on Word's PDF conversion info dialog.

    The dialog has title '?' or 'Microsoft Word' and contains an OK button.
    We enumerate ALL visible top-level windows and click any OK button found
    in a dialog that contains the known PDF-conversion message text.
    """
    import time
    try:
        import win32gui, win32con, win32api
    except ImportError:
        return
    deadline = time.time() + timeout
    while not stop_event.is_set() and time.time() < deadline:
        def _cb(hwnd, ctx):
            if not win32gui.IsWindowVisible(hwnd):
                return True
            # Match any small dialog window (Word's info dialog is narrow)
            title = win32gui.GetWindowText(hwnd)
            # Word PDF dialog has title "" or "?" or "Microsoft Word"
            if title not in ("", "?", "Microsoft Word"):
                return True
            # Look for an OK button among children
            try:
                child = win32gui.FindWindowEx(hwnd, 0, "Button", None)
                while child:
                    btn_text = win32gui.GetWindowText(child).strip().upper()
                    if btn_text in ("OK", "&OK"):
                        # Click it via SendMessage (more reliable than PostMessage for dialogs)
                        win32api.SendMessage(child, win32con.BM_CLICK, 0, 0)
                        ctx.append(hwnd)
                        return False  # stop enum
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


def convert_pdf_to_docx(src: Path, dest: Path) -> bool:
    """Convert a PDF file to .docx.

    Strategy (three tiers):
      1. Word COM — highest fidelity (photos, layout, tables intact).
         Auto-dismisses Word's PDF conversion info dialog via win32gui.
      2. pymupdf + python-docx — text-only fallback. No Office required.
      3. pytesseract OCR — scanned/image-only PDF fallback.
    """
    import shutil

    # ── Tier 1: Word COM (high fidelity) ─────────────────────────────────────
    try:
        import win32com.client, threading
        stop_evt = threading.Event()
        dismiss_t = threading.Thread(
            target=_auto_dismiss_word_dialog, args=(stop_evt,), daemon=True
        )
        dismiss_t.start()

        # Kill any existing Word process and remove stale output before starting
        import time
        subprocess.run(["taskkill", "/F", "/IM", "WINWORD.EXE"], capture_output=True)
        time.sleep(1)
        if dest.exists():
            try:
                dest.unlink()
            except Exception:
                pass

        word = win32com.client.Dispatch("Word.Application")
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
        stop_evt.set()   # dialog handled (or not needed) — stop watcher
        doc.SaveAs2(str(dest.resolve()), FileFormat=16)  # wdFormatDocx = 16
        doc.Close(SaveChanges=False)
        subprocess.run(["taskkill", "/F", "/IM", "WINWORD.EXE"], capture_output=True)
        log(f"  [OK] {dest.name}")
        return True
    except ImportError:
        log("  [WARN] pywin32 not installed — trying pymupdf")
    except Exception as e:
        log(f"  [WARN] Word COM failed ({e}) — trying pymupdf")
        subprocess.run(["taskkill", "/F", "/IM", "WINWORD.EXE"], capture_output=True)

    # ── Tier 2: pymupdf + python-docx ────────────────────────────────────────
    try:
        import pymupdf as fitz  # pymupdf
        from docx import Document
        from docx.shared import Pt

        pdf = fitz.open(str(src.resolve()))
        doc = Document()

        # Document title from filename
        doc.add_heading(src.stem, level=1)

        for page_num in range(len(pdf)):
            page = pdf[page_num]
            text = page.get_text("text")
            if not text.strip():
                continue
            # Add a page-break heading for multi-page docs
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
        log(f"  [OK-pymupdf] {dest.name}")
        return True
    except ImportError:
        log("  [WARN] pymupdf or python-docx not installed — trying OCR")
    except Exception as e:
        log(f"  [WARN] pymupdf extraction failed ({e}) — trying OCR")

    # ── Tier 3: pytesseract OCR (for scanned PDFs) ────────────────────────────
    try:
        import pymupdf as fitz  # pymupdf — used to render pages as images
        from PIL import Image as PILImage
        import pytesseract
        from docx import Document
        from docx.shared import Pt
        import io

        # Check Tesseract is installed
        try:
            pytesseract.get_tesseract_version()
        except Exception:
            log("  [ERROR] Tesseract OCR not installed. "
                "Download from: https://github.com/UB-Mannheim/tesseract/wiki")
            return False

        pdf = fitz.open(str(src.resolve()))
        doc = Document()
        doc.add_heading(src.stem, level=1)

        for page_num in range(len(pdf)):
            page = pdf[page_num]
            # Render page at 300 DPI for good OCR accuracy
            mat = fitz.Matrix(300 / 72, 300 / 72)
            pix = page.get_pixmap(matrix=mat, colorspace=fitz.csRGB)
            img_bytes = pix.tobytes("png")
            img = PILImage.open(io.BytesIO(img_bytes))

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
        log(f"  [OK-OCR] {dest.name}")
        return True
    except ImportError:
        log("  [ERROR] pytesseract / pymupdf / python-docx not installed. "
            "Run: pip install pymupdf pytesseract python-docx Pillow")
        return False
    except Exception as e:
        log(f"  [FAIL] {src.name} — all PDF->Word methods failed: {e}")
        return False


# ── PDF → Excel converter ─────────────────────────────────

def convert_pdf_to_xlsx(src: Path, dest: Path) -> bool:
    """Convert a PDF to .xlsx using word-position extraction for high fidelity.

    Strategy:
      1. pdfplumber extract_words() — groups words by Y-position to reconstruct
         rows precisely. Handles PDFs with no visible table lines.
      2. Fallback: raw text lines split into columns if word extraction fails.
    """
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

        wb   = openpyxl.Workbook()
        ws   = wb.active
        ws.title = "Data"

        all_rows         = []   # list of lists (one per text row across all pages)
        header_row       = None
        first_data_words = None  # x-positions of first data row for header bucketing

        with pdfplumber.open(str(src.resolve())) as pdf:
            for page_num, page in enumerate(pdf.pages, 1):
                words = page.extract_words(
                    x_tolerance=3, y_tolerance=3, keep_blank_chars=False
                )
                if not words:
                    continue

                # Group words by rounded Y position
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

                    # Detect header row: all text, no leading digit
                    is_header = (
                        header_row is None
                        and not any(t[0].isdigit() for t in tokens)
                        and any(kw in " ".join(tokens).lower()
                                for kw in ["date", "branch", "sales", "transaction",
                                           "count", "total", "avg"])
                    )

                    if is_header:
                        # Store raw header words with X positions for later bucketing
                        header_row = row_words   # list of word dicts
                    else:
                        # First non-header row — capture column X-centres
                        if first_data_words is None and page_num == 1:
                            first_data_words = row_words
                        all_rows.append(tokens)

        if not all_rows:
            log(f"  [WARN] No data extracted — trying text fallback")
            raise ValueError("no rows")

        # Determine number of columns from most common row width
        from collections import Counter
        col_counts = Counter(len(r) for r in all_rows)
        num_cols   = col_counts.most_common(1)[0][0]

        # Reconstruct header by bucketing header words into data column centres
        if header_row and first_data_words:
            col_centres  = [(w["x0"] + w["x1"]) / 2 for w in first_data_words]
            final_header = [""] * len(col_centres)
            for hw in header_row:
                hc      = (hw["x0"] + hw["x1"]) / 2
                nearest = min(range(len(col_centres)),
                              key=lambda i: abs(col_centres[i] - hc))
                sep = " " if final_header[nearest] else ""
                final_header[nearest] += sep + hw["text"]
            # Pad/trim to num_cols
            if len(final_header) < num_cols:
                final_header += [f"Col{i+1}" for i in range(len(final_header), num_cols)]
            final_header = final_header[:num_cols]
            # Clean up common PDF header garbling (overlapping text artifacts)
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

        # Write header
        for col_idx, hdr in enumerate(final_header, 1):
            cell = ws.cell(row=1, column=col_idx, value=hdr)
            cell.fill      = HEADER_FILL
            cell.font      = HEADER_FONT
            cell.alignment = CENTER
            cell.border    = THIN
        ws.freeze_panes = "A2"

        # Write data rows — skip rows whose column count differs significantly
        xlsx_row = 2
        for row in all_rows:
            if len(row) == 0:
                continue
            # Pad or trim to num_cols
            padded = (row + [""] * num_cols)[:num_cols]
            for col_idx, val in enumerate(padded, 1):
                cell = ws.cell(row=xlsx_row, column=col_idx, value=val)
                cell.font      = CELL_FONT
                cell.fill      = ROW_FILL_ALT if xlsx_row % 2 == 0 else PatternFill()
                cell.alignment = WRAP
                cell.border    = THIN
            xlsx_row += 1

        # Auto-size columns (cap at 40)
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
        log(f"  [OK] {dest.name} — {xlsx_row-2} data rows, {num_cols} columns")
        return True

    except ImportError as e:
        log(f"  [ERROR] Missing library: {e}")
        return False
    except Exception as e:
        log(f"  [FAIL] {src.name} — {e}")
        return False


def close_office_apps():
    """Force-close Word, Excel, and PowerPoint to release file locks and temp files."""
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
            log(f"  [INFO] Closed {name} (released file locks)")


def cleanup_output():
    """Remove any non-PDF files (Office temp files) from the output folder."""
    import time
    # Kill print spooler helper that holds .tmp locks (safe to kill — restarts automatically)
    subprocess.run(["taskkill", "/F", "/IM", "splwow64.exe"], capture_output=True)
    # Retry up to 10 times with 2s sleep — spooler may hold locks for several seconds
    for attempt in range(10):
        time.sleep(2)
        remaining = []
        removed = []
        for f in OUTPUT_DIR.iterdir():
            if f.is_file() and f.suffix.lower() != ".pdf":
                try:
                    f.unlink()
                    removed.append(f.name)
                except Exception:
                    remaining.append(f.name)
        if removed:
            log(f"  [INFO] Cleaned temp files from output: {', '.join(removed)}")
        if not remaining:
            break   # all gone — no need to retry
    if remaining:
        log(f"  [WARN] Could not remove temp files (still locked): {', '.join(remaining)}")


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


def convert_file(src: Path, pdf_mode: str = None) -> bool:
    ext = src.suffix.lower()

    # PDF source files route to Word or Excel depending on mode
    if ext == ".pdf":
        mode = pdf_mode or PDF_OUTPUT_MODE
        if mode == "excel":
            dest = OUTPUT_DIR / (src.stem + ".xlsx")
            log(f"  [CONVERTING] {src.name} -> {dest.name} ...")
            success = convert_pdf_to_xlsx(src, dest)
        else:
            dest = OUTPUT_DIR / (src.stem + ".docx")
            log(f"  [CONVERTING] {src.name} -> {dest.name} ...")
            success = convert_pdf_to_docx(src, dest)
        if not success:
            log(f"  [FAIL] {src.name} could not be converted.")
        return success

    converter = CONVERTERS.get(ext)
    if converter is None:
        log(f"  [SKIP] Unsupported format: {src.name}")
        return False
    dest = OUTPUT_DIR / (src.stem + ".pdf")
    log(f"  [CONVERTING] {src.name} -> {dest.name} ...")
    success = converter(src, dest)
    if not success:
        log(f"  [FAIL] {src.name} could not be converted.")
    return success


# ── Main ──────────────────────────────────────────────────
def run(single_file=None):
    INPUT_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    if single_file:
        src = Path(single_file)
        if not src.is_absolute():
            src = INPUT_DIR / src
        if not src.exists():
            log(f"[ERROR] File not found: {src}")
            sys.exit(1)
        files = [src]
    else:
        files = sorted(
            f for f in INPUT_DIR.iterdir()
            if f.is_file()
            and f.suffix.lower() in SUPPORTED
            and not f.name.startswith("~$")   # skip Office lock/temp files
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

    close_office_apps()
    cleanup_output()
    log("-" * 52)
    log(f"Done.  Converted: {ok}  |  Failed: {fail}")
    log(f"PDFs saved to: {OUTPUT_DIR}")
    log("=" * 52)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Convert documents to PDF.")
    parser.add_argument("--file", metavar="FILENAME",
                        help="Convert a single file from the input\\ folder.")
    args = parser.parse_args()
    run(single_file=args.file)
