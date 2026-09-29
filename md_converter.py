"""
=============================================================
  Document → Markdown Converter  (v1.0 — DocumentConverter suite)
  Usage:
    python md_converter.py                    # batch: all supported files in input\
    python md_converter.py --file foo.docx    # single file

  Input folder  : .\input\
  Output folder : same as converter.py (config.json → Desktop fallback)
  Log file      : .\logs\md_convert_log.txt

  Supported input formats:
    Documents    : .docx  .doc
    Spreadsheets : .xlsx  .xls
    Presentations: .pptx  .ppt
    Web          : .html  .htm
    Data         : .csv
    Diagram      : .mermaid  .bpmn
    Plain text   : .txt
    Markdown     : .md  (pass-through / re-saved)

  NOTES:
    .doc   — requires mammoth (pip install mammoth)
    .docx  — requires python-docx (pip install python-docx)
    .xlsx  — requires openpyxl (pip install openpyxl)
    .xls   — requires xlrd<2.0 (pip install xlrd==1.2.0)
    .pptx/.ppt — requires python-pptx (pip install python-pptx)
    .html  — requires html2text (pip install html2text)
    .csv   — requires pandas (pip install pandas)
    .bpmn  — XML text extraction only (no visual diagram render)
=============================================================
"""

import argparse
import csv
import io
import json
import re
import sys
import xml.etree.ElementTree as ET
from datetime import datetime
from pathlib import Path

# ── Folder layout ─────────────────────────────────────────────────────────────
BASE_DIR   = Path(__file__).parent
INPUT_DIR  = BASE_DIR / "input"
LOG_DIR    = BASE_DIR / "logs"
CONFIG_FILE = BASE_DIR / "config.json"

LOG_DIR.mkdir(parents=True, exist_ok=True)
INPUT_DIR.mkdir(parents=True, exist_ok=True)

LOG_FILE = LOG_DIR / "md_convert_log.txt"

SUPPORTED = {
    ".docx", ".doc",
    ".xlsx", ".xls",
    ".pptx", ".ppt",
    ".html", ".htm",
    ".csv",
    ".mermaid",
    ".bpmn",
    ".txt",
    ".md",
    ".pdf",
}


# ── Output folder resolution (mirrors converter_gui.py) ───────────────────────
def _load_output_dir() -> Path:
    """Read output folder from config.json. Falls back to Desktop."""
    desktop = Path.home() / "Desktop"
    try:
        if CONFIG_FILE.exists():
            data = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
            folder = Path(data.get("output_folder", ""))
            if folder and folder.exists():
                return folder
    except Exception:
        pass
    return desktop

OUTPUT_DIR = _load_output_dir()


# ── Logging ───────────────────────────────────────────────────────────────────
def log(msg: str):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line)
    with LOG_FILE.open("a", encoding="utf-8") as f:
        f.write(line + "\n")


# ── Helpers ───────────────────────────────────────────────────────────────────
def _slug(name: str) -> str:
    """Safe filename slug from document name."""
    return re.sub(r"[^\w\s-]", "", name).strip().replace(" ", "_")


# ══════════════════════════════════════════════════════════════════════════════
#  Per-format converters — each returns (markdown_str, success: bool)
# ══════════════════════════════════════════════════════════════════════════════

def _docx_to_md(src: Path) -> tuple[str, bool]:
    try:
        from docx import Document
        from docx.oxml.ns import qn
        doc = Document(src)
        lines = []
        for block in doc.element.body:
            tag = block.tag.split("}")[-1] if "}" in block.tag else block.tag

            if tag == "p":
                from docx.text.paragraph import Paragraph
                para = Paragraph(block, doc)
                text = para.text.strip()
                style = para.style.name if para.style else ""

                if not text:
                    lines.append("")
                    continue

                # Heading detection
                if style.startswith("Heading"):
                    level = "".join(filter(str.isdigit, style)) or "1"
                    lines.append(f"{'#' * int(level)} {text}")
                elif style == "Title":
                    lines.append(f"# {text}")
                elif style == "Subtitle":
                    lines.append(f"## {text}")
                elif para.runs and all(r.bold for r in para.runs if r.text.strip()):
                    lines.append(f"**{text}**")
                else:
                    # Inline formatting
                    inline = []
                    for run in para.runs:
                        t = run.text
                        if not t:
                            continue
                        if run.bold and run.italic:
                            t = f"***{t}***"
                        elif run.bold:
                            t = f"**{t}**"
                        elif run.italic:
                            t = f"*{t}*"
                        elif run.underline:
                            t = f"<u>{t}</u>"
                        inline.append(t)
                    lines.append("".join(inline) if inline else text)

            elif tag == "tbl":
                from docx.table import Table
                tbl = Table(block, doc)
                rows = [[cell.text.strip() for cell in row.cells] for row in tbl.rows]
                if not rows:
                    continue
                header = rows[0]
                lines.append("| " + " | ".join(header) + " |")
                lines.append("| " + " | ".join(["---"] * len(header)) + " |")
                for row in rows[1:]:
                    lines.append("| " + " | ".join(row) + " |")
                lines.append("")

        return "\n".join(lines), True

    except Exception as e:
        return f"<!-- ERROR: {e} -->", False


def _doc_to_md(src: Path) -> tuple[str, bool]:
    try:
        import mammoth
        with src.open("rb") as f:
            result = mammoth.convert_to_markdown(f)
        return result.value, True
    except Exception as e:
        return f"<!-- ERROR: {e} -->", False


def _xlsx_to_md(src: Path) -> tuple[str, bool]:
    try:
        import openpyxl
        wb = openpyxl.load_workbook(src, data_only=True)
        parts = []
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            parts.append(f"## Sheet: {sheet_name}\n")
            rows = list(ws.iter_rows(values_only=True))
            if not rows:
                parts.append("*(empty sheet)*\n")
                continue
            # Strip trailing all-None rows
            rows = [r for r in rows if any(c is not None for c in r)]
            if not rows:
                parts.append("*(empty sheet)*\n")
                continue
            # Determine column count
            col_count = max(len(r) for r in rows)
            header = [str(c) if c is not None else "" for c in rows[0]]
            parts.append("| " + " | ".join(header) + " |")
            parts.append("| " + " | ".join(["---"] * col_count) + " |")
            for row in rows[1:]:
                cells = [str(c) if c is not None else "" for c in row]
                # Pad if row is shorter than header
                while len(cells) < col_count:
                    cells.append("")
                parts.append("| " + " | ".join(cells) + " |")
            parts.append("")
        return "\n".join(parts), True
    except Exception as e:
        return f"<!-- ERROR: {e} -->", False


def _xls_to_md(src: Path) -> tuple[str, bool]:
    try:
        import xlrd
        wb = xlrd.open_workbook(src)
        parts = []
        for sheet in wb.sheets():
            parts.append(f"## Sheet: {sheet.name}\n")
            if sheet.nrows == 0:
                parts.append("*(empty sheet)*\n")
                continue
            rows = [sheet.row_values(i) for i in range(sheet.nrows)]
            header = [str(c) for c in rows[0]]
            col_count = len(header)
            parts.append("| " + " | ".join(header) + " |")
            parts.append("| " + " | ".join(["---"] * col_count) + " |")
            for row in rows[1:]:
                cells = [str(c) for c in row]
                parts.append("| " + " | ".join(cells) + " |")
            parts.append("")
        return "\n".join(parts), True
    except Exception as e:
        return f"<!-- ERROR: {e} -->", False


def _pptx_to_md(src: Path) -> tuple[str, bool]:
    try:
        from pptx import Presentation
        prs = Presentation(src)
        parts = []
        for i, slide in enumerate(prs.slides, 1):
            title_text = ""
            body_lines = []
            for shape in slide.shapes:
                if not shape.has_text_frame:
                    continue
                # Detect title placeholder
                if hasattr(shape, "placeholder_format") and shape.placeholder_format:
                    ph_idx = shape.placeholder_format.idx
                    if ph_idx == 0:  # Title
                        title_text = shape.text_frame.text.strip()
                        continue
                for para in shape.text_frame.paragraphs:
                    text = para.text.strip()
                    if not text:
                        continue
                    level = para.level or 0
                    indent = "  " * level
                    body_lines.append(f"{indent}- {text}")

            parts.append(f"## Slide {i}: {title_text or '(no title)'}")
            if body_lines:
                parts.extend(body_lines)
            parts.append("")
        return "\n".join(parts), True
    except Exception as e:
        return f"<!-- ERROR: {e} -->", False


def _html_to_md(src: Path) -> tuple[str, bool]:
    try:
        import html2text
        h = html2text.HTML2Text()
        h.ignore_links = False
        h.ignore_images = False
        h.body_width = 0  # No line wrapping
        content = src.read_text(encoding="utf-8", errors="replace")
        return h.handle(content), True
    except Exception as e:
        return f"<!-- ERROR: {e} -->", False


def _csv_to_md(src: Path) -> tuple[str, bool]:
    try:
        import pandas as pd
        df = pd.read_csv(src, dtype=str).fillna("")
        lines = []
        header = list(df.columns)
        lines.append("| " + " | ".join(header) + " |")
        lines.append("| " + " | ".join(["---"] * len(header)) + " |")
        for _, row in df.iterrows():
            lines.append("| " + " | ".join(str(v) for v in row) + " |")
        return "\n".join(lines), True
    except Exception as e:
        return f"<!-- ERROR: {e} -->", False


def _mermaid_to_md(src: Path) -> tuple[str, bool]:
    try:
        content = src.read_text(encoding="utf-8", errors="replace").strip()
        md = f"```mermaid\n{content}\n```"
        return md, True
    except Exception as e:
        return f"<!-- ERROR: {e} -->", False


def _bpmn_to_md(src: Path) -> tuple[str, bool]:
    """
    BPMN is XML. We extract process names, elements, and sequence flows.
    Visual diagram rendering is not possible without a BPMN engine.
    """
    try:
        content = src.read_text(encoding="utf-8", errors="replace")
        tree = ET.parse(io.StringIO(content))
        root = tree.getroot()

        # Strip namespace for easier parsing
        def strip_ns(tag):
            return tag.split("}")[-1] if "}" in tag else tag

        lines = [f"# BPMN: {src.stem}\n"]
        lines.append(f"> **To view diagram:** Open `{src.name}` in Draw.io desktop.\n")
        lines.append("> **Note:** Visual diagram not renderable in Markdown. Text extraction only.\n")

        # Raw XML block for reference
        lines.append("## Raw BPMN Source\n")
        lines.append(f"```xml\n{content.strip()}\n```\n")

        # Try to extract process elements
        lines.append("## Process Elements\n")
        for elem in root.iter():
            tag = strip_ns(elem.tag)
            name = elem.attrib.get("name", "").strip()
            eid  = elem.attrib.get("id", "")
            if tag in ("process", "subProcess", "task", "userTask",
                       "serviceTask", "startEvent", "endEvent",
                       "exclusiveGateway", "parallelGateway",
                       "sequenceFlow"):
                src_ref  = elem.attrib.get("sourceRef", "")
                tgt_ref  = elem.attrib.get("targetRef", "")
                if tag == "sequenceFlow":
                    lines.append(f"- **Flow**: {src_ref} → {tgt_ref}"
                                 + (f" *(name: {name})*" if name else ""))
                else:
                    lines.append(f"- **{tag}**: {name or eid}")

        return "\n".join(lines), True
    except Exception as e:
        # Fallback: dump raw XML in a code fence
        try:
            raw = src.read_text(encoding="utf-8", errors="replace")
            md = (f"# BPMN: {src.stem}\n\n"
                  f"> Parse error: {e}. Raw source below.\n\n"
                  f"```xml\n{raw}\n```")
            return md, True
        except Exception as e2:
            return f"<!-- ERROR: {e2} -->", False


def _pdf_to_md(src: Path) -> tuple[str, bool]:
    """
    Extract text and tables from PDF using pdfplumber.
    Tables are rendered as Markdown tables; body text as paragraphs.
    Formatting (fonts, colours, images) is not recoverable from PDF.
    """
    try:
        import pdfplumber
        parts = []
        with pdfplumber.open(src) as pdf:
            total = len(pdf.pages)
            for i, page in enumerate(pdf.pages, 1):
                parts.append(f"## Page {i} of {total}\n")

                # ── Tables first ──────────────────────────────────────────
                tables = page.extract_tables()
                table_bboxes = []
                if tables:
                    for tbl in tables:
                        if not tbl:
                            continue
                        rows = [[str(c).strip() if c else "" for c in row]
                                for row in tbl]
                        col_count = max(len(r) for r in rows)
                        header = rows[0]
                        # Pad header if needed
                        while len(header) < col_count:
                            header.append("")
                        parts.append("| " + " | ".join(header) + " |")
                        parts.append("| " + " | ".join(["---"] * col_count) + " |")
                        for row in rows[1:]:
                            while len(row) < col_count:
                                row.append("")
                            parts.append("| " + " | ".join(row) + " |")
                        parts.append("")

                # ── Body text ─────────────────────────────────────────────
                text = page.extract_text()
                if text:
                    # Normalize whitespace, preserve paragraph breaks
                    cleaned = re.sub(r"\n{3,}", "\n\n", text.strip())
                    parts.append(cleaned)
                    parts.append("")

        return "\n".join(parts), True

    except Exception as e:
        return f"<!-- ERROR: {e} -->", False


def _txt_to_md(src: Path) -> tuple[str, bool]:
    try:
        content = src.read_text(encoding="utf-8", errors="replace").strip()
        return content, True
    except Exception as e:
        return f"<!-- ERROR: {e} -->", False


def _md_passthrough(src: Path) -> tuple[str, bool]:
    try:
        return src.read_text(encoding="utf-8", errors="replace"), True
    except Exception as e:
        return f"<!-- ERROR: {e} -->", False


# ── Dispatch table ─────────────────────────────────────────────────────────────
CONVERTERS = {
    ".docx":    _docx_to_md,
    ".doc":     _doc_to_md,
    ".xlsx":    _xlsx_to_md,
    ".xls":     _xls_to_md,
    ".pptx":    _pptx_to_md,
    ".ppt":     _pptx_to_md,   # python-pptx handles both
    ".html":    _html_to_md,
    ".htm":     _html_to_md,
    ".csv":     _csv_to_md,
    ".mermaid": _mermaid_to_md,
    ".bpmn":    _bpmn_to_md,
    ".pdf":     _pdf_to_md,
    ".txt":     _txt_to_md,
    ".md":      _md_passthrough,
}


# ── Core convert function ──────────────────────────────────────────────────────
def convert_to_markdown(src: Path, output_dir: Path) -> bool:
    """Convert a single file to Markdown. Returns True on success."""
    ext = src.suffix.lower()
    fn  = CONVERTERS.get(ext)
    if fn is None:
        log(f"  SKIP [{src.name}] — unsupported format ({ext})")
        return False

    dest = output_dir / (src.stem + ".md")
    log(f"  Converting: {src.name}  →  {dest.name}")

    content, ok = fn(src)
    if not ok:
        log(f"  [FAIL] {src.name}: {content}")
        return False

    # Write header
    header = (
        f"---\n"
        f"source: {src.name}\n"
        f"converted: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
        f"---\n\n"
    )
    dest.write_text(header + content, encoding="utf-8")
    log(f"  [OK]   {dest.name}  ({len(content):,} chars)")
    return True


# ── Batch convert ──────────────────────────────────────────────────────────────
def batch_convert(input_dir: Path, output_dir: Path):
    """Convert all supported files in input_dir."""
    output_dir.mkdir(parents=True, exist_ok=True)
    files = sorted(
        f for f in input_dir.iterdir()
        if f.is_file()
        and f.suffix.lower() in SUPPORTED
        and not f.name.startswith("~$")   # skip Word lock files
    )
    if not files:
        log(f"No supported files found in {input_dir}")
        return

    log(f"\n{'='*54}")
    log(f"  Batch → Markdown  ({datetime.now().strftime('%Y-%m-%d %H:%M:%S')})")
    log(f"  Input:  {input_dir}")
    log(f"  Output: {output_dir}")
    log(f"  Files:  {len(files)}")
    log(f"{'='*54}")

    ok = fail = 0
    for src in files:
        success = convert_to_markdown(src, output_dir)
        if success:
            ok += 1
        else:
            fail += 1

    log(f"\n{'─'*54}")
    log(f"  Done.  Converted: {ok}  |  Failed/Skipped: {fail}")
    log(f"  Output: {output_dir}")
    log(f"{'='*54}\n")


# ── CLI entry point ────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(
        description="Convert documents to Markdown. "
                    "No --file argument = batch mode (all files in input\\)."
    )
    parser.add_argument(
        "--file", "-f",
        help="Single file to convert (full path or filename in input\\).",
        type=str, default=None
    )
    parser.add_argument(
        "--output", "-o",
        help="Output folder (default: config.json → Desktop).",
        type=str, default=None
    )
    args = parser.parse_args()

    output_dir = Path(args.output) if args.output else OUTPUT_DIR
    output_dir.mkdir(parents=True, exist_ok=True)

    if args.file:
        # Single file mode
        src = Path(args.file)
        if not src.is_absolute():
            src = INPUT_DIR / src
        if not src.exists():
            print(f"ERROR: File not found: {src}")
            sys.exit(1)
        ok = convert_to_markdown(src, output_dir)
        sys.exit(0 if ok else 1)
    else:
        # Batch mode
        batch_convert(INPUT_DIR, output_dir)


if __name__ == "__main__":
    main()
