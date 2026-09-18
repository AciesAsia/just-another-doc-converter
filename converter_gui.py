"""
=============================================================
  Convert to PDF — Desktop GUI  (Phase 5)
  Usage:
    python converter_gui.py

  Wraps converter.py — all conversion logic lives there.
  Converts: .docx .doc .xlsx .xls .pptx .ppt
            .html .htm .md .txt .csv  →  PDF

  For PDF → Word / Excel, use: pdf_to_office_gui.py
=============================================================
"""

import json
import os
import sys
import threading
import tkinter as tk
from tkinter import filedialog, scrolledtext, ttk
from pathlib import Path
from datetime import datetime

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE))

try:
    import converter as conv
except ImportError:
    import tkinter.messagebox as mb
    mb.showerror("Missing file", "converter.py not found next to converter_gui.py.")
    sys.exit(1)

FILETYPES = [
    ("Supported documents",
     "*.docx *.doc *.xlsx *.xls *.pptx *.ppt *.html *.htm *.md *.txt *.csv"),
    ("Word documents",   "*.docx *.doc"),
    ("Excel files",      "*.xlsx *.xls"),
    ("PowerPoint files", "*.pptx *.ppt"),
    ("HTML files",       "*.html *.htm"),
    ("Markdown",         "*.md"),
    ("Text files",       "*.txt"),
    ("CSV files",        "*.csv"),
    ("All files",        "*.*"),
]
SUPPORTED_TO_PDF = {".docx", ".doc", ".xlsx", ".xls", ".pptx", ".ppt",
                    ".html", ".htm", ".md", ".txt", ".csv"}

# ── NZB CP 002 palette ───────────────────────────────────────────────────────
CLR_BG       = "#E0E0E0"
CLR_SURFACE  = "#F4F4F4"
CLR_SURFACE2 = "#EBEBEB"
CLR_HEADER   = "#3A606E"
CLR_HEADER_HV= "#2d4d59"
CLR_MID      = "#607B7D"
CLR_MID_HV   = "#506869"
CLR_TEXT     = "#1E2D30"
CLR_SUBTEXT  = "#607B7D"
CLR_BORDER   = "#C8CACB"
CLR_SUCCESS  = "#5a8a6a"
CLR_ERROR    = "#b04040"
CLR_WARN     = "#b07a3a"

CONFIG_FILE = HERE / "config.json"


def _load_output_dir() -> Path:
    """Read output folder from config.json. Fall back to Desktop."""
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


class ConverterApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Convert to PDF")
        self.geometry("780x640")
        self.minsize(640, 520)
        self.configure(bg=CLR_BG)
        self.resizable(True, True)

        self._files: list[Path] = []
        self._running = False
        self._output_dir: Path = _load_output_dir()

        self._build_ui()
        self._refresh_queue()

    def _build_ui(self):
        # ── header bar ────────────────────────────────────────────────────────
        hdr_bar = tk.Frame(self, bg=CLR_HEADER, height=46)
        hdr_bar.pack(fill="x")
        hdr_bar.pack_propagate(False)
        tk.Label(
            hdr_bar, text="Convert to PDF",
            font=("Segoe UI", 13, "bold"),
            fg="#ffffff", bg=CLR_HEADER
        ).pack(side="left", padx=18, pady=10)
        tk.Label(
            hdr_bar, text="Word · Excel · PowerPoint · HTML · Markdown · CSV · TXT",
            font=("Segoe UI", 9),
            fg="#c8dde0", bg=CLR_HEADER
        ).pack(side="left", padx=4, pady=10)

        # ── add files panel ───────────────────────────────────────────────────
        add_frame = tk.Frame(self, bg=CLR_SURFACE, pady=12, padx=16,
                             highlightthickness=1, highlightbackground=CLR_BORDER)
        add_frame.pack(fill="x", padx=16, pady=(12, 0))

        tk.Label(
            add_frame, text="INPUT FILES",
            bg=CLR_SURFACE, fg=CLR_SUBTEXT,
            font=("Segoe UI", 8, "bold")
        ).pack(anchor="w", pady=(0, 6))

        btn_row = tk.Frame(add_frame, bg=CLR_SURFACE)
        btn_row.pack(fill="x")

        self._btn_files = self._make_btn(
            btn_row, "＋ Add Files", CLR_HEADER, CLR_HEADER_HV, self._pick_files)
        self._btn_files.pack(side="left", padx=(0, 8))

        self._btn_folder = self._make_btn(
            btn_row, "＋ Add Folder", CLR_MID, CLR_MID_HV, self._pick_folder)
        self._btn_folder.pack(side="left", padx=(0, 8))

        self._btn_clear = self._make_btn(
            btn_row, "✕ Clear", CLR_SURFACE, CLR_BORDER,
            self._clear_queue, fg=CLR_SUBTEXT, outline=True)
        self._btn_clear.pack(side="left")

        # ── output folder row ─────────────────────────────────────────────────
        out_frame = tk.Frame(self, bg=CLR_SURFACE, pady=8, padx=16,
                             highlightthickness=1, highlightbackground=CLR_BORDER)
        out_frame.pack(fill="x", padx=16, pady=(8, 0))

        tk.Label(
            out_frame, text="OUTPUT FOLDER",
            bg=CLR_SURFACE, fg=CLR_SUBTEXT,
            font=("Segoe UI", 8, "bold")
        ).pack(anchor="w", pady=(0, 4))

        out_row = tk.Frame(out_frame, bg=CLR_SURFACE)
        out_row.pack(fill="x")

        self._lbl_outdir = tk.Label(
            out_row, text=str(self._output_dir),
            bg=CLR_SURFACE2, fg=CLR_TEXT,
            font=("Segoe UI", 9), anchor="w", padx=8,
            relief="flat", highlightthickness=1,
            highlightbackground=CLR_BORDER, cursor="arrow"
        )
        self._lbl_outdir.pack(side="left", fill="x", expand=True, ipady=5)

        self._make_btn(
            out_row, "Browse…", CLR_MID, CLR_MID_HV,
            self._browse_output_dir
        ).pack(side="left", padx=(8, 0))

        # ── queue ─────────────────────────────────────────────────────────────
        queue_frame = tk.Frame(self, bg=CLR_BG, padx=16, pady=4)
        queue_frame.pack(fill="x")

        tk.Label(
            queue_frame, text="QUEUE",
            bg=CLR_BG, fg=CLR_SUBTEXT,
            font=("Segoe UI", 8, "bold")
        ).pack(anchor="w", pady=(6, 3))

        self._queue_box = tk.Listbox(
            queue_frame, bg=CLR_SURFACE2, fg=CLR_TEXT,
            selectbackground=CLR_HEADER,
            font=("Segoe UI", 9), height=5, bd=0,
            highlightthickness=1, highlightbackground=CLR_BORDER,
            activestyle="none")
        self._queue_box.pack(fill="x")

        self._lbl_queue_count = tk.Label(
            queue_frame, text="0 files queued",
            bg=CLR_BG, fg=CLR_SUBTEXT, font=("Segoe UI", 8))
        self._lbl_queue_count.pack(anchor="e", pady=(2, 4))

        # ── action row ────────────────────────────────────────────────────────
        mid = tk.Frame(self, bg=CLR_BG, padx=16, pady=6)
        mid.pack(fill="x")

        self._btn_convert = self._make_btn(
            mid, "▶  Convert All", CLR_HEADER, CLR_HEADER_HV,
            self._start_conversion,
            font=("Segoe UI", 11, "bold"), pady=10, padx=22)
        self._btn_convert.pack(side="left")

        self._btn_open = self._make_btn(
            mid, "📂  Open Output Folder", CLR_SURFACE, CLR_BORDER,
            self._open_output, fg=CLR_SUBTEXT, outline=True)
        self._btn_open.pack(side="left", padx=10)

        self._progress = ttk.Progressbar(mid, mode="indeterminate", length=160)
        self._progress.pack(side="right")

        # ── log ───────────────────────────────────────────────────────────────
        log_frame = tk.Frame(self, bg=CLR_BG, padx=16, pady=4)
        log_frame.pack(fill="both", expand=True)

        log_hdr = tk.Frame(log_frame, bg=CLR_BG)
        log_hdr.pack(fill="x")
        tk.Label(
            log_hdr, text="LOG",
            bg=CLR_BG, fg=CLR_SUBTEXT,
            font=("Segoe UI", 8, "bold")
        ).pack(side="left")
        self._make_btn(
            log_hdr, "Clear Log", CLR_BG, CLR_BG,
            self._clear_log, fg=CLR_SUBTEXT,
            font=("Segoe UI", 8), relief="flat"
        ).pack(side="right")

        self._log = scrolledtext.ScrolledText(
            log_frame, bg=CLR_SURFACE2, fg=CLR_TEXT,
            font=("Consolas", 9), bd=0, highlightthickness=1,
            highlightbackground=CLR_BORDER, state="disabled", wrap="word")
        self._log.pack(fill="both", expand=True, pady=(4, 0))

        self._log.tag_config("ok",      foreground=CLR_SUCCESS)
        self._log.tag_config("fail",    foreground=CLR_ERROR)
        self._log.tag_config("warn",    foreground=CLR_WARN)
        self._log.tag_config("info",    foreground=CLR_SUBTEXT)
        self._log.tag_config("heading", foreground=CLR_HEADER,
                              font=("Consolas", 9, "bold"))

        # ── status bar ────────────────────────────────────────────────────────
        bar = tk.Frame(self, bg=CLR_SURFACE, height=24,
                       highlightthickness=1, highlightbackground=CLR_BORDER)
        bar.pack(fill="x", side="bottom")
        self._lbl_status = tk.Label(
            bar, text="Ready.",
            bg=CLR_SURFACE, fg=CLR_SUBTEXT,
            font=("Segoe UI", 8), anchor="w", padx=10)
        self._lbl_status.pack(side="left", fill="y")

    def _make_btn(self, parent, text, bg, hover_bg, cmd,
                  fg=None, font=("Segoe UI", 9, "bold"),
                  relief="flat", pady=6, padx=12, outline=False):
        if fg is None:
            fg = "#ffffff" if bg in (CLR_HEADER, CLR_MID) else CLR_TEXT
        btn = tk.Button(
            parent, text=text, bg=bg, fg=fg, font=font,
            relief=relief, bd=0, cursor="hand2",
            activebackground=hover_bg, activeforeground=fg,
            pady=pady, padx=padx, command=cmd
        )
        if outline:
            btn.config(highlightthickness=1, highlightbackground=CLR_BORDER)
        btn.bind("<Enter>", lambda e: btn.config(bg=hover_bg))
        btn.bind("<Leave>", lambda e: btn.config(bg=bg))
        return btn

    def _browse_output_dir(self):
        folder = filedialog.askdirectory(
            title="Select output folder for converted PDFs",
            initialdir=str(self._output_dir))
        if folder:
            self._output_dir = Path(folder)
            self._lbl_outdir.config(text=str(self._output_dir))

    def _pick_files(self):
        paths = filedialog.askopenfilenames(
            title="Select files to convert to PDF",
            filetypes=FILETYPES,
            initialdir=str(conv.INPUT_DIR))
        for p in paths:
            path = Path(p)
            if path not in self._files:
                self._files.append(path)
        self._refresh_queue()

    def _pick_folder(self):
        folder = filedialog.askdirectory(
            title="Select folder to scan",
            initialdir=str(conv.INPUT_DIR))
        if not folder:
            return
        added = 0
        for f in sorted(Path(folder).iterdir()):
            if (f.is_file() and f.suffix.lower() in SUPPORTED_TO_PDF
                    and f not in self._files
                    and not f.name.startswith("~$")):
                self._files.append(f)
                added += 1
        self._log_write(f"Folder scanned — {added} supported file(s) added.\n", "info")
        self._refresh_queue()

    def _clear_queue(self):
        self._files.clear()
        self._refresh_queue()

    def _refresh_queue(self):
        self._queue_box.delete(0, "end")
        for f in self._files:
            self._queue_box.insert("end", f"  {f.name}  ({f.parent})")
        count = len(self._files)
        self._lbl_queue_count.config(
            text=f"{count} file{'s' if count != 1 else ''} queued")
        self._btn_convert.config(
            state="normal" if count > 0 and not self._running else "disabled")

    def _start_conversion(self):
        if not self._files or self._running:
            return
        self._running = True
        self._btn_convert.config(state="disabled")
        self._btn_files.config(state="disabled")
        self._btn_folder.config(state="disabled")
        self._progress.start(12)
        self._lbl_status.config(text="Converting…")
        files_snapshot = list(self._files)
        threading.Thread(target=self._run_conversions,
                         args=(files_snapshot,), daemon=True).start()

    def _run_conversions(self, files):
        output_dir = self._output_dir
        output_dir.mkdir(parents=True, exist_ok=True)
        total = len(files)
        ok = fail = 0

        self._log_write(
            f"\n{'='*54}\n"
            f"  Convert → PDF  —  {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
            f"  Files: {total}\n"
            f"  Output: {output_dir}\n"
            f"{'='*54}\n", "heading")

        for i, src in enumerate(files, 1):
            self._lbl_status.config(text=f"Converting {i}/{total}: {src.name}")
            self._log_write(f"\n[{i}/{total}] {src.name}\n", "info")

            dest = output_dir / (src.stem + ".pdf")
            ext  = src.suffix.lower()
            converter_fn = conv.CONVERTERS.get(ext)

            if converter_fn is None:
                self._log_write("  SKIP — unsupported format\n", "warn")
                fail += 1
                continue

            original_log = conv.log
            def gui_log(msg, _orig=original_log):
                self._log_write(msg + "\n",
                    "ok"   if "[OK]"   in msg else
                    "fail" if "[FAIL]" in msg or "[ERROR]" in msg else
                    "warn" if "[WARN]" in msg else "info")
                _orig(msg)
            conv.log = gui_log

            success = converter_fn(src, dest)
            conv.log = original_log

            if success:
                ok += 1
            else:
                fail += 1

        self._log_write(
            f"\n{'─'*54}\n"
            f"  Done.  Converted: {ok}  |  Failed: {fail}\n"
            f"  PDFs saved to: {output_dir}\n"
            f"{'='*54}\n", "heading")

        conv.close_office_apps()
        conv.cleanup_output()
        self.after(0, self._conversion_done, ok, fail)

    def _conversion_done(self, ok, fail):
        self._running = False
        self._progress.stop()
        self._files.clear()
        self._refresh_queue()
        self._btn_files.config(state="normal")
        self._btn_folder.config(state="normal")
        self._lbl_status.config(text=f"Done — {ok} converted, {fail} failed.")

    def _log_write(self, text, tag=""):
        self._log.config(state="normal")
        if tag:
            self._log.insert("end", text, tag)
        else:
            self._log.insert("end", text)
        self._log.see("end")
        self._log.config(state="disabled")

    def _clear_log(self):
        self._log.config(state="normal")
        self._log.delete("1.0", "end")
        self._log.config(state="disabled")

    def _open_output(self):
        os.startfile(str(self._output_dir))


if __name__ == "__main__":
    app = ConverterApp()
    app.mainloop()
