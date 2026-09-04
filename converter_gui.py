"""
=============================================================
  Convert to PDF — Desktop GUI  (Phase 5)
  Usage:
    python converter_gui.py       (standalone; launcher preferred)

  Wraps converter.py — all conversion logic lives there.
  Converts: .docx .doc .xlsx .xls .pptx .ppt
            .html .htm .md .txt .csv  ->  PDF

  For PDF -> Word / Excel, use: pdf_to_office_gui.py
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

# F-19: TkinterDnD for drag-and-drop; gracefully degrade if not installed
try:
    from tkinterdnd2 import DND_FILES
    _DND_AVAILABLE = True
except ImportError:
    _DND_AVAILABLE = False

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


# F-03: ConverterApp is now a tk.Frame (not tk.Tk).
# The launcher creates a TkinterDnD.Toplevel and packs this frame inside it.
class ConverterApp(tk.Frame):
    def __init__(self, master, **kwargs):
        super().__init__(master, bg=CLR_BG, **kwargs)

        # Configure the host window (Toplevel or Tk root from launcher / standalone run)
        master.title("Convert to PDF")
        master.geometry("780x640")
        master.minsize(640, 520)
        master.configure(bg=CLR_BG)
        master.resizable(True, True)

        self.pack(fill="both", expand=True)  # fill the master window

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

        queue_hdr = tk.Frame(queue_frame, bg=CLR_BG)
        queue_hdr.pack(fill="x")
        tk.Label(
            queue_hdr, text="QUEUE",
            bg=CLR_BG, fg=CLR_SUBTEXT,
            font=("Segoe UI", 8, "bold")
        ).pack(side="left", pady=(6, 3))

        # F-10/F-19: DnD hint shown when TkinterDnD is available
        if _DND_AVAILABLE:
            tk.Label(
                queue_hdr, text="(drag & drop files here)",
                bg=CLR_BG, fg=CLR_SUBTEXT,
                font=("Segoe UI", 7, "italic")
            ).pack(side="left", padx=(6, 0), pady=(6, 3))

        self._queue_box = tk.Listbox(
            queue_frame, bg=CLR_SURFACE2, fg=CLR_TEXT,
            selectbackground=CLR_HEADER,
            font=("Segoe UI", 9), height=5, bd=0,
            highlightthickness=1, highlightbackground=CLR_BORDER,
            activestyle="none")
        self._queue_box.pack(fill="x")

        # F-10/F-19: register the queue listbox as a DnD drop target
        if _DND_AVAILABLE:
            self._queue_box.drop_target_register(DND_FILES)
            self._queue_box.dnd_bind('<<Drop>>', self._on_dnd_drop)

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

        self._logbox = scrolledtext.ScrolledText(
            log_frame, bg=CLR_SURFACE2, fg=CLR_TEXT,
            font=("Consolas", 9), bd=0, highlightthickness=1,
            highlightbackground=CLR_BORDER, state="disabled", wrap="word")
        self._logbox.pack(fill="both", expand=True, pady=(4, 0))

        self._logbox.tag_config("ok",      foreground=CLR_SUCCESS)
        self._logbox.tag_config("fail",    foreground=CLR_ERROR)
        self._logbox.tag_config("warn",    foreground=CLR_WARN)
        self._logbox.tag_config("info",    foreground=CLR_SUBTEXT)
        self._logbox.tag_config("heading", foreground=CLR_HEADER,
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

    # F-10/F-19: handle file drop from TkinterDnD
    def _on_dnd_drop(self, event):
        raw = event.data
        # TkinterDnD wraps paths with spaces in {braces}; parse them out
        paths = []
        i = 0
        while i < len(raw):
            if raw[i] == '{':
                end = raw.index('}', i)
                paths.append(raw[i+1:end])
                i = end + 1
            elif raw[i] == ' ':
                i += 1
            else:
                j = raw.find(' ', i)
                if j == -1:
                    j = len(raw)
                paths.append(raw[i:j])
                i = j
        added = 0
        for p in paths:
            path = Path(p)
            if path.is_file() and path.suffix.lower() in SUPPORTED_TO_PDF:
                if path not in self._files:
                    self._files.append(path)
                    added += 1
            elif path.is_dir():
                for f in sorted(path.iterdir()):
                    if (f.is_file() and f.suffix.lower() in SUPPORTED_TO_PDF
                            and f not in self._files
                            and not f.name.startswith("~$")):
                        self._files.append(f)
                        added += 1
        if added:
            self._log_write(f"Dropped {added} file(s) into queue.\n", "info")
        self._refresh_queue()

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
        # F-05: status update from the main thread — safe
        self._lbl_status.config(text="Converting…")
        files_snapshot = list(self._files)
        threading.Thread(target=self._run_conversions,
                         args=(files_snapshot,), daemon=True).start()

    def _run_conversions(self, files):
        # F-05: all widget mutations from this worker thread go through self.after()
        output_dir = self._output_dir
        output_dir.mkdir(parents=True, exist_ok=True)
        total = len(files)
        ok = fail = 0

        # F-05: helpers that schedule UI updates on the main thread
        def _set_status(msg):
            self.after(0, self._lbl_status.config, {"text": msg})

        def _append_log(text, tag=""):
            self.after(0, self._log_write, text, tag)

        # F-20: log_fn callback — no monkey-patching of conv.log.
        # Routes to the GUI via self.after() AND calls the original logger for
        # file/console output so nothing is lost.
        _original_log = conv.log

        def _gui_log(msg):
            tag = ("ok"   if "[OK]"   in msg else
                   "fail" if "[FAIL]" in msg or "[ERROR]" in msg else
                   "warn" if "[WARN]" in msg else "info")
            # F-05: schedule the GUI write on the main thread
            self.after(0, self._log_write, msg + "\n", tag)
            _original_log(msg)   # preserve file/console logging

        _append_log(
            f"\n{'='*54}\n"
            f"  Convert -> PDF  —  {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
            f"  Files: {total}\n"
            f"  Output: {output_dir}\n"
            f"{'='*54}\n", "heading")

        for i, src in enumerate(files, 1):
            # F-05: schedule status update via after()
            _set_status(f"Converting {i}/{total}: {src.name}")
            _append_log(f"\n[{i}/{total}] {src.name}\n", "info")

            ext = src.suffix.lower()

            if ext not in SUPPORTED_TO_PDF:
                _append_log("  SKIP — unsupported format\n", "warn")
                fail += 1
                continue

            # F-20: pass log_fn; converter routes dest via _safe_dest internally
            # F-16: pass output_dir so converter saves to the user-chosen folder
            conv.OUTPUT_DIR = output_dir   # point module-level dir at chosen folder
            success = conv.convert_file(src, log_fn=_gui_log)

            if success:
                ok += 1
            else:
                fail += 1

        _append_log(
            f"\n{'─'*54}\n"
            f"  Done.  Converted: {ok}  |  Failed: {fail}\n"
            f"  PDFs saved to: {output_dir}\n"
            f"{'='*54}\n", "heading")

        # F-24: close_office_apps() removed from routine batch end;
        #        it is registered only as atexit emergency handler in converter.py
        conv.cleanup_output(output_dir)  # F-16: pass output_dir

        # F-05: schedule the UI-reset callback on the main thread
        self.after(0, self._conversion_done, ok, fail)

    def _conversion_done(self, ok, fail):
        # Called on main thread via self.after(); safe to touch widgets here
        self._running = False
        self._progress.stop()
        self._files.clear()
        self._refresh_queue()
        self._btn_files.config(state="normal")
        self._btn_folder.config(state="normal")
        self._lbl_status.config(text=f"Done — {ok} converted, {fail} failed.")

    def _log_write(self, text, tag=""):
        # Must only be called from the main thread (or via self.after)
        self._logbox.config(state="normal")
        if tag:
            self._logbox.insert("end", text, tag)
        else:
            self._logbox.insert("end", text)
        self._logbox.see("end")
        self._logbox.config(state="disabled")

    def _clear_log(self):
        self._logbox.config(state="normal")
        self._logbox.delete("1.0", "end")
        self._logbox.config(state="disabled")

    def _open_output(self):
        os.startfile(str(self._output_dir))


# F-03: standalone entry point — creates a TkinterDnD root if available,
#        plain Tk otherwise; packs ConverterApp (now a Frame) into it.
if __name__ == "__main__":
    if _DND_AVAILABLE:
        from tkinterdnd2 import TkinterDnD
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
    app = ConverterApp(root)
    root.mainloop()
