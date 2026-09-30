"""
gdrive_sync_gui.py
------------------
GUI for Google Drive PDF → Word sync.
Reads config from gdrive_sync.py — place both files in the same folder.

Usage:
  python gdrive_sync_gui.py
"""

import sys
import threading
import tkinter as tk
from tkinter import scrolledtext, filedialog
from pathlib import Path
import os

HERE = Path(__file__).parent

# ── Palette ───────────────────────────────────────────────────────────────────
BG         = "#F7F8FA"
SURFACE    = "#FFFFFF"
BORDER     = "#E4E6EA"
TEXT       = "#111827"
SUBTEXT    = "#6B7280"
PRIMARY    = "#2563EB"
PRIMARY_HV = "#1D4ED8"
SUCCESS    = "#059669"
ERROR      = "#DC2626"
WARN       = "#D97706"
LOG_BG     = "#F0F2F5"
LOG_TEXT   = "#374151"

FONT_UI    = ("Segoe UI", 10)
FONT_TITLE = ("Segoe UI", 18, "bold")
FONT_SUB   = ("Segoe UI", 9)
FONT_LOG   = ("Consolas", 9)
FONT_BTN   = ("Segoe UI", 11, "bold")
FONT_LABEL = ("Segoe UI", 8, "bold")

DEFAULT_DRIVE_FOLDER = "19RVhksgAetnFtdjQdT7NEgTn3OzI8yfk"
DEFAULT_LOCAL_DEST   = r"C:\Users\User\Documents\PROJECTS\LRM_DFS\02_PROJECT_FILES\Source_Docs"


class GDriveSyncApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Drive Sync")
        self.geometry("580x660")
        self.minsize(480, 560)
        self.configure(bg=BG)
        self.resizable(True, True)
        self._running = False
        self._build_ui()

    # ── UI construction ───────────────────────────────────────────────────────

    def _build_ui(self):
        tk.Frame(self, bg=BG, height=32).pack(fill="x")

        # title block
        title_frame = tk.Frame(self, bg=BG)
        title_frame.pack(fill="x", padx=36)

        tk.Label(
            title_frame, text="Drive Sync",
            font=FONT_TITLE, bg=BG, fg=TEXT, anchor="w"
        ).pack(anchor="w")

        tk.Label(
            title_frame,
            text="Converts PDFs in your Google Drive folder to .docx and saves them locally.",
            font=FONT_SUB, bg=BG, fg=SUBTEXT, anchor="w", wraplength=500, justify="left"
        ).pack(anchor="w", pady=(2, 0))

        tk.Frame(self, bg=BORDER, height=1).pack(fill="x", padx=36, pady=20)

        # Drive folder row
        self._build_field(
            label="GOOGLE DRIVE FOLDER",
            hint="Paste the Drive folder URL or ID",
            default=DEFAULT_DRIVE_FOLDER,
            attr="_entry_drive",
            browse=False,
        )

        tk.Frame(self, bg=BG, height=12).pack(fill="x")

        # Output folder row
        self._build_field(
            label="SAVE CONVERTED FILES TO",
            hint="Choose a local folder",
            default=DEFAULT_LOCAL_DEST,
            attr="_entry_dest",
            browse=True,
        )

        tk.Frame(self, bg=BG, height=20).pack(fill="x")

        # sync button
        self._btn_sync = tk.Button(
            self,
            text="\u27f3  Sync Now",
            font=FONT_BTN,
            bg=PRIMARY, fg="#ffffff",
            activebackground=PRIMARY_HV, activeforeground="#ffffff",
            relief="flat", bd=0, cursor="hand2",
            padx=28, pady=12,
            command=self._start_sync
        )
        self._btn_sync.pack(padx=36, anchor="w")
        self._btn_sync.bind("<Enter>", lambda e: self._btn_sync.config(bg=PRIMARY_HV))
        self._btn_sync.bind("<Leave>", lambda e: self._btn_sync.config(
            bg=PRIMARY if not self._running else PRIMARY_HV))

        # status label
        self._lbl_status = tk.Label(
            self, text="",
            font=FONT_SUB, bg=BG, fg=SUBTEXT, anchor="w"
        )
        self._lbl_status.pack(fill="x", padx=36, pady=8)

        tk.Frame(self, bg=BORDER, height=1).pack(fill="x", padx=36, pady=16)

        # log header
        log_hdr = tk.Frame(self, bg=BG)
        log_hdr.pack(fill="x", padx=36)

        tk.Label(log_hdr, text="LOG", font=FONT_LABEL, bg=BG, fg=SUBTEXT).pack(side="left")

        self._btn_open = tk.Button(
            log_hdr, text="Open Folder",
            font=("Segoe UI", 8), bg=BG, fg=PRIMARY,
            activebackground=BG, activeforeground=PRIMARY_HV,
            relief="flat", bd=0, cursor="hand2",
            command=self._open_folder
        )
        self._btn_open.pack(side="right")

        tk.Button(
            log_hdr, text="Clear",
            font=("Segoe UI", 8), bg=BG, fg=SUBTEXT,
            activebackground=BG, activeforeground=TEXT,
            relief="flat", bd=0, cursor="hand2",
            command=self._clear_log
        ).pack(side="right", padx=(0, 12))

        # log box
        log_frame = tk.Frame(self, bg=BG)
        log_frame.pack(fill="both", expand=True, padx=36, pady=(6, 24))

        self._log = scrolledtext.ScrolledText(
            log_frame,
            bg=LOG_BG, fg=LOG_TEXT,
            font=FONT_LOG,
            bd=0, relief="flat",
            highlightthickness=1, highlightbackground=BORDER,
            state="disabled", wrap="word"
        )
        self._log.pack(fill="both", expand=True)
        self._log.tag_config("ok",   foreground=SUCCESS)
        self._log.tag_config("err",  foreground=ERROR)
        self._log.tag_config("warn", foreground=WARN)
        self._log.tag_config("dim",  foreground=SUBTEXT)

    def _build_field(self, label, hint, default, attr, browse):
        frame = tk.Frame(self, bg=BG)
        frame.pack(fill="x", padx=36)

        tk.Label(frame, text=label, font=FONT_LABEL, bg=BG, fg=SUBTEXT).pack(anchor="w", pady=(0, 4))

        row = tk.Frame(frame, bg=BG)
        row.pack(fill="x")

        entry = tk.Entry(
            row,
            font=("Segoe UI", 9), bg=SURFACE, fg=TEXT,
            insertbackground=TEXT,
            highlightthickness=1, highlightbackground=BORDER,
            relief="flat"
        )
        entry.insert(0, default)
        entry.pack(side="left", fill="x", expand=True, ipady=7)
        setattr(self, attr, entry)

        if browse:
            tk.Button(
                row, text="Browse",
                font=("Segoe UI", 9), bg=PRIMARY, fg="#ffffff",
                activebackground=PRIMARY_HV, activeforeground="#ffffff",
                relief="flat", bd=0, cursor="hand2",
                padx=12, pady=0,
                command=lambda e=entry: self._browse_folder(e)
            ).pack(side="left", padx=(6, 0), ipady=7)

    def _browse_folder(self, entry):
        folder = filedialog.askdirectory(title="Select output folder")
        if folder:
            entry.delete(0, "end")
            entry.insert(0, folder)

    # ── Sync logic ────────────────────────────────────────────────────────────

    def _start_sync(self):
        if self._running:
            return

        drive_val = self._entry_drive.get().strip()
        dest_val  = self._entry_dest.get().strip()

        if not drive_val:
            self._lbl_status.config(text="Enter a Google Drive folder URL or ID.", fg=ERROR)
            return
        if not dest_val:
            self._lbl_status.config(text="Choose a local folder to save files.", fg=ERROR)
            return

        self._running = True
        self._btn_sync.config(text="\u27f3  Syncing\u2026", state="disabled", bg=PRIMARY_HV)
        self._lbl_status.config(text="Connecting to Google Drive\u2026", fg=SUBTEXT)
        self._clear_log()
        threading.Thread(
            target=self._run_sync,
            args=(drive_val, dest_val),
            daemon=True
        ).start()

    def _run_sync(self, drive_folder, local_dest):
        try:
            sys.path.insert(0, str(HERE))
            import importlib
            import gdrive_sync as gs

            original_print = print
            counts = {"converted": 0, "failed": 0}

            def gui_print(*args, **kwargs):
                msg = " ".join(str(a) for a in args)
                end = kwargs.get("end", "\n")
                tag = "dim"
                if "OK" in msg:
                    tag = "ok"
                    counts["converted"] += 1
                elif "FAILED" in msg or "ERROR" in msg:
                    tag = "err"
                    if "FAILED" in msg:
                        counts["failed"] += 1
                elif "WARN" in msg:
                    tag = "warn"
                self.after(0, self._append_log, msg + ("" if end == "" else ""), tag)

            import builtins
            builtins.print = gui_print

            try:
                importlib.reload(gs)
                gs.run(drive_folder, local_dest)
            finally:
                builtins.print = original_print

            self.after(0, self._sync_done, counts, local_dest)

        except Exception as e:
            self.after(0, self._append_log, f"[ERROR] {e}", "err")
            self.after(0, self._sync_done, {"converted": 0, "failed": 1}, local_dest)

    def _sync_done(self, counts, local_dest):
        self._running = False
        self._btn_sync.config(text="\u27f3  Sync Now", state="normal", bg=PRIMARY)
        converted = counts.get("converted", 0)
        failed    = counts.get("failed", 0)
        if failed == 0 and converted > 0:
            self._lbl_status.config(text=f"Done \u2014 {converted} file(s) converted.", fg=SUCCESS)
        elif failed > 0:
            self._lbl_status.config(text=f"Done \u2014 {converted} converted, {failed} failed.", fg=WARN)
        else:
            self._lbl_status.config(text="Done \u2014 no PDF files found in Drive folder.", fg=SUBTEXT)
        self._current_dest = local_dest

    # ── Log helpers ───────────────────────────────────────────────────────────

    def _append_log(self, text, tag=""):
        self._log.config(state="normal")
        if tag:
            self._log.insert("end", text + "\n", tag)
        else:
            self._log.insert("end", text + "\n")
        self._log.see("end")
        self._log.config(state="disabled")

    def _clear_log(self):
        self._log.config(state="normal")
        self._log.delete("1.0", "end")
        self._log.config(state="disabled")

    def _open_folder(self):
        folder = Path(getattr(self, "_current_dest", self._entry_dest.get().strip()))
        if folder.exists():
            os.startfile(str(folder))
        else:
            self._lbl_status.config(text="Output folder not found.", fg=ERROR)


if __name__ == "__main__":
    app = GDriveSyncApp()
    app.mainloop()
