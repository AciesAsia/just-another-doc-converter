"""
gdrive_sync_gui.py
------------------
Minimal GUI for Google Drive PDF → Word sync.
Reads config from gdrive_sync.py — place both files in the same folder.

Usage:
  python gdrive_sync_gui.py
"""

import sys
import threading
import tkinter as tk
from tkinter import scrolledtext
from pathlib import Path
import subprocess
import os

HERE = Path(__file__).parent

# ── Palette ───────────────────────────────────────────────────────────────────
BG          = "#F7F8FA"
SURFACE     = "#FFFFFF"
BORDER      = "#E4E6EA"
TEXT        = "#111827"
SUBTEXT     = "#6B7280"
PRIMARY     = "#2563EB"
PRIMARY_HV  = "#1D4ED8"
SUCCESS     = "#059669"
ERROR       = "#DC2626"
WARN        = "#D97706"
LOG_BG      = "#F0F2F5"
LOG_TEXT    = "#374151"

FONT_UI     = ("Segoe UI", 10)
FONT_TITLE  = ("Segoe UI", 18, "bold")
FONT_SUB    = ("Segoe UI", 9)
FONT_LOG    = ("Consolas", 9)
FONT_BTN    = ("Segoe UI", 11, "bold")
FONT_LABEL  = ("Segoe UI", 8, "bold")


class GDriveSyncApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Drive Sync")
        self.geometry("560x580")
        self.minsize(480, 480)
        self.configure(bg=BG)
        self.resizable(True, True)
        self._running = False
        self._build_ui()

    # ── UI construction ───────────────────────────────────────────────────────

    def _build_ui(self):
        # top spacer
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

        # divider
        tk.Frame(self, bg=BORDER, height=1).pack(fill="x", padx=36, pady=20)

        # destination row
        dest_frame = tk.Frame(self, bg=BG)
        dest_frame.pack(fill="x", padx=36)

        tk.Label(
            dest_frame, text="SAVING TO",
            font=FONT_LABEL, bg=BG, fg=SUBTEXT
        ).pack(anchor="w", pady=(0, 4))

        dest_path = r"C:\Users\User\Documents\PROJECTS\LRM_DFS\02_PROJECT_FILES\Source_Docs"
        self._lbl_dest = tk.Label(
            dest_frame, text=dest_path,
            font=("Segoe UI", 9), bg=SURFACE, fg=TEXT,
            anchor="w", padx=10,
            highlightthickness=1, highlightbackground=BORDER,
            relief="flat", cursor="arrow"
        )
        self._lbl_dest.pack(fill="x", ipady=7)

        tk.Frame(self, bg=BG, height=16).pack(fill="x")

        # sync button
        self._btn_sync = tk.Button(
            self,
            text="⟳  Sync Now",
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
        self._lbl_status.pack(fill="x", padx=36, pady=(8, 0))

        # divider
        tk.Frame(self, bg=BORDER, height=1).pack(fill="x", padx=36, pady=16)

        # log header row
        log_hdr = tk.Frame(self, bg=BG)
        log_hdr.pack(fill="x", padx=36)

        tk.Label(
            log_hdr, text="LOG",
            font=FONT_LABEL, bg=BG, fg=SUBTEXT
        ).pack(side="left")

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
        log_frame = tk.Frame(self, bg=BG, padx=36, pady=(6, 0))
        log_frame.pack(fill="both", expand=True, pady=(0, 24))

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

    # ── Sync logic ────────────────────────────────────────────────────────────

    def _start_sync(self):
        if self._running:
            return
        self._running = True
        self._btn_sync.config(text="⟳  Syncing…", state="disabled", bg=PRIMARY_HV)
        self._lbl_status.config(text="Connecting to Google Drive…", fg=SUBTEXT)
        self._clear_log()
        threading.Thread(target=self._run_sync, daemon=True).start()

    def _run_sync(self):
        try:
            # Import gdrive_sync inline so we capture its print output
            sys.path.insert(0, str(HERE))

            # Redirect stdout so we can capture log lines
            import io
            from contextlib import redirect_stdout

            # Patch gdrive_sync to use our logger
            import importlib
            import gdrive_sync as gs

            # Monkey-patch print inside gdrive_sync
            original_print = __builtins__["print"] if isinstance(__builtins__, dict) else print

            counts = {"converted": 0, "failed": 0}

            def gui_print(*args, **kwargs):
                msg = " ".join(str(a) for a in args)
                end = kwargs.get("end", "\n")
                flush = kwargs.get("flush", False)
                tag = "dim"
                if "OK" in msg:
                    tag = "ok"
                    if ".docx" in msg.lower() or "saved" in msg.lower() or "deleted" in msg.lower():
                        counts["converted"] += 1
                elif "FAILED" in msg or "ERROR" in msg:
                    tag = "err"
                    if "FAILED" in msg and "Converting:" in self._last_line:
                        counts["failed"] += 1
                elif "WARN" in msg:
                    tag = "warn"
                self._last_line = msg
                self.after(0, self._append_log, msg + ("" if end == "" else ""), tag)

            self._last_line = ""

            import builtins
            builtins.print = gui_print

            try:
                importlib.reload(gs)
                gs.run()
            finally:
                builtins.print = original_print

            # Count from log
            self.after(0, self._sync_done, counts)

        except Exception as e:
            self.after(0, self._append_log, f"[ERROR] {e}", "err")
            self.after(0, self._sync_done, {"converted": 0, "failed": 1})

    def _sync_done(self, counts):
        self._running = False
        self._btn_sync.config(text="⟳  Sync Now", state="normal", bg=PRIMARY)
        converted = counts.get("converted", 0)
        failed = counts.get("failed", 0)
        if failed == 0 and converted > 0:
            self._lbl_status.config(
                text=f"Done — {converted} file(s) converted.", fg=SUCCESS)
        elif failed > 0:
            self._lbl_status.config(
                text=f"Done — {converted} converted, {failed} failed.", fg=WARN)
        else:
            self._lbl_status.config(
                text="Done — no files found in Drive folder.", fg=SUBTEXT)

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
        folder = Path(r"C:\Users\User\Documents\PROJECTS\LRM_DFS\02_PROJECT_FILES\Source_Docs")
        if folder.exists():
            os.startfile(str(folder))
        else:
            self._lbl_status.config(text="Output folder not found.", fg=ERROR)


if __name__ == "__main__":
    app = GDriveSyncApp()
    app.mainloop()
