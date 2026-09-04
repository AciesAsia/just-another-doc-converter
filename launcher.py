"""
launcher.py — DocumentConverter Main Menu

F-03 (Fix B): Opens each converter GUI as a TkinterDnD.Toplevel within this
single process rather than spawning a subprocess.  No subprocess, no
multiprocessing, no freeze_support() needed — works in both frozen (.exe) and
development environments.

Architecture:
  root            TkinterDnD.Tk      — the launcher menu window
  converter win   TkinterDnD.Toplevel — ConverterApp (tk.Frame) packed inside
  pdf2office win  TkinterDnD.Toplevel — PdfToOfficeApp (tk.Frame) packed inside
"""

import sys
import tkinter as tk
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

# F-03/F-19: import TkinterDnD; degrade gracefully to plain Tk if not installed
try:
    from tkinterdnd2 import TkinterDnD
    _DND_AVAILABLE = True
except ImportError:
    TkinterDnD = None
    _DND_AVAILABLE = False

# F-03: import the converter GUIs as modules — no subprocess needed
from converter_gui import ConverterApp
from pdf_to_office_gui import PdfToOfficeApp

# ── NZB CP 002 palette ───────────────────────────────────────────────────────
CLR_BG        = "#E0E0E0"
CLR_SURFACE   = "#F4F4F4"
CLR_HEADER    = "#3A606E"
CLR_HEADER_HV = "#2d4d59"
CLR_MID       = "#607B7D"
CLR_TEXT      = "#1E2D30"
CLR_SUBTEXT   = "#607B7D"
CLR_BORDER    = "#C8CACB"


def _make_toplevel(root):
    # F-03/F-19: create a DnD-capable Toplevel; fall back to plain Toplevel
    if _DND_AVAILABLE:
        return TkinterDnD.Toplevel(root)
    return tk.Toplevel(root)


def _open_converter(root):
    """Open the 'Convert to PDF' window as a Toplevel in this process."""
    win = _make_toplevel(root)
    ConverterApp(win)   # ConverterApp.__init__ configures win title/size/etc.
    win.lift()
    win.focus_force()


def _open_pdf_to_office(root):
    """Open the 'Convert PDF to Office' window as a Toplevel in this process."""
    win = _make_toplevel(root)
    PdfToOfficeApp(win)   # PdfToOfficeApp.__init__ configures win title/size/etc.
    win.lift()
    win.focus_force()


def _build_button(parent, text, sub, color, hover_color, command, fg="#ffffff"):
    """Create a styled launcher card button."""
    card = tk.Frame(
        parent, bg=CLR_SURFACE, bd=0, relief="flat",
        highlightthickness=1, highlightbackground=CLR_BORDER
    )
    card.pack(fill="x", pady=6, padx=2)

    btn = tk.Button(
        card,
        text=text,
        font=("Segoe UI", 12, "bold"),
        fg=fg,
        bg=color,
        activebackground=hover_color,
        activeforeground=fg,
        relief="flat",
        cursor="hand2",
        bd=0,
        padx=20,
        pady=13,
        command=command,
        anchor="w",
    )
    btn.pack(fill="x", padx=12, pady=(12, 4))

    tk.Label(
        card,
        text=sub,
        font=("Segoe UI", 9),
        fg=CLR_SUBTEXT,
        bg=CLR_SURFACE,
        anchor="w",
    ).pack(fill="x", padx=14, pady=(0, 12))

    btn.bind("<Enter>", lambda e: btn.config(bg=hover_color))
    btn.bind("<Leave>", lambda e: btn.config(bg=color))

    return card


def main():
    # F-03/F-19: root is TkinterDnD.Tk so child Toplevels support drag-and-drop
    if _DND_AVAILABLE:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()

    root.title("DocumentConverter")
    root.configure(bg=CLR_BG)
    root.resizable(False, False)

    w, h = 460, 310
    sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
    root.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

    # ── header bar ────────────────────────────────────────────────────────────
    hdr_bar = tk.Frame(root, bg=CLR_HEADER, height=46)
    hdr_bar.pack(fill="x")
    hdr_bar.pack_propagate(False)
    tk.Label(
        hdr_bar, text="DocumentConverter",
        font=("Segoe UI", 14, "bold"),
        fg="#ffffff", bg=CLR_HEADER
    ).pack(side="left", padx=18, pady=10)

    # ── subtitle ──────────────────────────────────────────────────────────────
    sub_frame = tk.Frame(root, bg=CLR_BG)
    sub_frame.pack(fill="x", padx=20, pady=(14, 4))
    tk.Label(
        sub_frame, text="Local · Free · No cloud dependency",
        font=("Segoe UI", 9),
        fg=CLR_SUBTEXT, bg=CLR_BG
    ).pack(anchor="w")

    # ── divider ───────────────────────────────────────────────────────────────
    tk.Frame(root, bg=CLR_BORDER, height=1).pack(fill="x", padx=20, pady=(6, 2))

    # ── launcher cards ────────────────────────────────────────────────────────
    body = tk.Frame(root, bg=CLR_BG)
    body.pack(fill="both", expand=True, padx=18, pady=8)

    _build_button(
        body,
        text="▶   Convert Files to PDF",
        sub="DOCX · XLSX · PPTX · HTML · Markdown · CSV · TXT",
        color=CLR_HEADER,
        hover_color=CLR_HEADER_HV,
        fg="#ffffff",
        command=lambda: _open_converter(root),     # F-03: Toplevel, not subprocess
    )

    _build_button(
        body,
        text="▶   Convert PDF to Office",
        sub="PDF → Word (.docx)  ·  PDF → Excel (.xlsx)",
        color=CLR_HEADER,
        hover_color=CLR_HEADER_HV,
        fg="#ffffff",
        command=lambda: _open_pdf_to_office(root), # F-03: Toplevel, not subprocess
    )

    # ── footer ────────────────────────────────────────────────────────────────
    tk.Label(
        root, text="v2.0  ·  Windows",
        font=("Segoe UI", 8),
        fg=CLR_SUBTEXT, bg=CLR_BG
    ).pack(pady=(0, 10))

    root.mainloop()


if __name__ == "__main__":
    main()
