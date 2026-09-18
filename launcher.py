"""
launcher.py — DocumentConverter Main Menu
Launches converter_gui.py (Convert to PDF) or pdf_to_office_gui.py (PDF to Office)
as separate subprocess windows.
"""

import sys
import subprocess
import tkinter as tk
from pathlib import Path

# ── NZB CP 002 palette ───────────────────────────────────────────────────────
CLR_BG        = "#E0E0E0"   # window background
CLR_SURFACE   = "#F4F4F4"   # card / panel surface
CLR_HEADER    = "#3A606E"   # title bar / primary button
CLR_HEADER_HV = "#2d4d59"   # title bar hover
CLR_MID       = "#607B7D"   # secondary elements, borders
CLR_TEXT      = "#1E2D30"   # primary text
CLR_SUBTEXT   = "#607B7D"   # muted text
CLR_BORDER    = "#C8CACB"   # dividers / outlines

BASE_DIR = Path(__file__).resolve().parent


def launch(script: str):
    """Open a converter GUI as a separate process — no console window."""
    target = BASE_DIR / script
    kwargs = {}
    if sys.platform == "win32":
        kwargs["creationflags"] = 0x08000000  # CREATE_NO_WINDOW
    subprocess.Popen(
        [sys.executable, str(target)],
        cwd=str(BASE_DIR),
        **kwargs,
    )


def build_button(parent, text, sub, color, hover_color, command, fg="#ffffff"):
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
    root = tk.Tk()
    root.title("DocumentConverter")
    root.configure(bg=CLR_BG)
    root.resizable(False, False)

    w, h = 460, 310
    root.geometry(f"{w}x{h}+{(root.winfo_screenwidth()-w)//2}+{(root.winfo_screenheight()-h)//2}")

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

    build_button(
        body,
        text="▶   Convert Files to PDF",
        sub="DOCX · XLSX · PPTX · HTML · Markdown · CSV · TXT",
        color=CLR_HEADER,
        hover_color=CLR_HEADER_HV,
        fg="#ffffff",
        command=lambda: launch("converter_gui.py"),
    )

    build_button(
        body,
        text="▶   Convert PDF to Office",
        sub="PDF → Word (.docx)  ·  PDF → Excel (.xlsx)",
        color=CLR_HEADER,
        hover_color=CLR_HEADER_HV,
        fg="#ffffff",
        command=lambda: launch("pdf_to_office_gui.py"),
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
