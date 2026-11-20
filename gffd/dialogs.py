"""About and Setup Guide dialogs."""

from __future__ import annotations

import tkinter as tk
import webbrowser

from . import APP_TITLE, __version__
from .widgets import (ACCENT, AMBER, BG, BLUE, BORDER, GREEN, PURPLE, SURFACE,
                      TEXT_DIM, TEXT_SEC, Btn, apply_icon, label, px, scrollbar)

AUTHOR = "Soumya Subham Rout"
AUTHOR_EMAIL = "soumyasubham.ssr@gmail.com"
REPO_URL = "https://github.com/soumyasubham-ssr/Google-Form-File-Downloader"
PROFILE_URL = "https://github.com/soumyasubham-ssr"


def _dialog(parent: tk.Misc, title: str, resizable_y=False) -> tk.Toplevel:
    dialog = tk.Toplevel(parent)
    dialog.withdraw()
    dialog.title(title)
    dialog.configure(bg=BG)
    dialog.resizable(False, resizable_y)
    dialog.transient(parent)
    apply_icon(dialog)
    dialog.bind("<Escape>", lambda _e: dialog.destroy())
    return dialog


def _place(dialog: tk.Toplevel, parent: tk.Misc, width: int, height: int):
    px = parent.winfo_x() + (parent.winfo_width() - width) // 2
    py = parent.winfo_y() + (parent.winfo_height() - height) // 2
    dialog.geometry(f"{width}x{height}+{max(0, px)}+{max(0, py)}")
    dialog.deiconify()
    dialog.grab_set()
    dialog.focus_set()


def _link(parent, text, url):
    widget = label(parent, text, fg=ACCENT, anchor="w", cursor="hand2")
    widget.bind("<Button-1>", lambda _e: webbrowser.open(url))
    return widget


# ── About ──────────────────────────────────────────────────────────

def show_about(parent: tk.Misc):
    d = _dialog(parent, f"About {APP_TITLE}")
    f = tk.Frame(d, bg=BG, padx=32, pady=26)
    f.pack(fill="both", expand=True)

    label(f, "ABOUT THIS APP", 8, True, ACCENT, anchor="w").pack(fill="x", pady=(0, 4))
    label(f, APP_TITLE, 18, True, anchor="w").pack(fill="x")
    label(f, f"Version {__version__}", fg=ACCENT, anchor="w").pack(fill="x", pady=(1, 14))
    label(f, "Designed & Developed by", 9, True, TEXT_SEC, anchor="w").pack(fill="x")
    label(f, AUTHOR, 14, True, anchor="w").pack(fill="x", pady=(1, 10))
    label(f, "Building productivity software, automation tools, desktop\n"
             "applications, and AI-powered solutions with a focus on\n"
             "simplicity, performance, and reliability.",
          fg=TEXT_SEC, anchor="w", justify="left").pack(fill="x", pady=(0, 14))
    label(f, "Contact", 9, True, anchor="w").pack(fill="x")
    _link(f, f"Email: {AUTHOR_EMAIL}", f"mailto:{AUTHOR_EMAIL}").pack(fill="x", pady=(1, 0))
    _link(f, "GitHub: github.com/soumyasubham-ssr", PROFILE_URL).pack(fill="x", pady=(1, 0))
    _link(f, "Documentation & updates", REPO_URL).pack(fill="x", pady=(1, 10))
    label(f, f"© 2026 {AUTHOR}. All Rights Reserved.", 8, fg=TEXT_DIM, anchor="w").pack(fill="x")

    row = tk.Frame(f, bg=BG)
    row.pack(fill="x", pady=(12, 0))
    Btn(row, text="OK", w=80, h=30, r=5, command=d.destroy).pack(side="right")

    d.update_idletasks()
    _place(d, parent, f.winfo_reqwidth(), f.winfo_reqheight())


# ── Setup guide ────────────────────────────────────────────────────

GUIDE_STEPS = [
    ("Step 1: Share the Google Drive files",
     "Google Form uploads are stored in a Drive folder owned by the\n"
     "form owner. Open that folder and click Share, then set\n"
     "General access to \"Anyone with the link\" (Viewer is enough)."),
    ("Step 2: Prepare the spreadsheet",
     "Use an .xlsx or .csv file with a header row and two columns:\n\n"
     "  Column A – New filename (extension optional; blank keeps\n"
     "                    the original name).\n"
     "  Column B – File link(s) copied from the responses sheet.\n"
     "                    Several comma-separated links are allowed.\n\n"
     "Click Template to save a ready-made workbook."),
    ("Step 3: Choose input and output",
     "Click Browse next to \"Spreadsheet\" and select your file,\n"
     "then choose the output folder for the downloads."),
    ("Step 4: Start the download",
     "Click Start Download. Files are downloaded in parallel,\n"
     "renamed using Column A and saved to the output folder.\n"
     "Click Cancel at any time to stop the run safely."),
    ("Step 5: Review the results",
     "A CSV log (download_log_<date>.csv) is written to the output\n"
     "folder listing every row, its status and any error message."),
]

GUIDE_NOTES = (
    "• Files that are not shared publicly cannot be downloaded.\n"
    "• Duplicate links are downloaded once and reported as Skipped.\n"
    "• Existing files are never overwritten; a suffix like _2 is added.\n"
    "• Keep the column order: filename first, link second."
)

STEP_COLORS = [ACCENT, GREEN, BLUE, AMBER, PURPLE]


def show_guide(parent: tk.Misc):
    d = _dialog(parent, "Setup Guide", resizable_y=True)

    canvas = tk.Canvas(d, bg=BG, highlightthickness=0, bd=0)
    bar = scrollbar(d, canvas.yview)
    canvas.configure(yscrollcommand=bar.set)
    bar.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)

    inner = tk.Frame(canvas, bg=BG, padx=28, pady=20)
    win_id = canvas.create_window((0, 0), window=inner, anchor="nw")

    def on_configure(_e=None):
        canvas.configure(scrollregion=canvas.bbox("all"))
        canvas.itemconfigure(win_id, width=canvas.winfo_width())

    inner.bind("<Configure>", on_configure)
    canvas.bind("<Configure>", on_configure)
    d.bind("<MouseWheel>", lambda e: canvas.yview_scroll(int(-e.delta / 120), "units"))

    label(inner, "How to Set Up", 15, True, anchor="w").pack(fill="x")
    label(inner, APP_TITLE, 15, True, ACCENT, anchor="w").pack(fill="x", pady=(0, 14))

    for index, (title, body) in enumerate(GUIDE_STEPS):
        color = STEP_COLORS[index % len(STEP_COLORS)]
        step = tk.Frame(inner, bg=SURFACE, highlightbackground=BORDER,
                        highlightthickness=1, padx=14, pady=10)
        step.pack(fill="x", pady=(0, 8))
        header = tk.Frame(step, bg=SURFACE)
        header.pack(fill="x", pady=(0, 4))
        size = px(header, 10)
        dot = tk.Canvas(header, width=size, height=size, bg=SURFACE, highlightthickness=0)
        dot.pack(side="left", padx=(0, 8), pady=2)
        dot.create_oval(1, 1, size - 1, size - 1, fill=color, outline="")
        label(header, title, 10, True, color, anchor="w").pack(side="left")
        label(step, body, fg=TEXT_SEC, anchor="w", justify="left").pack(fill="x", padx=(18, 0))

    notes = tk.Frame(inner, bg="#1a1520", highlightbackground="#3d2e50",
                     highlightthickness=1, padx=14, pady=10)
    notes.pack(fill="x", pady=(6, 10))
    label(notes, "Important Notes", 10, True, AMBER, anchor="w").pack(fill="x", pady=(0, 4))
    label(notes, GUIDE_NOTES, fg=TEXT_SEC, anchor="w", justify="left").pack(fill="x")

    row = tk.Frame(inner, bg=BG)
    row.pack(fill="x", pady=(4, 0))
    Btn(row, text="Got it", w=80, h=30, r=5, command=d.destroy).pack(side="right")

    d.update_idletasks()
    width = min(max(px(parent, 500), inner.winfo_reqwidth() + px(parent, 20)), parent.winfo_screenwidth() - 100)
    height = min(inner.winfo_reqheight() + 4, parent.winfo_screenheight() - 80)
    _place(d, parent, width, height)
