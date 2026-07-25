"""Google Form File Downloader – download files linked in an Excel workbook."""

import csv
import ctypes
import mimetypes
import os
import queue
import re
import shutil
import sys
import threading
import time
import uuid
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path
from urllib.parse import unquote, urlparse

import tkinter as tk
from tkinter import filedialog, messagebox

import pandas as pd
import requests


APP_TITLE = "Google Form File Downloader"
APP_VERSION = "2.0.0"
MAX_WORKERS = 20
INVALID_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')

# ── Theme ──────────────────────────────────────────────────────────
BG            = "#0b0f17"
SURFACE       = "#131a27"
SURFACE_LIGHT = "#1c2638"
ACCENT        = "#8b7cff"
ACCENT_HOVER  = "#a69cff"
ACCENT_DIM    = "#6659d6"
TEXT          = "#f2f4f8"
TEXT_SEC      = "#a6b0c2"
TEXT_DIM      = "#68758a"
GREEN         = "#45d483"
RED           = "#ff6b7a"
AMBER         = "#f7c75a"
BORDER        = "#263249"
ENTRY_BG      = "#0f1623"
LOG_BG        = "#0a101a"


def _resource_path(relative: str) -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / relative


def _enable_windows_dpi_awareness():
    """Request per-monitor DPI awareness before Tk creates its first window."""
    if os.name != "nt":
        return
    try:
        # DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2
        ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4))
        return
    except (AttributeError, OSError, OverflowError):
        pass
    try:
        # Fallback for older Windows versions.
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
        return
    except (AttributeError, OSError):
        pass
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except (AttributeError, OSError):
        pass


# ── Business logic (preserved exactly) ─────────────────────────────

def remove_unreachable_local_proxy():
    """Discard only the known unusable localhost:9 proxy configuration."""
    removed = False
    for key in ("HTTP_PROXY", "HTTPS_PROXY", "ALL_PROXY", "http_proxy", "https_proxy", "all_proxy"):
        value = os.environ.get(key, "")
        parsed = urlparse(value if "://" in value else f"//{value}")
        if parsed.hostname in {"127.0.0.1", "localhost", "::1"} and parsed.port == 9:
            os.environ.pop(key, None)
            removed = True
    return removed


INVALID_LOCAL_PROXY_REMOVED = remove_unreachable_local_proxy()


def is_blank(value):
    return value is None or pd.isna(value) or not str(value).strip()


def extract_file_id(link):
    """Return a Drive ID when *link* is a Google Drive sharing URL."""
    for pattern in (r"/d/([a-zA-Z0-9_-]+)", r"[?&]id=([a-zA-Z0-9_-]+)"):
        match = re.search(pattern, link)
        if match:
            return match.group(1)
    return None


def safe_filename(value, fallback="download"):
    value = Path(str(value).strip()).name
    value = INVALID_FILENAME_CHARS.sub("_", value).rstrip(". ")
    return value[:180] or fallback


def filename_from_headers(headers, final_url):
    """Get a sensible filename from Content-Disposition, then the URL path."""
    disposition = headers.get("content-disposition", "")
    match = re.search(r"filename\*=UTF-8''([^;]+)", disposition, re.I)
    if not match:
        match = re.search(r'filename\s*=\s*"?([^";]+)', disposition, re.I)
    if match:
        return safe_filename(unquote(match.group(1)), "download")
    path_name = Path(unquote(urlparse(final_url).path)).name
    return safe_filename(path_name, "download") if path_name else "download"


# ── Custom widgets ─────────────────────────────────────────────────

def _round_rect(canvas, x1, y1, x2, y2, r=8, **kw):
    pts = [
        x1+r, y1, x2-r, y1, x2, y1, x2, y1+r,
        x2, y2-r, x2, y2, x2-r, y2,
        x1+r, y2, x1, y2, x1, y2-r,
        x1, y1+r, x1, y1, x1+r, y1,
    ]
    return canvas.create_polygon(pts, smooth=True, **kw)


class Btn(tk.Canvas):
    """Flat rounded button."""

    def __init__(self, parent, text="", command=None, w=120, h=32,
                 bg_color=ACCENT, hv=ACCENT_HOVER, pr=ACCENT_DIM,
                 fg="#fff", font=("Segoe UI Semibold", 9), r=6, **kw):
        pbg = parent.cget("bg") if hasattr(parent, "cget") else BG
        super().__init__(parent, width=w, height=h, bg=pbg,
                         highlightthickness=0, cursor="hand2", **kw)
        self._t, self._c = text, command
        self._bg, self._hv, self._pr = bg_color, hv, pr
        self._fg, self._f, self._r = fg, font, r
        self._dis = False
        self._draw()
        self.bind("<Enter>", lambda e: self._draw(self._hv))
        self.bind("<Leave>", lambda e: self._draw())
        self.bind("<ButtonPress-1>", lambda e: self._draw(self._pr))
        self.bind("<ButtonRelease-1>", self._click)

    def _draw(self, fill=None):
        self.delete("all")
        w, h = int(self["width"]), int(self["height"])
        c = TEXT_DIM if self._dis else (fill or self._bg)
        _round_rect(self, 0, 0, w, h, self._r, fill=c,
                    outline=BORDER if c == SURFACE_LIGHT else "")
        self.create_text(w//2, h//2, text=self._t,
                         fill=TEXT_SEC if self._dis else self._fg, font=self._f)

    def _click(self, _e):
        if not self._dis and self._c:
            self._draw(self._hv)
            self._c()

    def set_enabled(self, on):
        self._dis = not on
        self.config(cursor="hand2" if on else "arrow")
        self._draw()


class PBar(tk.Canvas):
    """Flat progress bar."""

    def __init__(self, parent, h=5, color=ACCENT, track=SURFACE_LIGHT, **kw):
        pbg = parent.cget("bg") if hasattr(parent, "cget") else SURFACE
        super().__init__(parent, height=h, bg=pbg, highlightthickness=0, **kw)
        self._c, self._t, self._v = color, track, 0
        self.bind("<Configure>", lambda e: self._draw())

    def set(self, pct):
        self._v = max(0, min(100, pct))
        self._draw()

    def _draw(self):
        self.delete("all")
        w, h = self.winfo_width(), self.winfo_height()
        if w < 4:
            return
        r = h // 2
        _round_rect(self, 0, 0, w, h, r, fill=self._t, outline="")
        if self._v > 0:
            fw = max(r*2, int(w * self._v / 100))
            _round_rect(self, 0, 0, fw, h, r, fill=self._c, outline="")


# ── About dialog ───────────────────────────────────────────────────

def _show_about(parent):
    d = tk.Toplevel(parent)
    d.title(f"About {APP_TITLE}")
    d.configure(bg=BG)
    d.resizable(False, False)
    d.transient(parent)
    d.grab_set()

    ico = _resource_path("assets/icon.ico")
    if ico.exists():
        try:
            d.iconbitmap(str(ico))
        except tk.TclError:
            pass

    f = tk.Frame(d, bg=BG, padx=32, pady=26)
    f.pack(fill="both", expand=True)

    tk.Label(f, text="ABOUT THIS APP", font=("Segoe UI Semibold", 8),
             fg=ACCENT, bg=BG, anchor="w").pack(fill="x", pady=(0, 4))
    tk.Label(f, text=APP_TITLE, font=("Segoe UI Semibold", 18),
             fg=TEXT, bg=BG, anchor="w").pack(fill="x")
    tk.Label(f, text=f"Version {APP_VERSION}", font=("Segoe UI", 9),
             fg=ACCENT, bg=BG, anchor="w").pack(fill="x", pady=(1, 14))
    tk.Label(f, text="Designed & Developed by", font=("Segoe UI Semibold", 9),
             fg=TEXT_SEC, bg=BG, anchor="w").pack(fill="x")
    tk.Label(f, text="Soumya Subham Rout", font=("Segoe UI Semibold", 14),
             fg=TEXT, bg=BG, anchor="w").pack(fill="x", pady=(1, 10))
    tk.Label(f, text="Building productivity software, automation tools, desktop\n"
             "applications, and AI-powered solutions with a focus on\n"
             "simplicity, performance, and reliability.",
             font=("Segoe UI", 9), fg=TEXT_SEC, bg=BG,
             anchor="w", justify="left").pack(fill="x", pady=(0, 14))
    tk.Label(f, text="Contact", font=("Segoe UI Semibold", 9),
             fg=TEXT, bg=BG, anchor="w").pack(fill="x")
    tk.Label(f, text="Email: soumyasubham.ssr@gmail.com", font=("Segoe UI", 9),
             fg=ACCENT, bg=BG, anchor="w", cursor="hand2").pack(fill="x", pady=(1, 0))
    tk.Label(f, text="GitHub: github.com/soumyasubham-ssr", font=("Segoe UI", 9),
             fg=ACCENT, bg=BG, anchor="w", cursor="hand2").pack(fill="x", pady=(1, 10))
    tk.Label(f, text="© 2026 Soumya Subham Rout. All Rights Reserved.",
             font=("Segoe UI", 8), fg=TEXT_DIM, bg=BG, anchor="w").pack(fill="x")

    br = tk.Frame(f, bg=BG)
    br.pack(fill="x", pady=(12, 0))
    Btn(br, text="OK", w=80, h=30, r=5, command=d.destroy).pack(side="right")

    d.update_idletasks()
    dw, dh = f.winfo_reqwidth() + 56, f.winfo_reqheight() + 44
    px = parent.winfo_x() + (parent.winfo_width() - dw) // 2
    py = parent.winfo_y() + (parent.winfo_height() - dh) // 2
    d.geometry(f"{dw}x{dh}+{max(0,px)}+{max(0,py)}")


# ── Guide dialog ──────────────────────────────────────────────────

_GUIDE_STEPS = [
    ("Step 1: Share the Google Drive Folder",
     "Open the Google Drive folder where all Google Form\n"
     "uploaded files are stored.\n\n"
     "  1. Click Share on the folder.\n"
     "  2. Change access to Anyone with the link.\n"
     "  3. Set permission to Editor.\n"
     "  4. Copy the folder link if required."),

    ("Step 2: Prepare the Excel File",
     "Create an Excel (.xlsx) file with two columns:\n\n"
     "  Column A – New filename to assign after download\n"
     "                    (extension is optional).\n"
     "  Column B – Google Form uploaded file link\n"
     "                    copied from the response spreadsheet.\n\n"
     "Example:\n"
     "  Student_001  |  https://drive.google.com/...\n"
     "  Student_002  |  https://drive.google.com/..."),

    ("Step 3: Select the Excel File",
     "Open the application, click Browse next to\n"
     "\"Excel File\", and select the prepared .xlsx file."),

    ("Step 4: Select the Output Folder",
     "Click Browse next to \"Output Folder\" and choose\n"
     "the destination where downloaded files will be saved."),

    ("Step 5: Start Download",
     "Click Start Download. The application will:\n\n"
     "  • Read the Excel file.\n"
     "  • Download each Google Drive file.\n"
     "  • Rename it using the value from Column A.\n"
     "  • Save it to the selected output folder."),
]

_GUIDE_NOTES = (
    "• Ensure every Google Drive file is accessible with the\n"
    "   required sharing permissions.\n"
    "• Do not modify the column order in the Excel file.\n"
    "• Supports bulk downloading and automatic renaming.\n"
    "• Invalid or inaccessible links are skipped and reported\n"
    "   in the download log."
)

STEP_COLORS = [ACCENT, GREEN, "#42a5f5", AMBER, "#ab47bc"]


def _show_guide(parent):
    d = tk.Toplevel(parent)
    d.title("Setup Guide")
    d.configure(bg=BG)
    d.resizable(False, True)
    d.transient(parent)
    d.grab_set()

    ico = _resource_path("assets/icon.ico")
    if ico.exists():
        try:
            d.iconbitmap(str(ico))
        except tk.TclError:
            pass

    # Scrollable canvas
    wrapper = tk.Frame(d, bg=BG)
    wrapper.pack(fill="both", expand=True)

    canvas = tk.Canvas(wrapper, bg=BG, highlightthickness=0, bd=0)
    scrollbar = tk.Scrollbar(wrapper, orient="vertical", command=canvas.yview,
                             bg=SURFACE_LIGHT, troughcolor=BG,
                             activebackground=ACCENT, relief="flat", bd=0,
                             highlightthickness=0)
    canvas.configure(yscrollcommand=scrollbar.set)
    scrollbar.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)

    inner = tk.Frame(canvas, bg=BG, padx=28, pady=20)
    win_id = canvas.create_window((0, 0), window=inner, anchor="nw")

    def _on_configure(_e=None):
        canvas.configure(scrollregion=canvas.bbox("all"))
        canvas.itemconfigure(win_id, width=canvas.winfo_width())

    inner.bind("<Configure>", _on_configure)
    canvas.bind("<Configure>", _on_configure)

    def _on_mousewheel(e):
        canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")

    canvas.bind_all("<MouseWheel>", _on_mousewheel)
    d.bind("<Destroy>", lambda e: canvas.unbind_all("<MouseWheel>") if e.widget is d else None)

    # Title
    tk.Label(inner, text="How to Set Up", font=("Segoe UI Semibold", 15),
             fg=TEXT, bg=BG, anchor="w").pack(fill="x")
    tk.Label(inner, text=APP_TITLE, font=("Segoe UI Semibold", 15),
             fg=ACCENT, bg=BG, anchor="w").pack(fill="x", pady=(0, 14))

    # Steps
    for i, (title, body) in enumerate(_GUIDE_STEPS):
        color = STEP_COLORS[i % len(STEP_COLORS)]

        step_frame = tk.Frame(inner, bg=SURFACE, highlightbackground=BORDER,
                              highlightthickness=1, padx=14, pady=10)
        step_frame.pack(fill="x", pady=(0, 8))

        hdr = tk.Frame(step_frame, bg=SURFACE)
        hdr.pack(fill="x", pady=(0, 4))

        dot = tk.Canvas(hdr, width=10, height=10, bg=SURFACE, highlightthickness=0)
        dot.pack(side="left", padx=(0, 8), pady=2)
        dot.create_oval(1, 1, 9, 9, fill=color, outline="")

        tk.Label(hdr, text=title, font=("Segoe UI Semibold", 10),
                 fg=color, bg=SURFACE, anchor="w").pack(side="left")

        tk.Label(step_frame, text=body, font=("Segoe UI", 9),
                 fg=TEXT_SEC, bg=SURFACE, anchor="w",
                 justify="left").pack(fill="x", padx=(18, 0))

    # Important Notes
    notes_frame = tk.Frame(inner, bg="#1a1520", highlightbackground="#3d2e50",
                           highlightthickness=1, padx=14, pady=10)
    notes_frame.pack(fill="x", pady=(6, 10))

    tk.Label(notes_frame, text="Important Notes", font=("Segoe UI Semibold", 10),
             fg=AMBER, bg="#1a1520", anchor="w").pack(fill="x", pady=(0, 4))
    tk.Label(notes_frame, text=_GUIDE_NOTES, font=("Segoe UI", 9),
             fg=TEXT_SEC, bg="#1a1520", anchor="w",
             justify="left").pack(fill="x")

    # Close button
    btn_row = tk.Frame(inner, bg=BG)
    btn_row.pack(fill="x", pady=(4, 0))
    Btn(btn_row, text="Got it", w=80, h=30, r=5,
        command=d.destroy).pack(side="right")

    d.update_idletasks()
    sw, sh = parent.winfo_screenwidth(), parent.winfo_screenheight()
    dw = min(460, sw - 100)
    dh = min(inner.winfo_reqheight() + 40, sh - 80)
    px = parent.winfo_x() + (parent.winfo_width() - dw) // 2
    py = parent.winfo_y() + max(0, (parent.winfo_height() - dh) // 2)
    d.geometry(f"{dw}x{dh}+{max(0, px)}+{max(0, py)}")


# ── Main application ───────────────────────────────────────────────

class FileDownloaderApp:
    def __init__(self, root):
        self.root = root
        root.title(APP_TITLE)
        root.configure(bg=BG)

        ico = _resource_path("assets/icon.ico")
        if ico.exists():
            try:
                root.iconbitmap(str(ico))
            except tk.TclError:
                pass

        self.excel_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.status_text = tk.StringVar(value="Ready to download")
        self.speed_text = tk.StringVar(value="Speed: --")
        self.progress_text = tk.StringVar(value="0 of 0 processed")
        names = ("Total links", "Completed", "Failed", "Link Missing", "Remaining")
        self.stat_vars = {n: tk.StringVar(value="0") for n in names}
        self.events = queue.Queue()
        self.running = False
        self.name_lock = threading.Lock()
        self.reserved_names = set()
        self.started_at = None
        self.total = self.completed = self.failed = self.missing = 0
        self.processed = self.active = self.downloaded_bytes = 0
        self.log_rows = []
        self.log_path = None

        self._build_ui()

        root.update_idletasks()
        sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
        ww = max(640, min(980, sw - 40))
        wh = max(460, min(720, sh - 60))
        root.geometry(f"{ww}x{wh}+{(sw-ww)//2}+{max(0,(sh-wh)//2-20)}")
        root.minsize(640, 460)

        root.after(100, self._process_events)

    def _card(self, parent, **kw):
        return tk.Frame(parent, bg=SURFACE, highlightbackground=BORDER,
                        highlightthickness=1, **kw)

    def _build_ui(self):
        outer = tk.Frame(self.root, bg=BG, padx=24, pady=18)
        outer.pack(fill="both", expand=True)

        # ── Header ─────────────────────────────────────────────────
        hdr = tk.Frame(outer, bg=BG)
        hdr.pack(fill="x", pady=(0, 18))

        brand = tk.Frame(hdr, bg=BG)
        brand.pack(side="left", fill="x", expand=True)
        mark = tk.Canvas(brand, width=34, height=34, bg=BG, highlightthickness=0)
        mark.pack(side="left", padx=(0, 11))
        _round_rect(mark, 1, 1, 33, 33, 10, fill=ACCENT, outline="")
        mark.create_text(17, 17, text="↓", fill="#ffffff", font=("Segoe UI Semibold", 20))
        title_box = tk.Frame(brand, bg=BG)
        title_box.pack(side="left", anchor="w")
        tk.Label(title_box, text=APP_TITLE, font=("Segoe UI Semibold", 16),
                 fg=TEXT, bg=BG, anchor="w").pack(anchor="w")
        tk.Label(title_box, text="Reliable bulk file transfer for your form responses",
                 font=("Segoe UI", 9), fg=TEXT_SEC, bg=BG, anchor="w").pack(anchor="w", pady=(2, 0))
        tk.Label(hdr, text="●  READY", font=("Segoe UI Semibold", 8),
                 fg=GREEN, bg=BG).pack(side="left", padx=(0, 18))

        Btn(hdr, text="Template", w=84, h=30,
            bg_color=SURFACE_LIGHT, hv="#2a3050", pr=SURFACE,
            r=7, font=("Segoe UI Semibold", 8),
            command=self._save_template).pack(side="right", padx=(5, 0))
        Btn(hdr, text="Guide", w=62, h=30,
            bg_color=SURFACE_LIGHT, hv="#2a3050", pr=SURFACE,
            r=7, font=("Segoe UI Semibold", 8),
            command=lambda: _show_guide(self.root)).pack(side="right", padx=(5, 0))
        Btn(hdr, text="About", w=62, h=30,
            bg_color=SURFACE_LIGHT, hv="#2a3050", pr=SURFACE,
            r=7, font=("Segoe UI Semibold", 8),
            command=lambda: _show_about(self.root)).pack(side="right")

        # ── File selection ─────────────────────────────────────────
        fc = self._card(outer, padx=18, pady=16)
        fc.pack(fill="x", pady=(0, 12))
        fc.columnconfigure(1, weight=1)

        tk.Label(fc, text="SOURCE & DESTINATION", font=("Segoe UI Semibold", 8),
                 fg=ACCENT, bg=SURFACE).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 10))
        tk.Label(fc, text="Excel File", font=("Segoe UI Semibold", 9),
                 fg=TEXT, bg=SURFACE).grid(row=1, column=0, sticky="w", padx=(0, 14), pady=5)
        self.excel_entry = tk.Entry(fc, textvariable=self.excel_path,
                                    font=("Segoe UI", 9), bg=ENTRY_BG, fg=TEXT,
                                    insertbackground=TEXT, relief="flat",
                                    highlightthickness=1, highlightbackground=BORDER,
                                    highlightcolor=ACCENT)
        self.excel_entry.grid(row=1, column=1, sticky="ew", pady=5, ipady=6)
        self.excel_btn = Btn(fc, text="Browse", w=72, h=30,
                             bg_color=SURFACE_LIGHT, hv="#2a3050", pr=SURFACE,
                             r=7, font=("Segoe UI Semibold", 8), command=self.browse_excel)
        self.excel_btn.grid(row=1, column=2, padx=(8, 0), pady=5)

        tk.Label(fc, text="Output Folder", font=("Segoe UI Semibold", 9),
                 fg=TEXT, bg=SURFACE).grid(row=2, column=0, sticky="w", padx=(0, 14), pady=5)
        self.output_entry = tk.Entry(fc, textvariable=self.output_path,
                                     font=("Segoe UI", 9), bg=ENTRY_BG, fg=TEXT,
                                     insertbackground=TEXT, relief="flat",
                                     highlightthickness=1, highlightbackground=BORDER,
                                     highlightcolor=ACCENT)
        self.output_entry.grid(row=2, column=1, sticky="ew", pady=5, ipady=6)
        self.output_btn = Btn(fc, text="Browse", w=72, h=30,
                              bg_color=SURFACE_LIGHT, hv="#2a3050", pr=SURFACE,
                              r=7, font=("Segoe UI Semibold", 8), command=self.browse_folder)
        self.output_btn.grid(row=2, column=2, padx=(8, 0), pady=5)

        tk.Label(fc, text="Col 1 = filename, Col 2 = URL, Row 1 = header",
                 font=("Segoe UI", 8), fg=TEXT_DIM, bg=SURFACE
                 ).grid(row=3, column=1, columnspan=2, sticky="w", pady=(3, 0))

        # ── Stats ──────────────────────────────────────────────────
        srow = tk.Frame(outer, bg=BG)
        srow.pack(fill="x", pady=(0, 12))
        colors = {"Total links": TEXT, "Completed": GREEN, "Failed": RED,
                  "Link Missing": AMBER, "Remaining": TEXT_SEC}
        for i, name in enumerate(self.stat_vars):
            srow.columnconfigure(i, weight=1)
            c = self._card(srow, padx=12, pady=10)
            c.grid(row=0, column=i, sticky="nsew", padx=(0 if i == 0 else 3, 0))
            tk.Label(c, textvariable=self.stat_vars[name],
                     font=("Segoe UI Semibold", 17), fg=colors[name],
                     bg=SURFACE).pack(anchor="w")
            tk.Label(c, text=name.upper(), font=("Segoe UI Semibold", 7),
                     fg=TEXT_DIM, bg=SURFACE).pack(anchor="w")

        # ── Progress + Start ───────────────────────────────────────
        pc = self._card(outer, padx=18, pady=14)
        pc.pack(fill="x", pady=(0, 12))
        pc.columnconfigure(0, weight=1)

        r0 = tk.Frame(pc, bg=SURFACE)
        r0.grid(row=0, column=0, sticky="ew")
        r0.columnconfigure(0, weight=1)
        tk.Label(r0, textvariable=self.status_text, font=("Segoe UI Semibold", 10),
                 fg=TEXT_SEC, bg=SURFACE).grid(row=0, column=0, sticky="w")
        tk.Label(r0, textvariable=self.speed_text, font=("Segoe UI", 9),
                 fg=TEXT_DIM, bg=SURFACE).grid(row=0, column=1, sticky="e")

        self.progress = PBar(pc, h=8, color=ACCENT, track=SURFACE_LIGHT)
        self.progress.grid(row=1, column=0, sticky="ew", pady=(5, 3))

        r2 = tk.Frame(pc, bg=SURFACE)
        r2.grid(row=2, column=0, sticky="ew")
        r2.columnconfigure(0, weight=1)
        tk.Label(r2, textvariable=self.progress_text, font=("Segoe UI", 8),
                 fg=TEXT_DIM, bg=SURFACE).pack(side="left")
        self.start_btn = Btn(r2, text="Start Download  →", w=154, h=36,
                             font=("Segoe UI Semibold", 9), r=8,
                             command=self.start_download)
        self.start_btn.pack(side="right", pady=(3, 0))

        # ── Activity log ───────────────────────────────────────────
        lc = self._card(outer, padx=14, pady=12)
        lc.pack(fill="both", expand=True)

        tk.Label(lc, text="ACTIVITY LOG", font=("Segoe UI Semibold", 8),
                 fg=TEXT_SEC, bg=SURFACE).pack(anchor="w", pady=(0, 3))
        lf = tk.Frame(lc, bg=SURFACE)
        lf.pack(fill="both", expand=True)
        self.log = tk.Text(lf, wrap="word", relief="flat", bg="#0a101a",
                           fg=TEXT_SEC, font=("Cascadia Mono", 9),
                           state="disabled", insertbackground=TEXT,
                           selectbackground=ACCENT, highlightthickness=0,
                           padx=10, pady=8)
        sb = tk.Scrollbar(lf, orient="vertical", command=self.log.yview,
                          bg=SURFACE_LIGHT, troughcolor=LOG_BG,
                          activebackground=ACCENT, relief="flat", bd=0,
                          highlightthickness=0)
        self.log.configure(yscrollcommand=sb.set)
        self.log.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

    # ── Template ───────────────────────────────────────────────────

    def _save_template(self):
        src = _resource_path("assets/template.xlsx")
        if not src.exists():
            messagebox.showerror("Template not found",
                                 "The template file could not be located.")
            return
        dest = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile="template.xlsx",
            title="Save Excel Template",
        )
        if dest:
            try:
                shutil.copy2(str(src), dest)
                messagebox.showinfo("Template saved", f"Template saved to:\n{dest}")
            except Exception as err:
                messagebox.showerror("Save failed", str(err))

    # ── File dialogs ───────────────────────────────────────────────

    def browse_excel(self):
        selection = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if selection:
            self.excel_path.set(selection)

    def browse_folder(self):
        selection = filedialog.askdirectory()
        if selection:
            self.output_path.set(selection)

    # ── Log ────────────────────────────────────────────────────────

    def _append_log(self, text):
        self.log.configure(state="normal")
        self.log.insert("end", f"[{datetime.now():%H:%M:%S}] {text}\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    # ── Control state ──────────────────────────────────────────────

    def _set_controls(self, enabled):
        st = "normal" if enabled else "disabled"
        for w in (self.excel_entry, self.output_entry):
            w.configure(state=st)
        for b in (self.excel_btn, self.output_btn, self.start_btn):
            b.set_enabled(enabled)

    # ── Download orchestration (unchanged logic) ───────────────────

    def start_download(self):
        excel_file, output_dir = Path(self.excel_path.get()), Path(self.output_path.get())
        if not excel_file.is_file() or not output_dir.is_dir():
            messagebox.showerror("Missing information", "Select a valid Excel file and output folder.")
            return
        try:
            frame = pd.read_excel(excel_file)
        except Exception as error:
            messagebox.showerror("Cannot read Excel file", str(error))
            return
        if len(frame.columns) < 2:
            messagebox.showerror("Invalid Excel file", "The file must have at least two columns: filename and download URL.")
            return
        if frame.empty:
            messagebox.showerror("No download rows", "The Excel file contains no rows after the header.")
            return
        self.running, self.started_at = True, time.monotonic()
        self.total = len(frame)
        self.completed = self.failed = self.missing = self.processed = self.active = self.downloaded_bytes = 0
        self.log_rows, self.reserved_names = [], set()
        self.progress.set(0)
        self.log.configure(state="normal"); self.log.delete("1.0", "end"); self.log.configure(state="disabled")
        self._set_controls(False); self._update_stats()
        if INVALID_LOCAL_PROXY_REMOVED:
            self._append_log("Ignored an invalid localhost:9 proxy setting so downloads can connect directly.")
        self._append_log(f"Queue started: {self.total} row(s), maximum {MAX_WORKERS} simultaneous downloads.")
        threading.Thread(target=self._run_queue, args=(frame.iloc[:, :2], output_dir), daemon=True).start()

    def _run_queue(self, frame, output_dir):
        jobs, seen_urls = [], set()
        for sequence, (_, row) in enumerate(frame.iterrows(), start=1):
            requested_name = "" if is_blank(row.iloc[0]) else str(row.iloc[0]).strip()
            source_url = "" if is_blank(row.iloc[1]) else str(row.iloc[1]).strip()
            if not source_url:
                self.events.put(("result", self._record(source_url, "Link Missing", "", "URL cell is blank", 0)))
            elif source_url in seen_urls:
                self.events.put(("result", self._record(source_url, "Skipped", "", "Duplicate URL skipped", 0)))
            else:
                seen_urls.add(source_url)
                jobs.append((source_url, requested_name, sequence, output_dir))
        with ThreadPoolExecutor(max_workers=MAX_WORKERS, thread_name_prefix="file-download") as executor:
            futures = [executor.submit(self._download_one, *job) for job in jobs]
            for future in as_completed(futures):
                try:
                    self.events.put(("result", future.result()))
                except Exception as error:
                    self.events.put(("result", self._record("", "Failed", "", f"Unexpected worker error: {error}", 0)))
        self.events.put(("finished", None))

    def _reserve_target(self, output_dir, filename):
        filename = safe_filename(filename, "download")
        stem, suffix = Path(filename).stem, Path(filename).suffix
        with self.name_lock:
            candidate, number = filename, 2
            while candidate.lower() in self.reserved_names or (output_dir / candidate).exists():
                candidate = f"{stem}_{number}{suffix}"
                number += 1
            self.reserved_names.add(candidate.lower())
        return output_dir / candidate

    def _download_one(self, source_url, requested_name, sequence, output_dir):
        self.events.put(("worker_started", None))
        try:
            drive_id = extract_file_id(source_url)
            download_url = source_url
            if drive_id:
                download_url = f"https://drive.usercontent.google.com/download?id={drive_id}&export=download&confirm=t"

            session = requests.Session()
            session.trust_env = not INVALID_LOCAL_PROXY_REMOVED
            with session.get(download_url, stream=True, timeout=(15, 180), allow_redirects=True,
                             headers={"User-Agent": "Mozilla/5.0 (compatible; GoogleFormFileDownloader/2.0)"}) as response:
                response.raise_for_status()
                server_name = filename_from_headers(response.headers, response.url)
                filename = requested_name or server_name
                if not Path(filename).suffix:
                    extension = Path(server_name).suffix or mimetypes.guess_extension(response.headers.get("content-type", "").split(";")[0]) or ""
                    filename += extension
                target = self._reserve_target(output_dir, filename)
                temporary = output_dir / f".{uuid.uuid4().hex}.part"
                byte_count = 0
                with temporary.open("wb") as output:
                    for chunk in response.iter_content(chunk_size=256 * 1024):
                        if chunk:
                            output.write(chunk)
                            byte_count += len(chunk)
                if byte_count == 0:
                    raise RuntimeError("The server returned an empty file.")
                os.replace(temporary, target)
                return self._record(source_url, "Success", target.name, "", byte_count)
        except Exception as error:
            try:
                if "temporary" in locals() and temporary.exists():
                    temporary.unlink()
            except OSError:
                pass
            if "drive_id" in locals() and drive_id:
                error = f"Google Drive direct download failed: {error}"
            return self._record(source_url, "Failed", "", str(error), 0)

    @staticmethod
    def _record(source_url, status, filename, error, byte_count):
        return {"Timestamp": datetime.now().isoformat(timespec="seconds"), "Source URL": source_url,
                "Status": status, "Saved filename": filename, "Error message": error, "Bytes downloaded": byte_count}

    # ── Event processing (unchanged logic) ─────────────────────────

    def _process_events(self):
        try:
            while True:
                event, payload = self.events.get_nowait()
                if event == "worker_started":
                    self.active += 1; self._update_stats()
                elif event == "result":
                    self.active = max(0, self.active - 1) if payload["Status"] in {"Success", "Failed"} else self.active
                    self.log_rows.append(payload); self.processed += 1; self.downloaded_bytes += payload["Bytes downloaded"]
                    if payload["Status"] == "Success":
                        self.completed += 1; self._append_log(f"Downloaded: {payload['Saved filename']}")
                    elif payload["Status"] == "Link Missing":
                        self.missing += 1; self._append_log("Link Missing: URL cell is blank")
                    else:
                        self.failed += 1; self._append_log(f"{payload['Status']}: {payload['Error message']}")
                    self._update_stats()
                elif event == "finished":
                    self._finish_run()
        except queue.Empty:
            pass
        self.root.after(100, self._process_events)

    def _update_stats(self):
        remaining = max(0, self.total - self.processed)
        for variable, value in zip(self.stat_vars.values(), (self.total, self.completed, self.failed, self.missing, remaining)):
            variable.set(str(value))
        percent = self.processed / self.total * 100 if self.total else 0
        self.progress.set(percent)
        self.progress_text.set(f"{self.processed} of {self.total} processed ({percent:.0f}%)")
        elapsed = max(time.monotonic() - self.started_at, 0.001) if self.started_at else 0.001
        self.speed_text.set(f"Speed: {self.downloaded_bytes / elapsed / 1024 / 1024:.2f} MB/s  •  {self.active}/{MAX_WORKERS} active")
        self.status_text.set("Downloading files…" if remaining else "Finalizing download log…")

    def _finish_run(self):
        if not self.running:
            return
        elapsed = time.monotonic() - self.started_at
        self.log_path = Path(self.output_path.get()) / f"download_log_{datetime.now():%Y%m%d_%H%M%S}.csv"
        try:
            with self.log_path.open("w", newline="", encoding="utf-8-sig") as log_file:
                writer = csv.DictWriter(log_file, fieldnames=["Timestamp", "Source URL", "Status", "Saved filename", "Error message", "Bytes downloaded"])
                writer.writeheader(); writer.writerows(self.log_rows)
            log_location = str(self.log_path)
        except Exception as error:
            log_location = f"Could not save log: {error}"
        self.running = False; self._set_controls(True)
        self.status_text.set("Download run completed")
        self.speed_text.set(f"Completed in {elapsed:.1f} seconds")
        self._append_log(f"Finished. Success: {self.completed}; failed: {self.failed}; link missing: {self.missing}. Log: {log_location}")
        messagebox.showinfo("Download complete", f"Total rows processed: {self.processed}\nSuccessfully downloaded: {self.completed}\nFailed downloads: {self.failed}\nLink Missing: {self.missing}\nTotal execution time: {elapsed:.1f} seconds\n\nLog file:\n{log_location}")


if __name__ == "__main__":
    _enable_windows_dpi_awareness()
    window = tk.Tk()
    FileDownloaderApp(window)
    window.mainloop()
