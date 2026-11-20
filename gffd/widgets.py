"""Theme constants and small custom Tk widgets."""

from __future__ import annotations

import ctypes
import os
import sys
import tkinter as tk
from pathlib import Path

# ── Theme ──────────────────────────────────────────────────────────
BG            = "#0b0f17"
SURFACE       = "#131a27"
SURFACE_LIGHT = "#1c2638"
SURFACE_HOVER = "#2a3050"
ACCENT        = "#8b7cff"
ACCENT_HOVER  = "#a69cff"
ACCENT_DIM    = "#6659d6"
TEXT          = "#f2f4f8"
TEXT_SEC      = "#a6b0c2"
TEXT_DIM      = "#68758a"
GREEN         = "#45d483"
RED           = "#ff6b7a"
AMBER         = "#f7c75a"
BLUE          = "#42a5f5"
PURPLE        = "#ab47bc"
BORDER        = "#263249"
ENTRY_BG      = "#0f1623"
LOG_BG        = "#0a101a"

FONT          = "Segoe UI"
FONT_BOLD     = "Segoe UI Semibold"
FONT_MONO     = "Cascadia Mono"


# ── Platform helpers ───────────────────────────────────────────────

def resource_path(relative: str) -> Path:
    """Locate a bundled resource both from source and inside a PyInstaller exe."""
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent.parent))
    return base / relative


ICON_PATH = "assets/icon.ico"


def apply_icon(window: tk.Misc):
    icon = resource_path(ICON_PATH)
    if icon.exists():
        try:
            window.iconbitmap(str(icon))
        except tk.TclError:
            pass


def enable_windows_dpi_awareness():
    """Request per-monitor DPI awareness before Tk creates its first window."""
    if os.name != "nt":
        return
    attempts = (
        lambda: ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4)),
        lambda: ctypes.windll.shcore.SetProcessDpiAwareness(2) == 0,  # HRESULT S_OK
        lambda: ctypes.windll.user32.SetProcessDPIAware(),
    )
    for attempt in attempts:
        try:
            if attempt():
                return
        except (AttributeError, OSError, OverflowError):
            continue


def ui_scale(widget: tk.Misc) -> float:
    """Factor between the current display DPI and 96 DPI (1.0 at 100 %)."""
    try:
        return max(1.0, float(widget.tk.call("tk", "scaling")) * 72 / 96)
    except (tk.TclError, ValueError):
        return 1.0


def px(widget: tk.Misc, value: float) -> int:
    """Convert a 96-DPI pixel size to physical pixels for *widget*'s display."""
    return int(round(value * ui_scale(widget)))


def style_scrollbars(root: tk.Misc):
    """Dark ttk scrollbars (classic tk scrollbars ignore colours on Windows)."""
    from tkinter import ttk
    style = ttk.Style(root)
    style.theme_use("clam")
    style.configure("Dark.Vertical.TScrollbar", gripcount=0, background=SURFACE_LIGHT,
                    darkcolor=SURFACE_LIGHT, lightcolor=SURFACE_LIGHT, troughcolor=LOG_BG,
                    bordercolor=LOG_BG, arrowcolor=TEXT_DIM, relief="flat")
    style.map("Dark.Vertical.TScrollbar", background=[("active", ACCENT), ("pressed", ACCENT_DIM)])


def scrollbar(parent, command):
    from tkinter import ttk
    return ttk.Scrollbar(parent, orient="vertical", command=command,
                         style="Dark.Vertical.TScrollbar")


def round_rect(canvas: tk.Canvas, x1, y1, x2, y2, r=8, **kw):
    pts = [x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r,
           x2, y2 - r, x2, y2, x2 - r, y2,
           x1 + r, y2, x1, y2, x1, y2 - r,
           x1, y1 + r, x1, y1, x1 + r, y1]
    return canvas.create_polygon(pts, smooth=True, **kw)


# ── Widgets ────────────────────────────────────────────────────────

class Btn(tk.Canvas):
    """Flat rounded button drawn on a canvas."""

    def __init__(self, parent, text="", command=None, w=120, h=32,
                 bg_color=ACCENT, hv=ACCENT_HOVER, pr=ACCENT_DIM,
                 fg="#ffffff", font=(FONT_BOLD, 9), r=6, **kw):
        super().__init__(parent, width=px(parent, w), height=px(parent, h), bg=parent.cget("bg"),
                         highlightthickness=0, cursor="hand2", **kw)
        r = px(parent, r)
        self._text, self._command = text, command
        self._colors = (bg_color, hv, pr)
        self._fg, self._font, self._r = fg, font, r
        self._disabled = False
        self._hover = False
        self._draw()
        self.bind("<Enter>", lambda _e: self._set_hover(True))
        self.bind("<Leave>", lambda _e: self._set_hover(False))
        self.bind("<ButtonPress-1>", lambda _e: self._draw(self._colors[2]))
        self.bind("<ButtonRelease-1>", self._click)

    @classmethod
    def secondary(cls, parent, text, command, w=72, h=30):
        return cls(parent, text=text, command=command, w=w, h=h,
                   bg_color=SURFACE_LIGHT, hv=SURFACE_HOVER, pr=SURFACE,
                   r=7, font=(FONT_BOLD, 8))

    def _set_hover(self, on):
        self._hover = on
        self._draw()

    def _draw(self, fill=None):
        self.delete("all")
        w, h = int(self["width"]), int(self["height"])
        if self._disabled:
            color = TEXT_DIM
        else:
            color = fill or (self._colors[1] if self._hover else self._colors[0])
        round_rect(self, 0, 0, w, h, self._r, fill=color,
                   outline=BORDER if color == SURFACE_LIGHT else "")
        self.create_text(w // 2, h // 2, text=self._text,
                         fill=TEXT_SEC if self._disabled else self._fg, font=self._font)

    def _click(self, event):
        inside = 0 <= event.x <= int(self["width"]) and 0 <= event.y <= int(self["height"])
        self._draw()
        if inside and not self._disabled and self._command:
            self._command()

    def configure_button(self, text=None, command=None, colors=None):
        if text is not None:
            self._text = text
        if command is not None:
            self._command = command
        if colors is not None:
            self._colors = colors
        self._draw()

    def set_enabled(self, on: bool):
        self._disabled = not on
        self.config(cursor="hand2" if on else "arrow")
        self._draw()


class PBar(tk.Canvas):
    """Flat rounded progress bar."""

    def __init__(self, parent, h=5, color=ACCENT, track=SURFACE_LIGHT, **kw):
        super().__init__(parent, height=px(parent, h), bg=parent.cget("bg"), highlightthickness=0, **kw)
        self._color, self._track, self._value = color, track, 0.0
        self.bind("<Configure>", lambda _e: self._draw())

    def set(self, pct: float):
        pct = max(0.0, min(100.0, pct))
        if pct != self._value:
            self._value = pct
            self._draw()

    def _draw(self):
        self.delete("all")
        w, h = self.winfo_width(), self.winfo_height()
        if w < 4:
            return
        r = h // 2
        round_rect(self, 0, 0, w, h, r, fill=self._track, outline="")
        if self._value > 0:
            round_rect(self, 0, 0, max(r * 2, int(w * self._value / 100)), h, r,
                       fill=self._color, outline="")


def card(parent, **kw) -> tk.Frame:
    return tk.Frame(parent, bg=SURFACE, highlightbackground=BORDER, highlightthickness=1, **kw)


def label(parent, text="", size=9, bold=False, fg=TEXT, **kw) -> tk.Label:
    kw.setdefault("bg", parent.cget("bg"))
    return tk.Label(parent, text=text, font=(FONT_BOLD if bold else FONT, size), fg=fg, **kw)
